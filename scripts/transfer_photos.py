"""
transfer_photos.py
------------------
Standalone helper for Windows 11: pull photos from an iPhone connected over
USB-C, group them by date + location (fuzzy clustering + reverse-geocode),
and drop them into a OneDrive folder.

Usage (from an activated venv with requirements installed):

    python transfer_photos.py --dest "C:\\Users\\<you>\\OneDrive\\Pictures\\iPhone"

Flags:
    --dest PATH         Target OneDrive folder (required)
    --staging PATH      Temp folder for raw copies (default: %TEMP%\\iphone_photos)
    --device NAME       Substring match for the iPhone name (default: "Apple iPhone")
    --radius-m FLOAT    DBSCAN cluster radius in meters (default: 500)
    --min-samples INT   DBSCAN min samples per cluster (default: 1)
    --move              Delete staged copies after filing them
    --dry-run           Analyze and print the plan, but don't write to --dest
    --no-geocode        Skip reverse-geocoding; use coordinates only
    --email EMAIL       Contact email for the Nominatim User-Agent (recommended)

What it does:
  1. Enumerates the iPhone as a Windows Portable Device via the Shell COM API
     (iPhones don't mount as a drive letter, so we can't use normal file I/O
     on the phone itself).
  2. Recursively copies everything under DCIM\\ into a local staging folder.
  3. Reads EXIF DateTimeOriginal + GPS for each image (HEIC supported via
     pillow-heif). Falls back to file mtime if EXIF date is missing.
  4. Groups photos by calendar day, then runs DBSCAN (haversine metric) on
     the GPS points within each day to find location clusters.
  5. Reverse-geocodes each cluster centroid via OpenStreetMap Nominatim
     (rate-limited to 1 req/sec per their usage policy) to get a human name.
  6. Files each photo into <dest>\\YYYY-MM-DD_<Place>\\ with SHA-256 dedupe.

Requirements (see requirements.txt):
    pywin32, Pillow, pillow-heif, scikit-learn, numpy, requests
"""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import shutil
import sys
import tempfile
import time
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

# --- Third-party imports (fail fast with a helpful message) ------------------
try:
    import numpy as np
    import requests
    from PIL import Image, ExifTags
    from sklearn.cluster import DBSCAN
except ImportError as e:
    sys.exit(
        f"Missing dependency: {e.name}. Install with:\n"
        f"    pip install -r requirements.txt"
    )

try:
    import pillow_heif  # type: ignore

    pillow_heif.register_heif_opener()
except ImportError:
    # HEIC files will be skipped for EXIF, but can still be copied as bytes.
    pillow_heif = None

# win32com is only needed for the phone-reading step. If we're just re-processing
# an existing staging folder we can skip it.
try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # type: ignore


# -----------------------------------------------------------------------------
# Phone enumeration via Windows Shell (MTP / WPD)
# -----------------------------------------------------------------------------

SHELL_MY_COMPUTER = 17  # ssfDRIVES


def find_phone_folder(device_substring: str):
    """Return the Shell Folder object for the connected iPhone, or None."""
    if win32com is None:
        sys.exit(
            "pywin32 is not installed — required to read from the iPhone. "
            "Run: pip install pywin32"
        )
    shell = win32com.client.Dispatch("Shell.Application")
    my_computer = shell.Namespace(SHELL_MY_COMPUTER)
    for item in my_computer.Items():
        if device_substring.lower() in item.Name.lower():
            return item.GetFolder
    return None


def walk_shell_folder(folder) -> Iterable:
    """Recursively yield every non-folder Shell item below `folder`."""
    for item in folder.Items():
        if item.IsFolder:
            yield from walk_shell_folder(item.GetFolder)
        else:
            yield item


def copy_from_phone(device_substring: str, staging: Path) -> int:
    """
    Copy every file from the iPhone's DCIM tree into `staging`.
    Returns the number of new files copied.
    """
    phone_root = find_phone_folder(device_substring)
    if phone_root is None:
        sys.exit(
            f"Could not find a portable device matching '{device_substring}'. "
            f"Unlock the phone, tap 'Trust This Computer', and make sure it "
            f"appears in File Explorer under 'This PC'."
        )

    # iPhones expose 'Internal Storage' -> 'DCIM' -> APPLE folders.
    dcim = None
    for top in phone_root.Items():  # 'Internal Storage'
        if not top.IsFolder:
            continue
        for sub in top.GetFolder.Items():
            if sub.IsFolder and sub.Name.upper() == "DCIM":
                dcim = sub.GetFolder
                break
        if dcim:
            break

    if dcim is None:
        sys.exit("Found the iPhone but no DCIM folder inside it.")

    staging.mkdir(parents=True, exist_ok=True)
    shell = win32com.client.Dispatch("Shell.Application")
    dest_folder = shell.Namespace(str(staging))

    copied = 0
    for item in walk_shell_folder(dcim):
        target = staging / item.Name
        if target.exists() and target.stat().st_size > 0:
            continue  # already staged
        # CopyHere flags: 16 = Yes to all, 256 = no progress, 512 = no confirm dir
        dest_folder.CopyHere(item, 16 | 256 | 512)
        copied += 1

    # CopyHere is async — wait for the file count to settle.
    last = -1
    for _ in range(60):
        count = sum(1 for _ in staging.iterdir())
        if count == last:
            break
        last = count
        time.sleep(1)

    return copied


# -----------------------------------------------------------------------------
# EXIF extraction
# -----------------------------------------------------------------------------

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".tiff", ".dng"}
VIDEO_EXTS = {".mov", ".mp4", ".m4v"}

EXIF_TAGS = {v: k for k, v in ExifTags.TAGS.items()}
GPS_TAGS = {v: k for k, v in ExifTags.GPSTAGS.items()}


@dataclass
class PhotoMeta:
    path: Path
    taken: datetime
    lat: float | None = None
    lon: float | None = None
    cluster: int = -1  # filled in later
    place: str = ""    # filled in later


def _to_degrees(value) -> float:
    d, m, s = value
    return float(d) + float(m) / 60 + float(s) / 3600


def read_meta(path: Path) -> PhotoMeta | None:
    """Return PhotoMeta for `path`, or None if the file isn't a supported image."""
    ext = path.suffix.lower()
    taken: datetime | None = None
    lat = lon = None

    if ext in IMAGE_EXTS:
        try:
            with Image.open(path) as img:
                exif = img.getexif()
                # DateTimeOriginal lives in the ExifIFD sub-block
                ifd = exif.get_ifd(0x8769) if exif else {}
                dto = ifd.get(EXIF_TAGS.get("DateTimeOriginal", 0))
                if dto:
                    try:
                        taken = datetime.strptime(dto, "%Y:%m:%d %H:%M:%S")
                    except ValueError:
                        pass
                gps = exif.get_ifd(0x8825) if exif else {}
                if gps:
                    lat_ref = gps.get(GPS_TAGS["GPSLatitudeRef"])
                    lat_val = gps.get(GPS_TAGS["GPSLatitude"])
                    lon_ref = gps.get(GPS_TAGS["GPSLongitudeRef"])
                    lon_val = gps.get(GPS_TAGS["GPSLongitude"])
                    if lat_val and lon_val:
                        lat = _to_degrees(lat_val)
                        if lat_ref in ("S", b"S"):
                            lat = -lat
                        lon = _to_degrees(lon_val)
                        if lon_ref in ("W", b"W"):
                            lon = -lon
        except Exception:
            pass
    elif ext in VIDEO_EXTS:
        pass  # Videos keep their mtime; no EXIF parse.
    else:
        return None

    if taken is None:
        taken = datetime.fromtimestamp(path.stat().st_mtime)

    return PhotoMeta(path=path, taken=taken, lat=lat, lon=lon)


# -----------------------------------------------------------------------------
# Clustering + reverse-geocode
# -----------------------------------------------------------------------------

EARTH_RADIUS_M = 6_371_000.0


def cluster_by_location(
    photos: list[PhotoMeta], radius_m: float, min_samples: int
) -> None:
    """
    Assign `.cluster` on each photo in-place. Runs DBSCAN per calendar day so
    photos from the same place on different days get distinct clusters.
    Photos without GPS get cluster = -1 ("unknown").
    """
    # Group indices by date
    by_day: dict[str, list[int]] = defaultdict(list)
    for i, p in enumerate(photos):
        by_day[p.taken.strftime("%Y-%m-%d")].append(i)

    next_cluster_id = 0
    for day, idxs in by_day.items():
        geo_idxs = [i for i in idxs if photos[i].lat is not None]
        if not geo_idxs:
            continue
        coords = np.radians(
            [[photos[i].lat, photos[i].lon] for i in geo_idxs]
        )
        eps = radius_m / EARTH_RADIUS_M
        labels = DBSCAN(
            eps=eps, min_samples=min_samples, metric="haversine"
        ).fit_predict(coords)

        # Remap to globally-unique ids (DBSCAN restarts labels at 0 per call)
        local_to_global: dict[int, int] = {}
        for local, idx in zip(labels, geo_idxs):
            if local == -1:
                photos[idx].cluster = -1
                continue
            if local not in local_to_global:
                local_to_global[local] = next_cluster_id
                next_cluster_id += 1
            photos[idx].cluster = local_to_global[local]


def reverse_geocode(lat: float, lon: float, email: str | None) -> str:
    """Call Nominatim and return a short place label like 'Paris, Montmartre'."""
    headers = {
        "User-Agent": f"PaveCapture-photo-transfer/1.0 ({email or 'no-contact'})"
    }
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/reverse",
            params={"lat": lat, "lon": lon, "format": "json", "zoom": 14},
            headers=headers,
            timeout=10,
        )
        r.raise_for_status()
        data = r.json()
    except Exception:
        return ""
    addr = data.get("address", {})
    parts = [
        addr.get("suburb") or addr.get("neighbourhood") or addr.get("village"),
        addr.get("city") or addr.get("town") or addr.get("municipality"),
        addr.get("country_code", "").upper() or None,
    ]
    label = ", ".join(p for p in parts if p)
    return label or data.get("display_name", "").split(",")[0]


def label_clusters(photos: list[PhotoMeta], do_geocode: bool, email: str | None) -> None:
    """Compute a human-readable `place` string per cluster."""
    by_cluster: dict[int, list[PhotoMeta]] = defaultdict(list)
    for p in photos:
        if p.cluster != -1:
            by_cluster[p.cluster].append(p)

    for cid, members in by_cluster.items():
        lat = float(np.mean([m.lat for m in members]))
        lon = float(np.mean([m.lon for m in members]))
        if do_geocode:
            place = reverse_geocode(lat, lon, email)
            time.sleep(1.1)  # Nominatim requires <=1 req/sec
        else:
            place = f"{lat:.4f}_{lon:.4f}"
        if not place:
            place = f"{lat:.4f}_{lon:.4f}"
        for m in members:
            m.place = place


# -----------------------------------------------------------------------------
# Filing into OneDrive
# -----------------------------------------------------------------------------

_SAFE = re.compile(r"[^A-Za-z0-9._-]+")


def slug(s: str) -> str:
    return _SAFE.sub("_", s).strip("_") or "Unknown"


def sha256(path: Path, chunk: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def file_photos(
    photos: list[PhotoMeta], dest: Path, move: bool, dry_run: bool
) -> None:
    seen_hashes: set[str] = set()
    summary: dict[str, int] = defaultdict(int)

    for p in photos:
        day = p.taken.strftime("%Y-%m-%d")
        place = slug(p.place) if p.place else "NoLocation"
        folder_name = f"{day}_{place}"
        target_dir = dest / folder_name
        target = target_dir / p.path.name

        action = "PLAN" if dry_run else ("MOVE" if move else "COPY")
        print(f"[{action}] {p.path.name} -> {folder_name}/")
        summary[folder_name] += 1

        if dry_run:
            continue

        target_dir.mkdir(parents=True, exist_ok=True)

        if target.exists():
            # Dedupe by hash; if contents match, skip; otherwise add a suffix.
            try:
                if sha256(target) == sha256(p.path):
                    if move:
                        p.path.unlink(missing_ok=True)
                    continue
            except OSError:
                pass
            stem, suffix = target.stem, target.suffix
            n = 1
            while target.exists():
                target = target_dir / f"{stem}__{n}{suffix}"
                n += 1

        if move:
            shutil.move(str(p.path), target)
        else:
            shutil.copy2(p.path, target)

    print("\n--- Summary ---")
    for folder, n in sorted(summary.items()):
        print(f"  {folder}: {n}")
    print(f"  Total: {sum(summary.values())}")


# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dest", required=True, type=Path)
    parser.add_argument(
        "--staging",
        type=Path,
        default=Path(tempfile.gettempdir()) / "iphone_photos",
    )
    parser.add_argument("--device", default="Apple iPhone")
    parser.add_argument("--radius-m", type=float, default=500.0)
    parser.add_argument("--min-samples", type=int, default=1)
    parser.add_argument("--move", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-geocode", action="store_true")
    parser.add_argument("--email", default=None)
    parser.add_argument(
        "--skip-copy",
        action="store_true",
        help="Skip pulling from the phone; just re-process --staging.",
    )
    args = parser.parse_args()

    if not args.skip_copy:
        print(f"Copying from iPhone to {args.staging} ...")
        n = copy_from_phone(args.device, args.staging)
        print(f"  staged {n} new file(s)")
    else:
        print(f"Using existing staging folder: {args.staging}")

    print("Reading EXIF ...")
    photos: list[PhotoMeta] = []
    for path in sorted(args.staging.rglob("*")):
        if path.is_file():
            meta = read_meta(path)
            if meta:
                photos.append(meta)
    print(f"  parsed {len(photos)} media file(s)")

    print(f"Clustering by location (radius={args.radius_m} m) ...")
    cluster_by_location(photos, args.radius_m, args.min_samples)

    print("Labeling clusters ...")
    label_clusters(photos, do_geocode=not args.no_geocode, email=args.email)

    print(f"Filing into {args.dest} ...")
    file_photos(photos, args.dest, move=args.move, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
