"""video2excel core: turn a video recording into an .xlsx with its frames.

Pipeline:
  1. ffprobe  -> read width/height/fps/duration/frame count
  2. ffmpeg   -> sample frames at a chosen rate into JPEGs
  3. Pillow   -> downscale each frame to a thumbnail
  4. openpyxl -> embed thumbnails + frame no. + timestamp into a workbook

No third-party video libs are needed beyond ffmpeg/ffprobe on PATH.
"""

from __future__ import annotations

import json
import math
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, asdict
from pathlib import Path

from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------- probing

@dataclass
class VideoInfo:
    width: int
    height: int
    fps: float
    duration: float
    nb_frames: int


class ConversionError(RuntimeError):
    """Raised for any user-facing failure (bad file, no ffmpeg, etc.)."""


def _require_tool(name: str) -> str:
    path = shutil.which(name)
    if not path:
        raise ConversionError(
            f"'{name}' was not found on PATH. Install ffmpeg (e.g. `brew install ffmpeg`)."
        )
    return path


def probe_video(video_path: Path) -> VideoInfo:
    ffprobe = _require_tool("ffprobe")
    cmd = [
        ffprobe, "-v", "error",
        "-select_streams", "v:0",
        "-show_entries", "stream=width,height,r_frame_rate,nb_frames",
        "-show_entries", "format=duration",
        "-of", "json",
        str(video_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise ConversionError(
            "Could not read the video. It may be corrupt or not a video file.\n"
            + proc.stderr.strip()
        )
    data = json.loads(proc.stdout or "{}")
    streams = data.get("streams") or []
    if not streams:
        raise ConversionError("No video stream found in the uploaded file.")
    s = streams[0]

    num, _, den = (s.get("r_frame_rate") or "0/1").partition("/")
    try:
        fps = float(num) / float(den) if float(den) else 0.0
    except (ValueError, ZeroDivisionError):
        fps = 0.0

    duration = float(data.get("format", {}).get("duration") or 0.0)
    try:
        nb_frames = int(s.get("nb_frames") or 0)
    except ValueError:
        nb_frames = 0
    if nb_frames == 0 and fps and duration:
        nb_frames = int(round(fps * duration))

    return VideoInfo(
        width=int(s.get("width") or 0),
        height=int(s.get("height") or 0),
        fps=round(fps, 3),
        duration=round(duration, 3),
        nb_frames=nb_frames,
    )


# ---------------------------------------------------------------- extraction

def _fmt_timestamp(seconds: float) -> str:
    if seconds < 0:
        seconds = 0.0
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:06.3f}"


def extract_frames(
    video_path: Path,
    out_dir: Path,
    sample_fps: float,
    max_frames: int,
) -> list[dict]:
    """Sample frames at `sample_fps` per second. Returns frame metadata rows.

    Each row: {index, time_s, time_str, path}
    """
    ffmpeg = _require_tool("ffmpeg")
    out_dir.mkdir(parents=True, exist_ok=True)

    pattern = str(out_dir / "frame_%05d.jpg")
    cmd = [
        ffmpeg, "-hide_banner", "-loglevel", "error",
        "-i", str(video_path),
        "-vf", f"fps={sample_fps}",
        "-q:v", "3",
        pattern,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise ConversionError("Frame extraction failed.\n" + proc.stderr.strip())

    frames = sorted(out_dir.glob("frame_*.jpg"))
    if not frames:
        raise ConversionError("No frames could be extracted from this video.")

    # Safety cap: evenly drop frames if we overshoot the limit.
    if len(frames) > max_frames:
        step = len(frames) / max_frames
        kept = [frames[min(int(i * step), len(frames) - 1)] for i in range(max_frames)]
        # de-dupe while preserving order
        seen, trimmed = set(), []
        for f in kept:
            if f not in seen:
                seen.add(f)
                trimmed.append(f)
        for f in frames:
            if f not in seen:
                f.unlink(missing_ok=True)
        frames = trimmed

    rows = []
    for i, f in enumerate(frames):
        # frame i was sampled at t = i / sample_fps seconds (first frame at t=0)
        t = i / sample_fps if sample_fps else 0.0
        rows.append({
            "index": i + 1,
            "time_s": round(t, 3),
            "time_str": _fmt_timestamp(t),
            "path": f,
        })
    return rows


def _make_thumbnail(src: Path, dst: Path, thumb_width: int) -> tuple[int, int]:
    with PILImage.open(src) as im:
        im = im.convert("RGB")
        w, h = im.size
        if w > thumb_width:
            new_h = max(1, round(h * thumb_width / w))
            im = im.resize((thumb_width, new_h), PILImage.LANCZOS)
        im.save(dst, "JPEG", quality=85)
        return im.size


# ---------------------------------------------------------------- workbook

# openpyxl geometry helpers: Excel column width ~= px/7, row height is in points (px*0.75)
_PX_PER_WIDTH_UNIT = 7.0
_PT_PER_PX = 0.75

_HEADER_FILL = PatternFill("solid", fgColor="1F2937")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_TITLE_FONT = Font(bold=True, size=14, color="111827")
_LABEL_FONT = Font(bold=True, color="374151")


def build_workbook(
    out_xlsx: Path,
    info: VideoInfo,
    rows: list[dict],
    source_name: str,
    sample_fps: float,
    thumb_width: int,
    thumb_dir: Path,
) -> None:
    wb = Workbook()

    # --- Summary sheet ---
    summary = wb.active
    summary.title = "Summary"
    summary["A1"] = "Video → Excel conversion"
    summary["A1"].font = _TITLE_FONT
    meta = [
        ("Source file", source_name),
        ("Resolution", f"{info.width} × {info.height}"),
        ("Original frame rate", f"{info.fps} fps"),
        ("Duration", f"{_fmt_timestamp(info.duration)}  ({info.duration} s)"),
        ("Total frames in video", info.nb_frames),
        ("Sampled at", f"{sample_fps} frame(s)/second"),
        ("Frames extracted", len(rows)),
        ("Thumbnail width", f"{thumb_width}px"),
    ]
    for r, (label, value) in enumerate(meta, start=3):
        summary.cell(row=r, column=1, value=label).font = _LABEL_FONT
        summary.cell(row=r, column=2, value=value)
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 48

    # --- Frames sheet ---
    ws = wb.create_sheet("Frames")
    headers = ["Frame #", "Time (s)", "Timestamp", "Frame image"]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=text)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 16
    img_col = 4
    ws.column_dimensions[get_column_letter(img_col)].width = thumb_width / _PX_PER_WIDTH_UNIT

    thumb_dir.mkdir(parents=True, exist_ok=True)
    center = Alignment(horizontal="center", vertical="center")

    for r, row in enumerate(rows, start=2):
        ws.cell(row=r, column=1, value=row["index"]).alignment = center
        ws.cell(row=r, column=2, value=row["time_s"]).alignment = center
        ws.cell(row=r, column=3, value=row["time_str"]).alignment = center

        thumb_path = thumb_dir / f"thumb_{row['index']:05d}.jpg"
        tw, th = _make_thumbnail(row["path"], thumb_path, thumb_width)
        ws.row_dimensions[r].height = th * _PT_PER_PX

        img = XLImage(str(thumb_path))
        img.width, img.height = tw, th
        ws.add_image(img, f"{get_column_letter(img_col)}{r}")

    wb.save(out_xlsx)


# ---------------------------------------------------------------- orchestration

@dataclass
class ConvertResult:
    xlsx_path: Path
    info: VideoInfo
    frames_extracted: int
    sample_fps: float


def convert_video_to_excel(
    video_path: Path,
    out_xlsx: Path,
    sample_fps: float = 1.0,
    max_frames: int = 400,
    thumb_width: int = 220,
    workdir: Path | None = None,
) -> ConvertResult:
    """Full pipeline. `workdir` holds temp frames; a temp dir is used if omitted."""
    video_path = Path(video_path)
    out_xlsx = Path(out_xlsx)
    if not video_path.exists():
        raise ConversionError(f"Video not found: {video_path}")
    if sample_fps <= 0:
        raise ConversionError("Sample rate must be greater than 0.")

    info = probe_video(video_path)

    owns_workdir = workdir is None
    workdir = Path(workdir) if workdir else Path(tempfile.mkdtemp(prefix="v2x_"))
    frames_dir = workdir / "frames"
    thumbs_dir = workdir / "thumbs"
    try:
        rows = extract_frames(video_path, frames_dir, sample_fps, max_frames)
        build_workbook(
            out_xlsx, info, rows, video_path.name,
            sample_fps, thumb_width, thumbs_dir,
        )
    finally:
        if owns_workdir:
            shutil.rmtree(workdir, ignore_errors=True)

    return ConvertResult(
        xlsx_path=out_xlsx,
        info=info,
        frames_extracted=len(rows),
        sample_fps=sample_fps,
    )


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="Convert a video recording into an Excel file of frames.")
    p.add_argument("video", help="path to the input video")
    p.add_argument("-o", "--output", help="output .xlsx path", default=None)
    p.add_argument("--fps", type=float, default=1.0, help="frames to sample per second (default 1)")
    p.add_argument("--max-frames", type=int, default=400, help="safety cap on extracted frames")
    p.add_argument("--thumb-width", type=int, default=220, help="thumbnail width in px")
    args = p.parse_args()

    src = Path(args.video)
    out = Path(args.output) if args.output else src.with_suffix(".xlsx")
    res = convert_video_to_excel(
        src, out,
        sample_fps=args.fps,
        max_frames=args.max_frames,
        thumb_width=args.thumb_width,
    )
    print(f"Wrote {res.xlsx_path}  ({res.frames_extracted} frames, {res.info.duration}s @ {res.sample_fps}fps sample)")
