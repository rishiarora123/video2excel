"""video2excel web app — upload a video, download an .xlsx.

Two modes:
  - data   : auto-extract the TEXT shown in the recording (Name/Position/Org)
             into rows. Engine = 'vision' (Claude, accurate, needs a key) or
             'ocr' (free local tesseract, lower accuracy).
  - frames : embed each sampled frame as an image (the original behaviour).

Run:
    .venv/bin/uvicorn video2excel.app:app --host 127.0.0.1 --port 8070
"""

from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates

try:
    from .converter import convert_video_to_excel, ConversionError
    from .data_extractor import convert_video_to_data
except ImportError:  # run as a script
    from converter import convert_video_to_excel, ConversionError
    from data_extractor import convert_video_to_data

BASE_DIR = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

app = FastAPI(title="video2excel")

MAX_UPLOAD_BYTES = 500 * 1024 * 1024  # 500 MB


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    has_key = bool(os.environ.get("ANTHROPIC_API_KEY"))
    has_tesseract = bool(shutil.which("tesseract"))
    return templates.TemplateResponse(
        request, "index.html",
        {"has_key": has_key, "has_tesseract": has_tesseract},
    )


@app.post("/convert")
async def convert(
    video: UploadFile = File(...),
    mode: str = Form("data"),
    # data-mode options
    engine: str = Form("vision"),
    api_key: str = Form(""),
    model: str = Form(""),
    data_fps: float = Form(1.0),
    # frames-mode options
    fps: float = Form(1.0),
    max_frames: int = Form(400),
    thumb_width: int = Form(220),
):
    workdir = Path(tempfile.mkdtemp(prefix="v2x_web_"))
    src_name = Path(video.filename or "upload.mp4").name
    src_path = workdir / src_name
    try:
        size = 0
        with src_path.open("wb") as f:
            while chunk := await video.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise ConversionError("File too large (limit 500 MB).")
                f.write(chunk)

        stem = Path(src_name).stem
        if mode == "frames":
            out_name = f"{stem}_frames.xlsx"
            out_path = workdir / out_name
            convert_video_to_excel(
                src_path, out_path,
                sample_fps=max(0.01, float(fps)),
                max_frames=max(1, int(max_frames)),
                thumb_width=max(40, int(thumb_width)),
                workdir=workdir / "work",
            )
        else:  # data
            out_name = f"{stem}_data.xlsx"
            out_path = workdir / out_name
            convert_video_to_data(
                src_path, out_path,
                engine=engine,
                sample_fps=max(0.1, float(data_fps)),
                model=(model.strip() or None),
                api_key=(api_key.strip() or None),
                workdir=workdir / "work",
            )

        return FileResponse(
            out_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=out_name,
            background=_Cleanup(workdir),
        )
    except ConversionError as e:
        shutil.rmtree(workdir, ignore_errors=True)
        return HTMLResponse(f"<h3>Conversion error</h3><p>{e}</p><a href='/'>Back</a>", status_code=400)
    except Exception as e:  # pragma: no cover
        shutil.rmtree(workdir, ignore_errors=True)
        return HTMLResponse(f"<h3>Unexpected error</h3><p>{e}</p><a href='/'>Back</a>", status_code=500)


class _Cleanup:
    """Starlette BackgroundTask-compatible callable that removes the temp dir."""

    def __init__(self, path: Path):
        self.path = path

    async def __call__(self) -> None:
        shutil.rmtree(self.path, ignore_errors=True)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8070)
