# video2excel

Upload a screen/phone **video recording** → download an **Excel file**, in one
of two modes:

- **Data** — extracts the **text shown in the recording** into rows. Built for
  recordings that scroll through a list where each item has a few lines (e.g. a
  conference *Attendees* list: **Name / Position / Organization**). Duplicate
  cards seen across the scroll are merged automatically.
- **Frames** — embeds each sampled frame as an **image** with its frame number
  and timestamp.

## Run the web app

```bash
cd "/Users/rishiarora/RI HACK/REA"
.venv/bin/uvicorn video2excel.app:app --host 127.0.0.1 --port 8070
```

Open <http://127.0.0.1:8070>, pick a mode, drop in a video, click **Convert to Excel**.

## Data mode — two engines

| Engine | Accuracy | Needs | Notes |
|--------|----------|-------|-------|
| **Local OCR** (tesseract) | Good, free | `brew install tesseract` | Tuned for this card layout — crops the text column, upscales, drops icon/button glyphs, splits wrapped position/organization lines, and keeps the most complete read per person. On the sample clip it matches the hand-transcribed data exactly. |
| **Claude vision** | Highest, most general | `ANTHROPIC_API_KEY` (per-video API cost) | Best for messy/unknown layouts or non-card recordings. |

> The OCR engine is tuned for a vertically-scrolling list of cards where each
> record is **Name (bold) / Position / Organization**, like the sample. For a
> very different layout, use the vision engine.
>
> **Sample rate (data mode).** Default is **2 fps**, which captures everyone —
> including people only briefly on screen at the start or end of the scroll.
> The OCR parser filters the extra noise finer sampling brings (icon/button
> glyphs, title/organization lines mistaken for names, gibberish) and the
> records are clustered with order-independent union-find, so the row set is
> stable across sample rates. Drop to **1 fps** for the cleanest text if you
> don't need those edge rows.
>
> De-duplication keeps two people apart when their names differ only by case or
> punctuation (two deliberate cards in the list), and merges a single read of
> the same card across frames plus clear 1-character OCR typos of one name.
> A person whose card is cut off at the very end of the recording (so their
> organisation never fully scrolls into view) is still listed, with whatever
> fields were visible.

### Using the vision engine

Give the app your Anthropic API key one of two ways:

```bash
export ANTHROPIC_API_KEY=sk-ant-...        # then start the server
```

…or paste it into the key field in the UI (used for that request only, never stored).

Default model is `claude-sonnet-5-5` (a cost-effective vision model for this
bulk, per-frame job). Override with the **Model** field or:

```bash
export VIDEO2EXCEL_MODEL=claude-opus-5-5
```

## Command line

```bash
cd "/Users/rishiarora/RI HACK/REA"

# Data (text in the video)
.venv/bin/python video2excel/data_extractor.py INPUT.mp4 --engine vision --fps 1
.venv/bin/python video2excel/data_extractor.py INPUT.mp4 --engine ocr   --fps 1

# Frames (images)
.venv/bin/python video2excel/converter.py INPUT.mp4 -o OUTPUT.xlsx --fps 1
```

## How data mode works

1. `ffmpeg` samples frames at the chosen rate.
2. Near-identical frames (static holds during a scroll) are skipped via a
   pixel-difference check, so the same screen isn't processed twice.
3. Each remaining frame is transcribed into records by the chosen engine
   (only cards fully visible in that frame; cut-off cards are picked up in the
   next frame).
4. Records are de-duplicated by name, preserving scroll order.
5. openpyxl writes a **Data** sheet (Name / Position / Organization) plus a
   **Summary** sheet.

> The columns are currently Name / Position / Organization (the three lines on
> each attendee card). To change the fields, edit `FIELDS` / `FIELD_LABELS` and
> the extraction prompt in `data_extractor.py`.

## Requirements

`ffmpeg` + `ffprobe` on PATH, and the Python deps in `requirements.txt`
(already installed in the project `.venv`). `tesseract` only for the OCR engine.

## Files

- `data_extractor.py` — data mode: frames → records → xlsx (vision + OCR engines).
- `converter.py` — frames mode: frames → images in xlsx.
- `app.py` — FastAPI upload/download web app (both modes).
- `templates/index.html` — upload UI.
