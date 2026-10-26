"""video2excel data extractor: read the TEXT shown in a recording into rows.

Built for screen recordings that scroll through a list of records (e.g. a
conference "Attendees" list where each card = Name / Position / Organization).

Pipeline:
  1. ffmpeg samples frames (reuses converter.extract_frames)
  2. near-duplicate frames are skipped (average-hash) so static stretches of a
     scroll aren't processed twice
  3. an *engine* turns each frame image into a list of record dicts
       - VisionEngine : Claude vision (accurate; needs ANTHROPIC_API_KEY)
       - OcrEngine    : local tesseract via pytesseract (free; lower accuracy)
  4. records are de-duplicated across overlapping frames, preserving scroll order
  5. openpyxl writes the rows to an .xlsx

Engines only transcribe the single image they are given; the model/OCR output
is treated as data, never as instructions.
"""

from __future__ import annotations

import base64
import io
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

from PIL import Image as PILImage

try:
    from .converter import extract_frames, probe_video, VideoInfo, ConversionError
except ImportError:  # run as a script
    from converter import extract_frames, probe_video, VideoInfo, ConversionError


# ---- the fields we pull out of each card -------------------------------------
# Order matters: this is the column order in the output sheet.
FIELDS = ["name", "position", "organization"]
FIELD_LABELS = {"name": "Name", "position": "Position", "organization": "Organization"}


# ---------------------------------------------------------------- frame dedup

def _frame_signature(path: Path) -> bytes:
    """Downscaled grayscale pixels (status/tab bars cropped) for a cheap diff.

    A perceptual hash can't tell scroll positions apart (the whole list looks
    similar), so we compare actual pixels via MSE and skip only frames that are
    essentially static holds of the previous one.
    """
    with PILImage.open(path) as im:
        im = im.convert("L")
        w, h = im.size
        im = im.crop((0, int(h * 0.12), w, int(h * 0.91))).resize((64, 128), PILImage.LANCZOS)
        return im.tobytes()


def _mse(a: bytes, b: bytes) -> float:
    return sum((x - y) * (x - y) for x, y in zip(a, b)) / len(a)


def dedupe_frames(frame_rows: list[dict], threshold: float = 40.0) -> list[dict]:
    """Drop frames that barely differ from the last kept one (static holds).

    Measured MSE on real scroll recordings: static holds sit under ~5, any real
    movement is >250, so the default threshold of 40 cleanly separates them.
    """
    kept: list[dict] = []
    last_sig: bytes | None = None
    for row in frame_rows:
        sig = _frame_signature(row["path"])
        if last_sig is not None and _mse(sig, last_sig) < threshold:
            continue
        kept.append(row)
        last_sig = sig
    return kept


# ---------------------------------------------------------------- record dedup

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _name_tokens(name: str) -> set[str]:
    out = set()
    for t in name.split():
        t = re.sub(r"[^a-z]", "", t.lower())
        if t:
            out.add(t)
    return out


def _completeness(rec: dict) -> tuple[int, int]:
    """Rank reads of the same person: more filled fields, then more text."""
    filled = sum(1 for f in FIELDS if rec.get(f))
    length = sum(len(rec.get(f, "")) for f in ("position", "organization"))
    return (filled, length)


def _name_ratio(a: str, b: str) -> float:
    import difflib
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


def _name_flat(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _same_person(name_a: str, toks_a: set[str], name_b: str, toks_b: set[str],
                 similarity: float) -> bool:
    # partial read of the same card: one name's tokens are a subset of the other
    # (e.g. "Adel OROZCO" ⊂ "Adel CHAOUCH OROZCO"). Identical reads also land here.
    if len(toks_a & toks_b) >= 2 and (toks_a <= toks_b or toks_b <= toks_a):
        return True
    # Same letters but different case/punctuation = two DELIBERATE cards shown in
    # the list (e.g. "Adel CHAOUCH OROZCO" vs "Adel Chaouch-Orozco"). Keep both.
    if _name_flat(name_a) == _name_flat(name_b):
        return False
    # Otherwise, only a genuine letter-level OCR typo of one card merges
    # (e.g. "CHAOUCH" vs "CHAOQUCH").
    return _name_ratio(name_a, name_b) >= similarity


def merge_records(per_frame: list[list[dict]], similarity: float = 0.96) -> list[dict]:
    """Collapse the same person seen across overlapping frames into one row.

    1. Reject "names" that are actually an organization or position seen on some
       other card — these come from top-cut cards whose real name scrolled off
       (e.g. "QS ImpACT", "UAE University", "Deputy Vice Rector").
    2. Cluster the remaining reads with union-find (order-independent, so the
       result is stable regardless of sample rate). Two reads are the same
       person when their names share a token subset (partial reads like
       "Adel OROZCO" ⊂ "Adel CHAOUCH OROZCO") or are a near-identical string
       (a 1-char OCR typo, ratio >= 0.96). Names with the SAME letters but
       different case/punctuation are kept apart — they are two deliberate
       cards shown in the list ("Adel CHAOUCH OROZCO" vs "Adel Chaouch-Orozco").
    3. Each person keeps the MOST COMPLETE field read, is displayed with the
       name spelling seen in the most frames, and the rows come out in the
       order the people first appear.
    """
    from collections import Counter

    recs = []
    known_org: set[str] = set()
    known_pos: set[str] = set()
    for cards in per_frame:
        for rec in cards:
            rec = {f: _norm(rec.get(f, "")) for f in FIELDS}
            if rec["organization"]:
                known_org.add(rec["organization"].lower())
            if rec["position"]:
                known_pos.add(rec["position"].lower())
            if rec["name"]:
                recs.append(rec)
    recs = [r for r in recs
            if r["name"].lower() not in known_org and r["name"].lower() not in known_pos]
    if not recs:
        return []

    toks = [_name_tokens(r["name"]) for r in recs]

    # union-find over records
    parent = list(range(len(recs)))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[max(ri, rj)] = min(ri, rj)   # keep earliest-appearing root

    for i in range(len(recs)):
        for j in range(i + 1, len(recs)):
            if _same_person(recs[i]["name"], toks[i], recs[j]["name"], toks[j], similarity):
                union(i, j)

    clusters: dict[int, list[int]] = {}
    for i in range(len(recs)):
        clusters.setdefault(find(i), []).append(i)

    out = []
    for root in sorted(clusters):                    # first-appearance order
        members = clusters[root]
        best = max((recs[m] for m in members), key=_completeness)
        names = Counter(recs[m]["name"] for m in members)
        # display name: most frequent, then fullest (most tokens), then longest
        disp = max(names, key=lambda n: (names[n], len(_name_tokens(n)), len(n)))
        rec = dict(best)
        rec["name"] = disp
        out.append(rec)
    return out


# ---------------------------------------------------------------- engines

class Engine:
    name = "base"

    def extract(self, image_path: Path) -> list[dict]:
        raise NotImplementedError


_VISION_PROMPT = (
    "This is ONE frame from a screen recording that scrolls through a conference "
    "\"Attendees\" list. Each attendee is shown as a card with three text lines:\n"
    "  - NAME (bold, top line)\n"
    "  - POSITION / job title (second line)\n"
    "  - ORGANIZATION / institution (third line; it may wrap onto two lines)\n\n"
    "Return a JSON array of the attendees whose card is FULLY visible in this image. "
    "Skip any card that is cut off at the very top or bottom edge (it will be fully "
    "visible in another frame). Transcribe the text EXACTLY as shown, including "
    "spelling and capitalization. If a field is not visible, use an empty string.\n\n"
    "Each element must be an object: "
    "{\"name\": \"\", \"position\": \"\", \"organization\": \"\"}\n"
    "Output ONLY the JSON array - no prose, no code fences."
)

DEFAULT_VISION_MODEL = os.environ.get("VIDEO2EXCEL_MODEL", "claude-sonnet-5-5")


class VisionEngine(Engine):
    """Transcribe each frame with Claude vision. Needs ANTHROPIC_API_KEY."""

    name = "vision"

    def __init__(self, model: str | None = None, api_key: str | None = None,
                 send_width: int = 540):
        try:
            import anthropic
        except ImportError as e:
            raise ConversionError("The 'anthropic' package is not installed.") from e

        key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not key:
            raise ConversionError(
                "No API key for the Claude-vision engine. Set ANTHROPIC_API_KEY "
                "in your environment (or paste a key in the form), or choose the "
                "free local OCR engine instead."
            )
        self._anthropic = anthropic
        self.client = anthropic.Anthropic(api_key=key)
        self.model = model or DEFAULT_VISION_MODEL
        self.send_width = send_width

    def _encode(self, image_path: Path) -> str:
        with PILImage.open(image_path) as im:
            im = im.convert("RGB")
            if im.width > self.send_width:
                h = round(im.height * self.send_width / im.width)
                im = im.resize((self.send_width, h), PILImage.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=80)
        return base64.standard_b64encode(buf.getvalue()).decode("ascii")

    def extract(self, image_path: Path) -> list[dict]:
        b64 = self._encode(image_path)
        try:
            resp = self.client.messages.create(
                model=self.model,
                max_tokens=2000,
                output_config={"effort": "low"},
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {
                            "type": "base64", "media_type": "image/jpeg", "data": b64}},
                        {"type": "text", "text": _VISION_PROMPT},
                    ],
                }],
            )
        except self._anthropic.AuthenticationError as e:
            raise ConversionError("Claude API rejected the key (authentication failed).") from e
        except self._anthropic.APIStatusError as e:
            raise ConversionError(f"Claude API error ({e.status_code}): {e.message}") from e
        except self._anthropic.APIConnectionError as e:
            raise ConversionError("Could not reach the Claude API (network error).") from e

        text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
        return _parse_json_records(text)


def _parse_json_records(text: str) -> list[dict]:
    text = text.strip()
    if not text:
        return []
    # tolerate code fences or stray prose around the array
    m = re.search(r"\[.*\]", text, re.DOTALL)
    if m:
        text = m.group(0)
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return []
    if not isinstance(data, list):
        return []
    out = []
    for item in data:
        if isinstance(item, dict):
            out.append({f: str(item.get(f, "") or "") for f in FIELDS})
    return out


# words that mark the line where the organization/institution begins
_ORG_KEYWORDS = (
    "university", "universitas", "universiti", "universidad", "université",
    "college", "institute", "institut", "school", "academy", "polytechnic",
    "ministry", "hospital", "foundation", "council", "authority",
    "company", "corporation",
)


def _clean_ocr_line(words: list[str]) -> str:
    """Join OCR words into a line, dropping icon glyphs and symbol-only tokens."""
    toks = []
    for w in words:
        if any(c.isalpha() for c in w) or w == "&":
            toks.append(w.strip("©®|[]{}:;·•«»"))
    return re.sub(r"\s+", " ", " ".join(toks)).strip(" :;,-|")


_GLYPH_RE = re.compile(r"[jiloue]{1,4}")
_GLYPH_CHARS = set("jilouet")


def _is_button_noise(text: str) -> bool:
    # the two action buttons (calendar / chat) OCR as blobs of J/I/L/U/o/e/t
    # (e.g. "JL", "JU", "JLo)", "OO", "JLeJte3"); real content lines have >2
    # letters and contain letters outside that set. We only flag a longer blob
    # when it is made purely of those glyph letters AND contains a j or l (the
    # button strokes), so institution words like "University" are never dropped.
    core = re.sub(r"[^A-Za-z]", "", text).lower()
    if not core:
        return True
    if len(core) <= 2 or _GLYPH_RE.fullmatch(core):
        return True
    return set(core) <= _GLYPH_CHARS and ("j" in core or "l" in core)


def _strip_leading_glyphs(name: str) -> str:
    # remove a leaked button glyph that OCR stuck in front of a name ("OO ALIA")
    toks = name.split()
    while toks:
        core = re.sub(r"[^A-Za-z]", "", toks[0]).lower()
        if core and len(core) <= 2 and _GLYPH_RE.fullmatch(core):
            toks.pop(0)
        else:
            break
    return " ".join(toks)


def _strip_trailing_glyphs(text: str, max_strip: int = 3) -> str:
    # the action buttons trail a field as short blobs ("it oO", "JL"); strip up
    # to a few trailing tokens made only of glyph letters j/i/l/o/u/e/t.
    toks = text.split()
    stripped = 0
    while toks and stripped < max_strip:
        core = re.sub(r"[^A-Za-z]", "", toks[-1]).lower()
        if core and len(core) <= 3 and set(core) <= _GLYPH_CHARS:
            toks.pop()
            stripped += 1
        else:
            break
    return " ".join(toks)


# whole words that mark a line as a job title, not a person's name. Used to
# reject top-cut partial cards whose name scrolled off, leaving the position
# line as the first visible line.
_ROLE_WORDS = {
    "director", "manager", "rector", "professor", "chancellor", "officer",
    "dean", "president", "deputy", "vice", "assistant", "associate", "chief",
    "coordinator", "managing", "administrator", "secretary", "consultant",
    "lead", "founder", "provost", "registrar", "principal", "supervisor",
    "specialist", "executive", "analyst", "engineer", "staff", "head", "pro",
    "programme", "program", "affairs", "affair", "operations", "communication",
}
# whole words that mark a line as an institution, not a person's name.
_ORG_WORDS = {
    "university", "universitas", "universiti", "universidad", "college",
    "institute", "institut", "school", "academy", "polytechnic", "ministry",
    "hospital", "foundation", "council", "authority", "company", "corporation",
    "sciences", "technology",
}


def _words(text: str) -> list[str]:
    return [re.sub(r"[^a-z]", "", t.lower()) for t in text.split()]


def _looks_like_person_name(name: str) -> bool:
    toks = name.split()
    if len(toks) < 2:
        return False
    capitalised = sum(1 for t in toks if t[:1].isupper() and any(c.isalpha() for c in t))
    return capitalised >= 2 and sum(c.isalpha() for c in name) >= 4


def _is_valid_name_line(name: str) -> bool:
    """True only for lines that plausibly ARE a person's name.

    Rejects: title lines ("Deputy Vice Rector"), institution lines ("UAE
    University"), and OCR gibberish ("CS a ON ee ee? ae") — all of which show up
    as the first visible line of a card whose real name scrolled off an edge.
    """
    if not _looks_like_person_name(name):
        return False
    words = [w for w in _words(name) if w]
    if any(w in _ROLE_WORDS or w in _ORG_WORDS for w in words):
        return False
    # need >=2 real word tokens and at least one >=3 letters (kills "CS ON ee ae")
    real = [w for w in words if len(w) >= 2]
    if len(real) < 2 or not any(len(w) >= 3 for w in real):
        return False
    return True


def _norm_name(name: str) -> str:
    name = re.sub(r"\.{2,}", "...", name)          # normalise ellipsis (M... / M....)
    return re.sub(r"\s+", " ", name).strip(" :;,-|")


def _split_position_org(rest: list[str]) -> tuple[str, str]:
    """Split a card's non-name lines into (position, organization).

    Positions and orgs both wrap onto two lines, and font sizes don't separate
    them, so we anchor on the last line that names an institution; everything
    before it is the position. Without a keyword, assume the last line is the org.
    """
    if not rest:
        return "", ""
    if len(rest) == 1:
        return rest[0], ""
    last_kw = None
    for i, line in enumerate(rest):
        if any(k in line.lower() for k in _ORG_KEYWORDS):
            last_kw = i
    if last_kw is None:
        return " ".join(rest[:-1]), rest[-1]
    return " ".join(rest[:last_kw]), " ".join(rest[last_kw:])


class OcrEngine(Engine):
    """Free local OCR with tesseract, tuned for scrolling card lists.

    Crops to the text column (dropping the avatar's clock badge and the
    scrollbar), upscales for legibility, groups words into lines and lines into
    cards by vertical gaps, drops action-button glyphs, and splits each card
    into name / position / organization.
    """

    name = "ocr"

    def __init__(self):
        if not shutil.which("tesseract"):
            raise ConversionError(
                "tesseract is not installed (needed for the free OCR engine). "
                "Install it with `brew install tesseract`, or use the Claude-vision engine."
            )
        try:
            import pytesseract  # noqa: F401
        except ImportError as e:
            raise ConversionError("The 'pytesseract' package is not installed.") from e

    def extract(self, image_path: Path) -> list[dict]:
        import pytesseract
        from pytesseract import Output

        with PILImage.open(image_path) as im:
            im = im.convert("L")
            w, h = im.size
            # text column only: drop status bar, tab bar, avatar column, scrollbar
            left = int(w * 0.292)
            right = int(w * 0.972)
            im = im.crop((left, int(h * 0.12), right, int(h * 0.91)))
            im = im.resize((im.width * 2, im.height * 2), PILImage.LANCZOS)
            data = pytesseract.image_to_data(im, output_type=Output.DICT)

        words = []
        for i, txt in enumerate(data["text"]):
            if txt.strip() and int(data["conf"][i]) >= 35:
                words.append((data["top"][i], data["left"][i], txt))
        words.sort()

        # group words into text lines (same line if tops are within 18px, upscaled)
        raw_lines: list[dict] = []
        for top, lft, txt in words:
            if raw_lines and abs(top - raw_lines[-1]["top"]) <= 18:
                raw_lines[-1]["words"].append((lft, txt))
                raw_lines[-1]["top"] = min(raw_lines[-1]["top"], top)
            else:
                raw_lines.append({"top": top, "words": [(lft, txt)]})
        lines = []
        for L in raw_lines:
            text = _clean_ocr_line([t for _, t in sorted(L["words"])])
            if text and not _is_button_noise(text):
                lines.append({"top": L["top"], "text": text})

        # group lines into cards (new card when the vertical gap > 110px, upscaled)
        cards: list[list[str]] = []
        cur: list[str] = []
        prev: int | None = None
        for L in lines:
            if prev is not None and L["top"] - prev > 110:
                if cur:
                    cards.append(cur)
                cur = []
            cur.append(L["text"])
            prev = L["top"]
        if cur:
            cards.append(cur)

        recs = []
        for c in cards:
            if not c:
                continue
            name = _strip_leading_glyphs(_norm_name(c[0]))
            if not _is_valid_name_line(name):
                continue
            pos, org = _split_position_org(c[1:])
            recs.append({
                "name": name,
                "position": _strip_trailing_glyphs(pos),
                "organization": _strip_trailing_glyphs(org),
            })
        return recs


def make_engine(engine: str, model: str | None = None, api_key: str | None = None) -> Engine:
    engine = (engine or "vision").lower()
    if engine == "vision":
        return VisionEngine(model=model, api_key=api_key)
    if engine == "ocr":
        return OcrEngine()
    raise ConversionError(f"Unknown engine '{engine}'. Use 'vision' or 'ocr'.")


# ---------------------------------------------------------------- workbook

def write_data_workbook(out_xlsx: Path, records: list[dict], info: VideoInfo,
                        source_name: str, engine_name: str) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Data"

    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["#"] + [FIELD_LABELS[f] for f in FIELDS]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for i, rec in enumerate(records, 1):
        ws.cell(i + 1, 1, i).border = border
        ws.cell(i + 1, 1).alignment = Alignment(horizontal="center", vertical="center")
        for c, f in enumerate(FIELDS, start=2):
            cell = ws.cell(i + 1, c, rec.get(f, ""))
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    widths = {"#": 5, "Name": 30, "Position": 44, "Organization": 40}
    from openpyxl.utils import get_column_letter
    for c, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(h, 24)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(records)+1}"

    # Summary sheet
    sm = wb.create_sheet("Summary")
    meta = [
        ("Source file", source_name),
        ("Engine", engine_name),
        ("Resolution", f"{info.width} × {info.height}"),
        ("Duration (s)", info.duration),
        ("Records extracted", len(records)),
    ]
    sm["A1"] = "video2excel — data extraction"
    sm["A1"].font = Font(bold=True, size=14)
    for r, (k, v) in enumerate(meta, start=3):
        sm.cell(r, 1, k).font = Font(bold=True)
        sm.cell(r, 2, v)
    sm.column_dimensions["A"].width = 22
    sm.column_dimensions["B"].width = 50

    wb.save(out_xlsx)


# ---------------------------------------------------------------- orchestration

@dataclass
class DataResult:
    xlsx_path: Path
    records: list[dict]
    frames_processed: int
    engine: str


def convert_video_to_data(
    video_path: Path,
    out_xlsx: Path,
    engine: str = "vision",
    sample_fps: float = 2.0,
    max_frames: int = 120,
    model: str | None = None,
    api_key: str | None = None,
    workdir: Path | None = None,
    frame_dedupe_threshold: float = 40.0,
    _engine_obj: Engine | None = None,
) -> DataResult:
    """Full pipeline: video -> deduped frames -> engine -> records -> xlsx.

    `_engine_obj` lets tests inject a fake engine.
    """
    video_path, out_xlsx = Path(video_path), Path(out_xlsx)
    if not video_path.exists():
        raise ConversionError(f"Video not found: {video_path}")

    eng = _engine_obj or make_engine(engine, model=model, api_key=api_key)
    info = probe_video(video_path)

    owns = workdir is None
    workdir = Path(workdir) if workdir else Path(tempfile.mkdtemp(prefix="v2x_data_"))
    frames_dir = workdir / "frames"
    try:
        frame_rows = extract_frames(video_path, frames_dir, sample_fps, max_frames)
        frame_rows = dedupe_frames(frame_rows, threshold=frame_dedupe_threshold)
        per_frame = [eng.extract(row["path"]) for row in frame_rows]
        records = merge_records(per_frame)
        write_data_workbook(out_xlsx, records, info, video_path.name, eng.name)
    finally:
        if owns:
            shutil.rmtree(workdir, ignore_errors=True)

    return DataResult(out_xlsx, records, len(frame_rows), eng.name)


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Extract the data shown in a recording into Excel.")
    p.add_argument("video")
    p.add_argument("-o", "--output", default=None)
    p.add_argument("--engine", choices=["vision", "ocr"], default="vision")
    p.add_argument("--fps", type=float, default=2.0)
    p.add_argument("--max-frames", type=int, default=120)
    p.add_argument("--model", default=None, help="vision model id (default claude-sonnet-5-5)")
    args = p.parse_args()

    src = Path(args.video)
    out = Path(args.output) if args.output else src.with_name(src.stem + "_data.xlsx")
    res = convert_video_to_data(
        src, out, engine=args.engine, sample_fps=args.fps,
        max_frames=args.max_frames, model=args.model,
    )
    print(f"Wrote {res.xlsx_path} — {len(res.records)} records "
          f"from {res.frames_processed} frames via {res.engine}")
    for r in res.records:
        print(f"  - {r['name']} | {r['position']} | {r['organization']}")
