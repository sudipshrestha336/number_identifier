"""
number_identifier.py
----------------------
Tkinter app: pick a Word (.docx) or PDF file, and it creates a copy where every
numerical figure is highlighted and has a comment asking the reviewer to verify it.
Figures found in an ignore list (typed, or loaded from Excel) are left alone.
The original file is never modified.

Requires:  pip install "python-docx>=1.2.0" openpyxl pymupdf     (pymupdf is only needed for PDFs)
Run:       python number_identifier.py
"""

import copy
import csv
import fnmatch
import os
import queue
import re
import shutil
import threading
import tkinter as tk
from decimal import Decimal
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx import Document
from openpyxl import load_workbook
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

# ============================ CORE LOGIC ==================================

NUMBER_RE = re.compile(
    r"""
    (?<![\d.,])
    (?:[$€£₹¥]\s?|USD\s?|EUR\s?|GBP\s?|INR\s?|NPR\s?|Rs\.?\s?)?
    \d+(?:,\d{2,3})*(?:\.\d+)?
    (?:\s?%|\s?(?:thousand|million|billion|trillion|crore|lakh|mn|bn|bps|k|m)\b)?
    """,
    re.VERBOSE | re.IGNORECASE,
)
SECTION_PREFIX_RE = re.compile(r"^\s*\d+(?:\.\d+)*[.)]\s")

COLORS = {
    "Yellow": WD_COLOR_INDEX.YELLOW,
    "Bright green": WD_COLOR_INDEX.BRIGHT_GREEN,
    "Turquoise": WD_COLOR_INDEX.TURQUOISE,
    "Pink": WD_COLOR_INDEX.PINK,
    "Gray": WD_COLOR_INDEX.GRAY_25,
}

DEFAULT_TEMPLATE = (
    'Please verify: "{value}"{kind}. Confirm the figure is correct, not '
    "misstated, and matches the source (value, units, decimals, period)."
)


def describe(value):
    v = value.strip()
    if "%" in v:
        return " (percentage)"
    if re.match(r"^[$€£₹¥]|^(USD|EUR|GBP|INR|NPR|Rs)", v, re.I):
        return " (monetary amount)"
    if re.search(r"(thousand|million|billion|trillion|crore|lakh|mn|bn|bps)$", v, re.I):
        return " (quantity/amount with scale)"
    if re.fullmatch(r"(19|20)\d{2}", v):
        return " (year?)"
    return ""


def _norm(text):
    """Lowercase and drop spaces/commas so '1,000' matches '1000' and '45 Million' matches '45million'."""
    return re.sub(r"[\s,]", "", text).lower()


def _strip_zeros(t):
    """12.50 -> 12.5, 20.0% -> 20%, 13,733,724.00 -> 13733724 (plain numbers only)."""
    m = re.fullmatch(r"(\d+)(?:\.(\d+))?(%?)", t)
    if not m:
        return t
    frac = (m.group(2) or "").rstrip("0")
    return m.group(1) + ("." + frac if frac else "") + m.group(3)


def canon(text):
    """Comparison form: no spaces/commas/case, no leading minus or brackets (Word figures are
    read without their sign, and (5) / -5 mean the same thing), no trailing decimal zeros."""
    t = _norm(text).replace("\u2212", "-")
    t = t.replace("(", "").replace(")", "").lstrip("-+")
    return _strip_zeros(t)


_ORIG = {}   # canonical entry -> the text as typed / read from Excel (for the CSV report)


def parse_ignore_list(raw):
    """Split on commas/semicolons/new lines. Entries may use * and ? wildcards."""
    out = []
    for x in (x.strip() for x in re.split(r"[,;\n]+", raw)):
        c = canon(x) if x else ""
        if c:
            _ORIG[c] = x
            out.append(c)
    return out


def _plain_number(v):
    """Excel numeric value -> clean text (2200.0 -> '2200', 0.1+0.2 noise avoided)."""
    if isinstance(v, int):
        return str(v)
    d = v if isinstance(v, Decimal) else Decimal(repr(v))
    return format(d.normalize(), "f")


def displayed_decimals(number_format):
    """How many decimals a cell's number format shows (None = can't tell, e.g. General)."""
    if not number_format or number_format.lower() == "general":
        return None
    fmt = number_format.split(";")[0]
    fmt = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', "", fmt)   # drop quoted text, [colors], escapes
    if "e+" in fmt.lower() or "/" in fmt:                           # scientific / fractions
        return None
    m = re.search(r"\.([0#?]+)", fmt)
    return len(m.group(1)) if m else 0


def _clean_excel_number(v, number_format, round_display):
    """Return the figure as text. With round_display, float noise is removed by rounding to
    the decimals shown in the cell (General format falls back to 6 decimals)."""
    is_pct = "%" in (number_format or "")
    d = v if isinstance(v, int) else Decimal(repr(v))
    if is_pct:
        d = d * 100
    if round_display and not isinstance(v, int):
        dec = displayed_decimals(number_format)
        d = round(d, 6 if dec is None else dec)
    text = _plain_number(d)
    return text + "%" if is_pct else text


def load_ignore_from_excel(path, skip_header=False, round_display=True):
    """Read every cell of every sheet; return normalized ignore patterns.
    - numbers become text (2200.0 -> 2200)
    - cells formatted as % are read as shown (0.2 -> 20%)
    - round_display: round to the decimals shown in the cell (removes float noise like 123.45000001)
    - text cells are used as typed (wildcards * ? allowed)"""
    wb = load_workbook(path, data_only=True)
    entries = []
    for ws in wb.worksheets:
        for r_idx, row in enumerate(ws.iter_rows(), 1):
            if skip_header and r_idx == 1:
                continue
            for cell in row:
                v = cell.value
                if v is None or isinstance(v, bool):
                    continue
                if isinstance(v, (int, float)):
                    entries.append(_clean_excel_number(v, cell.number_format, round_display))
                elif isinstance(v, str) and v.strip():
                    entries.append(v.strip())
    wb.close()
    seen, out = set(), []
    for e in entries:
        n = canon(e)
        if n and n not in seen:
            _ORIG[n] = e
            seen.add(n)
            out.append(n)
    return out


def match_ignore(value, patterns):
    """Return the ignore entry that matches the figure (or its numeric part), else None."""
    if not patterns:
        return None
    core = re.search(r"\d+(?:,\d{2,3})*(?:\.\d+)?", value)
    candidates = {canon(value)}
    if core:
        pct = "%" if re.search(r"\d\s?%", value) else ""
        candidates.add(canon(core.group(0)))
        if pct:
            candidates.add(canon(core.group(0) + pct))
    for pat in patterns:
        for c in candidates:
            if fnmatch.fnmatchcase(c, pat):
                return pat
    return None


def is_ignored(value, patterns):
    return match_ignore(value, patterns) is not None


def is_plain_run(r):
    kids = [c for c in r if c.tag != qn("w:rPr")]
    return len(kids) == 1 and kids[0].tag == qn("w:t")


def run_text(r):
    return "".join(t.text or "" for t in r.findall(qn("w:t")))


def set_run_text(run, text):
    run.text = text
    t = run._r.find(qn("w:t"))
    if t is not None:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def build_map(paragraph):
    full, spans, pos = [], [], 0
    for run in paragraph.runs:
        txt = run_text(run._r) if is_plain_run(run._r) else "\x00"
        spans.append((run, pos, pos + len(txt), txt == "\x00"))
        full.append(txt)
        pos += len(txt)
    return "".join(full), spans


def split_at(paragraph, index):
    _, spans = build_map(paragraph)
    for run, s, e, opaque in spans:
        if not opaque and s < index < e:
            text = run_text(run._r)
            k = index - s
            clone = copy.deepcopy(run._r)  # keeps all run formatting
            run._r.addnext(clone)
            set_run_text(run, text[:k])
            set_run_text(Run(clone, paragraph), text[k:])
            return


def runs_covering(paragraph, start, end):
    _, spans = build_map(paragraph)
    return [r for r, s, e, opaque in spans if not opaque and s >= start and e <= end and e > s]


def snippet(full, start, end, width=35):
    """Short whole-word context around a figure for the CSV report."""
    a0, b0 = max(0, start - width), min(len(full), end + width)
    ctx = full[a0:b0].replace("\x00", " ")
    if a0 > 0 and " " in ctx[:start - a0]:
        ctx = "\u2026" + ctx[ctx.index(" "):]
    if b0 < len(full) and " " in ctx[end - a0:]:
        ctx = ctx[:ctx.rindex(" ")] + "\u2026"
    return ctx.strip()


def locate(paragraph, cfg):
    """Human-readable position: 'Body text' or 'Table 2, row 5, col 3'."""
    p = paragraph._p
    tc = next((a for a in p.iterancestors() if a.tag == qn("w:tc")), None)
    if tc is None:
        return "Body text"
    tr = tc.getparent()
    tbl = tr.getparent()
    t_no = next((i for i, t in enumerate(cfg.get("_tables", []), 1) if t is tbl), "?")
    return (f"Table {t_no}, row {tbl.findall(qn('w:tr')).index(tr) + 1}, "
            f"col {tr.findall(qn('w:tc')).index(tc) + 1}")


def process_paragraph(doc, paragraph, cfg, idx=0):
    full, _ = build_map(paragraph)
    if not any(ch.isdigit() for ch in full):
        return 0, 0
    skip_until = 0
    if cfg["skip_sections"]:
        m = SECTION_PREFIX_RE.match(full)
        if m:
            skip_until = m.end()

    loc = f"{locate(paragraph, cfg)} (para {idx})"

    def rec(m, entry, status):
        ctx = snippet(full, m.start(), m.end())
        return {"figure": m.group(0).strip(), "entry": entry, "status": status,
                "location": loc, "context": ctx}

    matches = list(NUMBER_RE.finditer(full))
    recs = [None] * len(matches)
    count = skipped = 0
    # Work backwards so earlier character offsets stay valid after splitting runs.
    for i in range(len(matches) - 1, -1, -1):
        m = matches[i]
        if m.start() < skip_until:
            recs[i] = rec(m, "", "SKIPPED (section number)")
            continue
        value = m.group(0).strip()
        hit = match_ignore(value, cfg["ignore"])
        if hit is not None:
            cfg.setdefault("_used", set()).add(hit)
            recs[i] = rec(m, _ORIG.get(hit, hit), "IGNORED (matched)")
            skipped += 1
            continue
        start, end = m.start(), m.start() + len(m.group(0).rstrip())
        split_at(paragraph, end)
        split_at(paragraph, start)
        runs = runs_covering(paragraph, start, end)
        if not runs:
            recs[i] = rec(m, "", "NOT COMMENTED (could not anchor)")
            continue
        if cfg["highlight"]:
            for r in runs:
                r.font.highlight_color = cfg["color"]
        text = cfg["template"].replace("{value}", value).replace("{kind}", describe(value))
        doc.add_comment(runs=[runs[0], runs[-1]], text=text,
                        author=cfg["author"], initials=cfg["initials"])
        cfg.setdefault("_commented", []).append(value)
        recs[i] = rec(m, "", "COMMENTED (not matched)")
        count += 1
    cfg.setdefault("_records", []).extend(r for r in recs if r)
    return count, skipped


def iter_body_paragraphs(doc):
    for p in doc.element.body.iter(qn("w:p")):
        if any(a.tag == qn("w:txbxContent") for a in p.iterancestors()):
            continue  # Word doesn't allow comments inside text boxes
        yield Paragraph(p, doc._body)


def build_report(cfg, tolerance=Decimal("0.01"), limit=25):
    """Explain why figures stayed commented: ignore entries nothing matched, and commented
    figures that are almost (but not exactly) equal to an ignore entry."""
    entries = cfg.get("ignore", [])
    used = cfg.get("_used", set())
    unused = [e for e in entries if e not in used]

    numeric = []
    for e in entries:
        if re.fullmatch(r"\d+(?:\.\d+)?", e):
            numeric.append((Decimal(e), e))
    near, seen = [], set()
    for v in cfg.get("_commented", []):
        core = re.search(r"\d+(?:,\d{2,3})*(?:\.\d+)?", v)
        if not core or v in seen:
            continue
        seen.add(v)
        d = Decimal(core.group(0).replace(",", ""))
        for ev, et in numeric:
            diff = abs(d - ev)
            if 0 < diff <= tolerance:
                near.append((v, et, diff))
                break
    return {"unused": unused[:limit], "unused_total": len(unused),
            "near": near[:limit], "near_total": len(near),
            "no_text": cfg.get("_notext", [])}


def write_csv(cfg, path):
    """All figures in document order, then a separate list of the commented (unmatched) ones."""
    recs = cfg.get("_records", [])
    header = ["No.", "Figure in document", "Matched ignore entry", "Status", "Location", "Context"]

    def safe(x):                      # stop Excel treating text as a formula
        return "'" + x if x[:1] in ("=", "+", "-", "@") else x

    def row(n, r):
        return [n, r["figure"], r["entry"], r["status"], r["location"], safe(r["context"])]

    commented = [r for r in recs if r["status"].startswith(("COMMENTED", "NOT COMMENTED"))]
    ignored = sum(r["status"].startswith("IGNORED") for r in recs)
    other = len(recs) - len(commented) - ignored
    with open(path, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig so Excel opens it cleanly
        w = csv.writer(f)
        w.writerow(["ALL NUMERICAL FIGURES IN THE DOCUMENT (in document order)"])
        w.writerow(header)
        for n, r in enumerate(recs, 1):
            w.writerow(row(n, r))
        w.writerow([])
        w.writerow([f"COMMENTED FIGURES - NOT MATCHED BY THE IGNORE LIST ({len(commented)})"])
        w.writerow(header)
        for n, r in enumerate(commented, 1):
            w.writerow(row(n, r))
        w.writerow([])
        w.writerow(["SUMMARY"])
        w.writerow(["Total figures found", len(recs)])
        w.writerow(["Ignored (matched)", ignored])
        w.writerow(["Commented (not matched)", len(commented)])
        w.writerow(["Skipped (section numbers)", other])


PDF_COLORS = {
    "Yellow": (1, 1, 0), "Bright green": (0, 1, 0), "Turquoise": (0, 1, 1),
    "Pink": (1, 0.45, 0.8), "Gray": (0.75, 0.75, 0.75),
}


def _fitz():
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz
        except ImportError:
            raise RuntimeError("PDF support needs PyMuPDF.\nInstall it with:   pip install pymupdf")
    return fitz


def pdf_lines(page):
    """Yield (text, per-character boxes, is_first_line_of_block) in reading order."""
    data = page.get_text("rawdict", sort=True)
    for block in data["blocks"]:
        if block.get("type") != 0:
            continue
        for li, line in enumerate(block["lines"]):
            text, boxes = [], []
            for span in line["spans"]:
                for ch in span["chars"]:
                    for c in ch["c"]:
                        text.append(c)
                        boxes.append(ch["bbox"])
            yield "".join(text), boxes, li == 0


def annotate_pdf(src, dst, cfg, progress=None):
    """Highlight every figure and attach a PDF comment (annotation). Existing content is
    untouched: the file is copied and the new annotations are appended."""
    fitz = _fitz()
    shutil.copyfile(src, dst)
    doc = fitz.open(dst)
    if doc.needs_pass:
        doc.close()
        raise RuntimeError("This PDF is password-protected. Remove the password first.")
    cfg["_records"], cfg["_commented"], cfg["_used"], cfg["_notext"] = [], [], set(), []
    rgb = PDF_COLORS.get(cfg.get("color_name", "Yellow"), (1, 1, 0))
    total = skipped = 0

    for pno in range(doc.page_count):
        page = doc[pno]
        lines = list(pdf_lines(page))
        if not any(t.strip() for t, _, _ in lines):
            cfg["_notext"].append(pno + 1)
        for lno, (full, boxes, first) in enumerate(lines, 1):
            if not any(ch.isdigit() for ch in full):
                continue
            skip_until = 0
            if cfg["skip_sections"] and first:
                sm = SECTION_PREFIX_RE.match(full)
                if sm:
                    skip_until = sm.end()
            loc = f"Page {pno + 1}, line {lno}"

            def rec(m, entry, status):
                return {"figure": m.group(0).strip(), "entry": entry, "status": status,
                        "location": loc, "context": snippet(full, m.start(), m.end())}

            for m in NUMBER_RE.finditer(full):
                if m.start() < skip_until:
                    cfg["_records"].append(rec(m, "", "SKIPPED (section number)"))
                    continue
                value = m.group(0).strip()
                hit = match_ignore(value, cfg["ignore"])
                if hit is not None:
                    cfg["_used"].add(hit)
                    cfg["_records"].append(rec(m, _ORIG.get(hit, hit), "IGNORED (matched)"))
                    skipped += 1
                    continue
                start, end = m.start(), m.start() + len(m.group(0).rstrip())
                r = fitz.Rect(boxes[start])
                for b in boxes[start + 1:end]:
                    r.include_rect(fitz.Rect(b))
                text = cfg["template"].replace("{value}", value).replace("{kind}", describe(value))
                try:
                    if cfg["highlight"]:
                        annot = page.add_highlight_annot(r)
                        annot.set_colors(stroke=rgb)
                    else:
                        annot = page.add_text_annot(fitz.Point(r.x1, r.y0), text, icon="Comment")
                        annot.set_colors(stroke=rgb)
                    annot.set_info(content=text, title=cfg["author"])
                    annot.update()
                except Exception:  # noqa: BLE001
                    cfg["_records"].append(rec(m, "", "NOT COMMENTED (could not anchor)"))
                    continue
                cfg["_commented"].append(value)
                cfg["_records"].append(rec(m, "", "COMMENTED (not matched)"))
                total += 1
        if progress:
            progress(pno + 1, doc.page_count, total)

    try:
        doc.saveIncr()                       # append-only: original bytes stay exactly as they were
    except Exception:  # noqa: BLE001
        tmp = dst + ".tmp"
        doc.save(tmp, garbage=0, deflate=True)
        doc.close()
        os.replace(tmp, dst)
    else:
        doc.close()
    return total, skipped, build_report(cfg)


def annotate_file(kind, src, dst, cfg, progress=None):
    if kind == "PDF":
        return annotate_pdf(src, dst, cfg, progress)
    return annotate(src, dst, cfg, progress)


def annotate(src, dst, cfg, progress=None):
    doc = Document(src)
    paras = list(iter_body_paragraphs(doc))
    cfg["_tables"] = list(doc.element.body.iter(qn("w:tbl")))
    cfg["_records"], cfg["_commented"], cfg["_used"] = [], [], set()
    total = skipped = 0
    for i, p in enumerate(paras, 1):
        c, sk = process_paragraph(doc, p, cfg, i)
        total += c
        skipped += sk
        if progress:
            progress(i, len(paras), total)
    doc.save(dst)
    return total, skipped, build_report(cfg)


# ================================ GUI =====================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Number Verifier \u2013 Word & PDF")
        self.geometry("820x780")
        self.minsize(640, 560)
        self.q = queue.Queue()

        self.kind = tk.StringVar(value="Word")
        self.in_label = tk.StringVar(value="Input .docx")
        self.src = tk.StringVar()
        self.dst = tk.StringVar()
        self.author = tk.StringVar(value="Number Reviewer")
        self.initials = tk.StringVar(value="NR")
        self.highlight = tk.BooleanVar(value=True)
        self.color = tk.StringVar(value="Yellow")
        self.skip_sections = tk.BooleanVar(value=True)
        self.xl_path = tk.StringVar()
        self.xl_header = tk.BooleanVar(value=False)
        self.xl_round = tk.BooleanVar(value=True)
        self.save_csv = tk.BooleanVar(value=True)

        self._build()
        self.after(100, self._poll)

    # ---- layout ----
    def _build(self):
        pad = {"padx": 10, "pady": 5}
        root = ttk.Frame(self, padding=10)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)

        # Files
        files = ttk.LabelFrame(root, text="Files")
        files.grid(row=0, column=0, sticky="ew", **pad)
        files.columnconfigure(1, weight=1)
        ttk.Label(files, text="Document type").grid(row=0, column=0, sticky="w", **pad)
        kinds = ttk.Frame(files)
        kinds.grid(row=0, column=1, columnspan=2, sticky="w", **pad)
        ttk.Radiobutton(kinds, text="Word (.docx)", variable=self.kind, value="Word",
                        command=self._kind_changed).pack(side="left", padx=(0, 18))
        ttk.Radiobutton(kinds, text="PDF (.pdf)", variable=self.kind, value="PDF",
                        command=self._kind_changed).pack(side="left")
        ttk.Label(files, textvariable=self.in_label).grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(files, textvariable=self.src).grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(files, text="Browse…", command=self._pick_src).grid(row=1, column=2, **pad)
        ttk.Label(files, text="Save as").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(files, textvariable=self.dst).grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(files, text="Browse…", command=self._pick_dst).grid(row=2, column=2, **pad)

        # Options
        opt = ttk.LabelFrame(root, text="Options")
        opt.grid(row=1, column=0, sticky="ew", **pad)
        opt.columnconfigure(1, weight=1)
        opt.columnconfigure(3, weight=1)
        ttk.Label(opt, text="Comment author").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(opt, textvariable=self.author).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Label(opt, text="Initials").grid(row=0, column=2, sticky="w", **pad)
        ttk.Entry(opt, textvariable=self.initials, width=8).grid(row=0, column=3, sticky="w", **pad)

        ttk.Checkbutton(opt, text="Highlight numbers", variable=self.highlight,
                        command=self._toggle_color).grid(row=1, column=0, sticky="w", **pad)
        self.color_box = ttk.Combobox(opt, textvariable=self.color, values=list(COLORS),
                                      state="readonly", width=14)
        self.color_box.grid(row=1, column=1, sticky="w", **pad)
        ttk.Checkbutton(opt, text='Skip section numbers (e.g. "3." at paragraph start)',
                        variable=self.skip_sections).grid(row=2, column=0, columnspan=4,
                                                          sticky="w", **pad)
        ttk.Checkbutton(opt, text="Also save a CSV report (all figures, matched entry, status; "
                                  "commented figures listed at the end)",
                        variable=self.save_csv).grid(row=3, column=0, columnspan=4, sticky="w", **pad)

        # Ignore list
        ign = ttk.LabelFrame(root, text="Ignore list  (numbers that should NOT get a comment)")
        ign.grid(row=2, column=0, sticky="ew", **pad)
        ign.columnconfigure(1, weight=1)
        ttk.Label(ign, text="Excel file").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(ign, textvariable=self.xl_path).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(ign, text="Browse…", command=self._pick_xlsx).grid(row=0, column=2, padx=(0, 4), pady=5)
        ttk.Button(ign, text="Clear", command=lambda: self.xl_path.set("")).grid(row=0, column=3, padx=(0, 10), pady=5)
        ttk.Checkbutton(ign, text="First row is a header (skip it)",
                        variable=self.xl_header).grid(row=1, column=1, sticky="w", padx=10)
        ttk.Checkbutton(ign, text="Round Excel numbers to the decimals shown in the cell (fixes SUM float noise)",
                        variable=self.xl_round).grid(row=1, column=2, columnspan=2, sticky="w", padx=4)
        ttk.Label(ign, text="Extra entries typed here (optional; commas / new lines separate; * and ? wildcards):"
                  ).grid(row=2, column=0, columnspan=4, sticky="w", padx=10, pady=(6, 0))
        self.ignore = tk.Text(ign, height=2, wrap="word")
        self.ignore.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=(2, 2))
        ttk.Label(ign, foreground="gray",
                  text="Excel: every cell on every sheet is read (numbers, %, or text with wildcards)."
                  ).grid(row=4, column=0, columnspan=4, sticky="w", padx=10, pady=(0, 5))

        # Template
        tpl = ttk.LabelFrame(root, text="Comment text  ({value} = the number, {kind} = type hint)")
        tpl.grid(row=3, column=0, sticky="ew", **pad)
        tpl.columnconfigure(0, weight=1)
        self.template = tk.Text(tpl, height=4, wrap="word")
        self.template.insert("1.0", DEFAULT_TEMPLATE)
        self.template.grid(row=0, column=0, sticky="ew", **pad)

        # Run
        run = ttk.Frame(root)
        run.grid(row=4, column=0, sticky="ew", **pad)
        run.columnconfigure(0, weight=1)
        self.bar = ttk.Progressbar(run, mode="determinate")
        self.bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.btn = ttk.Button(run, text="Add comments", command=self._start)
        self.btn.grid(row=0, column=1)

        # Log
        log = ttk.LabelFrame(root, text="Log")
        log.grid(row=5, column=0, sticky="nsew", **pad)
        root.rowconfigure(5, weight=1)
        log.columnconfigure(0, weight=1)
        log.rowconfigure(0, weight=1)
        self.log = tk.Text(log, height=8, state="disabled", wrap="word")
        self.log.grid(row=0, column=0, sticky="nsew", padx=(5, 0), pady=5)
        sb = ttk.Scrollbar(log, command=self.log.yview)
        sb.grid(row=0, column=1, sticky="ns", pady=5)
        self.log.configure(yscrollcommand=sb.set)

    # ---- helpers ----
    def _toggle_color(self):
        self.color_box.configure(state="readonly" if self.highlight.get() else "disabled")

    def _write(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _ext(self):
        return ".pdf" if self.kind.get() == "PDF" else ".docx"

    def _types(self):
        return [("PDF files", "*.pdf")] if self.kind.get() == "PDF" else [("Word documents", "*.docx")]

    def _kind_changed(self):
        self.in_label.set(f"Input {self._ext()}")
        if self.src.get() and Path(self.src.get()).suffix.lower() != self._ext():
            self.src.set("")          # the chosen file no longer matches the selected type
            self.dst.set("")

    def _pick_src(self):
        f = filedialog.askopenfilename(title=f"Select {self.kind.get()} document", filetypes=self._types())
        if f:
            self.src.set(f)
            p = Path(f)
            self.dst.set(str(p.with_name(p.stem + "_commented" + self._ext())))

    def _write_report(self, rep):
        if rep.get("no_text"):
            pages = ", ".join(map(str, rep["no_text"][:30])) + (" ..." if len(rep["no_text"]) > 30 else "")
            self._write(f"\nWARNING: {len(rep['no_text'])} page(s) have no selectable text (scanned images?), "
                        f"so nothing could be checked there: {pages}\n   Run OCR on the PDF first, then retry.")
        if rep["near"]:
            self._write(f"\nNear misses ({rep['near_total']}): commented figures within 0.01 of an ignore entry "
                        "(usually hidden decimals in Excel):")
            for v, e, d in rep["near"]:
                self._write(f"   Word {v}  ~  entry {e}  (differs by {d})")
        if rep["unused"]:
            self._write(f"\nIgnore entries that matched nothing in the document ({rep['unused_total']}):")
            self._write("   " + ", ".join(rep["unused"]) + (" ..." if rep["unused_total"] > len(rep["unused"]) else ""))

    def _pick_xlsx(self):
        f = filedialog.askopenfilename(title="Select Excel file with numbers to ignore",
                                       filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if f:
            self.xl_path.set(f)
            try:
                n = len(load_ignore_from_excel(f, self.xl_header.get(), self.xl_round.get()))
                self._write(f"Ignore list loaded from Excel: {n} unique entries.")
            except Exception as e:  # noqa: BLE001
                messagebox.showerror("Excel error", f"Could not read the file:\n{e}")

    def _pick_dst(self):
        f = filedialog.asksaveasfilename(title="Save commented copy as",
                                         defaultextension=self._ext(),
                                         filetypes=self._types())
        if f:
            self.dst.set(f)

    # ---- run ----
    def _start(self):
        src, dst = self.src.get().strip(), self.dst.get().strip()
        if not src or not Path(src).is_file():
            return messagebox.showerror("Missing file", f"Please choose a valid input {self._ext()} file.")
        if Path(src).suffix.lower() != self._ext():
            return messagebox.showerror("Wrong file type", f"The selected type is {self.kind.get()}, "
                                        f"but the file is not a {self._ext()} file.")
        if not dst:
            return messagebox.showerror("Missing output", "Please choose where to save the new file.")
        if Path(src).resolve() == Path(dst).resolve():
            return messagebox.showerror("Same file", "Output must differ from the input file.")

        ignore = parse_ignore_list(self.ignore.get("1.0", "end"))
        xl = self.xl_path.get().strip()
        if xl:
            try:
                from_xl = load_ignore_from_excel(xl, self.xl_header.get(), self.xl_round.get())
            except Exception as e:  # noqa: BLE001
                return messagebox.showerror("Excel error", f"Could not read the ignore file:\n{e}")
            self._write(f"Ignoring {len(from_xl)} entries from Excel.")
            ignore = list(dict.fromkeys(from_xl + ignore))

        cfg = {
            "author": self.author.get() or "Number Reviewer",
            "initials": self.initials.get() or "NR",
            "highlight": self.highlight.get(),
            "color": COLORS[self.color.get()],
            "color_name": self.color.get(),
            "skip_sections": self.skip_sections.get(),
            "ignore": ignore,
            "template": self.template.get("1.0", "end").strip() or DEFAULT_TEMPLATE,
        }
        self.btn.configure(state="disabled")
        self.bar.configure(value=0)
        self._write(f"Processing: {src}")
        csv_path = str(Path(dst).with_name(Path(dst).stem + "_report.csv")) if self.save_csv.get() else None
        threading.Thread(target=self._worker, args=(src, dst, cfg, csv_path, self.kind.get()), daemon=True).start()

    def _worker(self, src, dst, cfg, csv_path=None, kind="Word"):
        try:
            total, skipped, rep = annotate_file(kind, src, dst, cfg,
                                                progress=lambda i, n, t: self.q.put(("progress", i, n, t)))
            if csv_path:
                write_csv(cfg, csv_path)
            self.q.put(("done", total, skipped, dst, rep, csv_path))
        except Exception as e:  # noqa: BLE001
            self.q.put(("error", str(e)))

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, i, n, _t = msg
                    self.bar.configure(maximum=max(n, 1), value=i)
                elif msg[0] == "done":
                    _, total, skipped, dst, rep, csv_path = msg
                    self.btn.configure(state="normal")
                    self._write(f"Done. {total} figures commented, {skipped} ignored.\nSaved to: {dst}")
                    if csv_path:
                        self._write(f"CSV report: {csv_path}")
                    self._write_report(rep)
                    messagebox.showinfo("Finished", f"{total} figures commented.\n{skipped} ignored "
                                                    f"(ignore list).\n\nSaved to:\n{dst}"
                                                    + (f"\n\nCSV report:\n{csv_path}" if csv_path else ""))
                elif msg[0] == "error":
                    self.btn.configure(state="normal")
                    self._write(f"Error: {msg[1]}")
                    messagebox.showerror("Error", msg[1])
        except queue.Empty:
            pass
        self.after(100, self._poll)


if __name__ == "__main__":
    App().mainloop()
