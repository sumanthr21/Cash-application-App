"""
Cash Application report engine (Test Equity).

Reads the agents' EOD workbooks, cleans the rows and appends them to the data sheets of the
master report by editing the .xlsx XML directly. Editing the XML (instead of opening the file
with openpyxl and saving it) keeps every chart, comment, hyperlink and formula in the report.

Nothing here talks to Streamlit or GitHub, so it can be tested on its own.
"""
import datetime
import io
import re
import zipfile
from collections import defaultdict
from xml.sax.saxutils import escape

import openpyxl

# --------------------------------------------------------------------------------------
# SETTINGS you may need to change one day
# --------------------------------------------------------------------------------------
AGENTS = ["Varsha", "Gokulnath"]      # names exactly as they appear in the report's Name column
EPOCH = datetime.datetime(1899, 12, 30)
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Each target describes: which EOD sheet feeds it, which report sheet it goes to, the Task text
# written to the Task column, and which report column (letter) gets which field.
# Fields: month, date, week, name, task, + the fields below that come from the EOD sheet.
TARGETS = {
    "billtrust": dict(
        label="Billtrust (Varsha)", sheet="Billtrust", task="Billtrust Review",
        cols=dict(A="month", B="date", C="week", E="name", F="task", G="customer",
                  H="invoice", I="amount", J="tasktype", K="status", L="comment"),
        known=dict(J="tasktype", K="status"),
        required=("customer", "invoice", "amount")),
    "keyoff": dict(
        label="Key offs (Varsha)", sheet="Keyoff", task="Keyoff",
        cols=dict(A="month", B="date", C="week", E="name", F="task", G="customer",
                  H="account", I="amount", J="invoice", K="tasktype", L="status", M="comment"),
        known=dict(K="tasktype", L="status"),
        required=("customer", "invoice", "amount")),
    "unapplied": dict(
        label="Unapplied cash (Varsha)", sheet="Unapplied Cash", task="Unapplied Cash",
        cols=dict(A="month", B="date", C="week", E="name", F="task", G="invdate",
                  H="journal", I="amount", J="reference", K="sequence", L="status",
                  M="journal2", N="agent", O="notes"),
        known=dict(L="status"),
        required=("journal", "amount")),
    "ccp_hisco": dict(
        label="CCP_HISCO (Gokulnath)", sheet="Credit Card", task="Credit Card Processing (Hisco)",
        cols=dict(A="month", B="date", C="week", E="name", F="task", G="customer",
                  H="custno", I="invoice", J="amount", K="joffset", L="joeet"),
        known={}, required=("customer", "invoice", "amount")),
    "ccp_te": dict(
        label="CCP_TE (Gokulnath)", sheet="Credit Card", task="Credit Card Processing (Test Equity)",
        cols=dict(A="month", B="date", C="week", E="name", F="task", G="customer",
                  H="custno", I="invoice", J="amount", K="joffset", L="joeet"),
        known={}, required=("customer", "invoice", "amount")),
    "matchreview": dict(
        label="TE_MATCH REVIEW (Gokulnath)", sheet="Billtrust", task="Billtrust Review",
        cols=dict(A="month", B="date", C="week", E="name", F="task", G="customer",
                  H="invoice", I="amount", J="tasktype", K="status", L="comment"),
        known=dict(J="tasktype", K="status"),
        required=("customer", "invoice", "amount")),
}

# EOD sheet name (letters+digits only, lower case) -> target
SHEET_ALIASES = {
    "billtrust": "billtrust",
    "keyoffs": "keyoff", "keyoff": "keyoff",
    "unappliedcash": "unapplied",
    "ccphisco": "ccp_hisco",
    "ccpte": "ccp_te",
    "tematchreview": "matchreview", "matchreview": "matchreview",
}

# EOD header (letters+digits only, lower case) -> field
HEADER_ALIASES = {
    "date": "date",
    "agentname": "name", "employeename": "name", "agent": "name_agent",
    "customername": "customer",
    "invoicenumber": "invoice", "invoice": "invoice",
    "amount": "amount",
    "tasktype": "tasktype",
    "status": "status",
    "comment": "comment", "comments": "comment",
    "accountnumber": "account",
    "customernumber": "custno",
    "journaloffset": "joffset", "journaloeet": "joeet", "journaloffsetjournaloeet": "joeet",
    "invoicedate": "invdate",
    "journal": "journal", "journal2": "journal2",
    "reference": "reference", "sequence": "sequence", "notes": "notes",
    "matchreview": "flag",
}


# --------------------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------------------
def norm_key(s):
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def clean_text(v):
    if v is None:
        return None
    s = str(v).replace("\xa0", " ").strip()
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)
    return s or None


def to_serial(d):
    return (datetime.datetime(d.year, d.month, d.day) - EPOCH).days


def from_serial(n):
    return EPOCH + datetime.timedelta(days=int(n))


def parse_date(v):
    """EOD dates arrive as real dates, Excel serial numbers or text like 16.11.2026."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, datetime.datetime):
        return v.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(v, datetime.date):
        return datetime.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 30000 < v < 80000:
            return from_serial(v)
        return None
    s = str(v).strip().replace("\xa0", "")
    m = re.fullmatch(r"(\d{1,2})[./-](\d{1,2})[./-](\d{4})", s)
    if m:
        d, mo, y = map(int, m.groups())
        try:
            return datetime.datetime(y, mo, d)
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?", s)
    if m:
        y, mo, d = map(int, m.groups())
        return datetime.datetime(y, mo, d)
    return None


def parse_amount(v):
    """'$   (8,763.86)' -> -8763.86 ; '306.18 ' -> 306.18 ; numbers pass through."""
    if v is None:
        return None, True
    if isinstance(v, bool):
        return None, False
    if isinstance(v, (int, float)):
        return (int(v) if float(v).is_integer() else float(v)), True
    s = str(v).replace("\xa0", " ").strip()
    if not s:
        return None, True
    neg = s.startswith("(") or "(" in s and ")" in s or s.startswith("-")
    t = re.sub(r"[^\d.]", "", s)
    if not t or t == ".":
        return s, False
    try:
        x = float(t)
    except ValueError:
        return s, False
    if neg:
        x = -x
    x = round(x, 2)
    return (int(x) if x.is_integer() else x), True


def tidy_id(v):
    """Account / customer numbers: 604030 or '604030\\xa0' -> 604030 ; '1262 - 109257' stays text."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float)):
        return v
    s = clean_text(v)
    if s is None:
        return None
    if re.fullmatch(r"\d{1,15}", s):
        return int(s)
    return s


def tidy_invoice(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    return clean_text(v)


def week_label(d):
    """Weeks run Monday-Sunday. Week 1 is the week that contains the month's first weekday."""
    first = d.replace(day=1)
    while first.weekday() >= 5:
        first += datetime.timedelta(days=1)
    monday1 = first - datetime.timedelta(days=first.weekday())
    return "Week %d" % ((d - monday1).days // 7 + 1)


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


# --------------------------------------------------------------------------------------
# reading the EOD files
# --------------------------------------------------------------------------------------
def read_eod(files):
    """
    files: list of (file_name, bytes).
    Returns (records, warnings, errors)
      records: list of dict(target=..., fields...) in file order, each with a real date
    """
    records, warnings, errors = [], [], []
    for fname, data in files:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except Exception as e:  # noqa
            errors.append(f"{fname}: cannot open this file ({e}).")
            continue
        used = 0
        for ws in wb.worksheets:
            key = SHEET_ALIASES.get(norm_key(ws.title))
            if key is None:
                warnings.append(f"{fname}: sheet '{ws.title.strip()}' is not one the report uses, so it was ignored.")
                continue
            recs, w, e = _read_sheet(fname, ws, key)
            records += recs
            warnings += w
            errors += e
            used += 1
        if used == 0:
            errors.append(f"{fname}: no known sheets found. Expected sheets such as Billtrust, Key offs, "
                          "Unapplied cash, CCP_HISCO, CCP_TE, TE_ MATCH REVIEW.")
    return records, warnings, errors


def _read_sheet(fname, ws, key):
    tgt = TARGETS[key]
    where = f"{fname} / {ws.title.strip()}"
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], [], []
    # header -> field (first occurrence wins; a second 'Journal' becomes journal2)
    colmap, seen = {}, defaultdict(int)
    for i, h in enumerate(rows[0]):
        if h is None:
            continue
        k = norm_key(h)
        if k == "journal":
            seen["journal"] += 1
            if seen["journal"] == 2:
                k = "journal2"
        f = HEADER_ALIASES.get(k)
        if f and f not in colmap.values():
            colmap[i] = f
    fields = set(colmap.values())
    if key == "unapplied":
        # this sheet has an 'Agent' column; the report's Name comes from it
        if "name_agent" in fields:
            colmap = {i: ("name" if f == "name_agent" else f) for i, f in colmap.items()}
    else:
        colmap = {i: f for i, f in colmap.items() if f != "name_agent"}
    fields = set(colmap.values())
    errs = []
    for need in ("date", "name"):
        if need not in fields:
            errs.append(f"{where}: column for '{need}' not found in the header row.")
    for need in tgt["required"]:
        if need not in fields:
            errs.append(f"{where}: column for '{need}' not found in the header row.")
    if errs:
        return [], [], errs

    recs, warns = [], []
    last_date = None
    skipped_flag = 0
    for rno, row in enumerate(rows[1:], start=2):
        rec = {colmap[i]: row[i] for i in colmap if i < len(row)}
        if all(rec.get(f) in (None, "") for f in tgt["required"]):
            continue                                   # empty / serial-number-only rows
        if key == "matchreview":
            flag = rec.get("flag")
            if flag is not True and str(flag).strip().lower() not in ("true", "yes", "1", "x"):
                skipped_flag += 1
                continue
        d = parse_date(rec.get("date"))
        if d is None:
            d = last_date                              # dates that are blank or formulas: fill down
            if d is None:
                errs.append(f"{where} row {rno}: no date and nothing above to copy it from.")
                continue
        last_date = d
        out = dict(target=key, date=d, src=f"{where} row {rno}")
        # name
        raw_name = clean_text(rec.get("name"))
        name, known = normalise_name(raw_name)
        if not known:
            warns.append(f"{where}: agent '{raw_name}' is not one of {', '.join(AGENTS)}; kept as typed.")
        out["name"] = name
        # values
        out["customer"] = clean_text(rec.get("customer"))
        out["invoice"] = tidy_invoice(rec.get("invoice"))
        amt, ok = parse_amount(rec.get("amount"))
        if not ok:
            warns.append(f"{where} row {rno}: amount '{rec.get('amount')}' is not a number; kept as text.")
        out["amount"] = amt
        for f in ("tasktype", "status", "comment", "reference", "sequence", "notes"):
            out[f] = clean_text(rec.get(f))
        out["account"] = tidy_id(rec.get("account"))
        out["custno"] = tidy_id(rec.get("custno"))
        out["journal"] = tidy_id(rec.get("journal"))
        out["journal2"] = tidy_id(rec.get("journal2"))
        out["joffset"] = tidy_id(rec.get("joffset"))
        out["joeet"] = tidy_id(rec.get("joeet"))
        idate = parse_date(rec.get("invdate"))
        out["invdate"] = idate
        if key == "matchreview":
            out["tasktype"] = "Match Review"
            out["status"] = "Completed"
        if key == "unapplied":
            out["agent"] = name
        recs.append(out)
    if skipped_flag:
        warns.append(f"{where}: {skipped_flag} row(s) skipped because 'Match Review' is not ticked.")
    return recs, warns, errs


def normalise_name(raw):
    if not raw:
        return raw, False
    first = raw.split()[0].lower()
    for a in AGENTS:
        if a.lower() == first or a.lower() == raw.lower():
            return a, True
    return raw, False


# --------------------------------------------------------------------------------------
# reading the report (XML level)
# --------------------------------------------------------------------------------------
ROW_RE = re.compile(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', re.S)
CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def _si_text(s):
    return "".join(re.findall(r"<t[^>]*>(.*?)</t>", s, flags=re.S))


def _unesc(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))


class Book:
    def __init__(self, data):
        self.data = data
        self.zin = zipfile.ZipFile(io.BytesIO(data))
        self.names = self.zin.namelist()
        self.changed = {}
        wbxml = self.zin.read("xl/workbook.xml").decode("utf8")
        rels = self.zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
        rid = {}
        for m in re.finditer(r"<Relationship ([^>]*?)/?>", rels):
            a = m.group(1)
            i = re.search(r'Id="([^"]+)"', a)
            t = re.search(r'Target="([^"]+)"', a)
            if i and t:
                rid[i.group(1)] = t.group(1)
        self.sheet_path = {}
        for m in re.finditer(r"<sheet ([^>]*?)/>", wbxml):
            a = m.group(1)
            n = _unesc(re.search(r'name="([^"]+)"', a).group(1))
            r = re.search(r'r:id="([^"]+)"', a).group(1)
            t = rid[r]
            self.sheet_path[n] = ("xl/" + t.lstrip("/")) if not t.startswith("/xl/") else t.lstrip("/")
        self._ss = None

    @property
    def shared(self):
        if self._ss is None:
            self._ss = []
            if "xl/sharedStrings.xml" in self.names:
                d = self.zin.read("xl/sharedStrings.xml").decode("utf8")
                self._ss = [_unesc(_si_text(s)) for s in re.findall(r"<si>(.*?)</si>|<si/>", d, flags=re.S)]
        return self._ss

    def xml(self, sheet):
        p = self.sheet_path[sheet]
        if p in self.changed:
            return self.changed[p]
        return self.zin.read(p).decode("utf8")

    def set_xml(self, sheet, text):
        self.changed[self.sheet_path[sheet]] = text

    def cells_of(self, body):
        out = {}
        for m in CELL_RE.finditer(body or ""):
            col, _, attrs, inner = m.groups()
            if not inner:
                continue
            t = re.search(r' t="(\w+)"', attrs)
            t = t.group(1) if t else "n"
            if t == "inlineStr":
                out[col] = (_unesc(_si_text(inner)), "s", re.search(r' s="(\d+)"', attrs))
                continue
            v = re.search(r"<v>(.*?)</v>", inner, re.S)
            if not v:
                continue
            v = v.group(1)
            if t == "s":
                out[col] = (self.shared[int(v)], "s", re.search(r' s="(\d+)"', attrs))
            elif t in ("str", "e"):
                out[col] = (_unesc(v), "s", re.search(r' s="(\d+)"', attrs))
            elif t == "b":
                out[col] = (v == "1", "n", re.search(r' s="(\d+)"', attrs))
            else:
                try:
                    out[col] = (float(v), "n", re.search(r' s="(\d+)"', attrs))
                except ValueError:
                    pass
        return out

    def sheet_rows(self, sheet):
        xml = self.xml(sheet)
        for m in ROW_RE.finditer(xml):
            yield int(m.group(1)), self.cells_of(m.group(2))

    def last_row_number(self, sheet):
        xml = self.xml(sheet)
        nums = re.findall(r'<row r="(\d+)"', xml[-400000:]) or re.findall(r'<row r="(\d+)"', xml)
        return max(int(x) for x in nums)


def existing_keys(book, sheet):
    """(date serial, name, task) for every row already in the sheet."""
    keys = set()
    for _, c in book.sheet_rows(sheet):
        if "B" in c and "E" in c and c["B"][1] == "n":
            keys.add((int(c["B"][0]), str(c["E"][0]).strip(), str(c.get("F", ("",))[0]).strip()))
    return keys


def known_values(book, sheet, letters, last_n=4000):
    """Spellings already used in the report for the given columns (from the newest rows)."""
    rows = []
    for r, c in book.sheet_rows(sheet):
        rows.append(c)
        if len(rows) > last_n * 3:
            rows = rows[-last_n:]
    rows = rows[-last_n:]
    out = {}
    for L in letters:
        vals = {}
        for c in rows:
            if L in c and c[L][1] == "s" and c[L][0]:
                vals[c[L][0].strip().lower()] = c[L][0].strip()
        out[L] = vals
    return out


def style_map(book, sheet, letters, depth=600):
    """For each column and value kind (s=text, n=number), the style id used by the newest rows."""
    xml = book.xml(sheet)
    tail = xml[: xml.rfind("</sheetData>")]
    chunk = tail[-max(200000, depth * 700):]
    rows = ROW_RE.findall(chunk)
    sm = defaultdict(dict)
    for _, body in reversed(rows):
        for m in CELL_RE.finditer(body or ""):
            col, _, attrs, inner = m.groups()
            if col not in letters or not inner:
                continue
            s = re.search(r' s="(\d+)"', attrs)
            if not s:
                continue
            t = re.search(r' t="(\w+)"', attrs)
            kind = "s" if (t and t.group(1) in ("s", "str", "inlineStr")) else "n"
            sm[col].setdefault(kind, s.group(1))
    return sm


# --------------------------------------------------------------------------------------
# writing
# --------------------------------------------------------------------------------------
def _cell_xml(ref, value, kind, style):
    s = f' s="{style}"' if style else ""
    if kind == "n":
        if isinstance(value, float):
            txt = repr(value)
        else:
            txt = str(value)
        return f'<c r="{ref}"{s}><v>{txt}</v></c>'
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'


def month_label(book, d):
    pref = MONTHS[d.month - 1]
    yy = "%02d" % (d.year % 100)
    for n in book.sheet_path:
        m = re.fullmatch(r"([A-Za-z]+)'(\d\d)", n)
        if m and m.group(2) == yy and m.group(1)[:3].lower() == pref.lower():
            return n, True
    return f"{pref}'{yy}", False


def plan_and_append(report_bytes, records):
    """
    Validate and append. Returns dict(ok, new_bytes, messages, warnings, errors, summary)
    Nothing is written if any error is found.
    """
    res = dict(ok=False, new_bytes=None, messages=[], warnings=[], errors=[], summary=[])
    if not records:
        res["errors"].append("No usable rows were found in the EOD files.")
        return res
    book = Book(report_bytes)
    for t in {TARGETS[r["target"]]["sheet"] for r in records}:
        if t not in book.sheet_path:
            res["errors"].append(f"The report has no '{t}' sheet.")
    if res["errors"]:
        return res

    # ---- duplicate guard (date + agent + task already in that sheet)
    keys_cache = {}
    dup = set()
    for r in records:
        tgt = TARGETS[r["target"]]
        sh = tgt["sheet"]
        if sh not in keys_cache:
            keys_cache[sh] = existing_keys(book, sh)
        k = (to_serial(r["date"]), r["name"], tgt["task"])
        if k in keys_cache[sh]:
            dup.add((r["date"].strftime("%d-%b-%Y"), r["name"], tgt["label"]))
    if dup:
        res["errors"].append("Already loaded (date / agent / sheet): " +
                             "; ".join(f"{a} {b} - {c}" for a, b, c in sorted(dup)) +
                             ". Nothing was changed. If this is a correction, restore the report from your backup first.")
        return res

    # ---- unknown month sheet
    labels = {}
    for r in records:
        lab, exists = month_label(book, r["date"])
        labels[r["date"].strftime("%Y-%m")] = (lab, exists)
        r["month"], r["week"] = lab, week_label(r["date"])
    for ym, (lab, exists) in labels.items():
        if not exists:
            res["warnings"].append(
                f"The report has no '{lab}' sheet yet, so the Daily date cell could not be set. Rows were added with Month = '{lab}'. "
                "Create the month sheet (see README: new month) and the Daily/Weekly sheets will pick the rows up.")

    # ---- append per sheet
    by_sheet = defaultdict(list)
    for r in records:
        by_sheet[TARGETS[r["target"]]["sheet"]].append(r)

    counts = defaultdict(int)
    for sh, recs in by_sheet.items():
        xml = book.xml(sh)
        end = xml.rfind("</sheetData>")
        last = book.last_row_number(sh)
        letters = set()
        for r in recs:
            letters |= set(TARGETS[r["target"]]["cols"].keys())
        sm = style_map(book, sh, letters)
        # spelling check against what the report already uses
        known_cache = {}
        new_rows = []
        n = last
        for r in recs:
            tgt = TARGETS[r["target"]]
            n += 1
            cells = []
            for L in sorted(tgt["cols"], key=col_to_num):
                f = tgt["cols"][L]
                val = {"month": r["month"], "week": r["week"], "task": tgt["task"]}.get(f, r.get(f))
                if f == "date":
                    val = to_serial(r["date"])
                if f == "invdate":
                    val = to_serial(r["invdate"]) if r.get("invdate") else None
                if val is None or val == "":
                    continue
                if f in tgt["known"].values():
                    key = (sh, L)
                    if key not in known_cache:
                        known_cache[key] = known_values(book, sh, [L])[L]
                    kv = known_cache[key]
                    if isinstance(val, str):
                        hit = kv.get(val.strip().lower())
                        if hit:
                            val = hit
                        elif kv:
                            w = f"New {f} '{val}' in {tgt['label']} is not used elsewhere in the report; kept as typed."
                            if w not in res["warnings"]:
                                res["warnings"].append(w)
                kind = "n" if isinstance(val, (int, float)) and not isinstance(val, bool) else "s"
                style = sm.get(L, {}).get(kind) or next(iter(sm.get(L, {}).values()), None)
                cells.append(_cell_xml(f"{L}{n}", val, kind, style))
            new_rows.append(f'<row r="{n}">' + "".join(cells) + "</row>")
            counts[(tgt["label"], r["name"], r["date"].strftime("%d-%b-%Y"))] += 1
        xml = xml[:end] + "".join(new_rows) + xml[end:]
        xml = re.sub(r'<dimension ref="A1:([A-Z]+)\d+"/>', lambda m: f'<dimension ref="A1:{m.group(1)}{n}"/>', xml, count=1)
        book.set_xml(sh, xml)
        res["messages"].append(f"{sh}: +{len(recs)} rows (now {n} rows)")

    # ---- Daily date cell on the month sheet
    latest = max(r["date"] for r in records)
    lab, exists = month_label(book, latest)
    if exists:
        xml = book.xml(lab)
        m = re.search(r'<c r="B6"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S)
        if m:
            cur = re.search(r"<v>(.*?)</v>", m.group(2) or "")
            cur_v = float(cur.group(1)) if cur else 0
            if to_serial(latest) >= cur_v:
                attrs = re.sub(r' t="\w+"', "", m.group(1))
                new = f'<c r="B6"{attrs}><v>{to_serial(latest)}</v></c>'
                xml = xml[:m.start()] + new + xml[m.end():]
                book.set_xml(lab, xml)
                res["messages"].append(f"{lab}: Daily date (B6) set to {latest.strftime('%d-%b-%Y')}")
            else:
                res["messages"].append(f"{lab}: Daily date (B6) left unchanged because the rows are older than its current date")
        else:
            res["warnings"].append(f"Could not find cell B6 on '{lab}', so the Daily date was not set.")

    # ---- recalc on open
    wbxml = book.zin.read("xl/workbook.xml").decode("utf8")
    if "fullCalcOnLoad" not in wbxml:
        wbxml = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml, count=1)
    book.changed["xl/workbook.xml"] = wbxml

    # ---- write the zip
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zo:
        for info in book.zin.infolist():
            data = book.changed.get(info.filename)
            if data is None:
                data = book.zin.read(info.filename)
            elif isinstance(data, str):
                data = data.encode("utf8")
            zo.writestr(info.filename, data)
    res["new_bytes"] = out.getvalue()
    res["summary"] = [(a, b, c, n) for (a, b, c), n in sorted(counts.items())]
    res["ok"] = True
    return res


def process(report_bytes, eod_files):
    """Main entry used by the app. eod_files: list of (name, bytes)."""
    records, warns, errs = read_eod(eod_files)
    if errs:
        return dict(ok=False, errors=errs, warnings=warns, messages=[], summary=[], new_bytes=None)
    res = plan_and_append(report_bytes, records)
    res["warnings"] = warns + res["warnings"]
    return res


def preview(eod_files):
    """Quick look for the app before the person clicks Update."""
    records, warns, errs = read_eod(eod_files)
    tab = defaultdict(int)
    for r in records:
        tab[(TARGETS[r["target"]]["label"], r["name"], r["date"].strftime("%d-%b-%Y"))] += 1
    return [(a, b, c, n) for (a, b, c), n in sorted(tab.items())], warns, errs
