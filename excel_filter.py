"""Shift-report summary: filter a spreadsheet by date range, keyword and hour.

Code is English, the UI is pt-BR on purpose. Run with: python excel_filter.py
"""

import io
import json
import queue
import calendar
import posixpath
import re
import threading
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

BUILD = "15/09 01:18"  # carimbo do último commit, para saber qual versão está rodando
DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y",
                "%d%m%Y", "%d%m%y")
EXCEL_EPOCH = date(1899, 12, 30)
HOUR_RE = re.compile(r"\s*\d{1,2}\s*(h|:\d{2})\s*[-–às]*\s*(\d{1,2}\s*(h|:\d{2}))?\s*", re.I)
NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
# elements that sit between </sheetData> and <autoFilter> in the schema
BEFORE_AUTOFILTER = (b"sheetCalcPr", b"sheetProtection", b"protectedRanges", b"scenarios")


def parse_date(value):
    """Cell (date/datetime/text) -> date, or None when it cannot be read."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            pass
    return None


def mask_date(text):
    """What was typed, as dd/mm/aaaa — the slashes appear on their own."""
    digits = re.sub(r"\D", "", text)[:8]
    return "/".join(part for part in (digits[:2], digits[2:4], digits[4:]) if part)


def read_typed_date(text, month=None):
    """Text from a date field -> date, or None.

    A bare day ("6", "15") means that day of the report's month, which is the
    common case here: the whole spreadsheet is one month.
    """
    text = text.strip()
    if not text:
        return None
    if month and re.fullmatch(r"\d{1,2}", text):
        year, number = month
        try:
            return date(year, number, int(text))
        except ValueError:
            return None
    return parse_date(text)


def _header(sheet):
    return [str(c.value).strip() if c.value is not None else "" for c in sheet[1]]


def read(path, sheet=None, book=None):
    """(sheet name, headers, [(row number, values)]) with calculated values.

    Without `sheet`, picks the one that is an actual table: filled header on row 1
    and the most rows — "Tabela de Dados" among the 35 sheets of the shift report.
    Reads once: the UI keeps the rows and refilters in memory, otherwise every
    keystroke would reopen a file that is 59MB here.
    """
    book = book or load_workbook(path, data_only=True)
    if sheet is None:
        ranked = [(sum(1 for h in _header(book[n]) if h), book[n].max_row, n) for n in book.sheetnames]
        sheet = max(ranked, key=lambda c: (c[0] >= 3, c[1]))[2]
    ws = book[sheet]
    headers = _header(ws)
    rows = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None and str(v).strip() != "" for v in values):
            rows.append((r, values))
    return sheet, headers, rows


# --- Shift report (the "Relatório de Turno" workbook) --------------------------
# Its "Tabela de Dados" sheet is broken: column A was dragged down one day per
# ROW instead of per day, so it drifts 01/09 -> 24/10 for a single month, and the
# INDIRECT that pulls the text picks the day sheet by TEXT(A,"d") — meaning each
# block of 24 hours reads 24 DIFFERENT days. The day sheets are the real source.
LINE_RE = re.compile(r"LINHA\s+([A-Z]?\d+[A-Z]?)", re.I)  # "LINHA D02 - 1º TURNO" -> D02
SHIFT_FIELD_NAMES = {"horario": "horário", "formato": "formato", "plan": "plan",
                     "produzido": "produzido", "apontado": "apontado",
                     "perdas": "principais perdas"}
SHIFT_HOURS = 8  # each shift block lists eight hours


class _Cell:
    """What the reader asks of a cell: its value and where it is."""
    __slots__ = ("value", "row", "column")

    def __init__(self, value, row, column):
        self.value, self.row, self.column = value, row, column


class _Range:
    """A merged range, the four numbers the reader uses."""
    __slots__ = ("min_row", "max_row", "min_col", "max_col")

    def __init__(self, min_row, max_row, min_col, max_col):
        self.min_row, self.max_row = min_row, max_row
        self.min_col, self.max_col = min_col, max_col


class _Merged:
    def __init__(self, ranges):
        self.ranges = ranges


class FastSheet:
    """A day sheet read in streaming mode: only the rows the report lives in.

    openpyxl in normal mode takes 25s to open one 58MB workbook and 0.4s to pull
    the data out of it — it parses every cell of every sheet. In read-only mode the
    same file opens in 0.1s, but read-only sheets have no random access and no
    merged ranges, which the reader needs. So the top rows are copied into a grid
    and the merges are read straight from the sheet XML.
    """
    ROWS = 70  # header block, hour rows and some slack

    def __init__(self, values, merges):
        self._values, self.merged_cells = values, _Merged(merges)

    def cell(self, row, column):
        return _Cell(self._values.get((row, column)), row, column)

    def iter_rows(self, min_row=1, max_row=None, **ignored):
        last = min(max_row or self.ROWS, self.ROWS)
        width = max((c for _, c in self._values), default=0)
        for r in range(min_row, last + 1):
            yield tuple(_Cell(self._values.get((r, c)), r, c) for c in range(1, width + 1))


def _merges_of(archive, sheet_xml):
    """Merged ranges of a sheet, straight from its XML — read-only mode drops them."""
    found = []
    for ref in re.findall(rb'<mergeCell ref="([A-Z]+\d+):([A-Z]+\d+)"', archive.read(sheet_xml)):
        (first, last) = (ref[0].decode(), ref[1].decode())
        first_col, first_row = coordinate_from_string(first)
        last_col, last_row = coordinate_from_string(last)
        found.append(_Range(first_row, last_row, column_index_from_string(first_col),
                            column_index_from_string(last_col)))
    return found


def fast_sheets(source, names):
    """{name: FastSheet} for the day sheets, without parsing the whole workbook."""
    book = load_workbook(source, data_only=True, read_only=True)
    sheets = {}
    try:
        with zipfile.ZipFile(source) as archive:
            for name in names:
                values = {}
                for r, row in enumerate(book[name].iter_rows(min_row=1, max_row=FastSheet.ROWS,
                                                             values_only=True), 1):
                    for c, value in enumerate(row, 1):
                        if value is not None:
                            values[(r, c)] = value
                sheets[name] = FastSheet(values, _merges_of(archive, _sheet_path(archive, name)))
    finally:
        book.close()
    if hasattr(source, "seek"):
        source.seek(0)
    return sheets


def shift_blocks(sheet):
    """[{field: column, "row": header row, "until": where the block ends}] per shift.

    Found by reading the labels, never by fixed coordinates: the workbooks come in
    several versions of the same model — November has the header on row 26 with
    blocks at B/AI/BP, February has row 27 with B/AJ/BR — and a hardcoded layout
    silently drops a whole month.
    """
    spots = sorted((c.row, c.column) for row in sheet.iter_rows(min_row=1, max_row=60)
                   for c in row if isinstance(c.value, str)
                   and c.value.strip().casefold() == SHIFT_FIELD_NAMES["horario"])
    blocks = []
    for i, (row, column) in enumerate(spots):
        until = spots[i + 1][1] if i + 1 < len(spots) and spots[i + 1][0] == row else column + 33
        block = {"row": row, "until": until, "horario": column}
        for field, label in SHIFT_FIELD_NAMES.items():
            for c in range(column, until):
                value = sheet.cell(row, c).value
                if isinstance(value, str) and value.strip().casefold() == label:
                    block[field] = c
                    break
        if "apontado" in block:
            block["diferenca"] = block["apontado"] + 2  # this one has no label of its own
        blocks.append(block)
    return blocks


def is_shift_report(book):
    """Whether this workbook is a shift report with one sheet per day."""
    days = [n for n in book.sheetnames if n.isdigit()]
    return len(days) >= 5 and bool(shift_blocks(book[days[0]]))


def reference_month(book):
    """(year, month, sheet that said so) of the report, or None.

    The day sheets are named "1".."31" and say nothing about the month. Two sheets
    can disagree — the broken "Tabela de Dados" starts at 01/09 while the planning
    sheet lists 01/11..30/11 — so only a sheet that reads like a calendar counts:
    each day exactly once, in a consecutive run.
    """
    for name in (n for n in book.sheetnames if not n.isdigit()):
        days = [v.date() for row in book[name].iter_rows(max_row=60, max_col=30, values_only=True)
                for v in row if isinstance(v, datetime)]
        unique = sorted(set(days))
        if len(unique) >= 5 and len(unique) == len(days) and \
                (unique[-1] - unique[0]).days == len(unique) - 1:
            return unique[0].year, unique[0].month, name
    return None


def _block_field(sheet, block, label):
    """Value of a labelled cell in a shift block's header area, e.g. "SKU DESCRIÇÃO".

    The label sits at a slightly different column in each block (B6, AJ6, BR6), so
    it is searched for instead of hardcoded; the value is the first filled cell to
    its right.
    """
    for r in range(1, block["row"]):
        for c in range(block["horario"], block["until"]):
            cell = sheet.cell(r, c).value
            if isinstance(cell, str) and label in cell.upper():
                for right in range(c + 1, min(c + 8, block["until"])):
                    value = sheet.cell(r, right).value
                    if value not in (None, ""):
                        return str(value).strip()
    return ""


def _block_line(sheet, block):
    """"LINHA D02 - 1º TURNO" over a block -> "D02", or "" when the block has no title."""
    for r in (1, 2, 3):
        for c in range(block["horario"], min(block["horario"] + 4, block["until"])):
            value = sheet.cell(r, c).value
            found = LINE_RE.search(str(value)) if isinstance(value, str) else None
            if found:
                return found.group(1).upper()
    return ""


def _span_label(hours):
    """["06h - 07h", ..., "09h - 10h"] -> "06h - 10h" — one occurrence, four merged hours."""
    first, last = hours[0], hours[-1]
    if first == last:
        return first
    end = last.split("-")[-1].strip() if "-" in last else last
    return f"{first.split('-')[0].strip()} - {end}"


def read_shift_report(book, year, month, sheets=None):
    """(headers, rows) built from the day sheets, with the date rebuilt from the sheet name.

    Sheet "6" is day 6 of the reference month; each sheet holds three shift blocks
    of eight hours side by side. The losses cell is usually MERGED over several
    hour rows — one report covering 06h to 10h — so the row that owns the text
    carries the whole span, and "Horas" lists every hour it covers so the hour
    filter finds it. Days the month does not have are skipped.
    """
    headers = ["Data", "Linha", "Dia", "Turno", "SKU", "Trocas", "Horário", "Período", "Horas",
               "Formato", "Plan", "Produzido", "Apontado", "Diferença", "Principais perdas"]
    rows, number, seen_line = [], 1, ""
    for name in sorted((n for n in book.sheetnames if n.isdigit()), key=int):
        day = int(name)
        try:
            when = date(year, month, day)
        except ValueError:
            continue  # 31 in a 30-day month
        sheet = sheets[name] if sheets else book[name]
        blocks = shift_blocks(sheet)
        # only the first shift is usually titled "LINHA D01 - 1º TURNO"; the others
        # belong to the same line, and leaving them nameless threw them into a
        # separate block at the end of the list, as if the day stopped at 14h
        titles = [_block_line(sheet, block) for block in blocks]
        seen_line = next((t for t in titles if t), seen_line)
        for shift, block in enumerate(blocks, 1):
            first_row = block["row"] + 1
            last_row = block["row"] + SHIFT_HOURS
            line = titles[shift - 1] or seen_line or "—"
            sku = _block_field(sheet, block, "SKU")
            swaps = [t for i in (1, 2, 3)
                     if (t := _block_field(sheet, block, f"TROCA {i}"))
                     and t.upper() not in ("N/A", "NA", "-", "0")]
            losses = block.get("perdas")
            merges = {r.min_row: min(r.max_row, last_row) for r in sheet.merged_cells.ranges
                      if losses and r.min_col <= losses <= r.max_col and r.min_row >= first_row}
            hour_of = {r: str(sheet.cell(r, block["horario"]).value or "").strip()
                       for r in range(first_row, last_row + 1)}
            for r in range(first_row, last_row + 1):
                cell = lambda field: (sheet.cell(r, block[field]).value
                                      if field in block else None)
                if not hour_of[r]:
                    continue
                covered = [hour_of[x] for x in range(r, merges.get(r, r) + 1) if hour_of.get(x)]
                number += 1
                rows.append((number, [when, line, day, shift, sku, " · ".join(swaps),
                                      hour_of[r], _span_label(covered),
                                      " | ".join(covered), cell("formato"), cell("plan"),
                                      cell("produzido"), cell("apontado"), cell("diferenca"),
                                      cell("perdas")]))
    return headers, rows


MONTH_NAMES = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
               "agosto", "setembro", "outubro", "novembro", "dezembro")


def month_from_name(name):
    """(year or None, month) read from a file name, or None.

    The workbooks in the bundle are named by month — "Relatório de Turno - Linha 1-
    Março.xlsx" — and carry no year, so the year has to come from somewhere else.
    """
    plain = unicodedata.normalize("NFKD", str(name).lower()).encode("ascii", "ignore").decode()
    stamped = re.search(r"(20\d{2})[-_/.]?(\d{1,2})|(\d{1,2})[-_/.](20\d{2})", plain)
    if stamped:
        year, month = (stamped.group(1), stamped.group(2)) if stamped.group(1) \
            else (stamped.group(4), stamped.group(3))
        if 1 <= int(month) <= 12:
            return int(year), int(month)
    for number, month_name in enumerate(MONTH_NAMES, 1):
        plain_name = unicodedata.normalize("NFKD", month_name).encode("ascii", "ignore").decode()
        if re.search(rf"\b{plain_name}\b", plain):
            return None, number
    return None


PLACEHOLDER_YEAR = 2000  # leap, so 29/02 survives until the real year is known


def _with_source(headers, rows, name):
    """Tag every row with the file it came from, so a day can say which report it is."""
    return headers + ["Arquivo"], [(number, list(values) + [name]) for number, values in rows]


def _with_year(headers, rows, year):
    """Rewrite the date of every row to the year that was only known at the end."""
    date_at = _at(headers, "Data")
    if date_at is None:
        return rows
    fixed = []
    for number, values in rows:
        day = values[date_at]
        try:
            values = list(values)
            values[date_at] = date(year, day.month, day.day)
        except ValueError:
            continue  # 29/02 in a year that has no 29/02
        fixed.append((number, values))
    return fixed


def load_bundle(path, progress=None):
    """Every .xlsx inside a zip read as one dataset, each with its own month.

    Each workbook names its own month (planning sheet, or the month in the file
    name), so the screen has no month field to fill: the date range does that job.
    """
    parts, months, headers = [], [], []
    with zipfile.ZipFile(path) as bundle:
        inside = sorted(n for n in bundle.namelist()
                        if n.lower().endswith(".xlsx") and not n.startswith("__MACOSX")
                        and not Path(n).name.startswith("~$"))
        if not inside:
            raise ValueError("O zip não tem nenhuma planilha .xlsx dentro.")
        for i, name in enumerate(inside, 1):
            short = Path(name).name
            if progress:
                progress(i, len(inside), short)
            data = io.BytesIO(bundle.read(name))
            book = load_workbook(data, data_only=True, read_only=True)
            try:
                if not is_shift_report(book):
                    months.append((short, None, "não é relatório de turno"))
                    continue
                found = reference_month(book)
                named = month_from_name(short)
                if found and named and named[1] != found[1]:
                    # the file name and the planning sheet disagree: the name is chosen
                    # per report, the sheet is often left over from the template. The
                    # name wins and the conflict is said out loud, never swallowed.
                    when = (named[0], named[1])
                    source = "nome do arquivo (a aba {} dizia {:02d}/{})".format(
                        found[2], found[1], found[0])
                elif found:
                    when, source = found[:2], f"aba {found[2]}"
                elif named and named[0]:
                    when, source = named, "nome do arquivo"
                elif named:
                    when, source = (None, named[1]), "nome do arquivo (sem ano)"
                else:
                    when, source = (None, None), "não encontrado"
                months.append((short, when, source))
                # the year can still be missing here — it is borrowed from the other
                # workbooks after the loop — so read with a placeholder and remap below
                days = [n for n in book.sheetnames if n.isdigit()]
                headers, rows = read_shift_report(book, when[0] or PLACEHOLDER_YEAR, when[1],
                                                  sheets=fast_sheets(data, days))
                headers, rows = _with_source(headers, rows, short)
                parts.append((short, headers, rows))
            finally:
                book.close()
    # a workbook that only named the month borrows the year the others agree on
    years = Counter(when[0] for _, when, _ in months if when and when[0])
    year = years.most_common(1)[0][0] if years else date.today().year
    borrowed = {short for short, when, _ in months if when and when[1] and not when[0]}
    for i, (short, when, source) in enumerate(months):
        if short in borrowed:
            months[i] = (short, (year, when[1]), source + f", ano {year} dos outros arquivos")
    if borrowed:
        parts = [(short, part_headers, _with_year(part_headers, rows, year)
                  if short in borrowed else rows)
                 for short, part_headers, rows in parts]
    return parts, months, headers


HISTORY_FILE = Path.home() / ".resumo-ocorrencias" / "historico.json"
HISTORY_LIMIT = 15


def history():
    """Spreadsheets opened before, newest first. A missing or broken file is no history."""
    try:
        saved = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return [e for e in saved if isinstance(e, dict) and e.get("path")]


def remember(path, entry):
    """Record this spreadsheet at the top of the history and keep the file small."""
    kept = [e for e in history() if e["path"] != str(path)]
    kept.insert(0, {"path": str(path), "when": datetime.now().isoformat(timespec="seconds"), **entry})
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE.write_text(json.dumps(kept[:HISTORY_LIMIT], ensure_ascii=False, indent=1),
                            encoding="utf-8")
    return kept[:HISTORY_LIMIT]


def forget(path):
    """Take one spreadsheet out of the history and give back what is left."""
    kept = [e for e in history() if e["path"] != str(path)]
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE.write_text(json.dumps(kept, ensure_ascii=False, indent=1), encoding="utf-8")
    return kept


def forget_all():
    """Wipe the history."""
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE.write_text("[]", encoding="utf-8")
    return []


def load(path, month=None, progress=None):
    """Open the file once and give back everything the screen needs.

    A .zip is every shift report inside it, merged. A single shift report is read
    from the day sheets (its "Tabela de Dados" has broken dates); anything else
    falls back to the plain table reader.
    """
    if str(path).lower().endswith(".zip"):
        parts, months, headers = load_bundle(path, progress)
        rows, number = [], 0
        for short, part_headers, part_rows in parts:
            for _, values in part_rows:
                number += 1
                rows.append((number, values))
        read_months = [when for _, when, _ in months if when and all(when)]
        return {"kind": "turno", "sheet": f"{len(parts)} planilhas do zip", "headers": headers,
                "rows": rows, "month": min(read_months) if read_months else None,
                "month_from": None, "month_guessed": False, "months": months, "path": str(path),
                "columns": {"date": "Data", "hour": "Horas", "label": "Período",
                            "text": "Principais perdas", "line": "Linha"}}
    book = load_workbook(path, data_only=True, read_only=True)
    try:
        if is_shift_report(book):
            found = reference_month(book)
            when = month or (found[:2] if found else (date.today().year, date.today().month))
            days = [n for n in book.sheetnames if n.isdigit()]
            headers, rows = read_shift_report(book, *when, sheets=fast_sheets(path, days))
            headers, rows = _with_source(headers, rows, Path(path).name)
            return {"kind": "turno", "sheet": "abas por dia", "headers": headers, "rows": rows,
                    "path": str(path), "month": when,
                    "month_from": found[2] if found and not month else None,
                    # no sheet named the month and nobody asked for one: it is a guess,
                    # and a guessed date must say so instead of passing for a read value
                    "month_guessed": not month and not found,
                    "columns": {"date": "Data", "hour": "Horas", "label": "Período",
                                "text": "Principais perdas", "line": "Linha"}}
        book.close()  # the plain table reader needs random access, so it opens its own
        book = load_workbook(path, data_only=True)
        sheet, headers, rows = read(path, book=book)
        date_col, hour_col, text_col = detect_columns(headers, rows)
        return {"kind": "tabela", "sheet": sheet, "headers": headers, "rows": rows,
                "path": str(path), "month": None, "month_from": None, "month_guessed": False,
                "columns": {"date": date_col, "hour": hour_col, "label": hour_col,
                            "text": text_col, "line": "Linha" if "Linha" in headers else ""}}
    finally:
        book.close()


def sentences_with(text, terms):
    """Only the sentences carrying one of the terms — a cell holds a whole shift log."""
    if not terms:
        return str(text).strip()
    parts = re.split(r"(?<=[.;!?])\s+|\n+", str(text))
    hits = [p.strip() for p in parts if p.strip() and any(t in p.casefold() for t in terms)]
    return "\n".join(hits) if hits else str(text).strip()


def distinct(headers, rows, column):
    """Sorted distinct values of a column — feeds the hour picker."""
    i = _column_index(headers, column) - 1
    return sorted({str(v[i]).strip() for _, v in rows if v[i] is not None and str(v[i]).strip()})


def detect_columns(headers, rows):
    """Find the date, hour and occurrence columns: (date_col, hour_col, text_col).

    Goes by content, not by header name — otherwise "Hora" (1..24) would beat
    "Horário" ("07h - 08h") in this spreadsheet.
    """
    def score(i, test):
        return sum(1 for _, v in rows if v[i] is not None and str(v[i]).strip() and test(v[i]))

    def best(scores):  # needs half the rows, so it never guesses a random column
        i, points = max(scores.items(), key=lambda x: x[1], default=(None, 0))
        return headers[i] if points >= max(1, len(rows) // 2) else ""

    used = [i for i in range(len(headers)) if headers[i]]
    dates = {i: score(i, lambda v: parse_date(v) is not None) for i in used}
    hours = {i: score(i, lambda v: bool(HOUR_RE.fullmatch(str(v)))) for i in used}
    # ponytail: the occurrence is the longest text column. If two long text
    # columns ever show up, this becomes a picker on screen.
    length = {i: sum(len(str(v[i])) for _, v in rows if isinstance(v[i], str)) for i in used}
    text = max(length.items(), key=lambda x: x[1], default=(None, 0))
    return best(dates), best(hours), headers[text[0]] if text[1] > 40 else ""


def _column_index(headers, name):
    found = [i for i, h in enumerate(headers, 1) if h == name]
    if len(found) != 1:
        raise ValueError(f"Coluna {name!r}: {len(found)} colunas com esse nome no cabeçalho.")
    return found[0]


def _matcher(headers, text_col, terms, date_col, start, end, hour_col, hour,
             line_col=None, lines=()):
    """One sieve, used by the summary and by the export — never two rulesets."""
    if start and end and start > end:
        raise ValueError("Data inicial maior que a data final.")
    if terms and not text_col:
        raise ValueError("A planilha não tem uma coluna de texto para procurar a palavra.")
    if (start or end) and not date_col:
        raise ValueError("A planilha não tem uma coluna de data.")
    if hour and not hour_col:
        raise ValueError("A planilha não tem uma coluna de hora.")
    if lines and not line_col:
        raise ValueError("A planilha não diz a linha de produção.")
    ti = _column_index(headers, text_col) - 1 if terms else None
    di = _column_index(headers, date_col) - 1 if (start or end) else None
    hi = _column_index(headers, hour_col) - 1 if hour else None
    li = _column_index(headers, line_col) - 1 if lines else None
    wanted = {str(l).strip().casefold() for l in lines}

    def matches(row_number, values):
        if li is not None and str(values[li]).strip().casefold() not in wanted:
            return False
        if ti is not None:
            value = values[ti]
            if value is None or not any(t in str(value).casefold() for t in terms):
                return False
        if hi is not None:
            value = values[hi]
            if value is None or hour.casefold() not in str(value).casefold():
                return False
        if di is not None:
            raw = values[di]
            if raw is None or str(raw).strip() == "":
                return False
            day = parse_date(raw)
            if day is None:
                raise ValueError(f"Data ilegível na linha {row_number}: {raw!r}")
            if (start and day < start) or (end and day > end):
                return False
        return True

    return matches, ti, di, hi


def _as_number(value):
    """The cell as a number, or None. 113 cells of "Produzido" hold the text "0"
    instead of the number, so a plain sum() blows up over the whole sheet."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value
    text = str(value).strip()
    return int(text) if re.fullmatch(r"-?\d+", text) else None


def _total(values):
    """(sum of the readable cells, whether every cell could be read).

    A cell nobody could read must never pass for a zero inside a total, so what
    is left out is said out loud instead of silently added as nothing.
    """
    numbers = [n for n in map(_as_number, values) if n is not None]
    return sum(numbers), len(numbers) == len(values)


def _summable(values):
    """Whether adding this column up means anything."""
    numbers = [n for n in map(_as_number, values) if n is not None]
    if not numbers or len(numbers) < 0.8 * len(values):
        return False
    # ponytail: day, shift, line and hour look numeric but are codes. Told apart
    # by being small non-negative integers; if a real 0..31 quantity ever shows
    # up, this becomes a picker on screen.
    return not all(float(n).is_integer() and 0 <= n <= 31 for n in numbers)


def summarize(headers, rows, text_col=None, terms=(), date_col=None, start=None, end=None,
              hour_col=None, hour=None, line_col=None, lines=(), only_with_text=False):
    """Summary of the rows passing the filters.

    Returns the columns that carry content (empty ones are dropped), the matching
    rows, the totals of the summable columns and the date range found. No filter
    at all is valid: it summarises the whole sheet.
    """
    terms = [t.strip().casefold() for t in terms if t.strip()]
    matches, *_ = _matcher(headers, text_col, terms, date_col, start, end, hour_col, hour,
                           line_col, lines)
    ti = _column_index(headers, text_col) - 1 if (only_with_text and text_col) else None
    di = _column_index(headers, date_col) - 1 if date_col else None
    found = [(r, v) for r, v in rows
             if matches(r, v) and (ti is None or str(v[ti] or "").strip())]
    if line_col and line_col in headers:  # keep each line's occurrences together
        by_line = _column_index(headers, line_col) - 1
        found.sort(key=lambda item: str(item[1][by_line]))

    used = [i for i in range(len(headers))
            if headers[i] and any(v[i] not in (None, "", 0) for _, v in found)]
    totals, partial = {}, set()
    for i in used:
        column = [v[i] for _, v in found if v[i] is not None and str(v[i]).strip() != ""]
        if column and _summable(column):
            totals[headers[i]], complete = _total(column)
            if not complete:
                partial.add(headers[i])
    days = sorted(d for d in ((parse_date(v[di]) if di is not None else None) for _, v in found) if d)

    return {
        "matched": found,  # raw rows, for the numbers behind the second tab
        "columns": [headers[i] for i in used],
        "rows": [[v[i] for i in used] for _, v in found],
        "totals": totals,
        "partial": partial,  # totals where some cell could not be read
        "read": len(rows),
        "found": len(found),
        "period": (days[0], days[-1]) if days else None,
    }


def _sheet_path(archive, sheet):
    """xl/worksheets/sheetN.xml of a sheet, via the package rels — openpyxl's
    ws.path answers sheetNone.xml when the workbook was read from disk."""
    book = ET.fromstring(archive.read("xl/workbook.xml"))
    rid = next((s.get(NS_REL + "id") for s in book.iter(NS_MAIN + "sheet") if s.get("name") == sheet), None)
    if rid is None:
        raise ValueError(f"Aba {sheet!r} não encontrada no arquivo.")
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)


def _filter_column(index, operators, match_all):
    """<filterColumn> with customFilters; colId is 0-based inside the range."""
    tests = "".join(f'<customFilter operator="{op}" val="{v}"/>' for op, v in operators)
    return (f'<filterColumn colId="{index}">'
            f'<customFilters and="{1 if match_all else 0}">{tests}</customFilters>'
            f"</filterColumn>")


def _patch_sheet(xml, ref, hidden, criteria):
    """Mark rows hidden and replace the sheet's <autoFilter>. Nothing else changes."""
    def row(match):
        attrs = re.sub(rb'\s+hidden="[^"]*"', b"", match.group(1))
        number = int(re.search(rb'\br="(\d+)"', attrs).group(1))
        if number in hidden:
            attrs += b' hidden="1"'
        return b"<row" + attrs + match.group(2) + b">"

    xml = re.sub(rb"<row\b([^>]*?)(/?)>", row, xml)
    xml = re.sub(rb"<autoFilter\b[^>]*/>|<autoFilter\b.*?</autoFilter>", b"", xml, flags=re.S)

    at = xml.index(b"</sheetData>") + len(b"</sheetData>")
    for tag in BEFORE_AUTOFILTER:  # autoFilter has a fixed slot in the schema
        found = re.compile(rb"\A<" + tag + rb"\b[^>]*?(/>|>.*?</" + tag + rb">)", re.S).match(xml, at)
        if found:
            at = found.end()
    return xml[:at] + f'<autoFilter ref="{ref}">{"".join(criteria)}</autoFilter>'.encode() + xml[at:]


def export(src, dst, sheet, headers, rows, text_col=None, terms=(), date_col=None, start=None,
           end=None, hour_col=None, hour=None, line_col=None, lines=()):
    """Copy of the .xlsx with AutoFilter on top of the sheet and non-matching rows hidden.

    Nothing is deleted: clearing the filter in Excel brings everything back. Only
    this sheet's XML is rewritten — images, formulas, conditional formatting, the
    sensitivity label and the rest of the package are copied byte for byte.
    """
    if Path(src) == Path(dst):
        raise ValueError("O destino não pode ser o arquivo de origem.")
    terms = [t.strip().casefold() for t in terms if t.strip()]
    matches, ti, di, hi = _matcher(headers, text_col, terms, date_col, start, end, hour_col,
                                   hour, line_col, lines)

    visible = {r for r, v in rows if matches(r, v)}
    last = max(r for r, _ in rows)
    hidden = set(range(2, last + 1)) - visible

    criteria, exact = [], True
    if terms:
        if len(terms) <= 2:  # Excel takes at most two criteria per column
            criteria.append(_filter_column(ti, [("equal", f"*{t}*") for t in terms], match_all=False))
        else:
            exact = False
    if hi is not None:
        criteria.append(f'<filterColumn colId="{hi}"><filters><filter val="{hour}"/></filters></filterColumn>')
    if di is not None:
        bounds = [("greaterThanOrEqual", (start - EXCEL_EPOCH).days) if start else None,
                  ("lessThanOrEqual", (end - EXCEL_EPOCH).days) if end else None]
        criteria.append(_filter_column(di, [b for b in bounds if b], match_all=True))
    ref = f"A1:{get_column_letter(len(headers))}{last}"

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        target = _sheet_path(zin, sheet)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == target:
                data = _patch_sheet(data, ref, hidden, criteria)
            zout.writestr(item, data)
    return len(visible), exact


def thousands(value):
    """Number the pt-BR way: 976430 -> 976.430."""
    return f"{value:,.0f}".replace(",", ".") if isinstance(value, (int, float)) else show(value)


def show(value):
    """Cell value as the UI shows it."""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y") if (value.hour, value.minute) == (0, 0) \
            else value.strftime("%d/%m/%Y %H:%M")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return "" if value is None else str(value)


def start_dir():
    """Where the file dialog opens: the Windows Downloads folder under WSL, else home."""
    def files(folder):  # system profiles ("Public", "Usuário Padrão") have none
        try:
            return sum(1 for _ in folder.iterdir())
        except OSError:
            return -1

    best = max([Path.home() / "Downloads", *Path("/mnt/c/Users").glob("*/Downloads")], key=files)
    return str(best) if files(best) > 0 else str(Path.home())


# --- Numbers behind the second tab -------------------------------------------
SKU_RE = re.compile(r"\b(\d{7,9})\b")  # the code is the stable part: the wording varies
EQUIPMENT_RE = re.compile(r"^\s*[*•]?\s*([^\n:–-]{3,40}?)\s*[:–-]\s+\S")


def _at(headers, name):
    """Index of a column, or None — the board must survive a sheet without it."""
    return headers.index(name) if name in headers else None


def _values(headers, rows, name):
    i = _at(headers, name)
    return [] if i is None else [v[i] for _, v in rows]


def _add(buckets, key, values, plan_at, made_at):
    """Collect the readable numbers of a bucket; what cannot be read is not a zero."""
    planned, produced = buckets.setdefault(key, ([], []))
    for column, into in ((plan_at, planned), (made_at, produced)):
        number = _as_number(values[column]) if column is not None else None
        if number is not None:
            into.append(number)


def _bucket(collected):
    """(planned, produced) of a bucket — planned is None when no hour of it said one.

    Five hours of February have an empty Plan cell. Summing them as zero would
    paint the bar green, as if production had beaten a target nobody wrote down.
    """
    planned, produced = collected
    return (sum(planned) if planned else None), sum(produced)


def production_by_day(headers, rows):
    """[(day, planned, produced)] — one point per day, in date order."""
    date_at, plan_at, made_at = (_at(headers, n) for n in ("Data", "Plan", "Produzido"))
    if date_at is None:
        return []
    per_day = {}
    for _, values in rows:
        day = parse_date(values[date_at])
        if day is None:
            continue
        _add(per_day, day, values, plan_at, made_at)
    return [(day, *_bucket(per_day[day])) for day in sorted(per_day)]


def production_by_hour(headers, rows):
    """[(hour, planned, produced)] in shift order — what a single day looks like."""
    hour_at, plan_at, made_at = (_at(headers, n) for n in ("Horário", "Plan", "Produzido"))
    if hour_at is None:
        return []
    per_hour = {}
    for _, values in rows:
        hour = str(values[hour_at] or "").strip()
        if not hour:
            continue
        _add(per_hour, hour, values, plan_at, made_at)
    return [(hour, *_bucket(per_hour[hour])) for hour in per_hour]


def production_by_slot(headers, rows):
    """[((day, hour), planned, produced)] — one bar per hour across the whole period.

    A few days plotted per day are four lonely bars in an empty chart; per hour
    they read like the shift panel. The rows already come in day/shift/hour order.
    """
    date_at, hour_at = _at(headers, "Data"), _at(headers, "Horário")
    plan_at, made_at = _at(headers, "Plan"), _at(headers, "Produzido")
    if date_at is None or hour_at is None:
        return []
    per_slot = {}
    for _, values in rows:
        day, hour = parse_date(values[date_at]), str(values[hour_at] or "").strip()
        if day is None or not hour:
            continue
        _add(per_slot, (day, hour), values, plan_at, made_at)
    return [(slot, *_bucket(per_slot[slot])) for slot in per_slot]


def production_by_month(headers, rows):
    """[(month, planned, produced)] — a period longer than a month is unreadable day by day."""
    months = {}
    for day, planned, produced in production_by_day(headers, rows):
        was_planned, was_produced = months.get(day.replace(day=1), (None, 0))
        months[day.replace(day=1)] = (
            None if planned is None and was_planned is None else (was_planned or 0) + (planned or 0),
            was_produced + produced)
    return [(month, *months[month]) for month in sorted(months)]


DURATION_RE = re.compile(
    r"(?:(\d+)\s*h(?:oras?|rs?)?)?\s*(?:(\d+)\s*m(?:inutos?|in)?)?\s*(?:(\d+)\s*s)?", re.I)


def changeover_minutes(text):
    """"31m19s" -> 31.3, "22min" -> 22, "0 Minutos (TROCA DE ORDEM)" -> 0, lixo -> None.

    The field is typed by hand and comes in every spelling the shift leader felt like.
    """
    said = str(text or "").strip()
    if not said:
        return None
    found = DURATION_RE.match(said)
    hours, minutes, seconds = (int(g) if g else 0 for g in found.groups()) if found else (0, 0, 0)
    if not any((hours, minutes, seconds)):
        return 0.0 if re.match(r"^0\b", said) else None
    return hours * 60 + minutes + seconds / 60


def changeover_time_by_day(headers, rows):
    """[(day, minutes of changeover)] — the field belongs to the shift, counted once."""
    date_at, shift_at, swap_at = (_at(headers, n) for n in ("Data", "Turno", "Trocas"))
    if date_at is None or swap_at is None:
        return []
    per_day, seen = {}, set()
    for _, values in rows:
        day = parse_date(values[date_at])
        key = (day, values[shift_at] if shift_at is not None else 0)
        if day is None or key in seen:
            continue
        seen.add(key)
        per_day[day] = per_day.get(day, 0) + (changeover_minutes(values[swap_at]) or 0)
    return [(day, per_day[day]) for day in sorted(per_day)]


def losses_by_equipment(headers, rows, top=10):
    """[(equipment, how many times it shows up)] — "Gasser - alarme..." counts as Gasser.

    One line of the losses cell is one occurrence: the shift log lists them one per
    line, each starting with the machine that failed.
    """
    counted, seen_as = Counter(), {}
    for text in _values(headers, rows, "Principais perdas"):
        for line in str(text or "").splitlines():
            found = EQUIPMENT_RE.match(line)
            if not found:
                continue
            name = " ".join(found.group(1).split())
            counted[name.casefold()] += 1
            seen_as.setdefault(name.casefold(), name)
    return [(seen_as[key], n) for key, n in counted.most_common(top)]


# the event is a line that STARTS with it; "atrasando o change over" in the middle of a
# sentence is context, not another changeover. "chang\w*" covers the typos found in the
# sheet: "change over", "changeover", "CHANGER OVER", "Chang Over".
CHANGEOVER_RE = re.compile(
    r"^\s*[-*•]?\s*(chang\w*\s?over|troca de (formato|variante|sku|produto))", re.I)


def _skus_of(text):
    """[(code, description)] of a SKU cell — a changeover day holds two in one cell."""
    found = []
    for line in str(text or "").splitlines():
        code = SKU_RE.search(line)
        if code:
            found.append((code.group(1), " ".join(line.split())))
    return found


def skus_by_day(headers, rows):
    """[(day, [(code, description)], [(hour, what was written, how long)])] — SKU and changeovers.

    The changeover hour only exists in the shift log ("Change over: troca de
    formato..."), while how long it took sits in the block header (TROCA 1 - TEMPO).
    """
    date_at, sku_at = _at(headers, "Data"), _at(headers, "SKU")
    hour_at, text_at = _at(headers, "Período") or _at(headers, "Horário"), _at(headers, "Principais perdas")
    swap_at, shift_at, file_at = _at(headers, "Trocas"), _at(headers, "Turno"), _at(headers, "Arquivo")
    if date_at is None or sku_at is None:
        return []
    per_day, timed, told, sources = {}, {}, set(), {}
    for _, values in rows:
        day = parse_date(values[date_at])
        if day is None:
            continue
        skus, swaps = per_day.setdefault(day, ([], {}))
        if file_at is not None and values[file_at]:
            sources.setdefault(day, str(values[file_at]))
        if swap_at is not None and str(values[swap_at] or "").strip():
            timed[(day, values[shift_at] if shift_at is not None else 0)] = str(values[swap_at])
        for sku in _skus_of(values[sku_at]):
            if sku[0] not in [code for code, _ in skus]:
                skus.append(sku)
        for line in str(values[text_at] or "" if text_at is not None else "").splitlines():
            if not CHANGEOVER_RE.match(line):
                continue
            # the same changeover is written again in every hour it spans: one entry,
            # every hour it was mentioned in
            said = " ".join(line.split())
            hour = str(values[hour_at]) if hour_at is not None else ""
            hours, _ = swaps.setdefault(said, ([], str(values[swap_at] or "")
                                               if swap_at is not None else ""))
            if hour and hour not in hours:
                hours.append(hour)
            told.add((day, values[shift_at] if shift_at is not None else 0))
    # a changeover the header timed but nobody wrote down still happened: show it
    # with the hour missing instead of dropping it
    for (day, shift), took in timed.items():
        if (day, shift) not in told:
            per_day[day][1][f"troca cronometrada no {shift}º turno, sem relato escrito"] = ([], took)
    return [{"day": day, "file": sources.get(day, ""), "skus": skus,
             "swaps": _in_hour_order([(hours, said, took)
                                      for said, (hours, took) in swaps.items()])}
            for day, (skus, swaps) in sorted(per_day.items())]


def _in_hour_order(swaps):
    """Changeovers as they happened, so the Nth one switches to the Nth SKU."""
    return sorted(swaps, key=lambda swap: swap[0][0] if swap[0] else "zz")


def skus_in(headers, rows):
    """[(code, description, hours)] for the period, busiest first."""
    hours, described = Counter(), {}
    for value in _values(headers, rows, "SKU"):
        found = SKU_RE.search(str(value or ""))
        if not found:
            continue
        code = found.group(1)
        hours[code] += 1
        text = " ".join(str(value).split())
        if len(text) > len(described.get(code, "")):
            described[code] = text
    return [(code, described[code], n) for code, n in hours.most_common()]


def production_summary(headers, rows):
    """Planned x produced for the filtered period, plus what is missing."""
    planned, planned_ok = _total(_values(headers, rows, "Plan"))
    produced, produced_ok = _total(_values(headers, rows, "Produzido"))
    return {"plan": planned, "produzido": produced, "faltou": planned - produced,
            "atingido": produced / planned if planned else None,
            "completo": planned_ok and produced_ok,
            "skus": skus_in(headers, rows), "skus_por_dia": skus_by_day(headers, rows), "dias": production_by_day(headers, rows),
            "horas": production_by_hour(headers, rows),
            "slots": production_by_slot(headers, rows),
            "meses": production_by_month(headers, rows),
            "trocas_por_dia": changeover_time_by_day(headers, rows),
            "equipamentos": losses_by_equipment(headers, rows)}


def merge(datasets):
    """Several files read as one: the rows of each, renumbered, in one dataset.

    Two bundles side by side answer "how do line 1 and line 2 compare", which is
    the whole point of keeping the line as a column.
    """
    kept = [d for d in datasets if d and d["rows"]]
    if not kept:
        raise ValueError("Nada para juntar.")
    headers = kept[0]["headers"]
    for other in kept[1:]:
        if other["headers"] != headers:
            raise ValueError("Só dá para juntar arquivos do mesmo tipo de relatório.")
    rows, number = [], 0
    for one in kept:
        for _, values in one["rows"]:
            number += 1
            rows.append((number, values))
    months = [m for one in kept for m in one.get("months", [])]
    read_months = [when for _, when, _ in months if when and all(when)]
    # a dataset that is itself a merge already lists its sources: keep all of them,
    # otherwise adding a third file erases the names of the first two
    sources = [name for one in kept
               for name in (one.get("sources") or [one.get("path", one.get("sheet", ""))])]
    return {**kept[0], "rows": rows, "months": months,
            "sheet": "{} arquivos".format(len(sources)),
            "month": min(read_months) if read_months else kept[0].get("month"),
            "sources": sources}


def draw_day_chart(canvas, days, over="#2e7d32", under="#c62828", target="#1565c0",
                   unknown="#9aa7bd"):
    """One bar per day: what was produced, against the planned line.

    Green when production reached the plan, red when it fell short, and a blue
    line at the planned level — a bar taller than the line produced above plan.

    ponytail: tk.Canvas instead of matplotlib, a dependency that triples the .exe
    to draw rectangles. Swap it the day these charts need axes, zoom or export.
    """
    canvas.delete("all")
    width = max(canvas.winfo_width(), 200)
    height = max(canvas.winfo_height(), 120)
    if not days:
        canvas.create_text(width / 2, height / 2, text="Sem dados no filtro", fill="#777")
        return
    top, floor, left = 14, height - 22, 44
    ceiling = max(max(planned or 0, produced) for _, planned, produced in days) or 1
    step = (width - left - 10) / len(days)

    def y_of(value):
        return floor - (value / ceiling) * (floor - top)

    for value in (0, ceiling / 2, ceiling):  # three reference lines, with their numbers
        canvas.create_line(left, y_of(value), width - 8, y_of(value), fill="#e6e8ec")
        canvas.create_text(left - 6, y_of(value), text=f"{value / 1000:.0f}k" if value else "0",
                           anchor="e", fill="#777", font=("TkDefaultFont", 7))
    # every label is the first of a month: this is a monthly series, not a daily one
    month_series = all(hasattr(d, "day") and d.day == 1 for d, *_ in days) and len(days) > 1
    previous = None
    for i, (day, planned, produced) in enumerate(days):
        x = left + i * step
        if isinstance(day, tuple):  # (dia, hora): o dia só aparece quando vira
            label, previous = (f"{day[0].day:02d}" if day[0] != previous else ""), day[0]
            if label and i:
                canvas.create_line(x, top, x, floor, fill="#e6e8ec")
        elif hasattr(day, "day"):
            label = f"{day.day:02d}" if not month_series else day.strftime("%m/%y")
        else:
            label = str(day).split("-")[0].strip()
        wide = min(max(step * 0.62, 2), 54)
        middle = x + (step - wide) / 2
        paint = unknown if planned is None else (over if produced >= planned else under)
        canvas.create_rectangle(middle, y_of(produced), middle + wide, floor,
                                fill=paint, width=0, tags=("slot", f"slot{i}"))
        if planned:
            canvas.create_line(x + 1, y_of(planned), x + step - 1, y_of(planned),
                               fill=target, width=3)
        if label and len(days) <= 200:
            canvas.create_text(x + step / 2, floor + 7, text=label, fill="#777",
                               font=("TkDefaultFont", 7))
        # invisible column so the whole slot answers the mouse, not only the bar
        canvas.create_rectangle(x, top, x + step, floor, outline="", fill="",
                                tags=("slot", f"slot{i}"))


def draw_time_chart(canvas, series, colour="#8a5a00"):
    """One bar per slot for a single number — here, minutes of changeover."""
    canvas.delete("all")
    width = max(canvas.winfo_width(), 200)
    height = max(canvas.winfo_height(), 90)
    if not series or not any(value for _, value in series):
        canvas.create_text(width / 2, height / 2, text="Sem troca no filtro", fill="#777")
        return
    top, floor, left = 10, height - 18, 44
    ceiling = max(value for _, value in series) or 1
    step = (width - left - 10) / len(series)
    for mark in (0, ceiling):
        y = floor - (mark / ceiling) * (floor - top)
        canvas.create_line(left, y, width - 8, y, fill="#e6e8ec")
        canvas.create_text(left - 6, y, text=f"{mark:.0f}min", anchor="e", fill="#777",
                           font=("TkDefaultFont", 7))
    month_series = all(hasattr(label, "day") and label.day == 1 for label, _ in series) \
        and len(series) > 1
    for i, (label, value) in enumerate(series):
        x = left + i * step
        wide = min(max(step * 0.62, 2), 54)
        y = floor - (value / ceiling) * (floor - top)
        canvas.create_rectangle(x + (step - wide) / 2, y, x + (step - wide) / 2 + wide, floor,
                                fill=colour, width=0, tags=("slot", f"slot{i}"))
        canvas.create_rectangle(x, top, x + step, floor, outline="", fill="",
                                tags=("slot", f"slot{i}"))
        if len(series) <= 31 and hasattr(label, "day"):
            canvas.create_text(x + step / 2, floor + 7,
                               text=label.strftime("%m/%y") if month_series else f"{label.day:02d}",
                               fill="#777", font=("TkDefaultFont", 7))


def draw_rank_chart(canvas, items, colour="#c62828", highlight=None):
    """Horizontal bars: how many times each machine shows up in the period.

    Each bar is tagged "bar<i>" so the screen can tell which machine the mouse is
    over; `highlight` is the index to emphasise.
    """
    canvas.delete("all")
    width = max(canvas.winfo_width(), 200)
    height = max(canvas.winfo_height(), 120)
    if not items:
        canvas.create_text(width / 2, height / 2, text="Sem ocorrências no filtro", fill="#777")
        return
    label_width, step = 150, min(22, height / max(len(items), 1))
    ceiling = max(n for _, n in items) or 1
    for i, (name, count) in enumerate(items):
        y = i * step + step / 2
        tag = f"bar{i}"
        on = i == highlight
        canvas.create_text(label_width - 6, y, text=name[:26], anchor="e", tags=("bar", tag),
                           font=("TkDefaultFont", 8, "bold" if on else "normal"))
        long = (count / ceiling) * (width - label_width - 46)
        canvas.create_rectangle(label_width, y - step / 3, label_width + long, y + step / 3,
                                fill="#8e1b1b" if on else colour, width=0, tags=("bar", tag))
        canvas.create_text(label_width + long + 6, y, text=str(count), anchor="w",
                           font=("TkDefaultFont", 8), fill="#444", tags=("bar", tag))
        # invisible strip so the whole row answers the mouse, not just the bar
        canvas.create_rectangle(0, y - step / 2, width, y + step / 2, outline="", fill="",
                                tags=("bar", tag))


def tag_index(canvas, x, y, prefix):
    """Index of the "<prefix><i>" item nearest to (x, y), or None.

    By point, not by Tk's "current" item: the same lookup then works for a real
    mouse and for a generated event, which is what the tests use.
    """
    for item in canvas.find_closest(x, y):
        for tag in canvas.gettags(item):
            if tag.startswith(prefix) and tag[len(prefix):].isdigit():
                return int(tag[len(prefix):])
    return None


def draw_tip(canvas, x, y, lines):
    """Small box next to the cursor with the exact numbers of the bar under it."""
    canvas.delete("tip")
    pad, line_height = 6, 14
    width = max(len(line) for line in lines) * 6.4 + pad * 2
    height = len(lines) * line_height + pad * 2
    left = min(x + 12, max(canvas.winfo_width(), 200) - width - 4)
    top = max(y - height - 10, 2)
    canvas.create_rectangle(left, top, left + width, top + height, fill="#333a45", outline="",
                            tags="tip")
    for i, line in enumerate(lines):
        canvas.create_text(left + pad, top + pad + i * line_height, text=line, anchor="nw",
                           fill="white", font=("TkDefaultFont", 8), tags="tip")


def draw_chips(frame, chosen, on_remove):
    """Redraw the picked lines as chips, each with an x that takes it out again.

    Module level so a test can build it on a real Tk root and click the x.
    """
    import tkinter as tk
    from tkinter import ttk

    for child in frame.winfo_children():
        child.destroy()
    if not chosen:
        ttk.Label(frame, text="todas as linhas", foreground="#777").pack(side="left")
        return []
    made = []
    for name in chosen:
        chip = tk.Frame(frame, background="#dce6f5", padx=6, pady=1)
        chip.pack(side="left", padx=(0, 4))
        tk.Label(chip, text=name, background="#dce6f5").pack(side="left")
        close = tk.Button(chip, text="✕", background="#dce6f5", activebackground="#c3d4ee",
                          relief="flat", borderwidth=0, highlightthickness=0, padx=3, pady=0,
                          cursor="hand2", foreground="#345", command=lambda n=name: on_remove(n))
        close.pack(side="left")
        made.append(close)
    return made


def main(path=None):
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title(f"Resumo de ocorrências — build {BUILD}")
    # a readonly combobox paints its whole field with the selection colour once it
    # has focus: white it out, or every picker looks like a blue block
    style = ttk.Style()
    style.map("TCombobox", fieldbackground=[("readonly", "white")],
              selectbackground=[("readonly", "white")], selectforeground=[("readonly", "black")])
    root.geometry("1040x700")
    book = {"path": "", "headers": [], "rows": [], "date": "", "hour": "", "text": "",
            "whole": None, "lines": [], "chosen": []}

    source, detected = tk.StringVar(value="Nenhuma planilha aberta"), tk.StringVar()
    start_date, end_date, hour, keyword = (tk.StringVar() for _ in range(4))
    recent = tk.StringVar()
    status = tk.StringVar(value="Abra uma planilha, ou escolha uma do histórico.")
    summary = tk.StringVar(value="—")

    top = ttk.Frame(root, padding=10)
    top.pack(fill="x")

    ttk.Button(top, text="Abrir...", command=lambda: open_file()).grid(row=0, column=0, sticky="w")
    ttk.Button(top, text="+ Adicionar", command=lambda: open_file(add=True)).grid(
        row=0, column=5, sticky="e")
    ttk.Label(top, textvariable=source, font=("TkDefaultFont", 10, "bold")).grid(
        row=0, column=1, columnspan=5, sticky="w", padx=8)

    ttk.Label(top, text="Histórico").grid(row=1, column=0, sticky="w", pady=(8, 0))
    recent_box = ttk.Combobox(top, textvariable=recent, state="readonly", width=70)
    recent_box.grid(row=1, column=1, columnspan=4, sticky="we", pady=(8, 0), padx=(0, 8))
    reopen_bar = ttk.Frame(top)
    reopen_bar.grid(row=1, column=5, pady=(8, 0), sticky="e")
    ttk.Button(reopen_bar, text="Reabrir", command=lambda: reopen()).pack(side="left")
    ttk.Button(reopen_bar, text="✕", width=3, command=lambda: drop_history()).pack(side="left",
                                                                                  padx=(4, 0))

    # what the app worked out on its own: without this the filters look like magic
    ttk.Label(top, textvariable=detected, foreground="#3a6ea5").grid(
        row=2, column=0, columnspan=6, sticky="w", pady=(6, 0))

    ttk.Label(top, text="Data inicial (dd/mm/aaaa)").grid(row=3, column=0, sticky="w", pady=(10, 0))
    ttk.Label(top, text="Data final").grid(row=3, column=1, sticky="w", pady=(10, 0))
    ttk.Label(top, text="Hora").grid(row=3, column=2, sticky="w", pady=(10, 0))

    first = ttk.Entry(top, textvariable=start_date, width=16, state="disabled")
    last = ttk.Entry(top, textvariable=end_date, width=16, state="disabled")
    hour_box = ttk.Combobox(top, textvariable=hour, width=16, state="disabled")

    first.grid(row=4, column=0, sticky="w", padx=(0, 8))
    last.grid(row=4, column=1, sticky="w", padx=(0, 8))
    hour_box.grid(row=4, column=2, sticky="w", padx=(0, 8))

    export_button = ttk.Button(top, text="Exportar Excel filtrado",
                               command=lambda: export_file(), state="disabled")
    export_button.grid(row=4, column=5, sticky="e")

    ttk.Label(top, text="Linhas (nenhuma = todas)").grid(row=3, column=4, sticky="w", pady=(10, 0))
    line_choice = tk.StringVar()
    line_box = ttk.Combobox(top, textvariable=line_choice, state="disabled", width=18)
    line_box.grid(row=4, column=4, sticky="w", padx=(0, 8))
    chips = ttk.Frame(top)
    chips.grid(row=5, column=4, sticky="w", pady=(4, 0))

    ttk.Label(top, text="Palavra-chave").grid(row=5, column=0, sticky="w", pady=(10, 0))
    words = ttk.Entry(top, textvariable=keyword, state="disabled")
    words.grid(row=6, column=0, columnspan=5, sticky="we")
    filter_button = ttk.Button(top, text="Filtrar", command=lambda: refresh(), state="disabled")
    filter_button.grid(row=6, column=5, padx=(8, 0))
    top.columnconfigure(4, weight=1)

    headline = ttk.Label(root, textvariable=summary, padding=(12, 6),
                         font=("TkDefaultFont", 10, "bold"))
    headline.pack(fill="x")
    headline.bind("<Configure>", lambda e: headline.configure(wraplength=e.width - 24))

    tabs = ttk.Notebook(root)
    tabs.pack(fill="both", expand=True, padx=8, pady=(0, 4))

    frame = ttk.Frame(tabs, padding=(6, 6, 6, 6))
    tabs.add(frame, text="  Ocorrências  ")
    listing = tk.Text(frame, wrap="word", padx=10, pady=8, spacing3=4, state="disabled")
    bar = ttk.Scrollbar(frame, orient="vertical", command=listing.yview)
    listing.configure(yscrollcommand=bar.set)
    listing.pack(side="left", fill="both", expand=True)
    bar.pack(side="right", fill="y")
    listing.tag_config("when", font=("TkDefaultFont", 10, "bold"), foreground="#1a4d80")
    listing.tag_config("none", foreground="#777")
    listing.tag_config("hit", background="#ffe680")
    listing.tag_config("line", font=("TkDefaultFont", 11, "bold"), foreground="#123",
                       spacing1=8, spacing3=6)

    board = ttk.Frame(tabs, padding=(10, 8, 10, 8))
    tabs.add(board, text="  Produção  ")
    tiles = ttk.Frame(board)
    tiles.pack(fill="x")
    tile_values = {}
    for i, (key, caption) in enumerate((("plan", "Planejado"), ("produzido", "Produzido"),
                                        ("faltou", "Faltou"), ("atingido", "Atingido"),
                                        ("ocorrencias", "Linhas de ocorrência"))):
        cell = tk.Frame(tiles, background="#eef1f6", padx=14, pady=8)
        cell.grid(row=0, column=i, sticky="we", padx=(0, 6))
        tiles.columnconfigure(i, weight=1)
        tile_values[key] = tk.Label(cell, text="—", background="#eef1f6",
                                    font=("TkDefaultFont", 14, "bold"))
        tile_values[key].pack(anchor="w")
        tk.Label(cell, text=caption, background="#eef1f6", foreground="#666").pack(anchor="w")

    board_title = tk.StringVar(value="Produzido por dia   (linha azul: planejado ·"
                                     " verde bateu o plano · vermelho ficou abaixo)")
    ttk.Label(board, textvariable=board_title, foreground="#444").pack(anchor="w", pady=(12, 2))
    day_chart = tk.Canvas(board, height=170, background="white", highlightthickness=1,
                          highlightbackground="#dfe3ea")
    day_chart.pack(fill="x")

    swap_title = tk.StringVar(value="Tempo de troca (changeover) por dia")
    ttk.Label(board, textvariable=swap_title, foreground="#444").pack(anchor="w", pady=(10, 2))
    swap_chart = tk.Canvas(board, height=96, background="white", highlightthickness=1,
                           highlightbackground="#dfe3ea")
    swap_chart.pack(fill="x")

    middle = ttk.Frame(board)
    middle.pack(fill="both", expand=True, pady=(12, 0))
    left_side = ttk.Frame(middle)
    left_side.pack(side="left", fill="both", expand=True)
    ttk.Label(left_side, text="Ocorrências por equipamento no período",
              foreground="#444").pack(anchor="w", pady=(0, 2))
    rank_chart = tk.Canvas(left_side, background="white", highlightthickness=1,
                           highlightbackground="#dfe3ea")
    rank_chart.pack(fill="both", expand=True)

    right_side = ttk.Frame(middle, width=430)
    right_side.pack(side="right", fill="y", padx=(12, 0))
    right_side.pack_propagate(False)
    sku_title = tk.StringVar(value="SKU e changeover por dia")
    ttk.Label(right_side, textvariable=sku_title, foreground="#444").pack(anchor="w", pady=(0, 2))
    sku_list = tk.Text(right_side, wrap="word", width=50, padx=8, pady=6, state="disabled",
                       background="white", highlightthickness=1, highlightbackground="#dfe3ea")
    sku_list.pack(fill="both", expand=True)
    sku_list.tag_config("day", font=("TkDefaultFont", 10, "bold"), spacing1=4)
    sku_list.tag_config("code", font=("TkDefaultFont", 9, "bold"), foreground="#1a4d80")
    sku_list.tag_config("swap", foreground="#8a5a00")
    sku_list.tag_config("quiet", foreground="#888")

    status_line = ttk.Label(root, textvariable=status, padding=(12, 4), foreground="#555")
    status_line.pack(fill="x")

    def say(message, tone="#555"):
        status.set(message)
        status_line.configure(foreground=tone)

    def show_history(entries=None):
        book["history"] = entries if entries is not None else history()
        recent_box["values"] = [
            "{}  ·  {}  ·  {} linhas{}".format(
                Path(e["path"]).name, e["when"][:16].replace("T", " "), e.get("rows", "?"),
                "  ·  {:02d}/{}".format(e["month"][1], e["month"][0]) if e.get("month") else "")
            for e in book["history"]]

    def reopen():
        picked = recent_box.current()
        if picked >= 0:
            entry = book["history"][picked]
            month = tuple(entry["month"]) if entry.get("month") else None
            open_file(entry["path"], month)

    def drop_history():
        """The x next to Reabrir: takes the chosen file out of the history."""
        picked = recent_box.current()
        if picked < 0:
            show_history(forget_all())
            say("Histórico apagado.")
            return
        gone = book["history"][picked]
        show_history(forget(gone["path"]))
        recent.set("")
        say(f"{Path(gone['path']).name} saiu do histórico.")

    def open_file(chosen=None, month=None, add=False):
        if book.get("reading"):
            say("Espere a leitura em andamento terminar.", "#b00020")
            return
        if add and not book["rows"]:
            say("Abra o primeiro arquivo antes de adicionar outro.", "#b00020")
            return
        if not chosen:
            opened = Path(book["path"]).parent if book["path"] else start_dir()
            chosen = filedialog.askopenfilename(
                filetypes=[("Excel ou pacote", "*.xlsx *.zip"), ("Excel", "*.xlsx"),
                           ("Pacote de relatórios", "*.zip")], initialdir=str(opened))
        if not chosen:
            return
        if book["rows"] and not add:
            # both buttons end up here; asking removes any doubt about which one was hit
            somar = messagebox.askyesnocancel(
                "Já há dados abertos",
                "Somar {} ao que já está aberto?\n\n"
                "Sim = somar     Não = substituir     Cancelar = deixar como está".format(
                    Path(chosen).name))
            if somar is None:
                say("Cancelado — nada mudou.")
                return
            add = somar
        if not Path(chosen).exists():
            say(f"{chosen} não está mais nesse caminho.", "#b00020")
            return
        source.set(f"Lendo {Path(chosen).name}...")
        say("Lendo em segundo plano — a janela continua respondendo.")
        book["reading"], book["adding"] = True, add
        for widget in (first, last, words, filter_button, export_button):
            widget.state(["disabled"])
        news = queue.Queue()

        def work():  # off the UI thread: Tk is touched only by the poller below
            try:
                news.put(("ok", load(chosen, month,
                                     progress=lambda i, n, name: news.put(("step", (i, n, name))))))
            except Exception as error:
                news.put(("erro", error))

        threading.Thread(target=work, daemon=True).start()

        def poll():
            try:
                kind, payload = news.get_nowait()
            except queue.Empty:
                root.after(100, poll)
                return
            if kind == "step":
                done, total, name = payload
                source.set(f"Lendo {done}/{total}: {name}")
                root.after(50, poll)
                return
            book["reading"] = False
            for widget in (first, last, words, filter_button, export_button):
                widget.state(["!disabled"])
            if kind == "erro":
                source.set("Nenhuma planilha aberta")
                say(f"Erro ao ler: {payload}", "#b00020")
                return
            finish(payload)

        root.after(100, poll)

    def finish(data):
        chosen = data["path"] if isinstance(data, dict) and data.get("path") else book["path"]
        add = book.get("adding", False)
        if add:
            try:
                data = merge([book["data"], data])
            except ValueError as error:
                say(f"Erro ao juntar: {error}", "#b00020")
                return
        book["just_loaded"] = ("Somado ao que já estava aberto: {} arquivos.".format(
            len(data.get("sources", []))) if add else "Abriu do zero — o que estava aberto"
            " foi substituído. Use + Adicionar para somar.")
        book["data"] = data
        book.update(path=data.get("path", chosen), headers=data["headers"],
                    rows=data["rows"], kind=data["kind"])
        picked = data["columns"]
        book["date"], book["hour"] = picked["date"], picked["hour"]
        book["label"], book["text"] = picked["label"], picked["text"]
        book["line"] = picked.get("line", "")
        book["month"] = data["month"]
        source.set("{}   ·   {} linhas".format(
            " + ".join(Path(p).name for p in data["sources"]) if data.get("sources")
            else Path(chosen).name, len(data["rows"])))
        if data.get("months"):  # zip: several workbooks, each with its own month
            lidos = [f"{when[1]:02d}/{when[0]}" for _, when, _ in data["months"] if when and all(when)]
            perdidos = [name for name, when, _ in data["months"] if not (when and all(when))]
            origem = "{} planilhas · {}".format(len(data["months"]), ", ".join(sorted(set(lidos))))
            if perdidos:
                origem += f"   ·   SEM mês: {', '.join(perdidos)}"
        elif data["kind"] == "turno":
            day_at = data["headers"].index("Dia")
            origem = "lido das {} abas por dia".format(len({r[day_at] for _, r in data["rows"]}))
        else:
            origem = f"aba {data['sheet']}"
        vindo = data.get("month_from")
        if vindo:
            sobre_o_mes = f"   ·   mês veio da aba {vindo}"
        elif data.get("month_guessed"):
            sobre_o_mes = "   ·   NENHUMA aba diz o mês — confira o campo"
        else:
            sobre_o_mes = ""
        detected.set("{}   ·   Data: {}   ·   Hora: {}   ·   Ocorrência: {}{}".format(
            origem, book["date"] or "não encontrada", book["label"] or "não encontrada",
            book["text"] or "não encontrada", sobre_o_mes))
        book["whole"] = summarize(book["headers"], book["rows"], date_col=book["date"],
                                  text_col=book["text"])
        book["whole"]["with_text"] = summarize(book["headers"], book["rows"],
                                               text_col=book["text"], only_with_text=True)["found"]
        # "Horas" holds every hour an occurrence covers; the picker offers them one by one
        found_hours = sorted({h.strip() for value in
                              (distinct(book["headers"], book["rows"], book["hour"]) if book["hour"] else [])
                              for h in value.split("|")})
        hour_box["values"] = [""] + [h for h in found_hours if HOUR_RE.fullmatch(h)]
        book["lines"] = (distinct(book["headers"], book["rows"], picked["line"])
                         if picked.get("line") else [])
        book["chosen"] = []
        line_box.configure(state="readonly" if book["lines"] else "disabled")
        redraw()
        for field in (start_date, end_date, hour, keyword):
            field.set("")
        for widget in (first, last, words, filter_button, export_button):
            widget.state(["!disabled"])
        hour_box.configure(state="readonly")
        if not add:
            show_history(remember(chosen, {"rows": len(data["rows"]), "month": data["month"]}))
        refresh()

    def as_date(text):
        if not text.strip():
            return None
        day = read_typed_date(text, book.get("month"))
        if day is None and book.get("month") and re.fullmatch(r"\d{1,2}", text.strip()):
            raise ValueError("O dia {} não existe em {:02d}/{}.".format(
                text.strip(), book["month"][1], book["month"][0]))
        if day is None:
            raise ValueError(f"Data {text.strip()!r} não dá para ler — use dd/mm/aaaa"
                             " ou só o dia.")
        return day

    def typing_date(field):
        """Put the slashes in while the user types, keeping the cursor at the end."""
        def on_key(event):
            if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Tab", "Return"):
                return
            masked = mask_date(event.widget.get())
            if masked != event.widget.get():
                event.widget.delete(0, "end")
                event.widget.insert(0, masked)
            event.widget.icursor("end")
        field.bind("<KeyRelease>", on_key)

    def terms():
        return [t.strip() for t in keyword.get().split(",") if t.strip()]

    def redraw():
        draw_chips(chips, book["chosen"], drop_line)
        line_box["values"] = [n for n in book["lines"] if n not in book["chosen"]]
        line_choice.set("")

    def add_line(event=None):
        if event:
            event.widget.selection_clear()
        name = line_choice.get().strip()
        if name and name not in book["chosen"]:
            book["chosen"].append(name)
            redraw()
            refresh()

    def drop_line(name):
        book["chosen"].remove(name)
        redraw()
        refresh()

    def chosen_lines():
        return list(book["chosen"])

    def filters():
        return dict(text_col=book["text"], terms=terms(), date_col=book["date"],
                    start=as_date(start_date.get()), end=as_date(end_date.get()),
                    hour_col=book["hour"], hour=hour.get().strip(),
                    line_col=book.get("line", ""), lines=chosen_lines())

    def refresh(event=None):
        if event is not None and isinstance(getattr(event, "widget", None), ttk.Combobox):
            event.widget.selection_clear()
        if not book["rows"]:
            return
        try:
            chosen = filters()
            result = summarize(book["headers"], book["rows"], only_with_text=True, **chosen)
            # production exists in every hour, even the ones with nothing written down
            hours_result = summarize(book["headers"], book["rows"], **chosen)
        except Exception as error:
            say(f"Erro: {error}", "#b00020")
            return
        columns = result["columns"]
        when = [columns.index(c) for c in (book["date"], book["label"]) if c in columns]
        body = columns.index(book["text"]) if book["text"] in columns else None
        marks = [t.casefold() for t in terms()]

        line_at = columns.index(book["line"]) if book.get("line") in columns else None
        listing.configure(state="normal")
        listing.delete("1.0", "end")
        current = None
        for row in result["rows"]:
            # one block per production line, in the order the lines appear
            if line_at is not None and row[line_at] != current:
                current = row[line_at]
                listing.insert("end", f"\nLINHA {show(current)}\n", "line")
            listing.insert("end", " · ".join(show(row[i]) for i in when) + "\n", "when")
            text = show(row[body]).strip() if body is not None else ""
            listing.insert("end", sentences_with(text, marks) + "\n\n")
        if not result["rows"]:
            listing.insert("end", "Nenhuma ocorrência com esses filtros.\n", "none")
        for term in terms():
            at = "1.0"
            while at := listing.search(term, at, nocase=True, stopindex="end"):
                listing.tag_add("hit", at, at := f"{at}+{len(term)}c")
        listing.configure(state="disabled")

        draw_board(hours_result)

        whole = book["whole"]
        totals = "   ".join(
            "{}: {}{} de {}".format(name, thousands(value), "+" if name in result["partial"] else "",
                                    thousands(whole["totals"].get(name, 0)))
            for name, value in result["totals"].items())
        period = " · {} a {}".format(*map(show, result["period"])) if result["period"] else ""
        summary.set(f"{result['found']} ocorrências de {whole['found']} linhas da planilha"
                    f" ({whole['with_text']} têm ocorrência escrita){period}\n{totals}")
        say(book.pop("just_loaded", None)
            or "Com palavra-chave, só a frase que a contém aparece. Enter refiltra.")

    def draw_board(result):
        """Second tab: the same filter, seen as production numbers."""
        numbers = production_summary(book["headers"], result["matched"])
        book["board"] = numbers
        about = ", ".join(terms())
        dias = len(numbers["dias"])
        quando = ("Produzido por hora em " + show(numbers["dias"][0][0]) if dias == 1
                  else "Produzido por hora" if len(numbers["slots"]) <= 120
                  else "Produzido por dia" if dias <= 31 else "Produzido por mês")
        board_title.set(quando + ("   ·   clique numa barra para abrir os dias"
                                  if dias > 1 else "")
                        + (f" — só horas com {about}" if about else "")
                        + "   (linha azul: planejado · verde bateu o plano · vermelho ficou abaixo)")
        sku_title.set(f"SKU e changeover por dia — {about}" if about
                      else "SKU e changeover por dia")
        tile_values["plan"].configure(text=thousands(numbers["plan"]))
        tile_values["produzido"].configure(text=thousands(numbers["produzido"]))
        tile_values["faltou"].configure(text=thousands(numbers["faltou"]))
        tile_values["atingido"].configure(
            text="—" if numbers["atingido"] is None else f"{numbers['atingido'] * 100:.1f}%")
        tile_values["ocorrencias"].configure(text=str(sum(n for _, n in numbers["equipamentos"])))
        # one day -> per hour; a short period -> per hour across the days; longer -> per day
        if len(numbers["dias"]) == 1:
            book["series"] = numbers["horas"]
        elif len(numbers["slots"]) <= 120:
            book["series"] = numbers["slots"]
        elif len(numbers["dias"]) <= 31:
            book["series"] = numbers["dias"]
        else:  # more than a month of bars is a picket fence: one bar per month
            book["series"] = numbers["meses"]
        book["monthly"] = book["series"] is numbers["meses"]
        draw_day_chart(day_chart, book["series"])
        por_mes = len(numbers["dias"]) > 31
        book["swaps"] = (_monthly(numbers["trocas_por_dia"]) if por_mes
                         else numbers["trocas_por_dia"])
        swap_title.set("Tempo de troca (changeover) por {} · total {:.0f} min".format(
            "mês" if por_mes else "dia", sum(v for _, v in numbers["trocas_por_dia"])))
        draw_time_chart(swap_chart, book["swaps"])
        draw_rank_chart(rank_chart, numbers["equipamentos"], highlight=book.get("over"))
        sku_list.configure(state="normal")
        sku_list.delete("1.0", "end")
        for day in numbers["skus_por_dia"]:
            skus, swaps = day["skus"], day["swaps"]
            head = show(day["day"]) + (f"   ·   {day['file']}" if day["file"] else "")
            sku_list.insert("end", head + "\n", "day")
            if not skus:
                sku_list.insert("end", "   sem SKU anotado\n", "quiet")
            else:
                sku_list.insert("end", "   ", "quiet")
                sku_list.insert("end", skus[0][0], "code")
                sku_list.insert("end", "   " + skus[0][1].split(" - ", 1)[-1] + "\n")
            # each changeover moves the day on to the next SKU written in the cell
            for i, (hours, said, took) in enumerate(swaps):
                when = ", ".join(hours) if hours else "hora não anotada"
                sku_list.insert("end", "   ⇄ houve um changeover {}{}\n".format(
                    when, f"  ({took})" if took else ""), "swap")
                if i + 1 < len(skus):
                    sku_list.insert("end", "      passa a ", "quiet")
                    sku_list.insert("end", skus[i + 1][0], "code")
                    sku_list.insert("end", "   " + skus[i + 1][1].split(" - ", 1)[-1] + "\n")
                elif skus:
                    sku_list.insert("end", "      a planilha não anotou o SKU seguinte\n", "quiet")
                sku_list.insert("end", f"      {said}\n")
            sku_list.insert("end", "\n")
        if not numbers["skus_por_dia"]:
            sku_list.insert("end", "Nenhum SKU no filtro.", "quiet")
        sku_list.configure(state="disabled")

    def _monthly(series):
        months = {}
        for day, value in series:
            first = day.replace(day=1)
            months[first] = months.get(first, 0) + value
        return [(month, months[month]) for month in sorted(months)]

    def swap_hover(event):
        series = book.get("swaps") or []
        slot = tag_index(swap_chart, event.x, event.y, "slot")
        if slot is None or slot >= len(series):
            swap_chart.delete("tip")
            return
        label, minutes = series[slot]
        draw_tip(swap_chart, event.x, event.y,
                 [show(label), f"Troca: {minutes:.0f} min"])

    def zoom_in(event):
        """Click a bar to open the level below: month -> its days, day -> its hours."""
        series = book.get("series") or []
        slot = tag_index(day_chart, event.x, event.y, "slot")
        if slot is None or slot >= len(series):
            return
        label = series[slot][0]
        day = label[0] if isinstance(label, tuple) else label
        if not hasattr(day, "day"):
            return  # already at hour level: nothing finer to open
        if book.get("monthly"):
            first = day.replace(day=1)
            last = first.replace(day=calendar.monthrange(first.year, first.month)[1])
        else:
            first = last = day
        start_date.set(show(first))
        end_date.set(show(last))
        refresh()
        say("Abrindo {}. Apague as datas para voltar ao período inteiro.".format(
            first.strftime("%m/%Y") if book.get("monthly") else show(first)))

    def redraw_board(*_):
        if book.get("board"):
            draw_day_chart(day_chart, book.get("series") or book["board"]["dias"])
            draw_rank_chart(rank_chart, book["board"]["equipamentos"], highlight=book.get("over"))
            draw_time_chart(swap_chart, book.get("swaps") or [])

    def slot_under(event):
        return tag_index(day_chart, event.x, event.y, "slot")

    def day_hover(event):
        series = book.get("series") or []
        slot = slot_under(event)
        if slot is None or slot >= len(series):
            day_chart.delete("tip")
            return
        label, planned, produced = series[slot]
        title = (f"{show(label[0])} · {label[1]}" if isinstance(label, tuple)
                 else show(label) if hasattr(label, "day") else str(label))
        lines = [title, f"Produzido: {thousands(produced)}"]
        if planned is None:
            lines.append("Planejado: não anotado")
        else:
            gap = produced - planned
            lines += [f"Planejado: {thousands(planned)}",
                      "{}{} ({:.0f}%)".format("+" if gap >= 0 else "", thousands(gap),
                                              produced / planned * 100 if planned else 0)]
        draw_tip(day_chart, event.x, event.y, lines)

    def machine_under(event):
        """Which machine the mouse is on."""
        return tag_index(rank_chart, event.x, event.y, "bar")

    def rank_hover(event):
        over = machine_under(event)
        if over == book.get("over"):
            return
        book["over"] = over
        rank_chart.configure(cursor="hand2" if over is not None else "")
        draw_rank_chart(rank_chart, book["board"]["equipamentos"], highlight=over)
        if over is not None:
            name = book["board"]["equipamentos"][over][0]
            say(f"Clique em {name} para ver só o que envolve esse equipamento.")

    def rank_click(event):
        """Clicking a machine takes the user to its occurrences, already filtered."""
        over = machine_under(event)
        if over is None or not book.get("board"):
            return
        name = book["board"]["equipamentos"][over][0]
        keyword.set(name)
        refresh()
        tabs.select(frame)
        listing.see("1.0")
        say(f"Ocorrências de {name}. Apague a palavra-chave para ver todas de novo.")

    def export_file():
        if not book["rows"]:
            return
        if book.get("kind") == "turno":
            say("A exportação só vale para planilha de tabela simples — esta é montada"
                " a partir das abas por dia.", "#b00020")
            return
        origin = Path(book["path"])
        try:
            target = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                initialfile=origin.stem + "_filtrado.xlsx", initialdir=str(origin.parent))
            if not target:
                return
            say("Gravando — arquivo grande leva cerca de um minuto...")
            root.update()
            visible, exact = export(book["path"], target, book["sheet"],
                                    book["headers"], book["rows"], **filters())
            warning = "" if exact else " — mais de 2 palavras: o filtro do Excel abre sem o critério"
            say(f"{visible} linhas visíveis em {target}{warning}", "#1b7a3d")
        except Exception as error:
            say(f"Erro: {error}", "#b00020")

    show_history()
    day_chart.bind("<Configure>", redraw_board)
    day_chart.bind("<Motion>", day_hover)
    day_chart.bind("<Button-1>", zoom_in)
    day_chart.bind("<Leave>", lambda _: day_chart.delete("tip"))
    rank_chart.bind("<Configure>", redraw_board)
    swap_chart.bind("<Configure>", redraw_board)
    swap_chart.bind("<Motion>", swap_hover)
    swap_chart.bind("<Leave>", lambda _: swap_chart.delete("tip"))
    rank_chart.bind("<Motion>", rank_hover)
    rank_chart.bind("<Leave>", lambda _: (book.update(over=None), redraw_board()))
    rank_chart.bind("<Button-1>", rank_click)
    typing_date(first)
    typing_date(last)
    root.bind("<Return>", refresh)
    hour_box.bind("<<ComboboxSelected>>", refresh)
    line_box.bind("<<ComboboxSelected>>", add_line)
    if path:
        root.after(50, lambda: open_file(path))
    root.mainloop()


if __name__ == "__main__":
    import sys

    main(*sys.argv[1:2])  # optional: path of a spreadsheet to open right away
