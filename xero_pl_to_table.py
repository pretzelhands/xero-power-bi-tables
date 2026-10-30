#!/usr/bin/env python3
"""Turn Xero Profit and Loss exports into one long table for Power BI.

Reads Profit and Loss reports exported from Xero (Excel .xlsx or CSV, one or
several files, e.g. one per company), drops headings, totals and calculated
rows (Gross Profit, Net Profit), and unpivots the period columns into rows:

    Organisation, Section, Account, Period, Period start, Amount

Standard library only (Python 3.8+). Usage:

    python3 xero_pl_to_table.py "Profit and Loss.xlsx" [more files...] -o pl.csv

Made by Reportpost (https://reportpost.app), which emails Xero reports on a schedule.
MIT licence.
"""
import argparse
import csv
import io
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
HEAD = ['Organisation', 'Section', 'Account', 'Period', 'Period start', 'Amount']
MONTHS = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 'jul': 7,
          'aug': 8, 'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12}
# Rows Xero calculates from other rows. They'd double count in Power BI.
CALC = re.compile(r'^((gross|net|operating|net operating) (profit|loss|income)|'
                  r'(profit|loss) (before|after) (income )?tax(ation)?|'
                  r'(net )?profit ?/? ?\(?loss\)?|(profit|loss) for the (period|year))$', re.I)
YEAR = re.compile(r'^(19|20)\d\d$')  # a column headed just '2026' (year comparisons) is a period, not an amount
NUMERIC = re.compile(r"^[\s(−-]*[$€£¥]?[\d.,\s']+\)?-?$")


class Row(list):
    """A list of cell texts plus the set of column indexes that hold a formula."""
    def __init__(self, cells=(), formulas=()):
        super().__init__(cells)
        self.formulas = set(formulas)


def col_index(ref):
    m = re.match(r'[A-Z]+', ref or '')
    if not m:
        return None
    n = 0
    for ch in m.group(0):
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(path):
    """First worksheet of an .xlsx file as a list of Rows."""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if 'xl/sharedStrings.xml' in names:
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).iter(NS + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
        sheet = 'xl/worksheets/sheet1.xml'
        if sheet not in names:
            sheets = sorted(n for n in names if re.match(r'^xl/worksheets/[^/]+\.xml$', n))
            if not sheets:
                raise ValueError('no worksheet in this file')
            sheet = sheets[0]
        root = ET.fromstring(z.read(sheet))
    rows = []
    for r in root.iter(NS + 'row'):
        row = Row()
        for c in r.iter(NS + 'c'):
            col = col_index(c.get('r'))
            if col is None:
                col = len(row)
            if col > 400:
                continue
            kind, v = c.get('t'), c.find(NS + 'v')
            if kind == 's':
                val = shared[int(v.text)] if v is not None and v.text and int(v.text) < len(shared) else ''
            elif kind == 'inlineStr':
                val = ''.join(t.text or '' for t in c.iter(NS + 't'))
            else:
                val = (v.text or '') if v is not None else ''
            if c.find(NS + 'f') is not None:
                row.formulas.add(col)
            while len(row) < col:
                row.append('')
            if len(row) == col:
                row.append(val)
            else:
                row[col] = val
        rows.append(row)
    return rows


def read_csv(path):
    with open(path, encoding='utf-8-sig', newline='') as f:
        text = f.read()
    head = '\n'.join(text.split('\n')[:10])
    sep = max(['\t', ';', ','], key=lambda c: len(head.split(c)))
    return [Row(r) for r in csv.reader(io.StringIO(text), delimiter=sep)]


def num(s, comma_decimals=False):
    """Parse an amount as Xero or Excel writes it: 1,234.50 / (1,234.50) / -1.234,50 / '-'. NaN if not a number.
    comma_decimals: the file writes decimals after a comma, so a lone '1.234' is one thousand two hundred and thirty-four."""
    s = str(s or '').strip().replace('−', '-')
    if not s:
        return float('nan')
    if s == '-':
        return 0.0
    if comma_decimals and re.fullmatch(r'-?\d{1,3}(\.\d{3})+', s):
        return float(s.replace('.', ''))
    if re.fullmatch(r'-?\d+(\.\d+)?(e[-+]?\d+)?', s, re.I):
        return float(s)
    if not re.search(r'\d', s):
        return float('nan')
    neg = bool(re.fullmatch(r'\(.*\)', s)) or '-' in re.sub(r'\d.*$', '', s) or s.endswith('-')
    s = re.sub(r'[^0-9.,]', '', s)
    if comma_decimals and re.fullmatch(r'\d{1,3}(\.\d{3})+', s):
        s = s.replace('.', '')
    elif s.rfind(',') > s.rfind('.') and re.search(r',\d{1,2}$', s):
        s = s.replace('.', '').replace(',', '.')
    else:
        s = s.replace(',', '')
    try:
        v = float(s)
    except ValueError:
        return float('nan')
    return -v if neg else v


def is_num(s):
    return num(s) == num(s)  # NaN != NaN


def period_start(label):
    """'Oct 2026' or 'October 2026' -> '2026-10-01'; anything else (tracking options, date ranges) -> ''."""
    m = re.fullmatch(r'([A-Za-z]{3,9})\.?\s+(\d{4})', label)
    if not m:
        return ''
    word = m.group(1).lower()
    mo = MONTHS.get(word[:4]) or MONTHS.get(word[:3])
    return '%s-%02d-01' % (m.group(2), mo) if mo else ''


def summary_column(label):
    """Columns that add up or compare the others (Total, YTD, Variance, %): left out so nothing counts twice."""
    return bool(re.fullmatch(r'year to date|ytd|total', label, re.I) or re.search(r'variance|change|%', label, re.I))


def cell(row, i):
    return str(row[i]).strip() if i < len(row) and row[i] is not None else ''


def analyse(rows, fallback_org=''):
    """Returns (records, info) or raises ValueError with a plain explanation."""
    head = -1
    for r in range(min(len(rows), 40)):
        if not any(cell(rows[r], c) and (not NUMERIC.match(cell(rows[r], c)) or YEAR.match(cell(rows[r], c)))
                   for c in range(1, len(rows[r]))):
            continue
        if any(is_num(cell(rows[b], d)) for b in range(r + 1, min(len(rows), r + 15)) for d in range(1, len(rows[b]))):
            head = r
            break
    if head < 0:
        raise ValueError("couldn't find the heading row with the periods (Account, Oct 2026, ...). "
                         'Export the Profit and Loss from Xero as Excel and try that file.')
    cols, skipped = [], []
    for k in range(1, len(rows[head])):
        label = re.sub(r'\s+', ' ', cell(rows[head], k))
        if not label:
            continue
        if summary_column(label):
            skipped.append(label)
            continue
        cols.append((k, label, period_start(label)))
    if not cols:
        raise ValueError('this file only has total columns. Run the Profit and Loss by month '
                         '(or compare periods) in Xero and export it again.')
    # Xero exports start with the report title, the organisation and the dates in A1:A3.
    org, title = fallback_org, ''
    if head >= 2 and cell(rows[0], 0) and cell(rows[1], 0):
        title, org = cell(rows[0], 0), cell(rows[1], 0)
    # Decimals after a comma (1.234,50) anywhere and never after a point: then '1.234' means 1234 too.
    values = [cell(row, c) for row in rows[head + 1:] for c, _, _ in cols]
    comma_decimals = (any(re.search(r'\d,\d{1,2}\)?-?$', v) for v in values)
                      and not any(re.search(r'\d\.\d{1,2}\)?-?$', v) for v in values))
    # Names that a later 'Total ...' row closes, to tell a heading from an account with no amounts.
    closed = {re.sub(r'^total\s+', '', cell(row, 0).strip(), flags=re.I).lower()
              for row in rows[head + 1:] if re.match(r'total\s', cell(row, 0), re.I)}

    def base(x):
        """A heading's name without Xero's 'Less ' / 'Plus ' prefix, as its 'Total ...' row writes it."""
        return re.sub(r'^(less|plus):?\s+', '', x, flags=re.I).lower()
    out, stack, accounts = [], [], set()
    after_account = False
    for row in rows[head + 1:]:
        name = re.sub(r'[\t\r\n]+', ' ', cell(row, 0)).strip()
        if not name:
            continue
        if not any(cell(row, c) for c, _, _ in cols):
            # A heading such as 'Operating Expenses'. Right after an account row with no 'Total <name>' below,
            # it's an account with no amounts in these periods instead: skip it, keep the section.
            # (Files without total rows: always a heading.)
            if after_account and closed and name.lower() not in closed and base(name) not in closed:
                continue
            stack.append(name)
            after_account = False
            continue
        after_account = False
        if name.lower() == 'total':
            continue
        # 'Total <heading>' closes that heading (and anything nested in it). An account called 'Total ...' stays.
        # e.g. 'Total Cost of Sales' closes 'Less Cost of Sales'
        if re.match(r'total\s', name, re.I):
            target = re.sub(r'^total\s+', '', name, flags=re.I).lower()
            lower = [base(x) for x in stack]
            if target in lower:
                del stack[len(lower) - 1 - lower[::-1].index(target):]
                continue
        if CALC.match(name):
            continue
        if any(c in getattr(row, 'formulas', ()) for c, _, _ in cols):
            continue
        section = stack[-1] if stack else ''
        after_account = True
        accounts.add((section, name))
        for c, label, start in cols:
            v = num(cell(row, c), comma_decimals)
            if v == v:
                out.append([org, section, name, label, start, round(v, 4)])
    if not out:
        raise ValueError('found the periods but no account rows below them. Is this a Profit and Loss export?')
    return out, {'org': org, 'title': title, 'accounts': len(accounts), 'cols': cols, 'skipped': skipped}


def csv_cell(v):
    """Text that Excel would run as a formula gets a leading apostrophe."""
    if isinstance(v, float):
        return ('%.4f' % v).rstrip('0').rstrip('.') if v != int(v) else str(int(v))
    v = str(v)
    if re.match(r'^[=+\-@\t\r]', v):
        try:
            float(v)
        except ValueError:
            v = "'" + v
    return v


def main(argv=None):
    p = argparse.ArgumentParser(description='Turn Xero Profit and Loss exports (xlsx/csv) into one long CSV table for Power BI.')
    p.add_argument('files', nargs='+', help='Profit and Loss exports from Xero (.xlsx or .csv)')
    p.add_argument('-o', '--output', help='CSV file to write (default: print to the terminal)')
    p.add_argument('--keep-zero', action='store_true', help='keep rows where the amount is 0')
    a = p.parse_args(argv)

    records, seen, failed = [], set(), 0
    for path in a.files:
        name = os.path.basename(path)
        try:
            if name.lower().endswith('.xls'):
                raise ValueError('old .xls files are not supported. Export again from Xero as Excel (.xlsx).')
            rows = read_xlsx(path) if name.lower().endswith('.xlsx') else read_csv(path)
            fallback = re.sub(r'\s*-\s*Profit and Loss$', '', os.path.splitext(name)[0].replace('_', ' '), flags=re.I).strip()
            out, info = analyse(rows, fallback)
        except zipfile.BadZipFile:
            print('%s: not an .xlsx file. Export the Profit and Loss from Xero as Excel and try again.' % name, file=sys.stderr)
            failed += 1
            continue
        except (ValueError, OSError, ET.ParseError, csv.Error, UnicodeDecodeError) as e:
            print('%s: %s' % (name, e), file=sys.stderr)
            failed += 1
            continue
        # The same company and period in two files (a single month plus a by-month export) would count twice: keep the first.
        dup = sorted({r[3] for r in out if (r[0], r[3]) in seen})
        mine = [r for r in out if (r[0], r[3]) not in seen]
        seen.update((r[0], r[3]) for r in mine)
        records += mine
        n, dated = len(info['cols']), sum(1 for c in info['cols'] if c[2])
        msg = '%s: %d accounts x %d columns' % (info['org'] or name, info['accounts'], n)
        if info['title'] and not re.search(r'profit|loss', info['title'], re.I):
            msg += " (title is '%s', not a Profit and Loss, so check the result)" % info['title']
        if dated < n:
            msg += '. %d without a month (e.g. tracking options), so Period start is blank there' % (n - dated)
        if info['skipped']:
            msg += '. Left out %s (totals of the other columns)' % ', '.join(info['skipped'])
        if dup:
            msg += '. Skipped %s, already read from an earlier file' % ', '.join(dup)
        print(msg, file=sys.stderr)

    if not a.keep_zero:
        records = [r for r in records if r[5] != 0]
    if not records:
        print('No rows to write.', file=sys.stderr)
        return 1
    f = open(a.output, 'w', encoding='utf-8-sig', newline='') if a.output else sys.stdout
    try:
        w = csv.writer(f, lineterminator='\r\n' if a.output else '\n')
        w.writerow(HEAD)
        for r in records:
            w.writerow([csv_cell(v) for v in r])
    finally:
        if a.output:
            f.close()
    if a.output:
        print('Wrote %d rows to %s' % (len(records), a.output), file=sys.stderr)
    return 1 if failed and failed == len(a.files) else 0


if __name__ == '__main__':
    sys.exit(main())
