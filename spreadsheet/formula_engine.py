"""
Formula Engine — converts Excel formulas to Python and recalculates cell dependencies.

Supported:
  Arithmetic, SUM, AVERAGE, MIN, MAX, IF, IFERROR, ROUND, ABS,
  COUNT, COUNTA, SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF,
  VLOOKUP, HLOOKUP, INDEX, MATCH, AND, OR, NOT, INT, MOD,
  CONCATENATE, TEXT, LEFT, RIGHT, MID, TRIM, UPPER, LOWER,
  LEN, FIND, SEARCH, SUBSTITUTE, REPLACE, VALUE, ISBLANK,
  ISERROR, ISNUMBER, ISTEXT, SQRT, POWER, LOG, EXP, PI,
  TODAY, NOW, YEAR, MONTH, DAY, DATE, NETWORKDAYS
"""

import re
import math
from datetime import date, datetime


def col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def excel_ref_to_rc(ref: str):
    m = re.match(r'^\$?([A-Z]+)\$?(\d+)$', ref.upper())
    if not m:
        return None, None
    col = col_letter_to_index(m.group(1))
    row = int(m.group(2)) - 1
    return row, col


def col_index_to_letter(index: int) -> str:
    result = ''
    n = index
    while True:
        result = chr(65 + n % 26) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result


def formula_to_python(formula: str) -> str:
    expr = formula.strip()
    if expr.startswith('='):
        expr = expr[1:]

    func_map = {
        'SUMIFS':       '_SUMIFS',
        'SUMIF':        '_SUMIF',
        'COUNTIFS':     '_COUNTIFS',
        'COUNTIF':      '_COUNTIF',
        'AVERAGEIF':    '_AVERAGEIF',
        'AVERAGEIFS':   '_AVERAGEIFS',
        'VLOOKUP':      '_VLOOKUP',
        'HLOOKUP':      '_HLOOKUP',
        'INDEX':        '_INDEX',
        'MATCH':        '_MATCH',
        'SUM':          '_SUM',
        'AVERAGE':      '_AVERAGE',
        'MIN':          '_MIN',
        'MAX':          '_MAX',
        'IF':           '_IF',
        'IFS':          '_IFS',
        'IFERROR':      '_IFERROR',
        'IFNA':         '_IFNA',
        'ROUND':        '_ROUND',
        'ROUNDUP':      '_ROUNDUP',
        'ROUNDDOWN':    '_ROUNDDOWN',
        'CEILING':      '_CEILING',
        'FLOOR':        '_FLOOR',
        'ABS':          'abs',
        'COUNT':        '_COUNT',
        'COUNTA':       '_COUNTA',
        'COUNTBLANK':   '_COUNTBLANK',
        'AND':          '_AND',
        'OR':           '_OR',
        'NOT':          '_NOT',
        'XOR':          '_XOR',
        'INT':          '_INT',
        'TRUNC':        '_TRUNC',
        'MOD':          '_MOD',
        'CONCATENATE':  '_CONCATENATE',
        'CONCAT':       '_CONCATENATE',
        'TEXTJOIN':     '_TEXTJOIN',
        'TEXT':         '_TEXT',
        'VALUE':        '_VALUE',
        'LEFT':         '_LEFT',
        'RIGHT':        '_RIGHT',
        'MID':          '_MID',
        'TRIM':         '_TRIM',
        'UPPER':        '_UPPER',
        'LOWER':        '_LOWER',
        'PROPER':       '_PROPER',
        'LEN':          'len',
        'FIND':         '_FIND',
        'SEARCH':       '_SEARCH',
        'SUBSTITUTE':   '_SUBSTITUTE',
        'REPLACE':      '_REPLACE',
        'REPT':         '_REPT',
        'ISBLANK':      '_ISBLANK',
        'ISERROR':      '_ISERROR',
        'ISNUMBER':     '_ISNUMBER',
        'ISTEXT':       '_ISTEXT',
        'ISNA':         '_ISNA',
        'SQRT':         'math.sqrt',
        'POWER':        '_POWER',
        'LOG':          '_LOG',
        'LOG10':        '_LOG10',
        'LN':           '_LN',
        'EXP':          'math.exp',
        'PI':           '_PI',
        'RAND':         '_RAND',
        'RANDBETWEEN':  '_RANDBETWEEN',
        'TODAY':        '_TODAY',
        'NOW':          '_NOW',
        'YEAR':         '_YEAR',
        'MONTH':        '_MONTH',
        'DAY':          '_DAY',
        'DATE':         '_DATE',
        'DAYS':         '_DAYS',
        'NETWORKDAYS':  '_NETWORKDAYS',
        'EOMONTH':      '_EOMONTH',
        'LARGE':        '_LARGE',
        'SMALL':        '_SMALL',
        'RANK':         '_RANK',
        'PERCENTILE':   '_PERCENTILE',
        'STDEV':        '_STDEV',
        'VAR':          '_VAR',
        'MEDIAN':       '_MEDIAN',
        'NPV':          '_NPV',
        'PMT':          '_PMT',
        'FV':           '_FV',
        'PV':           '_PV',
        'RATE':         '_RATE',
        'NPER':         '_NPER',
        'CHOOSE':       '_CHOOSE',
        'OFFSET':       '_OFFSET',
        'INDIRECT':     '_INDIRECT',
        'ROW':          '_ROW',
        'COLUMN':       '_COLUMN',
        'ROWS':         '_ROWS',
        'COLUMNS':      '_COLUMNS',
    }

    # Sort by length descending to avoid partial replacements (SUMIFS before SUMIF)
    for excel_fn in sorted(func_map.keys(), key=len, reverse=True):
        py_fn = func_map[excel_fn]
        expr = re.sub(r'(?<![_A-Za-z])' + excel_fn + r'(?=\s*\()', py_fn, expr, flags=re.IGNORECASE)

    # Replace range references A1:B3 → _RANGE(r1,c1,r2,c2)
    def replace_range(m):
        r1, c1 = excel_ref_to_rc(m.group(1))
        r2, c2 = excel_ref_to_rc(m.group(2))
        if r1 is None:
            return m.group(0)
        return f'_RANGE({r1},{c1},{r2},{c2})'
    expr = re.sub(r'\$?([A-Z]+\$?\d+):\$?([A-Z]+\$?\d+)', replace_range, expr, flags=re.IGNORECASE)

    # Replace single cell refs B3 → _cell(2,1)
    def replace_cell(m):
        full = m.group(0)
        r, c = excel_ref_to_rc(full.replace('$', ''))
        if r is None:
            return full
        return f'_cell({r},{c})'
    expr = re.sub(r'(?<![_A-Za-z(,])\$?[A-Z]+\$?\d+(?![A-Za-z(])', replace_cell, expr, flags=re.IGNORECASE)

    expr = expr.replace('^', '**')
    expr = re.sub(r'(?<![&])&(?![&])', '+', expr)  # & concatenation but not &&
    expr = re.sub(r'\bTRUE\b', 'True', expr, flags=re.IGNORECASE)
    expr = re.sub(r'\bFALSE\b', 'False', expr, flags=re.IGNORECASE)

    return expr


def extract_cell_refs(formula: str):
    refs = set()
    for m in re.finditer(r'\$?([A-Z]+\$?\d+):\$?([A-Z]+\$?\d+)', formula, re.IGNORECASE):
        r1, c1 = excel_ref_to_rc(m.group(1).replace('$',''))
        r2, c2 = excel_ref_to_rc(m.group(2).replace('$',''))
        if r1 is not None:
            for r in range(min(r1,r2), max(r1,r2)+1):
                for c in range(min(c1,c2), max(c1,c2)+1):
                    refs.add((r, c))
    for m in re.finditer(r'(?<![_A-Za-z:])\$?([A-Z]+\$?\d+)(?![A-Za-z:(])', formula, re.IGNORECASE):
        r, c = excel_ref_to_rc(m.group(1).replace('$',''))
        if r is not None:
            refs.add((r, c))
    return list(refs)


def make_eval_context(cell_values: dict):

    def _cell(r, c):
        v = cell_values.get((r, c), 0)
        try:
            return float(v) if v not in (None, '') else 0
        except (TypeError, ValueError):
            return v or 0

    def _cell_raw(r, c):
        return cell_values.get((r, c), None)

    def _RANGE(r1, c1, r2, c2):
        vals = []
        for r in range(min(r1,r2), max(r1,r2)+1):
            for c in range(min(c1,c2), max(c1,c2)+1):
                v = cell_values.get((r, c), 0)
                try:
                    vals.append(float(v) if v not in (None,'') else 0)
                except (TypeError, ValueError):
                    vals.append(v if v is not None else 0)
        return vals

    def _RANGE_RAW(r1, c1, r2, c2):
        """Range keeping original values (for SUMIF criteria matching)"""
        vals = []
        for r in range(min(r1,r2), max(r1,r2)+1):
            for c in range(min(c1,c2), max(c1,c2)+1):
                vals.append(cell_values.get((r, c), None))
        return vals

    def _flat(*args):
        result = []
        for a in args:
            if isinstance(a, list):
                result.extend(a)
            else:
                result.append(a)
        return result

    def _to_num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0

    def _SUM(*args):
        return sum(_to_num(x) for x in _flat(*args))

    def _AVERAGE(*args):
        flat = [_to_num(x) for x in _flat(*args)]
        return sum(flat)/len(flat) if flat else 0

    def _MIN(*args):
        flat = [_to_num(x) for x in _flat(*args)]
        return min(flat) if flat else 0

    def _MAX(*args):
        flat = [_to_num(x) for x in _flat(*args)]
        return max(flat) if flat else 0

    def _LARGE(arr, k):
        flat = sorted([_to_num(x) for x in (_flat(arr) if isinstance(arr, list) else [arr])], reverse=True)
        return flat[int(k)-1] if int(k) <= len(flat) else 0

    def _SMALL(arr, k):
        flat = sorted([_to_num(x) for x in (_flat(arr) if isinstance(arr, list) else [arr])])
        return flat[int(k)-1] if int(k) <= len(flat) else 0

    def _MEDIAN(*args):
        flat = sorted([_to_num(x) for x in _flat(*args)])
        n = len(flat)
        if not n: return 0
        return flat[n//2] if n%2 else (flat[n//2-1]+flat[n//2])/2

    def _STDEV(*args):
        flat = [_to_num(x) for x in _flat(*args)]
        if len(flat) < 2: return 0
        avg = sum(flat)/len(flat)
        return math.sqrt(sum((x-avg)**2 for x in flat)/(len(flat)-1))

    def _VAR(*args):
        flat = [_to_num(x) for x in _flat(*args)]
        if len(flat) < 2: return 0
        avg = sum(flat)/len(flat)
        return sum((x-avg)**2 for x in flat)/(len(flat)-1)

    def _PERCENTILE(arr, k):
        flat = sorted([_to_num(x) for x in (_flat(arr) if isinstance(arr,list) else [arr])])
        if not flat: return 0
        idx = (len(flat)-1) * _to_num(k)
        lo, hi = int(idx), min(int(idx)+1, len(flat)-1)
        return flat[lo] + (flat[hi]-flat[lo]) * (idx-lo)

    def _RANK(val, arr, order=0):
        flat = [_to_num(x) for x in (_flat(arr) if isinstance(arr,list) else [arr])]
        flat_sorted = sorted(flat, reverse=(order==0))
        try:
            return flat_sorted.index(_to_num(val)) + 1
        except ValueError:
            return 0

    def _match_criteria(val, criteria):
        """Match a cell value against an Excel criteria string like '>100', '<>0', 'text'"""
        if val is None:
            val = ''
        criteria_str = str(criteria).strip()
        # Operator criteria: ">100", "<=50", "<>0", "=text"
        for op in ('>=', '<=', '<>', '>', '<', '='):
            if criteria_str.startswith(op):
                crit_val = criteria_str[len(op):].strip('"').strip("'")
                try:
                    crit_num = float(crit_val)
                    cell_num = _to_num(val)
                    if op == '>':  return cell_num > crit_num
                    if op == '<':  return cell_num < crit_num
                    if op == '>=': return cell_num >= crit_num
                    if op == '<=': return cell_num <= crit_num
                    if op == '<>': return cell_num != crit_num
                    if op == '=':  return cell_num == crit_num
                except ValueError:
                    sv = str(val).lower()
                    cv = crit_val.lower()
                    if op == '<>': return sv != cv
                    if op == '=':  return sv == cv
                return False
        # Wildcard match
        pattern = criteria_str.strip('"').strip("'")
        if '*' in pattern or '?' in pattern:
            regex = re.escape(pattern).replace(r'\*', '.*').replace(r'\?', '.')
            return bool(re.fullmatch(regex, str(val), re.IGNORECASE))
        # Plain equality
        try:
            return _to_num(val) == float(pattern)
        except ValueError:
            return str(val).lower() == pattern.lower()

    def _SUMIF(range_vals, criteria, sum_range=None):
        r = range_vals if isinstance(range_vals, list) else [range_vals]
        s = sum_range if isinstance(sum_range, list) else r
        return sum(_to_num(sv) for rv, sv in zip(r, s) if _match_criteria(rv, criteria))

    def _SUMIFS(sum_range, *args):
        s = sum_range if isinstance(sum_range, list) else [sum_range]
        pairs = [(args[i], args[i+1] if isinstance(args[i], list) else [args[i]]) for i in range(0, len(args)-1, 2)]
        criteria_ranges = []
        for crit_range, crit_val in pairs:
            cr = crit_range if isinstance(crit_range, list) else [crit_range]
            criteria_ranges.append((cr, crit_val if not isinstance(crit_val, list) else crit_val[0]))
        total = 0
        for i, sv in enumerate(s):
            if all(_match_criteria(cr[i] if i < len(cr) else None, cv) for cr, cv in criteria_ranges):
                total += _to_num(sv)
        return total

    def _COUNTIF(range_vals, criteria):
        r = range_vals if isinstance(range_vals, list) else [range_vals]
        return sum(1 for v in r if _match_criteria(v, criteria))

    def _COUNTIFS(*args):
        pairs = [(args[i], args[i+1]) for i in range(0, len(args)-1, 2)]
        ranges = [(r if isinstance(r, list) else [r], c) for r, c in pairs]
        if not ranges: return 0
        n = len(ranges[0][0])
        count = 0
        for i in range(n):
            if all(_match_criteria(r[i] if i < len(r) else None, c) for r, c in ranges):
                count += 1
        return count

    def _AVERAGEIF(range_vals, criteria, avg_range=None):
        r = range_vals if isinstance(range_vals, list) else [range_vals]
        a = avg_range if isinstance(avg_range, list) else r
        matched = [_to_num(av) for rv, av in zip(r, a) if _match_criteria(rv, criteria)]
        return sum(matched)/len(matched) if matched else 0

    def _AVERAGEIFS(avg_range, *args):
        a = avg_range if isinstance(avg_range, list) else [avg_range]
        pairs = [(args[i] if isinstance(args[i], list) else [args[i]], args[i+1]) for i in range(0, len(args)-1, 2)]
        matched = []
        for i, av in enumerate(a):
            if all(_match_criteria(cr[i] if i < len(cr) else None, cv) for cr, cv in pairs):
                matched.append(_to_num(av))
        return sum(matched)/len(matched) if matched else 0

    def _VLOOKUP(lookup, table, col_index, approx=True):
        tbl = table if isinstance(table, list) else [table]
        col_idx = int(col_index) - 1
        # table is flat row-major; we need to know # of columns
        # Since _RANGE returns flat list, we can't easily do 2D lookup
        # Return first match in first column offset by col_index
        # This is a simplified implementation
        try:
            for i, val in enumerate(tbl):
                if approx:
                    if str(val).lower() == str(lookup).lower() or _to_num(val) == _to_num(lookup):
                        return tbl[i + col_idx] if i + col_idx < len(tbl) else 0
                else:
                    if str(val).lower() == str(lookup).lower():
                        return tbl[i + col_idx] if i + col_idx < len(tbl) else 0
        except Exception:
            pass
        return 0

    def _HLOOKUP(lookup, table, row_index, approx=True):
        return _VLOOKUP(lookup, table, row_index, approx)

    def _INDEX(arr, row_num, col_num=1):
        a = arr if isinstance(arr, list) else [arr]
        idx = int(row_num) - 1
        return a[idx] if 0 <= idx < len(a) else 0

    def _MATCH(lookup, arr, match_type=1):
        a = arr if isinstance(arr, list) else [arr]
        for i, v in enumerate(a):
            try:
                if _to_num(v) == _to_num(lookup):
                    return i + 1
            except Exception:
                if str(v).lower() == str(lookup).lower():
                    return i + 1
        return 0

    def _IF(condition, true_val, false_val=0):
        return true_val if condition else false_val

    def _IFS(*args):
        for i in range(0, len(args)-1, 2):
            if args[i]:
                return args[i+1]
        return 0

    def _IFERROR(value, fallback):
        try:
            return value if value is not None else fallback
        except Exception:
            return fallback

    def _IFNA(value, fallback):
        return fallback if value is None else value

    def _ROUND(value, digits=0):
        try:
            return round(float(value), int(digits))
        except (TypeError, ValueError):
            return 0

    def _ROUNDUP(value, digits=0):
        import math
        f = float(value)
        d = int(digits)
        return math.ceil(f * 10**d) / 10**d

    def _ROUNDDOWN(value, digits=0):
        import math
        f = float(value)
        d = int(digits)
        return math.floor(f * 10**d) / 10**d

    def _CEILING(value, significance=1):
        import math
        return math.ceil(float(value) / float(significance)) * float(significance)

    def _FLOOR(value, significance=1):
        import math
        return math.floor(float(value) / float(significance)) * float(significance)

    def _COUNT(*args):
        count = 0
        for a in _flat(*args):
            try:
                float(a)
                count += 1
            except (TypeError, ValueError):
                pass
        return count

    def _COUNTA(*args):
        return sum(1 for a in _flat(*args) if a not in (None, ''))

    def _COUNTBLANK(*args):
        return sum(1 for a in _flat(*args) if a in (None, ''))

    def _AND(*args): return all(bool(a) for a in _flat(*args))
    def _OR(*args):  return any(bool(a) for a in _flat(*args))
    def _NOT(v):     return not bool(v)
    def _XOR(*args): return sum(bool(a) for a in _flat(*args)) % 2 == 1

    def _INT(v):
        import math
        return int(math.floor(float(v)))

    def _TRUNC(v, digits=0):
        import math
        f = float(v)
        d = int(digits)
        return math.trunc(f * 10**d) / 10**d

    def _MOD(n, d): return float(n) % float(d) if float(d) != 0 else 0

    def _CONCATENATE(*args): return ''.join(str(a) for a in _flat(*args) if a is not None)

    def _TEXTJOIN(delim, ignore_empty, *args):
        vals = [str(a) for a in _flat(*args) if not (ignore_empty and a in (None, ''))]
        return str(delim).join(vals)

    def _TEXT(value, fmt=''):
        try:
            f = float(value)
            if '%' in str(fmt): return f"{f*100:.1f}%"
            if '#,##0' in str(fmt): return f"{f:,.0f}"
            if '0.00' in str(fmt): return f"{f:.2f}"
            return f"{f:,.2f}"
        except (TypeError, ValueError):
            return str(value)

    def _VALUE(v):
        try:
            return float(str(v).replace(',','.').replace(' ',''))
        except (TypeError, ValueError):
            return 0

    def _LEFT(text, n=1):   return str(text)[:int(n)]
    def _RIGHT(text, n=1):  return str(text)[-int(n):]
    def _MID(text, start, n): s=str(text); return s[int(start)-1:int(start)-1+int(n)]
    def _TRIM(text):        return ' '.join(str(text).split())
    def _UPPER(text):       return str(text).upper()
    def _LOWER(text):       return str(text).lower()
    def _PROPER(text):      return str(text).title()
    def _REPT(text, n):     return str(text) * int(n)

    def _FIND(find, within, start=1):
        try:
            return str(within).index(str(find), int(start)-1) + 1
        except ValueError:
            return 0

    def _SEARCH(find, within, start=1):
        try:
            return str(within).lower().index(str(find).lower(), int(start)-1) + 1
        except ValueError:
            return 0

    def _SUBSTITUTE(text, old, new, instance=None):
        t = str(text)
        if instance is None:
            return t.replace(str(old), str(new))
        count = 0
        result = ''
        idx = 0
        while idx < len(t):
            pos = t.find(str(old), idx)
            if pos == -1:
                result += t[idx:]
                break
            count += 1
            if count == int(instance):
                result += t[idx:pos] + str(new)
                idx = pos + len(str(old))
                result += t[idx:]
                break
            result += t[idx:pos+len(str(old))]
            idx = pos + len(str(old))
        return result

    def _REPLACE(text, start, num, new):
        t = str(text)
        s = int(start)-1
        return t[:s] + str(new) + t[s+int(num):]

    def _ISBLANK(v):  return v is None or str(v).strip() == ''
    def _ISERROR(v):  return False
    def _ISNA(v):     return v is None
    def _ISNUMBER(v):
        try: float(v); return True
        except: return False
    def _ISTEXT(v):   return isinstance(v, str)

    def _POWER(base, exp): return float(base) ** float(exp)
    def _LOG(v, base=10):  return math.log(float(v), float(base)) if float(v) > 0 else 0
    def _LOG10(v):         return math.log10(float(v)) if float(v) > 0 else 0
    def _LN(v):            return math.log(float(v)) if float(v) > 0 else 0
    def _PI():             return math.pi
    def _RAND():           import random; return random.random()
    def _RANDBETWEEN(lo, hi): import random; return random.randint(int(lo), int(hi))

    def _TODAY():  return date.today()
    def _NOW():    return datetime.now()
    def _YEAR(d):
        if isinstance(d, (date, datetime)): return d.year
        try: return datetime.fromisoformat(str(d)).year
        except: return 0
    def _MONTH(d):
        if isinstance(d, (date, datetime)): return d.month
        try: return datetime.fromisoformat(str(d)).month
        except: return 0
    def _DAY(d):
        if isinstance(d, (date, datetime)): return d.day
        try: return datetime.fromisoformat(str(d)).day
        except: return 0
    def _DATE(y, m, d): return date(int(y), int(m), int(d))
    def _DAYS(end, start):
        try:
            e = end if isinstance(end, date) else datetime.fromisoformat(str(end)).date()
            s = start if isinstance(start, date) else datetime.fromisoformat(str(start)).date()
            return (e - s).days
        except: return 0
    def _NETWORKDAYS(start, end, holidays=None):
        try:
            s = start if isinstance(start, date) else datetime.fromisoformat(str(start)).date()
            e = end if isinstance(end, date) else datetime.fromisoformat(str(end)).date()
            days = 0
            cur = s
            while cur <= e:
                if cur.weekday() < 5:
                    days += 1
                cur = date.fromordinal(cur.toordinal() + 1)
            return days
        except: return 0
    def _EOMONTH(start, months):
        import calendar
        try:
            d = start if isinstance(start, date) else datetime.fromisoformat(str(start)).date()
            m = d.month + int(months)
            y = d.year + (m-1)//12
            m = (m-1)%12 + 1
            return date(y, m, calendar.monthrange(y, m)[1])
        except: return 0

    # Financial functions
    def _NPV(rate, *args):
        vals = _flat(*args)
        return sum(v / (1+float(rate))**(i+1) for i, v in enumerate(vals) if v is not None)

    def _PMT(rate, nper, pv, fv=0, t=0):
        r, n, p = float(rate), float(nper), float(pv)
        if r == 0: return -(p + float(fv)) / n
        return (r * (p * (1+r)**n + float(fv))) / ((1+r)**n - 1) * (-1)

    def _FV(rate, nper, pmt, pv=0, t=0):
        r, n = float(rate), float(nper)
        if r == 0: return -(float(pv) + float(pmt)*n)
        return -(float(pv)*(1+r)**n + float(pmt)*(1+r*t)*((1+r)**n-1)/r)

    def _PV(rate, nper, pmt, fv=0, t=0):
        r, n = float(rate), float(nper)
        if r == 0: return -(float(pmt)*n + float(fv))
        return -(float(pmt)*(1+r*t)*(1-(1+r)**-n)/r + float(fv)*(1+r)**-n)

    def _RATE(nper, pmt, pv, fv=0, t=0, guess=0.1):
        # Newton-Raphson approximation
        r = float(guess)
        for _ in range(100):
            try:
                f = float(pv)*(1+r)**float(nper) + float(pmt)*(1+r*t)*((1+r)**float(nper)-1)/r + float(fv)
                df = float(nper)*float(pv)*(1+r)**(float(nper)-1) + float(pmt)*(((1+r)**float(nper)-1)/r + float(nper)*(1+r)**(float(nper)-1)/r)
                if df == 0: break
                r -= f/df
            except: break
        return r

    def _NPER(rate, pmt, pv, fv=0, t=0):
        r = float(rate)
        if r == 0: return -(float(pv)+float(fv))/float(pmt) if float(pmt) != 0 else 0
        return math.log((-float(fv)*r+float(pmt)*(1+r*t))/(float(pv)*r+float(pmt)*(1+r*t)))/math.log(1+r)

    def _CHOOSE(index, *vals):
        idx = int(index) - 1
        flat = _flat(*vals)
        return flat[idx] if 0 <= idx < len(flat) else 0

    def _OFFSET(ref, rows, cols, h=None, w=None): return 0  # complex to implement
    def _INDIRECT(ref): return 0
    def _ROW(ref=None): return 0
    def _COLUMN(ref=None): return 0
    def _ROWS(arr): return len(arr) if isinstance(arr, list) else 1
    def _COLUMNS(arr): return 1

    return {
        '_cell': _cell, '_RANGE': _RANGE,
        '_SUM': _SUM, '_AVERAGE': _AVERAGE, '_MIN': _MIN, '_MAX': _MAX,
        '_LARGE': _LARGE, '_SMALL': _SMALL, '_MEDIAN': _MEDIAN,
        '_STDEV': _STDEV, '_VAR': _VAR, '_PERCENTILE': _PERCENTILE, '_RANK': _RANK,
        '_SUMIF': _SUMIF, '_SUMIFS': _SUMIFS,
        '_COUNTIF': _COUNTIF, '_COUNTIFS': _COUNTIFS,
        '_AVERAGEIF': _AVERAGEIF, '_AVERAGEIFS': _AVERAGEIFS,
        '_VLOOKUP': _VLOOKUP, '_HLOOKUP': _HLOOKUP,
        '_INDEX': _INDEX, '_MATCH': _MATCH,
        '_IF': _IF, '_IFS': _IFS, '_IFERROR': _IFERROR, '_IFNA': _IFNA,
        '_ROUND': _ROUND, '_ROUNDUP': _ROUNDUP, '_ROUNDDOWN': _ROUNDDOWN,
        '_CEILING': _CEILING, '_FLOOR': _FLOOR,
        '_COUNT': _COUNT, '_COUNTA': _COUNTA, '_COUNTBLANK': _COUNTBLANK,
        '_AND': _AND, '_OR': _OR, '_NOT': _NOT, '_XOR': _XOR,
        '_INT': _INT, '_TRUNC': _TRUNC, '_MOD': _MOD,
        '_CONCATENATE': _CONCATENATE, '_TEXTJOIN': _TEXTJOIN,
        '_TEXT': _TEXT, '_VALUE': _VALUE,
        '_LEFT': _LEFT, '_RIGHT': _RIGHT, '_MID': _MID,
        '_TRIM': _TRIM, '_UPPER': _UPPER, '_LOWER': _LOWER, '_PROPER': _PROPER,
        '_REPT': _REPT, '_FIND': _FIND, '_SEARCH': _SEARCH,
        '_SUBSTITUTE': _SUBSTITUTE, '_REPLACE': _REPLACE,
        '_ISBLANK': _ISBLANK, '_ISERROR': _ISERROR, '_ISNA': _ISNA,
        '_ISNUMBER': _ISNUMBER, '_ISTEXT': _ISTEXT,
        '_POWER': _POWER, '_LOG': _LOG, '_LOG10': _LOG10, '_LN': _LN,
        '_PI': _PI, '_RAND': _RAND, '_RANDBETWEEN': _RANDBETWEEN,
        '_TODAY': _TODAY, '_NOW': _NOW,
        '_YEAR': _YEAR, '_MONTH': _MONTH, '_DAY': _DAY, '_DATE': _DATE,
        '_DAYS': _DAYS, '_NETWORKDAYS': _NETWORKDAYS, '_EOMONTH': _EOMONTH,
        '_NPV': _NPV, '_PMT': _PMT, '_FV': _FV, '_PV': _PV,
        '_RATE': _RATE, '_NPER': _NPER,
        '_CHOOSE': _CHOOSE, '_OFFSET': _OFFSET, '_INDIRECT': _INDIRECT,
        '_ROW': _ROW, '_COLUMN': _COLUMN, '_ROWS': _ROWS, '_COLUMNS': _COLUMNS,
        'abs': abs, 'int': int, 'len': len, 'str': str,
        'math': math, 'True': True, 'False': False,
    }


def recalculate_sheet(sheet):
    from .models import Cell
    cells = list(sheet.cells.all())
    cell_values = {}
    for c in cells:
        try:
            cell_values[(c.row, c.col)] = float(c.value) if c.value not in (None,'') else 0
        except (TypeError, ValueError):
            cell_values[(c.row, c.col)] = c.value or 0

    formula_cells = [c for c in cells if c.python_formula]
    changed = {}
    for _ in range(5):
        for c in formula_cells:
            ctx = make_eval_context(cell_values)
            try:
                result = eval(c.python_formula, {"__builtins__": {}}, ctx)
                if isinstance(result, float):
                    result = round(result, c.decimal_places)
                new_val = str(result) if result is not None else '0'
                if new_val != c.value:
                    changed[(c.row, c.col)] = new_val
                cell_values[(c.row, c.col)] = result
            except Exception:
                pass
    return changed, cell_values


def recalculate_dependents(sheet, changed_row, changed_col, new_value):
    from .models import Cell
    from collections import deque

    cells = list(sheet.cells.all())
    cell_values = {}
    for c in cells:
        try:
            cell_values[(c.row, c.col)] = float(c.value) if c.value not in (None,'') else 0
        except (TypeError, ValueError):
            cell_values[(c.row, c.col)] = c.value or 0
    try:
        cell_values[(changed_row, changed_col)] = float(new_value)
    except (TypeError, ValueError):
        cell_values[(changed_row, changed_col)] = new_value

    dep_map = {}
    for d in sheet.dependencies.all():
        dep_map.setdefault((d.source_row, d.source_col), []).append((d.target_row, d.target_col))

    queue = deque([(changed_row, changed_col)])
    visited = set()
    recalc_order = []
    while queue:
        src = queue.popleft()
        if src in visited: continue
        visited.add(src)
        for tgt in dep_map.get(src, []):
            if tgt not in visited:
                recalc_order.append(tgt)
            queue.append(tgt)

    seen = set()
    recalc_final = []
    for item in recalc_order:
        if item not in seen:
            seen.add(item)
            recalc_final.append(item)

    cell_map = {(c.row, c.col): c for c in cells}
    updates = []
    to_save = []

    for (r, c) in recalc_final:
        cell = cell_map.get((r, c))
        if not cell or not cell.python_formula: continue
        ctx = make_eval_context(cell_values)
        try:
            result = eval(cell.python_formula, {"__builtins__": {}}, ctx)
            if isinstance(result, float):
                result = round(result, cell.decimal_places)
            new_val = str(result) if result is not None else '0'
            cell_values[(r, c)] = result
            cell.value = new_val
            to_save.append(cell)
            updates.append({
                'row': r, 'col': c, 'value': new_val,
                'formatted': format_value(new_val, cell.format_type, cell.decimal_places),
            })
        except Exception:
            pass

    if to_save:
        Cell.objects.bulk_update(to_save, ['value'])
    return updates


def format_value(value, fmt_type, decimal_places=2):
    if value is None or value == '': return ''
    try:
        f = float(value)
        if fmt_type == 'currency': return f"{f:,.{decimal_places}f} ₸"
        if fmt_type == 'percent':  return f"{f:.{decimal_places}f}%"
        if fmt_type == 'number':   return f"{f:,.{decimal_places}f}"
        if f == int(f) and decimal_places == 0: return f"{int(f):,}"
        return f"{f:,.{decimal_places}f}"
    except (TypeError, ValueError):
        return str(value)
