"""Monthly report engine. Uploads are processed in memory; source files are never saved."""
from __future__ import annotations

import ast
import copy
import io
import math
import operator
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.datetime import from_excel
from openpyxl.workbook.properties import CalcProperties


@dataclass
class Procurement:
    batches: list[dict]
    pending: int = 0

    @property
    def company_times(self):
        return sum(x['家次'] for x in self.batches)

    @property
    def package_times(self):
        return sum(x['包次'] for x in self.batches)


def normalize(text):
    return re.sub(r'\s+', '', str(text or '')).replace('(', '（').replace(')', '）')


def suppliers(value):
    return [normalize(v) for v in re.split(r'[\r\n/／]+', str(value or '')) if v.strip()]


def numeric(value, label):
    if isinstance(value, bool):
        raise ValueError(f'{label}须为数字')
    try:
        value = float(value)
    except (TypeError, ValueError):
        raise ValueError(f'{label}须为数字') from None
    if not math.isfinite(value) or value < 0:
        raise ValueError(f'{label}须为非负有限数字')
    return value


def award_date(value, epoch):
    if isinstance(value, (date, datetime)):
        return value
    if isinstance(value, (int, float)):
        return from_excel(value, epoch)
    value = str(value or '').strip()
    if not value or value.upper() in {'TBD', '待定', '/'}:
        return None
    value = value.replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-')
    try:
        return datetime.strptime(value.split()[0], '%Y-%m-%d')
    except ValueError:
        raise ValueError(f'无法识别中标日期：{value}') from None


def count_procurement(content: bytes, year: int, month: int) -> Procurement:
    """Filter award month first; deduplicate (batch, supplier), not supplier alone."""
    wb = load_workbook(io.BytesIO(content), data_only=True)
    groups = defaultdict(lambda: {'names': set(), 'packages': 0, 'projects': 0})
    seen = {}
    pending = 0
    found = False
    for ws in wb:
        header = None
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10)):
            cols = {normalize(c.value): c.column for c in row if c.value is not None}
            if '投标供应商' in cols:
                header = row[0].row, cols
                break
        if header is None:
            continue
        found = True
        start, cols = header
        required = ['中标日期', '参与批次', '采购项目名称']
        if any(k not in cols for k in required):
            raise ValueError(f'{ws.title}缺少中标日期、参与批次或采购项目名称')
        for row in range(start + 1, ws.max_row + 1):
            def val(key):
                return ws.cell(row, cols[key]).value if key in cols else None
            names = suppliers(val('投标供应商'))
            if not names or names == ['TBD']:
                continue
            when = award_date(val('中标日期'), wb.epoch)
            if when is None:
                pending += 1
                continue
            if (when.year, when.month) != (year, month):
                continue
            batch = normalize(val('参与批次'))
            project = normalize(val('采购项目名称'))
            if not batch or not project:
                raise ValueError(f'{ws.title}第{row}行缺少批次或项目名称')
            package = normalize(val('分标名称') or val('标包名称') or val('包号'))
            key = batch, project, package
            signature = tuple(sorted(names))
            if key in seen:
                if seen[key] != signature:
                    raise ValueError(f'同一项目在多个位置的供应商名单不一致：{project}')
                continue
            seen[key] = signature
            # Repeated names within one package are not additional bids.
            names = set(names)
            groups[batch]['names'].update(names)
            groups[batch]['packages'] += len(names)
            groups[batch]['projects'] += 1
    if not found:
        raise ValueError('未找到投标供应商列')
    return Procurement([
        {'批次': batch, '项目行数': g['projects'], '家次': len(g['names']), '包次': g['packages']}
        for batch, g in sorted(groups.items())
    ], pending)


def electricity(content: bytes, scope=('AL4', 'AL5', 'AL6')):
    wb = load_workbook(io.BytesIO(content), data_only=True)
    readings = {}
    for ws in wb:
        for row in ws:
            for c in row:
                match = re.fullmatch(r'(AL\d+)(?:[（(].*[）)])?', str(c.value or '').strip())
                if not match or match[1] not in scope:
                    continue
                # In supplied meter layouts, two readings and multiplier follow the meter name.
                if c.column + 3 > ws.max_column:
                    continue
                old, new, factor = [ws.cell(c.row, c.column + n).value for n in (1, 2, 3)]
                if not isinstance(old, (int, float)) or not isinstance(new, (int, float)):
                    continue
                multiplier = 1 if factor in (None, '', '/', '-') else numeric(factor, '电表倍率')
                if not multiplier:
                    raise ValueError(f'{match[1]}倍率不能为0')
                usage = numeric(round((new - old) * multiplier, 8), f'{match[1]}用电量')
                if match[1] in readings and not math.isclose(readings[match[1]], usage):
                    raise ValueError(f'{match[1]}存在不一致的重复抄表数据')
                readings[match[1]] = usage
    missing = set(scope) - readings.keys()
    if missing:
        raise ValueError('用电文件缺少电表：' + '、'.join(sorted(missing)))
    return round(sum(readings.values()), 8), readings


def formula_number(wb, ws, address, stack=None):
    """Resolve the simple arithmetic used by the source cumulative schedules; never eval code."""
    stack = set() if stack is None else stack
    key = (ws.title, address)
    if key in stack:
        raise ValueError('累计单元格存在循环引用')
    value = ws[address].value
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if not isinstance(value, str) or not value.startswith('='):
        raise ValueError(f'无法读取累计值：{ws.title}!{address}')
    stack = stack | {key}
    expression = value[1:]
    def replace(m):
        sheet = wb[m[1].replace("''", "'")] if m[1] else ws
        return repr(formula_number(wb, sheet, m[2].replace('$', ''), stack))
    expression = re.sub(r"(?:'((?:[^']|'')+)'!)?(\$?[A-Z]{1,3}\$?\d+)", replace, expression)
    expression = expression.replace('^', '**')
    expression = re.sub(r'\bSUM\(([^(),]+)\)', r'(\1)', expression, flags=re.I)
    ops = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
           ast.Div: operator.truediv, ast.Pow: operator.pow}
    def calc(node):
        if isinstance(node, ast.Constant) and type(node.value) in (int, float):
            return node.value
        if isinstance(node, ast.BinOp) and type(node.op) in ops:
            return ops[type(node.op)](calc(node.left), calc(node.right))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            return -calc(node.operand) if isinstance(node.op, ast.USub) else calc(node.operand)
        raise ValueError(f'累计值公式超出支持范围：{value}；请在Excel中重算保存后上传')
    return float(calc(ast.parse(expression, mode='eval').body))


def build_report(previous: bytes, last_year: bytes, year: int, month: int,
                 report_date: date, gasoline_liters: float, electricity_kwh: float,
                 procurement: Procurement | None):
    date(year, month, 1)
    gasoline_liters = numeric(gasoline_liters, '本月油量')
    electricity_kwh = numeric(electricity_kwh, '本月用电量')
    wb = load_workbook(io.BytesIO(previous))
    cached = load_workbook(io.BytesIO(previous), data_only=True)
    last = load_workbook(io.BytesIO(last_year), data_only=True)
    main = wb.worksheets[0]
    old = cached.worksheets[0]
    if normalize(main['B23'].value) != normalize('其中：外购使用电力（MWh）'):
        raise ValueError('上月表不是支持的2026新版布局；请使用含“外购使用电力”的模板')
    # Prefer the actual month in sheet/file metadata, not the report's next-month submission date.
    periods = re.findall(r'(\d{1,2})月', main.title)
    expected_month = 12 if month == 1 else month - 1
    if periods and int(periods[-1]) != expected_month:
        raise ValueError('上月模板月份与目标月份不连续')
    last_date = last.worksheets[0]['A2'].value
    if isinstance(last_date, (date, datetime)) and last_date.year not in (year - 1, year):
        raise ValueError('去年同期文件年份不匹配')
    def prior(addr):
        if month == 1:
            return 0.0
        v = old[addr].value
        return float(v) if isinstance(v, (int, float)) else formula_number(wb, main, addr)
    start = {k: prior(k) for k in ('C9', 'C12', 'C23', 'C29')}
    gas = round(start['C29'] + gasoline_liters * .00073, 10)
    power = round(start['C23'] + electricity_kwh * .001, 10)
    warnings = []
    increments = (procurement.company_times, procurement.package_times) if procurement else (None, None)
    main['A2'] = report_date
    main['C9'] = start['C9'] + (increments[0] or 0)
    main['C12'] = start['C12'] + (increments[1] or 0)
    main['C23'] = power
    main['C29'] = gas
    if procurement is None:
        warnings.append('未提供采购清单：家次和包次仅保留上月累计，尚未确认本月新增。')
    elif procurement.pending:
        warnings.append(f'清单中有{procurement.pending}行中标日期待定，未计入确定月份。')
    labels = {normalize(r[1].value): r[2].value for r in last.worksheets[0] if len(r) >= 3 and r[1].value}
    aliases = {normalize('其中：外购使用电力（MWh）'): normalize('用电量（MWh）')}
    for row in [9,10,11,12,13,14,15,16,21,23,24,25,26,27,28,29,30,31,32,33,34]:
        label = normalize(main.cell(row, 2).value)
        key = label if label in labels else aliases.get(label)
        if key in labels:
            main.cell(row, 4).value = labels[key]
        else:
            main.cell(row, 4).value = None
            if row in (9,12,21,23,29):
                raise ValueError(f'去年同期缺少关键指标：{main.cell(row,2).value}')
    # Carry other optional cumulative energy/count inputs forward; this release exposes oil and power changes.
    mapping = {main.title: f'服务双碳目标贡献力-{month}月'}
    aux = wb.worksheets[1] if len(wb.worksheets) > 1 else None
    if aux:
        mapping[aux.title] = f'{month}月采购、油、电'
    for sheet in wb:
        for row in sheet:
            for cell in row:
                if cell.data_type == 'f':
                    for before, after in mapping.items():
                        cell.value = cell.value.replace("'" + before.replace("'", "''") + "'!", "'" + after + "'!")
    for before, after in mapping.items():
        wb[before].title = after
    if aux:
        month_rows = []
        for row in range(3, min(aux.max_row, 20) + 1):
            v = aux.cell(row, 2).value
            if type(v) in (int,float) and v == int(v) and 1 <= v <= 12:
                month_rows.append(row)
            elif month_rows:
                break
        row = month_rows[-1] + 1 if month_rows else 3
        if month == 1:
            row = month_rows[0] if month_rows else 3
            for r in month_rows:
                for col in range(2,15):
                    if not isinstance(aux.cell(r,col), MergedCell): aux.cell(r,col).value = None
        if any(isinstance(aux.cell(row,col),MergedCell) or aux.cell(row,col).value is not None for col in range(2,15)):
            raise ValueError('辅助页月份区没有空行，需先扩展模板月份区')
        for col in range(2,15):
            aux.cell(row,col)._style = copy.copy(aux.cell(max(3,row-1),col)._style)
        values = {'B':month,'C':start['C29'],'D':gasoline_liters,'G':start['C23'],'H':electricity_kwh,
                  'M':increments[0],'N':increments[1]}
        for col,v in values.items(): aux[f'{col}{row}'] = v
        formulas = {'E':f'=C{row}+D{row}*0.00073','I':f'=G{row}+H{row}*0.001',
                    'K':f'={start["C9"]}+M{row}','L':f'={start["C12"]}+N{row}'}
        for col,v in formulas.items(): aux[f'{col}{row}'] = v
        for addr,col in [('C9','K'),('C12','L'),('C23','I'),('C29','E')]:
            main[addr] = f"='{aux.title}'!{col}{row}"
    wb.calculation = CalcProperties(calcId=191029, fullCalcOnLoad=True, forceFullCalc=True, calcMode='auto')
    result = io.BytesIO()
    wb.save(result)
    summary = {'本月家次':increments[0], '本月包次':increments[1],
               '累计家次':start['C9']+(increments[0] or 0), '累计包次':start['C12']+(increments[1] or 0),
               '本月电量kWh':electricity_kwh,'累计外购电MWh':power,'本月汽油L':gasoline_liters,'累计汽油t':gas}
    return result.getvalue(), summary, warnings
