"""Sparse, read-only cell evidence; never loads drawings, macros or external links."""
import bisect
import io
import re
import zipfile
from decimal import Decimal, InvalidOperation
from worker.office_worker import WorkerError, _check_zip
from worker.xlsx_package_patch import _xml, _sheet_parts, XlsxPackageError
from worker.docx_prepare import explicit_unit
from worker.xlsx_value_evidence import evidence_units, temporal_xml

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
UNIT = {'元', '千元', '万元', '亿元', '%'}
CELL = re.compile(r'^([A-Z]{1,3})([1-9][0-9]{0,6})$')


def position(ref):
    match = CELL.fullmatch(ref)
    if not match:
        raise WorkerError('invalid_cell_reference', '来源物理引用无效')
    col = 0
    for letter in match[1]:
        col = col * 26 + ord(letter) - 64
    row = int(match[2])
    if col > 16384 or row > 1048576:
        raise WorkerError('invalid_cell_reference', '超出Excel物理坐标')
    return row, col


def scalar(cell, strings):
    value = cell.find('s:v', NS)
    text = value.text if value is not None else None
    kind = cell.get('t', 'n')
    if kind == 'inlineStr':
        return ''.join(cell.xpath('s:is//s:t/text()', namespaces=NS))
    if text is None:
        return None
    if kind == 's':
        try:
            index = int(text)
            if index < 0:
                raise ValueError()
            return strings[index]
        except (ValueError, IndexError) as exc:
            raise WorkerError('invalid_shared_string', '来源共享字符串索引无效') from exc
    if kind == 'b':
        if text not in {'0', '1'}:
            raise WorkerError('invalid_scalar', '布尔值无效')
        return text == '1'
    if kind in {'str', 'e', 'd'}:
        return text
    try:
        number = Decimal(text)
        if not number.is_finite():
            raise InvalidOperation()
        if len(text) > 256 or number.adjusted() > 308:
            raise WorkerError('numeric_limit', '来源数值文本或指数超过安全读取上限')
        if number == number.to_integral_value():
            return int(number)
        converted=float(number)
        if converted==0 and number!=0:
            raise WorkerError('numeric_limit','来源非零数值低于当前数值表示范围，不能悄悄读成0')
        return converted
    except WorkerError:
        raise
    except (InvalidOperation, ValueError) as exc:
        raise WorkerError('invalid_scalar', '来源数值无效') from exc


def read_source(data, scope=None, requested_fields=None, offset=0, requested_refs=None):
    from lxml.etree import XMLSyntaxError
    try:
        return _read_source(data, scope, requested_fields, offset, requested_refs)
    except XlsxPackageError as exc:
        raise WorkerError(exc.code, exc.message) from exc
    except XMLSyntaxError as exc:
        raise WorkerError('invalid_source_xml', '来源XML语法错误，不能继续读取') from exc
    except KeyError as exc:
        raise WorkerError('missing_source_part', '来源缺少必需工作簿部件') from exc


def _index(text):
    try:
        value = int(text)
        if value < 0:
            raise ValueError()
        return value
    except (TypeError, ValueError) as exc:
        raise WorkerError('invalid_style_index', '来源样式编号无效') from exc


def _read_source(data, scope=None, requested_fields=None, offset=0, requested_refs=None):
    members = _check_zip(data, 'source')
    if not isinstance(offset, int) or isinstance(offset, bool) or not 0 <= offset <= 100000:
        raise WorkerError('invalid_source_offset', 'source_offset需0至100000整数')
    if requested_fields is not None and (not isinstance(requested_fields, list) or not 1 <= len(requested_fields) <= 32
                                        or any(not isinstance(f, str) or not f.strip() or len(f) > 80 for f in requested_fields)):
        raise WorkerError('invalid_requested_fields', 'requested_fields需1至32个非空检索词，每项最多80字符')
    query = [re.sub(r'\s+', '', f) for f in requested_fields or []]
    exact_refs=None
    if requested_refs is not None:
        if not isinstance(requested_refs,list) or not 1<=len(requested_refs)<=10000 or any(not isinstance(ref,str) or '!' not in ref for ref in requested_refs):
            raise WorkerError('invalid_requested_refs','来源位置需要1至10000个明确物理引用')
        exact_refs=set(requested_refs)
    nodes, summaries, stored_total = [], [], 0
    ignored = sorted(n for n in members if n.startswith(('xl/drawings/', 'xl/media/', 'xl/externalLinks/')) or
                     n in {'xl/vbaProject.bin', 'xl/calcChain.xml'})
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        from openpyxl.utils.datetime import WINDOWS_EPOCH, MAC_EPOCH
        properties = _xml(archive.read('xl/workbook.xml')).find('s:workbookPr', NS)
        epoch = MAC_EPOCH if properties is not None and properties.get('date1904') in {'1', 'true'} else WINDOWS_EPOCH
        formats = {}
        if 'xl/styles.xml' in members:
            from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
            styles = _xml(archive.read('xl/styles.xml'))
            custom = {_index(n.get('numFmtId')): n.get('formatCode', '') for n in styles.findall('s:numFmts/s:numFmt', NS)}
            for index, xf in enumerate(styles.findall('s:cellXfs/s:xf', NS)):
                fmt_id = _index(xf.get('numFmtId', '0'))
                fmt = custom.get(fmt_id, BUILTIN_FORMATS.get(fmt_id, 'General'))
                formats[index] = (fmt, is_date_format(fmt))
        strings = []
        if 'xl/sharedStrings.xml' in members:
            strings = [''.join(si.xpath('.//s:t/text()', namespaces=NS)) for si in _xml(archive.read('xl/sharedStrings.xml'))]
        parts = _sheet_parts(archive)
        if len(parts) > 32:
            raise WorkerError('workbook_too_large', '来源工作表超过32张')
        limits = None
        if scope:
            from openpyxl.utils.cell import range_boundaries
            if set(scope) != {'sheet', 'range'} or scope['sheet'] not in parts:
                raise WorkerError('invalid_scope', '指定范围须包含存在的sheet和range')
            try:
                limits = range_boundaries(scope['range'])
                if (any(x is None or x < 1 for x in limits) or limits[0] > limits[2] or limits[1] > limits[3]
                        or limits[2] > 16384 or limits[3] > 1048576):
                    raise ValueError()
            except (ValueError, TypeError) as exc:
                raise WorkerError('invalid_scope', '来源范围无效') from exc
        for name, part in parts.items():
            if scope and scope['sheet'] != name:
                continue
            if '!' in name:
                raise WorkerError('unaddressable_sheet_name', '工作表名含!，当前引用协议不能无歧义表示')
            tree = _xml(archive.read(part))
            physical = tree.findall('s:sheetData/s:row/s:c', NS)
            stored_total += len(physical)
            if stored_total > 100000:
                raise WorkerError('structure_limit', '来源实际存储格超过100000；未按声明尺寸分配网格')
            values, formulas, refs, errors, cell_formats,number_literals = {}, {}, {}, {}, {}, {}
            temporal = {}
            seen = set()
            for cell in physical:
                ref = cell.get('r', '')
                r, c = position(ref)
                if ref in seen:
                    raise WorkerError('duplicate_cell_reference', '来源物理单元格重复')
                seen.add(ref)
                val = scalar(cell, strings)
                if cell.get('t','n')=='n' and cell.find('s:v',NS) is not None:
                    number_literals[r,c]=cell.find('s:v',NS).text
                formula = cell.find('s:f', NS)
                if val is None and formula is None:
                    continue
                refs[r, c] = ref
                if formula is not None:
                    formulas[r, c] = '=' + (formula.text or '')
                values[r, c] = val
                fmt, date_style = formats.get(_index(cell.get('s', '0')), ('General', False))
                cell_formats[r, c] = fmt
                if cell.get('t') == 'e':
                    errors[r, c] = 'error_value'
                elif isinstance(val, (int, float)) and not isinstance(val, bool) and (
                        abs(val) > 2**53 - 1 or cell.find('s:v', NS) is not None and
                        abs(Decimal(cell.find('s:v', NS).text)) > 2**53 - 1):
                    errors[r, c] = 'numeric_precision'
                    values[r, c] = cell.find('s:v', NS).text
                try:
                    item = temporal_xml(cell.find('s:v', NS).text if cell.find('s:v', NS) is not None else None,
                                        cell.get('t', 'n'), date_style, epoch, fmt)
                    if item:
                        temporal[r, c] = item
                        values[r, c] = item['value']
                except (ValueError, OverflowError, TypeError):
                    errors[r, c] = 'invalid_temporal_value'
            rows, columns = {}, {}
            for r, c in values:
                rows.setdefault(r, []).append(c)
                columns.setdefault(c, []).append(r)
            for sequence in [*rows.values(), *columns.values()]:
                sequence.sort()
            included = set(values)
            declarations = [str(v) for v in values.values() if isinstance(v, str) and re.search(r'单位\s*[:：]', v)]
            declared_units = {explicit_unit(v) for v in declarations if explicit_unit(v)}
            if query or exact_refs is not None:
                included = set()
                for (r, c), value in values.items():
                    if name+'!'+refs[r,c] in (exact_refs or set()) or any(term in re.sub(r'\s+', '', str(value)) for term in query):
                        # Return evidence around a matching label, never a confirmed pairing.
                        included.add((r, c))
                        ci = bisect.bisect_left(rows[r], c)
                        ri = bisect.bisect_left(columns[c], r)
                        included.update((r, x) for x in rows[r][max(0,ci-4):ci+5])
                        included.update((x, c) for x in columns[c][max(0,ri-2):ri+3])
            selected = [(r, c) for r, c in sorted(included) if not limits or
                        limits[0] <= c <= limits[2] and limits[1] <= r <= limits[3]]
            merged = [m.get('ref') for m in tree.findall('s:mergeCells/s:mergeCell', NS)]
            if len(merged) > 10000:
                raise WorkerError('structure_limit', '合并范围超过10000项')
            summaries.append({'sheet': name, 'declared_dimension': tree.find('s:dimension', NS).get('ref') if tree.find('s:dimension', NS) is not None else None,
                              'merged_ranges': merged,
                              'stored_cells': len(physical), 'nonempty_or_formula_cells': len(values), 'selected_cells': len(selected)})
            for r, c in selected:
                left_keys = rows[r][:bisect.bisect_left(rows[r], c)]
                up_keys = columns[c][:bisect.bisect_left(columns[c], r)]
                left = [str(values[r, x])[:200] for x in left_keys[-4:] if values[r, x] is not None]
                above = [str(values[x, c])[:200] for x in up_keys[-4:] if values[x, c] is not None]
                right = values.get((r, c+1))
                evidence = ([str(right)] if right in UNIT else []) + left[-3:] + above[-3:]
                units = evidence_units(evidence)
                if right in UNIT:
                    units.add(right)
                # A percentage format is a display rule, never a money unit.
                percent_style = '%' in re.sub(r'"[^"]*"|\\.', '', cell_formats[r,c])
                if percent_style:
                    units.add('%')
                if (r, c) in temporal:
                    units.clear()
                # A unique explicit sheet declaration is evidence of amount units,
                # not inferred from the target. Multiple declarations stay ambiguous.
                if not units and isinstance(values[r,c],(int,float)) and not errors.get((r,c)):
                    units.update(declared_units)
                    evidence += declarations[:4]
                value = formulas.get((r, c), values[r, c])
                ref = name+'!'+refs[r, c]
                nodes.append({'ref': ref, 'kind': 'cell', 'value': value, 'text': str(value) if value is not None else '',
                              'row_headers': left, 'column_headers': above, 'parent_context': [name],
                              'row_header_refs': [name+'!'+refs[r,x] for x in left_keys[-4:]],
                              'column_header_refs': [name+'!'+refs[x,c] for x in up_keys[-4:]],
                              'unit': next(iter(units)) if len(units) == 1 else '', 'unit_evidence': evidence,
                              'unit_ambiguous': len(units) > 1, 'formula': (r, c) in formulas,
                              'cached_value': values[r, c] if (r, c) in formulas else None,
                              'number_format':cell_formats[r,c],
                              'number_literal':number_literals.get((r,c)),
                              'value_type':temporal.get((r,c),{}).get('value_type'),
                              'temporal_evidence':temporal.get((r,c),{}).get('temporal_evidence'),
                              'value_error': errors.get((r, c)), 'physical_path': ref,
                              'context_truncated': len(left_keys) > 4 or len(up_keys) > 4,
                              'relation_status': 'unconfirmed'})
    total = len(nodes)
    page = nodes[offset:offset+1000]
    next_offset = offset + len(page) if offset + len(page) < total else None
    return {'nodes': page, 'slots': [], 'capabilities': {'read': True, 'write': False,
            'formula_values': False, 'scope': scope, 'requested_fields': requested_fields,
            'source_offset': offset, 'next_offset': next_offset, 'matched_records': total,
            'requested_refs':requested_refs,
            'complete_requested_range': next_offset is None and offset == 0 and not query and exact_refs is None,
            'complete': not scope and not query and exact_refs is None and offset == 0 and next_offset is None and not ignored,
            'read_ignored_parts': ignored, 'images_read': False, 'macros_executed': False,
            'sheet_inventory': summaries, 'pagination': '只有本页位置可引用；未返回位置不算已读，检索结果不是业务匹配'}}
