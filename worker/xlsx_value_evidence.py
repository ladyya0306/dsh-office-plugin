"""Shared explicit unit and typed temporal evidence; no business inference."""
from datetime import date, datetime, time
import re
from worker.docx_prepare import explicit_unit

UNITS = {'元', '千元', '万元', '亿元', '%'}


def evidence_units(evidence):
    result = set()
    for item in evidence:
        text = str(item).strip()
        unit = text if text in UNITS else explicit_unit(text)
        if unit:
            result.add(unit)
    return result


def temporal_value(value):
    if getattr(value, 'microsecond', 0) % 1000:
        raise ValueError('日期/时间精度超过毫秒，不能悄悄截断')
    # Date-only OOXML and midnight numeric dates share one datetime identity.
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            raise ValueError('Excel日期不能隐式丢弃时区')
        return value.isoformat(timespec='milliseconds'), 'datetime'
    if isinstance(value, date):
        return datetime.combine(value, time()).isoformat(timespec='milliseconds'), 'datetime'
    if isinstance(value, time):
        if value.tzinfo is not None:
            raise ValueError('Excel时间不能隐式丢弃时区')
        return value.isoformat(timespec='milliseconds'), 'time'
    raise ValueError('未支持的日期、时间或持续时长')


def temporal_xml(raw, kind, date_style, epoch, number_format):
    if kind != 'd' and not (kind == 'n' and date_style and raw is not None):
        return None
    from openpyxl.utils.datetime import from_excel, from_ISO8601, WINDOWS_EPOCH
    from openpyxl.styles.numbers import is_timedelta_format
    if kind == 'd':
        clock = r'[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,3})?)?'
        if not isinstance(raw, str) or not re.fullmatch(r'(?:[0-9]{4}-[0-9]{2}-[0-9]{2}(?:T'+clock+r')?|'+clock+r')', raw):
            raise ValueError('日期需完整无时区ISO值，精度不超过毫秒；不能截取有效前缀')
        value = from_ISO8601(raw)
    else:
        from decimal import Decimal
        number = Decimal(raw)
        if not number.is_finite() or is_timedelta_format(number_format):
            raise ValueError('持续时长暂不支持，不能按日历日期解释')
        if epoch == WINDOWS_EPOCH and 60 <= number < 61:
            raise ValueError('1900日期系统的第60天为不存在的1900-02-29')
        value = from_excel(float(number), epoch)
    rendered, value_type = temporal_value(value)
    return {'value': rendered, 'value_type': value_type, 'temporal_evidence': {
        'storage_type': kind, 'raw_value': raw, 'epoch': epoch.date().isoformat(),
        'number_format': number_format, 'precision': 'milliseconds'}}
