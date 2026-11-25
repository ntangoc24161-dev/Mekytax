import datetime
import re

# Excel serial dates commonly land in ~20000..60000 for years 1954..2064.
_EXCEL_SERIAL_MIN = 20000
_EXCEL_SERIAL_MAX = 60000


def _to_date(raw_val):
    """Convert datetime/date, dd/MM/yyyy text, or Excel serial number to date."""
    if isinstance(raw_val, datetime.datetime):
        return raw_val.date()
    if isinstance(raw_val, datetime.date):
        return raw_val
    if isinstance(raw_val, datetime.time):
        return None
    if isinstance(raw_val, (int, float)) and not isinstance(raw_val, bool):
        serial = float(raw_val)
        if _EXCEL_SERIAL_MIN <= serial <= _EXCEL_SERIAL_MAX:
            try:
                from openpyxl.utils.datetime import from_excel
                dt = from_excel(serial)
                if isinstance(dt, datetime.datetime):
                    return dt.date()
                if isinstance(dt, datetime.date):
                    return dt
            except Exception:
                return None
        return None
    if isinstance(raw_val, str):
        s = raw_val.strip()
        if "/" in s:
            parts = s.split("/")
            if len(parts) == 3:
                try:
                    return datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                except ValueError:
                    return None
        # Numeric string that is really an Excel serial
        try:
            serial = float(s.replace(",", "."))
            if serial.is_integer() and _EXCEL_SERIAL_MIN <= serial <= _EXCEL_SERIAL_MAX:
                return _to_date(serial)
        except ValueError:
            pass
    return None


def _format_date(dt, switch):
    if 'dd/MM/yyyy' in switch:
        return dt.strftime('%d/%m/%Y')
    if '"dd"' in switch or "'dd'" in switch:
        return dt.strftime('%d')
    if '"MM"' in switch or "'MM'" in switch:
        return dt.strftime('%m')
    if '"yyyy"' in switch or "'yyyy'" in switch:
        return dt.strftime('%Y')
    return dt.strftime('%d/%m/%Y')


def format_value(field_name, switch, raw_val):
    # Handle numbers with switch \# "#,##0"
    if switch and r'\#' in switch:
        try:
            num = float(raw_val)
            # Format with thousand dot separator (Vietnamese style)
            # e.g. 19719634 -> 19.719.634
            formatted = f"{num:,.0f}".replace(",", ".")
            return formatted
        except (ValueError, TypeError):
            return str(raw_val) if raw_val is not None else ""

    # Handle dates with switch \@ "..."
    if switch and r'\@' in switch:
        dt = _to_date(raw_val)
        if dt:
            return _format_date(dt, switch)
        if raw_val is None or str(raw_val).strip() in ["", "00:00:00", "0", "1899-12-30 00:00:00+00:00"]:
            if '"dd"' in switch or '"MM"' in switch:
                return "....."
            return "..... / ..... / ......"
        return str(raw_val)

    # General values
    if raw_val is None:
        return ""
    if isinstance(raw_val, datetime.time) and str(raw_val) == "00:00:00":
        return "....."

    # Date-named fields without switch still must not print Excel serials
    name = (field_name or "").lower()
    looks_like_date_field = (
        "ngay" in name or "ngày" in name or name.startswith("từ_") or name.startswith("đến_")
        or name.startswith("tu_") or name.startswith("den_")
    )
    dt = _to_date(raw_val)
    if dt is not None and (looks_like_date_field or isinstance(raw_val, (datetime.date, datetime.datetime))):
        return dt.strftime('%d/%m/%Y')

    if isinstance(raw_val, float) and raw_val.is_integer():
        return str(int(raw_val))
    return str(raw_val)



