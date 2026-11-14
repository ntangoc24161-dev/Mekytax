import docx
import re
import os
import datetime
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
        dt = None
        if isinstance(raw_val, (datetime.date, datetime.datetime)):
            dt = raw_val
        elif isinstance(raw_val, str) and "/" in raw_val:
            parts = raw_val.split("/")
            if len(parts) == 3:
                try:
                    dt = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
                except ValueError:
                    pass
        elif isinstance(raw_val, datetime.time):
            # empty time like 00:00:00
            dt = None
            
        if dt:
            if 'dd/MM/yyyy' in switch:
                return dt.strftime('%d/%m/%Y')
            elif '"dd"' in switch or "'dd'" in switch:
                return dt.strftime('%d')
            elif '"MM"' in switch or "'MM'" in switch:
                return dt.strftime('%m')
            elif '"yyyy"' in switch or "'yyyy'" in switch:
                return dt.strftime('%Y')
        else:
            if raw_val is None or str(raw_val).strip() in ["", "00:00:00", "0", "1899-12-30 00:00:00+00:00"]:
                if '"dd"' in switch or '"MM"' in switch:
                    return "....."
                return "..... / ..... / 2026"
            return str(raw_val)

    # General values
    if raw_val is None:
        return ""
    if isinstance(raw_val, datetime.time) and str(raw_val) == "00:00:00":
        return "....."
    if isinstance(raw_val, (datetime.date, datetime.datetime)):
        return raw_val.strftime('%d/%m/%Y')
    if isinstance(raw_val, float) and raw_val.is_integer():
        return str(int(raw_val))
    return str(raw_val)


