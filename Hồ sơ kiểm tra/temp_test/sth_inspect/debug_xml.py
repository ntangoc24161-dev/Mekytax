# -*- coding: utf-8 -*-
import sys
from lxml import etree
from pathlib import Path
sys.path.insert(0, str(Path(r"D:\SECOND BRAIN\javis-os\brains\Hồ sơ kiểm tra\scripts")))
import cap_nhat_sth_qldn_so_qd as m

xml = m.load_xml(m.TONG_IN)


def valid(label, x):
    try:
        etree.fromstring(x.encode("utf8"))
        print("VALID", label, "len", len(x))
        return True
    except etree.XMLSyntaxError as e:
        print("INVALID", label, e)
        return False


valid("orig", xml)
xml2 = xml
# replay patch_tong in chunks by importing internals - just call stepwise copies
import re, copy

def step(name, fn):
    global xml2
    before = xml2
    xml2 = fn(xml2)
    ok = valid(name, xml2)
    if not ok:
        print("  delta", len(xml2) - len(before))
        sys.exit(1)


# I'll manually step through patch_tong by executing file sections
# Simpler: monkeypatch and split

orig_rep = m.rep
orig_rename = m.rename_qldn
orig_kt = m.replace_kt_fields

calls = []


def tracked_rep(xml, old, new, expected=None, label=""):
    out = orig_rep(xml, old, new, expected, label)
    valid(f"rep:{label}", out)
    return out


m.rep = tracked_rep


def tracked_rename(xml):
    out = orig_rename(xml)
    valid("rename_qldn", out)
    print("  QLDN", out.count("quản lý doanh nghiệp"), "QLKinh", out.count("quản lý kinh doanh"))
    return out


m.rename_qldn = tracked_rename


def tracked_kt(xml):
    out = orig_kt(xml)
    valid("kt_fields", out)
    return out


m.replace_kt_fields = tracked_kt

xml3 = m.patch_tong(xml)
valid("full patch_tong", xml3)
