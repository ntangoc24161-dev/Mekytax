# -*- coding: utf-8 -*-
"""
Script cap nhat va xuat tron bo Ho so di kiem tra ke hoach (DI-01 den DI-07, TONG HOP HO SO DI.docx va hoso.zip):
- Thiet lap kho giay A4, canh le chuan Nghi dinh 30/2020/ND-CP: Trai 3.0cm, Phai 2.0cm, Tren 2.0cm, Duoi 2.0cm
- Font chu Times New Roman toan bo tai lieu
- Noi dung kiem tra chuan hoa 3 chuyen de:
  + Chi phi quan ly doanh nghiep.
  + Chi phi ban hang.
  + Chi phi tai chinh.
- Can cu phap ly chuan theo quy dinh moi (Luat 108/2025/QH15, QD 1161/QD-CT, QD 1412/QD-DON, QD 1614/QD-TCT, QD 2618/QD-TCS3...)
"""

import sys
import os

# Import directly from builder
from build_ho_so_di_van_phat_thanh import build_all

if __name__ == "__main__":
    build_all()
