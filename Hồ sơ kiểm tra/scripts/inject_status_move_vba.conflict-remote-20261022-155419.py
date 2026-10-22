# -*- coding: utf-8 -*-
"""Gan VBA vao file JAVIS PRO de tu chuyen dong khi doi trang thai.

Luu y: tranh chuoi tieng Viet trong VBA (COM hay lam vo encoding).
So khop trang thai bang tien to '1.' / '2.'.
"""

from pathlib import Path

import win32com.client as win32

XLSX_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan - JAVIS PRO.xlsx"
XLSM_PATH = r"G:\Other computers\My Laptop\1. CONG TAC KIEM TRA\1- KIEM TRA\1. KIEM TRA TAI TRU SO NNT\00. Theo doi chung\Danh sach kiem tra quyet toan - JAVIS PRO.xlsm"

XL_OPEN_XML_WORKBOOK_MACRO = 52
VB_EXT_CT_STD_MODULE = 1

# Toan bo VBA dung ASCII de tranh mojibake khi AddFromString
VBA_THISWORKBOOK = r'''
Option Explicit

Private Const COL_STATUS As Long = 2
Private Const COL_MST As Long = 4
Private Const COL_TEN As Long = 5
Private Const HEADER_ROW As Long = 4
Private Const COL_LAST As Long = 52
Private Const SHEET_BAN_HANH As String = "01_DA_BAN_HANH"
Private Const SHEET_DANG_THUC_HIEN As String = "02_DANG_THUC_HIEN"

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim area As Range
    Dim statusCell As Range
    Dim newStatus As String
    Dim srcRow As Long

    On Error GoTo CleanFail

    If Sh.Name <> SHEET_BAN_HANH And Sh.Name <> SHEET_DANG_THUC_HIEN Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge > 50 Then Exit Sub

    Set area = Intersect(Target, Sh.Columns(COL_STATUS))
    If area Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each statusCell In area.Cells
        srcRow = statusCell.Row
        If srcRow <= HEADER_ROW Then GoTo NextCell
        If Trim$(CStr(Sh.Cells(srcRow, COL_MST).Value & "")) = "" And _
           Trim$(CStr(Sh.Cells(srcRow, COL_TEN).Value & "")) = "" Then GoTo NextCell

        newStatus = Trim$(CStr(statusCell.Value & ""))
        Call JavisStatusMove.MoveRowByStatus(Sh.Name, srcRow, newStatus, False)

NextCell:
    Next statusCell

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.StatusBar = "Javis move status error: " & Err.Description
    Resume CleanExit
End Sub
'''

VBA_MODULE = r'''
Option Explicit

Public Const JAVIS_SHEET_BAN_HANH As String = "01_DA_BAN_HANH"
Public Const JAVIS_SHEET_DANG_THUC_HIEN As String = "02_DANG_THUC_HIEN"

Public Function IsBanHanh(ByVal statusText As String) As Boolean
    IsBanHanh = (Left$(Trim$(statusText), 2) = "1.")
End Function

Public Function IsDangThucHien(ByVal statusText As String) As Boolean
    IsDangThucHien = (Left$(Trim$(statusText), 2) = "2.")
End Function

Public Sub MoveRowByStatus(ByVal sheetName As String, ByVal srcRow As Long, ByVal newStatus As String, Optional ByVal setStatus As Boolean = True)
    Dim sh As Worksheet
    Dim destName As String
    Dim destWs As Worksheet
    Dim destRow As Long
    Dim c As Long
    Dim eventsWas As Boolean

    Set sh = ThisWorkbook.Worksheets(sheetName)
    eventsWas = Application.EnableEvents
    Application.EnableEvents = False

    If setStatus Then
        sh.Cells(srcRow, 2).Value = newStatus
    End If

    destName = ""
    If sheetName = JAVIS_SHEET_DANG_THUC_HIEN And IsBanHanh(newStatus) Then
        destName = JAVIS_SHEET_BAN_HANH
    ElseIf sheetName = JAVIS_SHEET_BAN_HANH And IsDangThucHien(newStatus) Then
        destName = JAVIS_SHEET_DANG_THUC_HIEN
    End If

    If destName = "" Then
        Application.EnableEvents = eventsWas
        Exit Sub
    End If

    If Trim$(CStr(sh.Cells(srcRow, 4).Value & "")) = "" And Trim$(CStr(sh.Cells(srcRow, 5).Value & "")) = "" Then
        Application.EnableEvents = eventsWas
        Exit Sub
    End If

    Set destWs = ThisWorkbook.Worksheets(destName)
    destRow = NextEmptyDataRow(destWs)

    For c = 1 To 52
        destWs.Cells(destRow, c).Value = sh.Cells(srcRow, c).Value
        destWs.Cells(destRow, c).NumberFormat = sh.Cells(srcRow, c).NumberFormat
    Next c

    destWs.Cells(destRow, 2).Value = newStatus
    destWs.Cells(destRow, 1).Formula = "=IF(D" & destRow & "<>"""",ROW()-4,"""")"
    destWs.Cells(destRow, 8).Formula = "=IFERROR(VLOOKUP(G" & destRow & ",DANH_MUC!$A$3:$B$15,2,FALSE),"""")"
    destWs.Cells(destRow, 42).Formula = "=IF(OR(AN" & destRow & "<>"""",AO" & destRow & "<>""""),AN" & destRow & "+AO" & destRow & ","""")"
    destWs.Cells(destRow, 45).Formula = "=IF(OR(AQ" & destRow & "<>"""",AR" & destRow & "<>""""),AQ" & destRow & "+AR" & destRow & ","""")"
    destWs.Cells(destRow, 48).Formula = "=IF(OR(AT" & destRow & "<>"""",AU" & destRow & "<>""""),AT" & destRow & "+AU" & destRow & ","""")"
    destWs.Cells(destRow, 49).Formula = "=IF(OR(AP" & destRow & "<>"""",AS" & destRow & "<>"""",AV" & destRow & "<>""""),IF(AP" & destRow & "="""",0,AP" & destRow & ")+IF(AS" & destRow & "="""",0,AS" & destRow & ")+IF(AV" & destRow & "="""",0,AV" & destRow & "),"""")"

    sh.Rows(srcRow).Delete
    Application.StatusBar = "Javis: moved to [" & destName & "] row " & destRow
    Application.EnableEvents = eventsWas
End Sub

Private Function NextEmptyDataRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    If lastRow < 5 Then
        NextEmptyDataRow = 5
        Exit Function
    End If
    For r = 5 To lastRow + 50
        If Trim$(CStr(ws.Cells(r, 4).Value & "")) = "" And Trim$(CStr(ws.Cells(r, 5).Value & "")) = "" Then
            NextEmptyDataRow = r
            Exit Function
        End If
    Next r
    NextEmptyDataRow = lastRow + 1
End Function

Public Sub TestMoveRow6DthToBanHanh()
    Dim st As String
    st = "1. " & ChrW(272) & "ã ban hành"
    ' ChrW(272) = Đ
    Call MoveRowByStatus(JAVIS_SHEET_DANG_THUC_HIEN, 6, st, True)
End Sub

Public Sub TestMoveLastBhToDth()
    Dim ws As Worksheet
    Dim r As Long
    Dim last As Long
    Dim st As String
    Set ws = ThisWorkbook.Worksheets(JAVIS_SHEET_BAN_HANH)
    last = ws.Cells(ws.Rows.Count, 5).End(xlUp).Row
    st = "2. " & "Dang thuc hien"
    ' Build proper Vietnamese with ChrW
    st = "2. " & ChrW(272) & "ang th" & ChrW(7921) & "c hi" & ChrW(7879) & "n"
    Call MoveRowByStatus(JAVIS_SHEET_BAN_HANH, last, st, True)
End Sub
'''


def update_dashboard_note(wb):
    ws = wb.Worksheets("00_DASHBOARD")
    # Excel COM nhan Unicode tot hon AddFromString VBA
    ws.Range("B24").Value = (
        "4. TỰ CHUYỂN DÒNG KHI ĐỔI TRẠNG THÁI: tại sheet 02_DANG_THUC_HIEN, "
        'chọn trạng thái "1. Đã ban hành" sẽ tự chuyển hồ sơ sang 01_DA_BAN_HANH '
        "(chiều ngược lại cũng vậy). Mở file .xlsm và bật Enable Content / macro."
    )
    ws.Range("B24").Font.Name = "Times New Roman"
    ws.Range("B24").Font.Size = 10
    try:
        ws.Range("B24:J24").Merge()
    except Exception:
        pass


def clear_component_code(comp):
    mod = comp.CodeModule
    if mod.CountOfLines > 0:
        mod.DeleteLines(1, mod.CountOfLines)


def ensure_std_module(wb, name):
    for i in range(1, wb.VBProject.VBComponents.Count + 1):
        comp = wb.VBProject.VBComponents.Item(i)
        if comp.Name == name:
            return comp
    comp = wb.VBProject.VBComponents.Add(VB_EXT_CT_STD_MODULE)
    comp.Name = name
    return comp


def inject():
    xlsx = Path(XLSX_PATH)
    xlsm = Path(XLSM_PATH)
    if not xlsx.exists():
        raise FileNotFoundError(f"Khong thay file nguon: {xlsx}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    try:
        _ = excel.VBE
    except Exception as exc:
        excel.Quit()
        raise RuntimeError(
            "Excel chua cho phep truy cap VBA project. "
            "File > Options > Trust Center > Macro Settings > "
            "bat 'Trust access to the VBA project object model'."
        ) from exc

    wb = excel.Workbooks.Open(str(xlsx))
    try:
        update_dashboard_note(wb)

        this_wb = wb.VBProject.VBComponents("ThisWorkbook")
        clear_component_code(this_wb)
        this_wb.CodeModule.AddFromString(VBA_THISWORKBOOK)

        # remove old module if any then recreate clean
        for i in range(wb.VBProject.VBComponents.Count, 0, -1):
            comp = wb.VBProject.VBComponents.Item(i)
            if comp.Name == "JavisStatusMove":
                wb.VBProject.VBComponents.Remove(comp)
                break
        mod = ensure_std_module(wb, "JavisStatusMove")
        clear_component_code(mod)
        mod.CodeModule.AddFromString(VBA_MODULE)

        if xlsm.exists():
            try:
                xlsm.unlink()
            except PermissionError:
                pass
        wb.SaveAs(str(xlsm), FileFormat=XL_OPEN_XML_WORKBOOK_MACRO)
        print(f"Saved macro workbook: {xlsm}")
    finally:
        wb.Close(SaveChanges=False)

    wb2 = excel.Workbooks.Open(str(xlsx))
    try:
        update_dashboard_note(wb2)
        wb2.Save()
        print(f"Updated note on: {xlsx}")
    finally:
        wb2.Close(SaveChanges=False)

    excel.Quit()


def run_roundtrip_test():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass
    excel.EnableEvents = True

    wb = excel.Workbooks.Open(str(XLSM_PATH))
    try:
        ws1 = wb.Worksheets("01_DA_BAN_HANH")
        ws2 = wb.Worksheets("02_DANG_THUC_HIEN")

        def count(ws):
            n = 0
            for r in range(5, 60):
                if ws.Cells(r, 5).Value:
                    n += 1
            return n

        before = (count(ws1), count(ws2))
        print("Before:", before)

        # Direct call with ASCII-safe status using ChrW-built string from VBA test sub
        excel.Run("JavisStatusMove.TestMoveRow6DthToBanHanh")
        mid = (count(ws1), count(ws2))
        print("After move to BH:", mid)

        excel.Run("JavisStatusMove.TestMoveLastBhToDth")
        after = (count(ws1), count(ws2))
        print("After restore to DTH:", after)

        ok = before == after and mid == (before[0] + 1, before[1] - 1)
        print("ROUNDTRIP_OK" if ok else "ROUNDTRIP_FAIL")

        # Also verify SheetChange path: change cell status with events on
        # Find a DTH row and set status via cell (user path)
        # Use MoveRowByStatus directly from Python-constructed unicode status
        status_ban = "1. Đã ban hành"
        status_dth = "2. Đang thực hiện"
        # move first DTH via Run with args
        # Excel Run can pass args:
        r = None
        for i in range(5, 50):
            if ws2.Cells(i, 5).Value:
                r = i
                break
        if r:
            excel.Run("JavisStatusMove.MoveRowByStatus", "02_DANG_THUC_HIEN", r, status_ban, True)
            print("Python-arg move to BH counts:", count(ws1), count(ws2))
            last = ws1.Cells(ws1.Rows.Count, 5).End(3).Row  # xlUp=3
            excel.Run("JavisStatusMove.MoveRowByStatus", "01_DA_BAN_HANH", last, status_dth, True)
            print("Python-arg restore counts:", count(ws1), count(ws2))

        wb.Save()
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        run_roundtrip_test()
    else:
        inject()
        run_roundtrip_test()
