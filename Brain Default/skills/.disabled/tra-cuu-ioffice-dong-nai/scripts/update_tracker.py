import datetime
import openpyxl

def main():
    wb_path = r'D:\03. Ung dung\TEST DASHBOARD BAO CAO THEO DOI CONG VAN CONG AN\Quan ly cong van-ok\Bao_cao_cong_van.xlsx'
    wb = openpyxl.load_workbook(wb_path)
    ws = wb['NHAP_HO_SO']

    r = 730
    ws.cell(row=r, column=2).value = '10052/CV-VPCQCSĐT-Đ3'
    ws.cell(row=r, column=3).value = datetime.datetime(2026, 9, 28, 0, 0)
    ws.cell(row=r, column=4).value = 'Cơ quan Cảnh sát điều tra - Công an TP Đồng Nai'
    ws.cell(row=r, column=5).value = None
    ws.cell(row=r, column=6).value = datetime.datetime(2026, 9, 30, 0, 0)
    ws.cell(row=r, column=7).value = 'Cung cấp hồ sơ kê khai thuế (tờ khai GTGT, BC SD HĐ, bảng kê chi tiết mua vào/bán ra), địa chỉ ĐKKD, vi phạm, nhà cung cấp HĐĐT của các Công ty do đương sự Nguyễn Thị Ngọc Bình (SN 1993, CCCD: 075193003154, HKTT P.Long Bình) thành lập phục vụ điều tra tin báo tội phạm "Lừa đảo chiếm đoạt tài sản". Hạn trước 07/10/2026 (ĐTV Lê Duy Hoàng 0813.241.115)'
    ws.cell(row=r, column=8).value = 1
    ws.cell(row=r, column=9).value = '075193003154'
    ws.cell(row=r, column=10).value = 'Đương sự NGUYỄN THỊ NGỌC BÌNH'
    ws.cell(row=r, column=11).value = '102/81/4, tổ 18, KP31, phường Long Bình'
    ws.cell(row=r, column=12).value = 'KT1'
    ws.cell(row=r, column=13).value = None
    ws.cell(row=r, column=14).value = 'Mới'
    ws.cell(row=r, column=15).value = None
    ws.cell(row=r, column=16).value = datetime.datetime(2026, 9, 30, 0, 0)
    ws.cell(row=r, column=17).value = datetime.datetime(2026, 10, 7, 0, 0)
    ws.cell(row=r, column=18).value = 'Không'
    ws.cell(row=r, column=19).value = 'Không'
    ws.cell(row=r, column=20).value = None
    ws.cell(row=r, column=21).value = None
    ws.cell(row=r, column=22).value = 'Rà soát toàn bộ các DN do cá nhân đứng tên/đại diện theo CCCD 075193003154'

    wb.save(wb_path)
    print('Updated Bao_cao_cong_van.xlsx successfully at row 730.')

if __name__ == '__main__':
    main()
