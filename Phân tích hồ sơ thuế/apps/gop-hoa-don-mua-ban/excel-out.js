(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory(require("./core.js"));
  else root.GopHoaDonExcel = factory(root.GopHoaDon);
})(typeof self !== "undefined" ? self : this, function (G) {
  "use strict";

  function colLetter(n) {
    var s = "";
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function applyStyleRange(ws, r, c1, c2, font, align) {
    for (var c = c1; c <= c2; c++) {
      var cell = ws.getCell(r, c);
      cell.font = font;
      cell.alignment = align;
    }
  }

  function writeSheet(ExcelJS, wb, sheetName, title, company, rows, cols, chieuLabel) {
    var ws = wb.addWorksheet(sheetName, {
      views: [{ state: "frozen", ySplit: 6, xSplit: 0, showGridLines: false }],
      pageSetup: {
        paperSize: 9,
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        blackAndWhite: false,
        horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.6, header: 0.25, footer: 0.25 }
      },
      headerFooter: {
        differentOddEven: true,
        oddFooter: '&L&"Times New Roman,Regular"&9' + G.safeText(company.ten) + " - " + sheetName + '&C&9Trang &P / &N&R&9MST ' + company.mst,
        evenFooter: '&L&"Times New Roman,Regular"&9Trang &P / &N&C&9' + G.safeText(company.ten) + '&R&9' + sheetName
      }
    });
    ws.pageSetup.printTitlesRow = "6:6";

    var nCol = cols.length;
    var navy = "1B365D";
    var teal = "0F6C8C";
    var titleFont = { name: "Times New Roman", size: 16, bold: true, color: { argb: "FF" + navy } };
    var infoFont = { name: "Times New Roman", size: 12, bold: true, color: { argb: "FF1C2833" } };
    var headFont = { name: "Times New Roman", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    var dataFont = { name: "Times New Roman", size: 10, color: { argb: "FF1C2833" } };
    var totFont = { name: "Times New Roman", size: 10, bold: true, color: { argb: "FF1C2833" } };
    var center = { vertical: "middle", horizontal: "center", wrapText: true };
    var left = { vertical: "middle", horizontal: "left", wrapText: true };
    var thin = { style: "thin", color: { argb: "FFBFC5CE" } };
    var border = { top: thin, left: thin, bottom: thin, right: thin };

    ws.mergeCells(1, 1, 1, nCol);
    ws.getCell(1, 1).value = title;
    applyStyleRange(ws, 1, 1, nCol, titleFont, center);
    ws.getRow(1).height = 24;

    ws.mergeCells(2, 1, 2, nCol);
    ws.getCell(2, 1).value = "Đơn vị: " + (company.ten || "");
    applyStyleRange(ws, 2, 1, nCol, infoFont, left);

    ws.mergeCells(3, 1, 3, nCol);
    ws.getCell(3, 1).value = "Mã số thuế: " + (company.mst || "");
    applyStyleRange(ws, 3, 1, nCol, infoFont, left);

    var from = company.from ? G.formatDate(company.from) : "";
    var to = company.to ? G.formatDate(company.to) : "";
    ws.mergeCells(4, 1, 4, nCol);
    ws.getCell(4, 1).value = "Thời gian: " + (from && to ? (from + " - " + to) : "Theo dữ liệu nguồn");
    applyStyleRange(ws, 4, 1, nCol, infoFont, left);

    ws.mergeCells(5, 1, 5, nCol);
    ws.getCell(5, 1).value = "Số hóa đơn " + chieuLabel + ": " + rows.length;
    applyStyleRange(ws, 5, 1, nCol, { name: "Times New Roman", size: 11, italic: true, color: { argb: "FF" + teal } }, left);

    var head = ws.getRow(6);
    head.height = 28;
    cols.forEach(function (col, i) {
      var cell = ws.getCell(6, i + 1);
      cell.value = G.OUT_TITLES_VI[col.key] || col.title;
      cell.font = headFont;
      cell.alignment = center;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + navy } };
      cell.border = border;
      ws.getColumn(i + 1).width = col.width;
    });

    rows.forEach(function (row, idx) {
      var r = 7 + idx;
      var excelRow = ws.getRow(r);
      excelRow.height = 18;
      cols.forEach(function (col, i) {
        var cell = excelRow.getCell(i + 1);
        cell.font = dataFont;
        cell.border = border;
        cell.alignment = col.type === "money" || col.type === "date" || col.type === "stt" || col.type === "rate" ? center : left;
        if (idx % 2 === 1) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4F6F7" } };
        }
        if (col.type === "stt") {
          cell.value = row.stt || (idx + 1);
        } else if (col.type === "date") {
          if (row.ngayLap) {
            cell.value = row.ngayLap;
            cell.numFmt = "dd/mm/yyyy";
          } else cell.value = "";
        } else if (col.type === "money") {
          cell.value = row[col.key] == null ? null : Number(row[col.key]);
          cell.numFmt = "#,##0";
        } else if (col.type === "rate") {
          var rate = row[col.key];
          if (rate == null || rate === "") cell.value = null;
          else {
            cell.value = rate > 1 && rate <= 100 ? (rate / 100) : rate;
            cell.numFmt = "0%";
          }
        } else {
          cell.value = G.safeText(row[col.key] || "");
        }
      });
    });

    var totR = 7 + rows.length;
    var totRow = ws.getRow(totR);
    totRow.height = 20;
    cols.forEach(function (col, i) {
      var cell = totRow.getCell(i + 1);
      cell.font = totFont;
      cell.border = border;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6EAF8" } };
      cell.alignment = center;
      if (i === 0) cell.value = "TỔNG CỘNG";
      else if (col.type === "money" && rows.length) {
        var letter = colLetter(i + 1);
        cell.value = { formula: "SUM(" + letter + "7:" + letter + (totR - 1) + ")" };
        cell.numFmt = "#,##0";
      } else cell.value = "";
    });
    if (nCol > 1) {
      ws.mergeCells(totR, 1, totR, 4);
      ws.getCell(totR, 1).value = "TỔNG CỘNG";
      ws.getCell(totR, 1).font = totFont;
      ws.getCell(totR, 1).alignment = { vertical: "middle", horizontal: "right" };
      ws.getCell(totR, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6EAF8" } };
    }
    ws.getCell(1, 1).font = titleFont;
    ws.getCell(1, 1).alignment = center;
    return ws;
  }

  async function buildWorkbook(ExcelJS, company) {
    var wb = new ExcelJS.Workbook();
    wb.creator = "Mini app gop hoa don";
    wb.created = new Date();
    writeSheet(ExcelJS, wb, "Mua vào", "DANH SÁCH HÓA ĐƠN MUA VÀO", company, company.mua, G.OUT_COLS_MUA, "mua vào");
    writeSheet(ExcelJS, wb, "Bán ra", "DANH SÁCH HÓA ĐƠN BÁN RA", company, company.ban, G.OUT_COLS_BAN, "bán ra");
    return wb.xlsx.writeBuffer();
  }

  return { writeSheet: writeSheet, buildWorkbook: buildWorkbook };
});
