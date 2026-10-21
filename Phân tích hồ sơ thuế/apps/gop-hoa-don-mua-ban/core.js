/* Gop hoa don mua vao / ban ra - logic thuan, dung chung trinh duyet va Node. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.GopHoaDon = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var CANON = [
    { key: "sttGoc", aliases: ["stt", "so thu tu"] },
    { key: "mauHd", aliases: ["ky hieu mau", "ky hieu mau so", "mau so", "mau so hd", "mau so hoa don", "ky hieu mau hoa don"] },
    { key: "kyHieu", aliases: ["ky hieu", "ky hieu hd", "ky hieu hoa don", "seri", "seri hd", "series"] },
    { key: "soHd", aliases: ["so hoa don", "so hd", "so h d", "so chung tu", "so ct"] },
    { key: "ngayLap", aliases: ["ngay lap", "ngay hd", "ngay hoa don", "ngay lap hoa don", "ngay phat hanh", "ngay ct", "ngay chung tu", "ngay"] },
    { key: "mstNguoiBan", aliases: ["mst nguoi ban", "ma so thue nguoi ban", "mst ben ban", "mst ncc", "mst nha cung cap", "mst nguoi xuat", "mst nguoi xuat hang", "ma so thue nnt ban"] },
    { key: "tenNguoiBan", aliases: ["ten nguoi ban", "nguoi ban", "ten ben ban", "ten ncc", "ten nha cung cap", "ten nguoi xuat", "ten nguoi xuat hang", "ten nnt ban"] },
    { key: "diaChiNguoiBan", aliases: ["dia chi nguoi ban", "dia chi ben ban", "dia chi ncc", "dia chi nguoi xuat"] },
    { key: "mstNguoiMua", aliases: ["mst nguoi mua", "ma so thue nguoi mua", "mst ben mua", "mst khach hang", "mst nguoi nhan", "mst nguoi nhan hang", "ma so thue nnt mua"] },
    { key: "tenNguoiMua", aliases: ["ten nguoi mua", "nguoi mua", "ten ben mua", "ten khach hang", "ten nguoi nhan", "ten nguoi nhan hang", "ten nnt mua"] },
    { key: "diaChiNguoiMua", aliases: ["dia chi nguoi mua", "dia chi ben mua", "dia chi khach hang", "dia chi nguoi nhan"] },
    { key: "mstDoiTac", aliases: ["mst doi tac", "ma so thue", "mst"] },
    { key: "tenDoiTac", aliases: ["ten doi tac", "ten don vi", "ten nnt"] },
    { key: "matHang", aliases: ["mat hang", "ten hang", "ten hang hoa", "ten hang hoa dich vu", "noi dung", "dien giai"] },
    { key: "doanhSo", aliases: ["doanh so", "doanh so chua thue", "tong tien chua thue", "tien hang", "gia tri chua thue", "doanh thu", "gia tri"] },
    { key: "thueSuat", aliases: ["thue suat", "ts gtgt", "thue suat gtgt"] },
    { key: "tienThue", aliases: ["thue gtgt", "tien thue", "tong tien thue", "tien thue gtgt"] },
    { key: "tongTien", aliases: ["tong tien thanh toan", "tong thanh toan", "tong tien", "tong cong thanh toan"] },
    { key: "tienTe", aliases: ["don vi tien", "dong tien", "loai tien", "tien te"] },
    { key: "tinhTrang", aliases: ["tinh trang", "tinh trang hoa don", "trang thai", "trang thai hoa don"] }
  ];

  var OUT_COLS_MUA = [
    { key: "stt", title: "STT", width: 6, type: "stt" },
    { key: "mauHd", title: "Ky hieu mau HD", width: 14, type: "text" },
    { key: "kyHieu", title: "Ky hieu HD", width: 14, type: "text" },
    { key: "soHd", title: "So hoa don", width: 14, type: "text" },
    { key: "ngayLap", title: "Ngay lap", width: 14, type: "date" },
    { key: "mstNguoiBan", title: "MST nguoi ban", width: 16, type: "text" },
    { key: "tenNguoiBan", title: "Ten nguoi ban", width: 42, type: "text" },
    { key: "diaChiNguoiBan", title: "Dia chi nguoi ban", width: 36, type: "text" },
    { key: "mstNguoiMua", title: "MST nguoi mua", width: 16, type: "text" },
    { key: "tenNguoiMua", title: "Ten nguoi mua", width: 42, type: "text" },
    { key: "matHang", title: "Mat hang", width: 28, type: "text" },
    { key: "doanhSo", title: "Tien chua thue", width: 16, type: "money" },
    { key: "thueSuat", title: "Thue suat", width: 12, type: "rate" },
    { key: "tienThue", title: "Tien thue GTGT", width: 16, type: "money" },
    { key: "tongTien", title: "Tong thanh toan", width: 16, type: "money" },
    { key: "tienTe", title: "Don vi tien", width: 12, type: "text" },
    { key: "tinhTrang", title: "Tinh trang HD", width: 18, type: "text" },
    { key: "fileNguon", title: "File nguon", width: 28, type: "text" }
  ];

  var OUT_COLS_BAN = [
    { key: "stt", title: "STT", width: 6, type: "stt" },
    { key: "mauHd", title: "Ky hieu mau HD", width: 14, type: "text" },
    { key: "kyHieu", title: "Ky hieu HD", width: 14, type: "text" },
    { key: "soHd", title: "So hoa don", width: 14, type: "text" },
    { key: "ngayLap", title: "Ngay lap", width: 14, type: "date" },
    { key: "mstNguoiMua", title: "MST nguoi mua", width: 16, type: "text" },
    { key: "tenNguoiMua", title: "Ten nguoi mua", width: 42, type: "text" },
    { key: "diaChiNguoiMua", title: "Dia chi nguoi mua", width: 36, type: "text" },
    { key: "mstNguoiBan", title: "MST nguoi ban", width: 16, type: "text" },
    { key: "tenNguoiBan", title: "Ten nguoi ban", width: 42, type: "text" },
    { key: "matHang", title: "Mat hang", width: 28, type: "text" },
    { key: "doanhSo", title: "Tien chua thue", width: 16, type: "money" },
    { key: "thueSuat", title: "Thue suat", width: 12, type: "rate" },
    { key: "tienThue", title: "Tien thue GTGT", width: 16, type: "money" },
    { key: "tongTien", title: "Tong thanh toan", width: 16, type: "money" },
    { key: "tienTe", title: "Don vi tien", width: 12, type: "text" },
    { key: "tinhTrang", title: "Tinh trang HD", width: 18, type: "text" },
    { key: "fileNguon", title: "File nguon", width: 28, type: "text" }
  ];

  var OUT_TITLES_VI = {
    stt: "STT",
    mauHd: "Ký hiệu mẫu HĐ",
    kyHieu: "Ký hiệu HĐ",
    soHd: "Số hóa đơn",
    ngayLap: "Ngày lập",
    mstNguoiBan: "MST người bán",
    tenNguoiBan: "Tên người bán",
    diaChiNguoiBan: "Địa chỉ người bán",
    mstNguoiMua: "MST người mua",
    tenNguoiMua: "Tên người mua",
    diaChiNguoiMua: "Địa chỉ người mua",
    matHang: "Mặt hàng",
    doanhSo: "Tiền chưa thuế",
    thueSuat: "Thuế suất",
    tienThue: "Tiền thuế GTGT",
    tongTien: "Tổng thanh toán",
    tienTe: "Đơn vị tiền",
    tinhTrang: "Tình trạng HĐ",
    fileNguon: "File nguồn"
  };

  function stripVi(s) {
    return String(s == null ? "" : s)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .replace(/Đ/g, "D")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function cellText(v) {
    if (v == null || v === "") return "";
    if (v instanceof Date) return v;
    if (typeof v === "object") {
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join("");
      if (v.text != null && v.hyperlink) return v.text;
      if (v.text != null && v.result == null && !v.formula) return v.text;
      if (v.result != null) return cellText(v.result);
      if (v.formula) return "";
      if (v.error) return "";
    }
    return v;
  }

  function asString(v) {
    var t = cellText(v);
    if (t instanceof Date) return formatDate(t);
    if (t == null) return "";
    var s = String(t).trim();
    if (/^\d+\.0$/.test(s)) s = s.slice(0, -2);
    return s;
  }

  function safeText(s) {
    s = String(s == null ? "" : s);
    if (!s) return "";
    if (s.charAt(0) === "=") return "'" + s;
    return s;
  }

  function digitsMst(s) {
    var raw = String(s == null ? "" : s).trim().replace(/\s+/g, "");
    raw = raw.replace(/\.0$/, "");
    var m = raw.match(/^(\d{8,14})(?:-\d{1,6})?/);
    return m ? m[0] : "";
  }

  function toNumber(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number" && isFinite(v)) return v;
    var t = cellText(v);
    if (t instanceof Date) return null;
    if (typeof t === "number" && isFinite(t)) return t;
    var s = String(t).trim().replace(/\s/g, "").replace(/%$/, "");
    if (!s) return null;
    if (s.indexOf(",") >= 0 && s.indexOf(".") >= 0) {
      if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
      else s = s.replace(/,/g, "");
    } else if (s.indexOf(",") >= 0 && s.indexOf(".") < 0) {
      s = s.replace(",", ".");
    }
    var n = parseFloat(s);
    return isFinite(n) ? n : null;
  }

  function excelSerialToDate(n) {
    if (typeof n !== "number" || !isFinite(n) || n < 20000 || n > 80000) return null;
    var utc = Date.UTC(1899, 11, 30) + Math.round(n) * 86400000;
    return new Date(utc);
  }

  function parseDate(v) {
    if (v == null || v === "") return null;
    var t = cellText(v);
    if (t instanceof Date && !isNaN(t.getTime())) {
      return new Date(Date.UTC(t.getFullYear(), t.getMonth(), t.getDate()));
    }
    if (typeof t === "number") return excelSerialToDate(t);
    var s = String(t).trim();
    var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return null;
  }

  function formatDate(d) {
    if (!d) return "";
    var dd = String(d.getUTCDate()).padStart(2, "0");
    var mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    return dd + "/" + mm + "/" + d.getUTCFullYear();
  }

  function padNum(n, w) {
    var s = String(n == null ? "" : n);
    while (s.length < w) s = "0" + s;
    return s;
  }

  function naturalSoHd(s) {
    s = String(s || "").trim();
    var m = s.match(/(\d+)\s*$/);
    return m ? padNum(m[1], 12) : s.toLowerCase();
  }

  function hintChieu(name) {
    var n = stripVi(name);
    if (/\b(mua vao|mua vao|bang ke mua|hoa don mua|hd mua|mv)\b/.test(n) || /(^|[^a-z])(mua|mv)([^a-z]|$)/.test(n)) {
      if (!/\bban ra\b/.test(n)) return "mua";
    }
    if (/\b(ban ra|bang ke ban|hoa don ban|hd ban|br)\b/.test(n) || /(^|[^a-z])(ban|br)([^a-z]|$)/.test(n)) {
      if (!/\bmua vao\b/.test(n)) return "ban";
    }
    if (/vaoq|mua vao|muavao/.test(n)) return "mua";
    if (/raq|ban ra|banra/.test(n)) return "ban";
    return "";
  }

  function scoreHeaderCell(norm) {
    if (!norm) return 0;
    var keys = [
      "so hoa don", "so hd", "ngay lap", "ky hieu", "ma so thue", "mst",
      "nguoi ban", "nguoi mua", "doanh so", "thue gtgt", "thue suat",
      "tong tien", "tinh trang", "mat hang", "seri", "stt"
    ];
    var s = 0;
    for (var i = 0; i < keys.length; i++) if (norm.indexOf(keys[i]) >= 0) s++;
    return s;
  }

  function findHeader(rows) {
    var best = { idx: -1, score: 0 };
    var limit = Math.min(rows.length, 25);
    for (var i = 0; i < limit; i++) {
      var row = rows[i] || [];
      var score = 0;
      for (var j = 0; j < row.length; j++) score += scoreHeaderCell(stripVi(asString(row[j])));
      if (score > best.score) best = { idx: i, score: score };
    }
    if (best.score < 3) return null;
    var header = (rows[best.idx] || []).map(function (c) { return asString(c); });
    var next = rows[best.idx + 1] || [];
    var nextScore = 0;
    for (var k = 0; k < next.length; k++) nextScore += scoreHeaderCell(stripVi(asString(next[k])));
    var nextHasDataDate = parseDate(next[1]) || parseDate(next[2]) || parseDate(next[3]);
    if (nextScore >= 2 && !nextHasDataDate) {
      var width = Math.max(header.length, next.length);
      var merged = [];
      for (var c = 0; c < width; c++) {
        var a = asString(header[c]);
        var b = asString(next[c]);
        merged[c] = a && b && stripVi(a) !== stripVi(b) ? (a + " " + b) : (a || b);
      }
      return { idx: best.idx, header: merged, dataStart: best.idx + 2 };
    }
    return { idx: best.idx, header: header, dataStart: best.idx + 1 };
  }

  function mapHeader(headerCells) {
    var colMap = {};
    var used = {};
    for (var i = 0; i < headerCells.length; i++) {
      var norm = stripVi(headerCells[i]);
      if (!norm) continue;
      var hit = null;
      var hitLen = 0;
      for (var c = 0; c < CANON.length; c++) {
        var spec = CANON[c];
        if (used[spec.key]) continue;
        for (var a = 0; a < spec.aliases.length; a++) {
          var al = spec.aliases[a];
          if (norm === al || norm.indexOf(al) >= 0) {
            if (al.length > hitLen) {
              hit = spec.key;
              hitLen = al.length;
            }
          }
        }
      }
      if (hit) {
        colMap[hit] = i;
        used[hit] = true;
      }
    }
    return colMap;
  }

  function rowHasInvoice(mapped) {
    return !!(mapped.soHd || mapped.ngayLap || mapped.mstNguoiBan || mapped.mstNguoiMua || mapped.mstDoiTac);
  }

  function isTongCong(rawRow, mapped) {
    var tenBan = stripVi(mapped.tenNguoiBan || "");
    var tenMua = stripVi(mapped.tenNguoiMua || "");
    var tenDt = stripVi(mapped.tenDoiTac || "");
    var ten = (tenBan + " " + tenMua + " " + tenDt).trim();
    if (/\btong cong ty\b/.test(ten)) return false;
    var so = String(mapped.soHd || "").trim();
    var mst = mapped.mstNguoiBan || mapped.mstNguoiMua || mapped.mstDoiTac || "";
    if (/^tong cong$/.test(tenBan) || /^tong cong$/.test(tenMua) || /^tong cong$/.test(tenDt)) return true;
    if (/(^|\s)tong cong(\s|$)/.test(ten) && !so) return true;
    var firstCells = [];
    for (var i = 0; i < Math.min(8, (rawRow || []).length); i++) firstCells.push(stripVi(asString(rawRow[i])));
    var blob = firstCells.join(" | ");
    if (/(^|\s|\|)tong cong(\s|\||$)/.test(blob) || /(^|\s|\|)cong ky(\s|\||$)/.test(blob)) return true;
    var tenAny = (mapped.tenNguoiBan || mapped.tenNguoiMua || mapped.tenDoiTac || "").trim();
    if (!so && !mst && !tenAny && (mapped.doanhSo || mapped.tienThue || mapped.tongTien)) return true;
    return false;
  }

  function pick(row, colMap, key) {
    if (colMap[key] == null) return "";
    return row[colMap[key]];
  }

  function mapRow(row, colMap, meta) {
    var ngay = parseDate(pick(row, colMap, "ngayLap"));
    var out = {
      mauHd: asString(pick(row, colMap, "mauHd")),
      kyHieu: asString(pick(row, colMap, "kyHieu")),
      soHd: asString(pick(row, colMap, "soHd")),
      ngayLap: ngay,
      mstNguoiBan: digitsMst(pick(row, colMap, "mstNguoiBan")),
      tenNguoiBan: asString(pick(row, colMap, "tenNguoiBan")),
      diaChiNguoiBan: asString(pick(row, colMap, "diaChiNguoiBan")),
      mstNguoiMua: digitsMst(pick(row, colMap, "mstNguoiMua")),
      tenNguoiMua: asString(pick(row, colMap, "tenNguoiMua")),
      diaChiNguoiMua: asString(pick(row, colMap, "diaChiNguoiMua")),
      mstDoiTac: digitsMst(pick(row, colMap, "mstDoiTac")),
      tenDoiTac: asString(pick(row, colMap, "tenDoiTac")),
      matHang: asString(pick(row, colMap, "matHang")),
      doanhSo: toNumber(pick(row, colMap, "doanhSo")),
      thueSuat: toNumber(pick(row, colMap, "thueSuat")),
      tienThue: toNumber(pick(row, colMap, "tienThue")),
      tongTien: toNumber(pick(row, colMap, "tongTien")),
      tienTe: asString(pick(row, colMap, "tienTe")) || "VND",
      tinhTrang: asString(pick(row, colMap, "tinhTrang")),
      fileNguon: meta.fileName,
      sheetNguon: meta.sheetName,
      chieuHint: meta.chieuHint
    };
    if (!out.mstNguoiBan && out.mstDoiTac && (meta.chieuHint === "mua" || (!out.mstNguoiMua && meta.chieuHint !== "ban"))) {
      out.mstNguoiBan = out.mstDoiTac;
      if (!out.tenNguoiBan) out.tenNguoiBan = out.tenDoiTac;
    }
    if (!out.mstNguoiMua && out.mstDoiTac && meta.chieuHint === "ban") {
      out.mstNguoiMua = out.mstDoiTac;
      if (!out.tenNguoiMua) out.tenNguoiMua = out.tenDoiTac;
    }
    return out;
  }

  function countMap(values) {
    var m = {};
    for (var i = 0; i < values.length; i++) {
      var k = values[i];
      if (!k) continue;
      m[k] = (m[k] || 0) + 1;
    }
    return m;
  }

  function argmax(map) {
    var bestK = "", bestV = 0;
    Object.keys(map).forEach(function (k) {
      if (map[k] > bestV) {
        bestK = k;
        bestV = map[k];
      }
    });
    return { key: bestK, n: bestV };
  }

  function detectOwner(rows) {
    var mua = countMap(rows.map(function (r) { return r.mstNguoiMua; }));
    var ban = countMap(rows.map(function (r) { return r.mstNguoiBan; }));
    var topMua = argmax(mua);
    var topBan = argmax(ban);
    var n = rows.length || 1;
    var muaShare = topMua.n / n;
    var banShare = topBan.n / n;
    var hint = "";
    for (var h = 0; h < rows.length; h++) {
      if (rows[h].chieuHint) { hint = rows[h].chieuHint; break; }
    }
    if (topMua.n === 0 && topBan.n === 0) return { mst: "", chieu: hint, ten: "", weak: true };
    if (muaShare < 0.5 && banShare < 0.5) {
      return { mst: "", chieu: hint, ten: "", weak: true, nMua: topMua.n, nBan: topBan.n };
    }
    var ownerMst;
    var chieu;
    if (muaShare >= banShare) {
      ownerMst = topMua.key;
      chieu = "mua";
    } else {
      ownerMst = topBan.key;
      chieu = "ban";
    }
    var ten = "";
    for (var i = 0; i < rows.length; i++) {
      if (chieu === "mua" && rows[i].mstNguoiMua === ownerMst && rows[i].tenNguoiMua) {
        ten = rows[i].tenNguoiMua;
        break;
      }
      if (chieu === "ban" && rows[i].mstNguoiBan === ownerMst && rows[i].tenNguoiBan) {
        ten = rows[i].tenNguoiBan;
        break;
      }
    }
    return { mst: ownerMst, chieu: chieu, ten: ten, nMua: topMua.n, nBan: topBan.n, weak: false };
  }

  function classifyRow(row, ownerMst, fileChieu) {
    if (ownerMst) {
      var isMua = row.mstNguoiMua === ownerMst;
      var isBan = row.mstNguoiBan === ownerMst;
      if (isMua && !isBan) return "mua";
      if (isBan && !isMua) return "ban";
      if (isMua && isBan) return fileChieu || row.chieuHint || "ban";
    }
    if (row.chieuHint) return row.chieuHint;
    if (fileChieu) return fileChieu;
    return "";
  }

  function fillCounterparty(row, chieu, ownerMst, ownerTen) {
    if (chieu === "mua") {
      if (ownerMst && !row.mstNguoiMua) row.mstNguoiMua = ownerMst;
      if (ownerTen && !row.tenNguoiMua) row.tenNguoiMua = ownerTen;
      if (!row.mstNguoiBan && row.mstDoiTac) row.mstNguoiBan = row.mstDoiTac;
      if (!row.tenNguoiBan && row.tenDoiTac) row.tenNguoiBan = row.tenDoiTac;
    }
    if (chieu === "ban") {
      if (ownerMst && !row.mstNguoiBan) row.mstNguoiBan = ownerMst;
      if (ownerTen && !row.tenNguoiBan) row.tenNguoiBan = ownerTen;
      if (!row.mstNguoiMua && row.mstDoiTac) row.mstNguoiMua = row.mstDoiTac;
      if (!row.tenNguoiMua && row.tenDoiTac) row.tenNguoiMua = row.tenDoiTac;
    }
    return row;
  }

  function skipSheetName(name) {
    var n = stripVi(name);
    return n === "ptdn" || n.indexOf("huong dan") >= 0 || n.indexOf("sheet") === 0 && n.length < 8;
  }

  function ingestSheet(rows, fileName, sheetName, opts) {
    opts = opts || {};
    var found = findHeader(rows);
    if (!found) {
      return { rows: [], warning: "Khong tim thay hang tieu de cot trong sheet " + sheetName + " (" + fileName + ")" };
    }
    var colMap = mapHeader(found.header);
    if (colMap.soHd == null && colMap.ngayLap == null) {
      return { rows: [], warning: "Sheet " + sheetName + " thieu cot So hoa don / Ngay lap (" + fileName + ")" };
    }
    var chieuHint = hintChieu(sheetName) || hintChieu(fileName);
    var out = [];
    var droppedTong = 0;
    var droppedEmpty = 0;
    for (var i = found.dataStart; i < rows.length; i++) {
      var raw = rows[i] || [];
      var empty = true;
      for (var c = 0; c < raw.length; c++) {
        if (asString(raw[c])) { empty = false; break; }
      }
      if (empty) continue;
      var mapped = mapRow(raw, colMap, { fileName: fileName, sheetName: sheetName, chieuHint: chieuHint });
      if (opts.locTongCong !== false && isTongCong(raw, mapped)) {
        droppedTong++;
        continue;
      }
      if (!rowHasInvoice(mapped)) {
        droppedEmpty++;
        continue;
      }
      if (!mapped.soHd && !mapped.ngayLap) {
        droppedEmpty++;
        continue;
      }
      out.push(mapped);
    }
    return { rows: out, droppedTong: droppedTong, droppedEmpty: droppedEmpty, chieuHint: chieuHint, colMap: colMap };
  }

  function dedupeKey(r) {
    return [
      r.soHd || "",
      r.kyHieu || "",
      r.ngayLap ? formatDate(r.ngayLap) : "",
      r.mstNguoiBan || "",
      r.mstNguoiMua || "",
      r.doanhSo == null ? "" : String(r.doanhSo)
    ].join("|");
  }

  function sortRows(rows) {
    rows.sort(function (a, b) {
      var da = a.ngayLap ? a.ngayLap.getTime() : 0;
      var db = b.ngayLap ? b.ngayLap.getTime() : 0;
      if (da !== db) return da - db;
      var ka = String(a.kyHieu || "").toLowerCase();
      var kb = String(b.kyHieu || "").toLowerCase();
      if (ka !== kb) return ka < kb ? -1 : 1;
      var sa = naturalSoHd(a.soHd);
      var sb = naturalSoHd(b.soHd);
      if (sa !== sb) return sa < sb ? -1 : 1;
      return 0;
    });
    return rows;
  }

  function dateRange(rows) {
    var min = null, max = null;
    for (var i = 0; i < rows.length; i++) {
      var d = rows[i].ngayLap;
      if (!d) continue;
      if (!min || d < min) min = d;
      if (!max || d > max) max = d;
    }
    return { from: min, to: max };
  }

  function sumField(rows, key) {
    var t = 0;
    var any = false;
    for (var i = 0; i < rows.length; i++) {
      var n = rows[i][key];
      if (n == null || n === "") continue;
      t += Number(n) || 0;
      any = true;
    }
    return any ? t : 0;
  }

  function slugTen(s) {
    var t = String(s || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/đ/g, "d")
      .replace(/Đ/g, "D")
      .replace(/[^A-Za-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    if (t.length > 40) t = t.slice(0, 40);
    return t || "Cong_ty";
  }

  function buildCompanies(allRows, opts) {
    opts = opts || {};
    var forced = digitsMst(opts.mst || "");
    var forcedTen = (opts.ten || "").trim();
    var groups = {};

    function ensure(mst, ten) {
      if (!groups[mst]) groups[mst] = { mst: mst, ten: ten || "", mua: [], ban: [], unknown: [] };
      if (ten && !groups[mst].ten) groups[mst].ten = ten;
      return groups[mst];
    }

    if (forced) {
      var g = ensure(forced, forcedTen);
      allRows.forEach(function (row) {
        var chieu = classifyRow(row, forced, row.chieuHint);
        fillCounterparty(row, chieu, forced, forcedTen || g.ten);
        if (chieu === "mua") g.mua.push(row);
        else if (chieu === "ban") g.ban.push(row);
        else g.unknown.push(row);
      });
    } else {
      var byFile = {};
      allRows.forEach(function (row) {
        var k = row.fileNguon || "_";
        (byFile[k] = byFile[k] || []).push(row);
      });
      Object.keys(byFile).forEach(function (fn) {
        var owner = detectOwner(byFile[fn]);
        var mst = owner.mst || "CHUA_RO_MST";
        var g2 = ensure(mst, owner.ten || forcedTen);
        byFile[fn].forEach(function (row) {
          var chieu = classifyRow(row, owner.mst, owner.chieu || row.chieuHint);
          fillCounterparty(row, chieu, owner.mst, owner.ten);
          if (chieu === "mua") g2.mua.push(row);
          else if (chieu === "ban") g2.ban.push(row);
          else g2.unknown.push(row);
        });
      });
    }

    var list = Object.keys(groups).map(function (mst) {
      var g = groups[mst];
      if (opts.khuTrung !== false) {
        g.mua = uniqueRows(g.mua);
        g.ban = uniqueRows(g.ban);
      }
      sortRows(g.mua);
      sortRows(g.ban);
      g.mua.forEach(function (r, i) { r.stt = i + 1; });
      g.ban.forEach(function (r, i) { r.stt = i + 1; });
      var all = g.mua.concat(g.ban);
      var rg = dateRange(all);
      return {
        mst: g.mst,
        ten: g.ten || forcedTen || ("MST " + g.mst),
        mua: g.mua,
        ban: g.ban,
        unknown: g.unknown,
        from: rg.from,
        to: rg.to,
        tongMua: sumField(g.mua, "doanhSo"),
        thueMua: sumField(g.mua, "tienThue"),
        ttMua: sumField(g.mua, "tongTien"),
        tongBan: sumField(g.ban, "doanhSo"),
        thueBan: sumField(g.ban, "tienThue"),
        ttBan: sumField(g.ban, "tongTien"),
        fileName: "HDDT_" + (g.mst || "MST") + "_" + slugTen(g.ten || forcedTen) + ".xlsx"
      };
    });
    list.sort(function (a, b) { return (b.mua.length + b.ban.length) - (a.mua.length + a.ban.length); });
    return list;
  }

  function uniqueRows(rows) {
    var seen = {};
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var k = dedupeKey(rows[i]);
      if (seen[k]) continue;
      seen[k] = true;
      out.push(rows[i]);
    }
    return out;
  }

  return {
    stripVi: stripVi,
    asString: asString,
    safeText: safeText,
    digitsMst: digitsMst,
    toNumber: toNumber,
    parseDate: parseDate,
    formatDate: formatDate,
    hintChieu: hintChieu,
    findHeader: findHeader,
    mapHeader: mapHeader,
    ingestSheet: ingestSheet,
    detectOwner: detectOwner,
    classifyRow: classifyRow,
    buildCompanies: buildCompanies,
    skipSheetName: skipSheetName,
    cellText: cellText,
    sumField: sumField,
    dateRange: dateRange,
    OUT_COLS_MUA: OUT_COLS_MUA,
    OUT_COLS_BAN: OUT_COLS_BAN,
    OUT_TITLES_VI: OUT_TITLES_VI
  };
});
