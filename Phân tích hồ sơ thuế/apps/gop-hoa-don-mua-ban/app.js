(function () {
  "use strict";
  var G = window.GopHoaDon;
  var files = [];
  var companies = [];
  var lastStats = null;
  var $ = function (id) { return document.getElementById(id); };

  function fmtNum(n) {
    if (n == null || n === "") return "0";
    return Number(n).toLocaleString("vi-VN");
  }

  function logLine(msg, cls) {
    var box = $("log");
    var p = document.createElement("div");
    p.className = "log-line" + (cls ? " " + cls : "");
    p.textContent = msg;
    box.appendChild(p);
    box.scrollTop = box.scrollHeight;
  }

  function setProgress(pct, text) {
    $("bar").style.width = Math.max(0, Math.min(100, pct)) + "%";
    $("progressText").textContent = text || "";
  }

  function addFiles(list) {
    var seen = {};
    files.forEach(function (f) { seen[f.name + ":" + f.size] = true; });
    var n = 0;
    Array.prototype.forEach.call(list, function (f) {
      var low = (f.name || "").toLowerCase();
      if (!/\.(xlsx|xlsm)$/i.test(low)) return;
      if (/^~\$/.test(f.name)) return;
      var k = f.name + ":" + f.size;
      if (seen[k]) return;
      seen[k] = true;
      files.push(f);
      n++;
    });
    renderFiles();
    if (n) logLine("Đã thêm " + n + " file Excel.", "ok");
  }

  function renderFiles() {
    $("fileCount").textContent = files.length ? (files.length + " file") : "Chưa chọn file";
    var ul = $("fileList");
    ul.innerHTML = "";
    files.slice(0, 40).forEach(function (f) {
      var li = document.createElement("li");
      li.textContent = f.webkitRelativePath || f.name;
      ul.appendChild(li);
    });
    if (files.length > 40) {
      var more = document.createElement("li");
      more.textContent = "... và " + (files.length - 40) + " file nữa";
      ul.appendChild(more);
    }
    $("btnClear").disabled = files.length === 0;
    $("btnScan").disabled = files.length === 0;
  }

  function sheetToRows(ws) {
    var rows = [];
    ws.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      while (rows.length < rowNumber) rows.push([]);
      var arr = [];
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        arr[colNumber - 1] = cell.value;
      });
      rows[rowNumber - 1] = arr;
    });
    return rows;
  }

  async function buildWorkbook(company) {
    return window.GopHoaDonExcel.buildWorkbook(ExcelJS, company);
  }

  function downloadBlob(name, blob) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1500);
  }

  async function scan() {
    if (!files.length) return;
    $("btnScan").disabled = true;
    $("btnZip").disabled = true;
    $("companies").innerHTML = "";
    $("log").innerHTML = "";
    logLine("Đang đọc " + files.length + " file...");
    var allRows = [];
    var warnings = [];
    var droppedTong = 0;
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      setProgress(((i + 1) / files.length) * 70, "Đọc " + f.name);
      try {
        var buf = await f.arrayBuffer();
        var wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf);
        wb.eachSheet(function (ws) {
          if (G.skipSheetName(ws.name)) return;
          var ingested = G.ingestSheet(sheetToRows(ws), f.name, ws.name, {
            locTongCong: $("optTongCong").checked
          });
          if (ingested.warning) warnings.push(ingested.warning);
          droppedTong += ingested.droppedTong || 0;
          allRows = allRows.concat(ingested.rows || []);
        });
      } catch (err) {
        warnings.push("Không đọc được " + f.name + ": " + (err && err.message ? err.message : err));
      }
    }
    setProgress(85, "Phân loại công ty...");
    companies = G.buildCompanies(allRows, {
      mst: $("mst").value,
      ten: $("tenCty").value,
      khuTrung: $("optDedupe").checked
    });
    lastStats = { rows: allRows.length, droppedTong: droppedTong, warnings: warnings };
    renderCompanies();
    warnings.slice(0, 12).forEach(function (w) { logLine(w, "warn"); });
    if (warnings.length > 12) logLine("... và " + (warnings.length - 12) + " cảnh báo nữa", "warn");
    logLine("Đã gom " + allRows.length + " dòng hóa đơn, bỏ " + droppedTong + " dòng Tổng cộng.", "ok");
    logLine("Phát hiện " + companies.length + " công ty.", "ok");
    setProgress(100, "Xong bước đọc");
    $("btnScan").disabled = false;
    $("btnZip").disabled = companies.length === 0;
  }

  function renderCompanies() {
    var box = $("companies");
    box.innerHTML = "";
    if (!companies.length) {
      box.innerHTML = '<p class="muted">Chưa có công ty nào. Kiểm tra file nguồn hoặc điền MST.</p>';
      return;
    }
    companies.forEach(function (c, idx) {
      var el = document.createElement("article");
      el.className = "card";
      el.innerHTML =
        '<label class="pick"><input type="checkbox" class="ck" data-i="' + idx + '" checked> ' +
        "<strong>" + escapeHtml(c.ten) + "</strong></label>" +
        '<div class="meta">MST ' + escapeHtml(c.mst) + (c.from ? " · " + G.formatDate(c.from) + " - " + G.formatDate(c.to) : "") + "</div>" +
        '<div class="nums"><span>Mua vào: <b>' + c.mua.length + "</b> HĐ · " + fmtNum(c.tongMua) + " đ</span>" +
        "<span>Bán ra: <b>" + c.ban.length + "</b> HĐ · " + fmtNum(c.tongBan) + " đ</span></div>" +
        (c.unknown.length ? '<div class="warn-mini">Chưa phân chiều: ' + c.unknown.length + " dòng</div>" : "") +
        '<button type="button" class="btn ghost" data-dl="' + idx + '">Tải file công ty này</button>';
      box.appendChild(el);
    });
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (ch) {
      return ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[ch];
    });
  }

  function selectedCompanies() {
    var cks = document.querySelectorAll(".ck");
    if (!cks.length) return companies.slice();
    var out = [];
    cks.forEach(function (ck) {
      if (ck.checked) out.push(companies[+ck.getAttribute("data-i")]);
    });
    return out;
  }

  async function downloadOne(company) {
    setProgress(40, "Đang tạo " + company.fileName);
    var buf = await buildWorkbook(company);
    downloadBlob(company.fileName, new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    setProgress(100, "Đã tải " + company.fileName);
    logLine("Đã tải " + company.fileName, "ok");
  }

  async function downloadZip() {
    var list = selectedCompanies();
    if (!list.length) {
      logLine("Chưa chọn công ty nào.", "warn");
      return;
    }
    if (list.length === 1) {
      await downloadOne(list[0]);
      return;
    }
    if (typeof JSZip === "undefined") {
      for (var i = 0; i < list.length; i++) await downloadOne(list[i]);
      return;
    }
    $("btnZip").disabled = true;
    var zip = new JSZip();
    for (var j = 0; j < list.length; j++) {
      setProgress(((j + 1) / list.length) * 90, "Gói " + list[j].fileName);
      var buf = await buildWorkbook(list[j]);
      zip.file(list[j].fileName, buf);
    }
    var blob = await zip.generateAsync({ type: "blob" });
    downloadBlob("HDDT_tong_hop_" + list.length + "_cong_ty.zip", blob);
    logLine("Đã tải zip " + list.length + " công ty.", "ok");
    setProgress(100, "Đã tải zip");
    $("btnZip").disabled = false;
  }

  function bind() {
    var drop = $("drop");
    ["dragenter", "dragover"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) {
        e.preventDefault();
        drop.classList.add("hot");
      });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) {
        e.preventDefault();
        drop.classList.remove("hot");
      });
    });
    drop.addEventListener("drop", function (e) {
      addFiles(e.dataTransfer.files);
    });
    $("filePick").addEventListener("change", function (e) {
      addFiles(e.target.files);
      e.target.value = "";
    });
    $("folderPick").addEventListener("change", function (e) {
      addFiles(e.target.files);
      e.target.value = "";
    });
    $("btnClear").addEventListener("click", function () {
      files = [];
      companies = [];
      $("companies").innerHTML = "";
      $("log").innerHTML = "";
      renderFiles();
      setProgress(0, "");
    });
    $("btnScan").addEventListener("click", function () {
      scan().catch(function (err) {
        logLine("Lỗi: " + (err && err.message ? err.message : err), "warn");
        $("btnScan").disabled = false;
      });
    });
    $("btnZip").addEventListener("click", function () {
      downloadZip().catch(function (err) {
        logLine("Lỗi khi tạo file: " + (err && err.message ? err.message : err), "warn");
        $("btnZip").disabled = false;
      });
    });
    fetch("/api/ping").then(function (res) {
      if (res.ok) $("btnQuit").style.display = "inline-block";
    }).catch(function () {});
    $("btnQuit").addEventListener("click", function () {
      fetch("/api/shutdown", { method: "POST" }).finally(function () {
        window.close();
      });
    });
    $("companies").addEventListener("click", function (e) {
      var btn = e.target.closest("[data-dl]");
      if (!btn) return;
      var c = companies[+btn.getAttribute("data-dl")];
      if (c) downloadOne(c).catch(function (err) {
        logLine("Lỗi: " + (err && err.message ? err.message : err), "warn");
      });
    });
  }

  document.addEventListener("DOMContentLoaded", bind);
})();
