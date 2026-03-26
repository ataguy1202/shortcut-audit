/* ================================================
   Shortcut Audit — Spreadsheet Intelligence Engine
   Client-side Excel analysis: dependency graphing,
   risk detection, model health scoring.
   ================================================ */

(function () {
  "use strict";

  var dropZone = document.getElementById("drop-zone");
  var fileInput = document.getElementById("file-input");
  var landingSection = document.getElementById("landing-section");
  var dashboardSection = document.getElementById("dashboard-section");
  var backBtn = document.getElementById("back-btn");
  var exportBtn = document.getElementById("export-btn");
  var loadDemoBtn = document.getElementById("load-demo");
  var sheetFilter = document.getElementById("sheet-filter");
  var resetZoomBtn = document.getElementById("reset-zoom");

  var workbookData = null;
  var graphSim = null;

  // ---- Animated preview on landing ----
  buildPreviewAnimation();

  // ---- File handling ----
  dropZone.addEventListener("click", function () { fileInput.click(); });
  fileInput.addEventListener("change", function (e) {
    if (e.target.files.length) handleFile(e.target.files[0]);
  });
  dropZone.addEventListener("dragover", function (e) {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });
  dropZone.addEventListener("dragleave", function () {
    dropZone.classList.remove("drag-over");
  });
  dropZone.addEventListener("drop", function (e) {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });
  backBtn.addEventListener("click", showLanding);
  loadDemoBtn.addEventListener("click", loadDemo);

  function showLanding() {
    dashboardSection.classList.add("hidden");
    landingSection.classList.remove("hidden");
    if (graphSim) graphSim.stop();
  }

  function showDashboard() {
    landingSection.classList.add("hidden");
    dashboardSection.classList.remove("hidden");
  }

  function handleFile(file) {
    var loader = showLoader();
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = XLSX.read(data, { type: "array", cellFormula: true });
        document.getElementById("file-name").textContent = file.name;
        processWorkbook(wb);
        showDashboard();
      } catch (err) {
        alert("Could not parse this file. Please upload a valid .xlsx file.");
        console.error(err);
      }
      loader.remove();
    };
    reader.readAsArrayBuffer(file);
  }

  function showLoader() {
    var el = document.createElement("div");
    el.className = "loading-overlay";
    el.innerHTML = '<div class="loading-spinner"></div>';
    document.body.appendChild(el);
    return el;
  }

  // ---- Preview animation ----
  function buildPreviewAnimation() {
    var svg = document.getElementById("preview-graph");
    if (!svg) return;
    var nodes = [
      { x: 60, y: 50, r: 6, c: "#217346" },
      { x: 140, y: 35, r: 8, c: "#217346" },
      { x: 100, y: 100, r: 5, c: "#888" },
      { x: 200, y: 70, r: 10, c: "#217346" },
      { x: 180, y: 140, r: 6, c: "#217346" },
      { x: 260, y: 50, r: 7, c: "#888" },
      { x: 300, y: 110, r: 9, c: "#217346" },
      { x: 240, y: 170, r: 5, c: "#888" },
      { x: 80, y: 180, r: 7, c: "#217346" },
      { x: 160, y: 210, r: 6, c: "#c0392b" },
      { x: 310, y: 190, r: 5, c: "#217346" },
      { x: 40, y: 130, r: 4, c: "#888" },
      { x: 120, y: 150, r: 5, c: "#217346" },
      { x: 280, y: 230, r: 6, c: "#888" },
    ];
    var edges = [
      [0,1],[0,2],[1,3],[2,4],[3,5],[3,6],[4,7],[6,7],
      [2,8],[8,9],[6,10],[11,0],[11,8],[12,4],[12,9],[7,13],[10,13]
    ];
    var html = "";
    edges.forEach(function (e, i) {
      var a = nodes[e[0]], b = nodes[e[1]];
      html += '<line class="pedge" x1="'+a.x+'" y1="'+a.y+'" x2="'+b.x+'" y2="'+b.y+'" style="animation-delay:'+(.1+i*.06)+'s"/>';
    });
    nodes.forEach(function (n, i) {
      html += '<circle class="pnode" cx="'+n.x+'" cy="'+n.y+'" r="'+n.r+'" fill="'+n.c+'" style="animation-delay:'+(.2+i*.07)+'s"/>';
    });
    svg.innerHTML = html;
  }

  // ---- Demo workbook ----
  function loadDemo() {
    var loader = showLoader();
    setTimeout(function () {
      var wb = buildDemoWorkbook();
      document.getElementById("file-name").textContent = "Financial_Model_Demo.xlsx";
      processWorkbook(wb);
      showDashboard();
      loader.remove();
    }, 250);
  }

  function buildDemoWorkbook() {
    var wb = XLSX.utils.book_new();

    var isData = [
      ["Income Statement","","FY2023","FY2024","FY2025E"],
      ["Revenue","",48500000,62300000,{f:"C2*1.22"}],
      ["COGS","",-19400000,-24920000,{f:"-E2*0.4"}],
      ["Gross Profit","",{f:"C2+C3"},{f:"D2+D3"},{f:"E2+E3"}],
      ["Operating Expenses","","","",""],
      ["  SGA","",-8730000,-10614000,{f:"D6*1.15"}],
      ["  R&D","",-6790000,-9345000,{f:"-E2*0.16"}],
      ["  D&A","",-2425000,-3115000,{f:"BS!E6/10"}],
      ["Total OpEx","",{f:"SUM(C6:C8)"},{f:"SUM(D6:D8)"},{f:"SUM(E6:E8)"}],
      ["EBIT","",{f:"C4+C9"},{f:"D4+D9"},{f:"E4+E9"}],
      ["Interest Expense","",-1200000,-1350000,{f:"-BS!E12*0.06"}],
      ["EBT","",{f:"C10+C11"},{f:"D10+D11"},{f:"E10+E11"}],
      ["Tax (25%)","",{f:"C12*0.25"},{f:"D12*0.25"},{f:"E12*0.25"}],
      ["Net Income","",{f:"C12+C13"},{f:"D12+D13"},{f:"E12+E13"}],
      ["","","","",""],
      ["Margins","","","",""],
      ["Gross Margin","",{f:"C4/C2"},{f:"D4/D2"},{f:"E4/E2"}],
      ["EBIT Margin","",{f:"C10/C2"},{f:"D10/D2"},{f:"E10/E2"}],
      ["Net Margin","",{f:"C14/C2"},{f:"D14/D2"},{f:"E14/E2"}],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(isData), "Income Statement");

    var bsData = [
      ["Balance Sheet","","FY2023","FY2024","FY2025E"],
      ["Assets","","","",""],
      ["  Cash","",12400000,15800000,{f:"CF!E16"}],
      ["  Receivables","",8100000,10380000,{f:"'Income Statement'!E2/365*45"}],
      ["  Inventory","",4850000,6230000,{f:"-'Income Statement'!E3/365*60"}],
      ["  PP&E (net)","",24250000,31150000,{f:"D6+5000000-'Income Statement'!E8"}],
      ["Total Assets","",{f:"SUM(C3:C6)"},{f:"SUM(D3:D6)"},{f:"SUM(E3:E6)"}],
      ["","","","",""],
      ["Liabilities","","","",""],
      ["  Payables","",3880000,4990000,{f:"-'Income Statement'!E3/365*30"}],
      ["  Accrued","",2420000,3110000,{f:"'Income Statement'!E2*0.045"}],
      ["  Long-term Debt","",20000000,22500000,25000000],
      ["Total Liabilities","",{f:"SUM(C10:C12)"},{f:"SUM(D10:D12)"},{f:"SUM(E10:E12)"}],
      ["","","","",""],
      ["Equity","","","",""],
      ["  Retained Earnings","",23220000,29760000,{f:"D16+'Income Statement'!E14"}],
      ["  Common Stock","",3500000,3500000,3500000],
      ["Total Equity","",{f:"C16+C17"},{f:"D16+D17"},{f:"E16+E17"}],
      ["Total L+E","",{f:"C13+C18"},{f:"D13+D18"},{f:"E13+E18"}],
      ["Balance Check","",{f:"C7-C19"},{f:"D7-D19"},{f:"E7-E19"}],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bsData), "BS");

    var cfData = [
      ["Cash Flow Statement","","FY2023","FY2024","FY2025E"],
      ["Operating Activities","","","",""],
      ["  Net Income","",{f:"'Income Statement'!C14"},{f:"'Income Statement'!D14"},{f:"'Income Statement'!E14"}],
      ["  D&A","",{f:"-'Income Statement'!C8"},{f:"-'Income Statement'!D8"},{f:"-'Income Statement'!E8"}],
      ["  Change in AR","",{f:"-(BS!D4-BS!C4)"},{f:"-(BS!D4-BS!C4)"},{f:"-(BS!E4-BS!D4)"}],
      ["  Change in Inv","",{f:"-(BS!D5-BS!C5)"},{f:"-(BS!D5-BS!C5)"},{f:"-(BS!E5-BS!D5)"}],
      ["  Change in AP","",{f:"BS!D10-BS!C10"},{f:"BS!D10-BS!C10"},{f:"BS!E10-BS!D10"}],
      ["Cash from Ops","",{f:"SUM(C3:C7)"},{f:"SUM(D3:D7)"},{f:"SUM(E3:E7)"}],
      ["","","","",""],
      ["Investing Activities","","","",""],
      ["  CapEx","",-5000000,-6000000,-5000000],
      ["Cash from Investing","",{f:"C11"},{f:"D11"},{f:"E11"}],
      ["","","","",""],
      ["Financing","","","",""],
      ["Net Change in Cash","",{f:"C8+C12"},{f:"D8+D12"},{f:"E8+E12"}],
      ["Beginning Cash","",9800000,{f:"C16"},{f:"D16"}],
      ["Ending Cash","",{f:"C15+C16"},{f:"D15+D16"},{f:"E15+E16"}],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cfData), "CF");

    var assData = [
      ["Assumptions","","Value"],
      ["Revenue Growth","",0.22],
      ["COGS %","",0.4],
      ["SGA Growth","",0.15],
      ["R&D % of Rev","",0.16],
      ["Tax Rate","",0.25],
      ["Interest Rate","",0.06],
      ["DSO (days)","",45],
      ["DIO (days)","",60],
      ["DPO (days)","",30],
      ["","",""],
      ["Manual Override","",1500000],
      ["Error cell","",{f:"1/0"}],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(assData), "Assumptions");

    return wb;
  }

  // ---- Core analysis ----
  function processWorkbook(wb) {
    workbookData = analyze(wb);
    renderStats(workbookData);
    renderIssues(workbookData.issues);
    renderSheetBreakdown(workbookData.sheets);
    renderComplexityMap(workbookData);
    renderFunctionChart(workbookData.functionCounts);
    populateSheetFilter(workbookData.sheets);
    renderGraph(workbookData);
    wireExport(workbookData);
  }

  function analyze(wb) {
    var sheets = [];
    var allCells = [];
    var formulaCells = [];
    var edges = [];
    var functionCounts = {};
    var issues = [];
    var crossSheetCount = 0;

    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      var ref = ws["!ref"];
      if (!ref) {
        sheets.push({ name: name, cellCount: 0, formulaCount: 0, errorCount: 0 });
        return;
      }
      var range = XLSX.utils.decode_range(ref);
      var cellCount = 0, formulaCount = 0, errorCount = 0;

      for (var r = range.s.r; r <= range.e.r; r++) {
        for (var c = range.s.c; c <= range.e.c; c++) {
          var addr = XLSX.utils.encode_cell({ r: r, c: c });
          var cell = ws[addr];
          if (!cell) continue;
          cellCount++;
          var full = name + "!" + addr;
          allCells.push({ sheet: name, addr: addr, fullAddr: full, cell: cell, row: r, col: c });

          if (cell.f) {
            formulaCount++;
            formulaCells.push({ sheet: name, addr: addr, fullAddr: full, formula: cell.f, cell: cell, row: r, col: c });
            var refs = extractRefs(cell.f, name);
            refs.forEach(function (ref) {
              edges.push({ source: ref, target: full });
              if (!ref.startsWith(name + "!")) crossSheetCount++;
            });
            var funcs = extractFunctions(cell.f);
            funcs.forEach(function (fn) {
              functionCounts[fn] = (functionCounts[fn] || 0) + 1;
            });
            if (cell.t === "e" || (cell.w && /^#[A-Z]+/.test(cell.w))) {
              errorCount++;
              issues.push({
                severity: "critical",
                title: "Formula error in " + full,
                detail: (cell.w || "Error") + ": =" + cell.f,
              });
            }
          } else if (cell.t === "e") {
            errorCount++;
          }
        }
      }
      sheets.push({ name: name, cellCount: cellCount, formulaCount: formulaCount, errorCount: errorCount });
    });

    runAuditChecks(formulaCells, allCells, edges, issues);
    var score = computeHealthScore(formulaCells, issues, edges, sheets);

    return {
      sheets: sheets,
      formulaCells: formulaCells,
      allCells: allCells,
      edges: edges,
      functionCounts: functionCounts,
      issues: issues,
      crossSheetCount: crossSheetCount,
      score: score,
    };
  }

  function extractRefs(formula, currentSheet) {
    var refs = [];
    var re = /(?:'([^']+)'|([A-Za-z_]\w*))!(\$?[A-Z]{1,3}\$?\d{1,7})|\b(\$?[A-Z]{1,3}\$?\d{1,7})\b/g;
    var m;
    while ((m = re.exec(formula))) {
      if (m[1] || m[2]) {
        var sheet = m[1] || m[2];
        refs.push(sheet + "!" + m[3].replace(/\$/g, ""));
      } else if (m[4]) {
        var c = m[4].replace(/\$/g, "");
        if (/\d/.test(c) && /[A-Z]/i.test(c)) {
          refs.push(currentSheet + "!" + c);
        }
      }
    }
    return refs;
  }

  function extractFunctions(formula) {
    var funcs = [];
    var re = /([A-Z][A-Z0-9_.]+)\s*\(/g;
    var m;
    while ((m = re.exec(formula))) funcs.push(m[1]);
    return funcs;
  }

  function runAuditChecks(formulaCells, allCells, edges, issues) {
    var formulaAddrs = {};
    formulaCells.forEach(function (f) { formulaAddrs[f.fullAddr] = true; });
    var inputAddrs = {};
    edges.forEach(function (e) {
      if (!formulaAddrs[e.source]) inputAddrs[e.source] = true;
    });

    // Hardcoded values
    allCells.forEach(function (c) {
      if (!c.cell.f && c.cell.t === "n" && Math.abs(c.cell.v) > 100000 && inputAddrs[c.fullAddr]) {
        issues.push({
          severity: "warning",
          title: "Potential hardcoded value: " + c.fullAddr,
          detail: fmtNum(c.cell.v) + " is referenced by formulas. Consider linking to assumptions.",
        });
      }
    });

    // Long formulas
    formulaCells.forEach(function (fc) {
      if (fc.formula.length > 120) {
        issues.push({
          severity: "warning",
          title: "Complex formula in " + fc.fullAddr,
          detail: fc.formula.length + " characters. Consider decomposing into helper cells.",
        });
      }
    });

    // Volatile functions
    var volatiles = { NOW: 1, TODAY: 1, RAND: 1, RANDBETWEEN: 1, OFFSET: 1, INDIRECT: 1 };
    formulaCells.forEach(function (fc) {
      extractFunctions(fc.formula).forEach(function (fn) {
        if (volatiles[fn]) {
          issues.push({
            severity: "info",
            title: "Volatile function " + fn + " in " + fc.fullAddr,
            detail: "Recalculates on every change. Can degrade performance in large models.",
          });
        }
      });
    });

    // Inconsistent column formulas
    var bySheetCol = {};
    formulaCells.forEach(function (fc) {
      var key = fc.sheet + "!" + fc.addr.replace(/\d+/g, "");
      if (!bySheetCol[key]) bySheetCol[key] = [];
      bySheetCol[key].push(fc);
    });
    Object.keys(bySheetCol).forEach(function (key) {
      var cells = bySheetCol[key];
      if (cells.length < 3) return;
      var patterns = cells.map(function (c) {
        return { addr: c.fullAddr, p: c.formula.replace(/\d+/g, "#") };
      });
      var counts = {};
      patterns.forEach(function (p) { counts[p.p] = (counts[p.p] || 0) + 1; });
      var entries = Object.keys(counts).map(function (k) { return [k, counts[k]]; });
      entries.sort(function (a, b) { return b[1] - a[1]; });
      var dominant = entries[0];
      if (dominant && dominant[1] >= 2) {
        patterns.forEach(function (p) {
          if (p.p !== dominant[0] && counts[p.p] === 1) {
            issues.push({
              severity: "warning",
              title: "Inconsistent formula at " + p.addr,
              detail: "Differs from " + dominant[1] + " similar formulas in this column.",
            });
          }
        });
      }
    });

    var sevOrder = { critical: 0, warning: 1, info: 2 };
    issues.sort(function (a, b) { return sevOrder[a.severity] - sevOrder[b.severity]; });
  }

  function computeHealthScore(formulaCells, issues, edges, sheets) {
    var score = 100;
    issues.forEach(function (i) {
      if (i.severity === "critical") score -= 12;
      else if (i.severity === "warning") score -= 4;
      else score -= 1;
    });
    var totalCells = sheets.reduce(function (s, sh) { return s + sh.cellCount; }, 0);
    var totalFormulas = sheets.reduce(function (s, sh) { return s + sh.formulaCount; }, 0);
    if (totalCells > 0 && totalFormulas / totalCells > 0.3) score += 5;
    var cross = edges.filter(function (e) { return e.source.split("!")[0] !== e.target.split("!")[0]; });
    if (cross.length > 3) score += 5;
    return Math.max(0, Math.min(100, Math.round(score)));
  }

  // ---- Rendering ----
  function renderStats(data) {
    var tF = data.sheets.reduce(function (s, sh) { return s + sh.formulaCount; }, 0);
    var tC = data.sheets.reduce(function (s, sh) { return s + sh.cellCount; }, 0);
    document.getElementById("stat-formulas").textContent = tF.toLocaleString();
    document.getElementById("stat-cells").textContent = tC.toLocaleString();
    document.getElementById("stat-sheets").textContent = data.sheets.length;
    document.getElementById("stat-xsheet").textContent = data.crossSheetCount;
    document.getElementById("stat-errors").textContent = data.issues.length;
    document.getElementById("stat-depth").textContent = computeMaxDepth(data.edges);

    var s = data.score;
    document.getElementById("health-score").textContent = s;
    var grade = s >= 85 ? "Excellent" : s >= 70 ? "Good" : s >= 50 ? "Fair" : "Needs Work";
    document.getElementById("health-grade").textContent = grade;

    var ring = document.getElementById("score-ring-fill");
    var circ = 2 * Math.PI * 38;
    ring.setAttribute("stroke-dasharray", circ);
    ring.setAttribute("stroke-dashoffset", circ);
    requestAnimationFrame(function () {
      ring.style.transition = "stroke-dashoffset 1.2s cubic-bezier(.4,0,.2,1)";
      ring.style.strokeDashoffset = circ * (1 - s / 100);
    });
    ring.setAttribute("stroke", s >= 70 ? "#217346" : s >= 50 ? "#c98a0c" : "#c0392b");
  }

  function computeMaxDepth(edges) {
    if (!edges.length) return 0;
    var adj = {}, inDeg = {};
    edges.forEach(function (e) {
      if (!adj[e.source]) adj[e.source] = [];
      adj[e.source].push(e.target);
      inDeg[e.target] = (inDeg[e.target] || 0) + 1;
      if (!(e.source in inDeg)) inDeg[e.source] = 0;
    });
    var allNodes = Object.keys(Object.assign({}, adj, inDeg));
    var depth = {};
    var queue = [];
    allNodes.forEach(function (n) {
      depth[n] = 0;
      if (!inDeg[n]) queue.push(n);
    });
    var maxD = 0, processed = 0;
    while (queue.length && processed < 10000) {
      var node = queue.shift();
      processed++;
      (adj[node] || []).forEach(function (t) {
        depth[t] = Math.max(depth[t], depth[node] + 1);
        if (depth[t] > maxD) maxD = depth[t];
        inDeg[t]--;
        if (inDeg[t] <= 0) queue.push(t);
      });
    }
    return maxD;
  }

  function renderIssues(issues) {
    var list = document.getElementById("issues-list");
    var counts = document.getElementById("issue-counts");
    if (!issues.length) {
      list.innerHTML = '<div class="issues-empty">No issues detected. Model looks clean.</div>';
      counts.innerHTML = "";
      return;
    }
    var crit = 0, warn = 0, info = 0;
    issues.forEach(function (i) {
      if (i.severity === "critical") crit++;
      else if (i.severity === "warning") warn++;
      else info++;
    });
    var ch = "";
    if (crit) ch += '<span class="ibadge critical">' + crit + " critical</span>";
    if (warn) ch += '<span class="ibadge warning">' + warn + " warning</span>";
    if (info) ch += '<span class="ibadge info">' + info + " info</span>";
    counts.innerHTML = ch;

    list.innerHTML = issues.map(function (i) {
      return '<div class="issue-item"><span class="issue-dot ' + i.severity + '"></span>' +
        '<div class="issue-content"><div class="issue-title">' + esc(i.title) + '</div>' +
        '<div class="issue-detail">' + esc(i.detail) + '</div></div></div>';
    }).join("");
  }

  function renderSheetBreakdown(sheets) {
    var el = document.getElementById("sheet-breakdown");
    var maxF = Math.max.apply(null, sheets.map(function (s) { return s.formulaCount; }).concat([1]));
    el.innerHTML = sheets.map(function (s) {
      var pct = (s.formulaCount / maxF * 100).toFixed(1);
      return '<div class="sheet-row">' +
        '<span class="sheet-name">' + esc(s.name) + '</span>' +
        '<div class="sheet-bar-bg"><div class="sheet-bar-fill" style="width:' + pct + '%"></div></div>' +
        '<span class="sheet-stat">' + s.formulaCount + ' formulas</span></div>';
    }).join("");
  }

  // ---- Complexity heatmap ----
  function renderComplexityMap(data) {
    var el = document.getElementById("complexity-map");
    if (!data.formulaCells.length) {
      el.innerHTML = '<div class="func-empty">No formulas to map.</div>';
      return;
    }
    // Group by sheet, build a grid of formula complexity (length)
    var bySheet = {};
    data.formulaCells.forEach(function (fc) {
      if (!bySheet[fc.sheet]) bySheet[fc.sheet] = [];
      bySheet[fc.sheet].push(fc);
    });
    var maxLen = Math.max.apply(null, data.formulaCells.map(function (f) { return f.formula.length; }).concat([1]));

    var html = "";
    Object.keys(bySheet).forEach(function (sheet) {
      var cells = bySheet[sheet];
      html += '<div class="cmap-sheet"><div class="cmap-label">' + esc(sheet) + '</div><div class="cmap-grid">';
      cells.forEach(function (fc) {
        var intensity = fc.formula.length / maxLen;
        var color;
        if (intensity > 0.7) color = "rgba(192,57,43," + (0.4 + intensity * 0.6) + ")";
        else if (intensity > 0.3) color = "rgba(201,138,12," + (0.3 + intensity * 0.5) + ")";
        else color = "rgba(33,115,70," + (0.25 + intensity * 0.5) + ")";
        html += '<div class="cmap-cell" style="background:' + color + '" title="' + esc(fc.fullAddr) + ': ' + fc.formula.length + ' chars"></div>';
      });
      html += '</div></div>';
    });
    el.innerHTML = html;
  }

  function renderFunctionChart(counts) {
    var el = document.getElementById("function-chart");
    var entries = Object.keys(counts).map(function (k) { return [k, counts[k]]; });
    entries.sort(function (a, b) { return b[1] - a[1]; });
    entries = entries.slice(0, 18);
    if (!entries.length) {
      el.innerHTML = '<div class="func-empty">No functions detected.</div>';
      return;
    }
    var maxC = Math.max.apply(null, entries.map(function (e) { return e[1]; }));
    var maxH = 110;
    el.innerHTML = entries.map(function (e) {
      var h = Math.max(4, e[1] / maxC * maxH);
      return '<div class="func-col"><span class="func-count">' + e[1] + '</span>' +
        '<div class="func-bar" style="height:' + h + 'px"></div>' +
        '<span class="func-name">' + esc(e[0]) + '</span></div>';
    }).join("");
  }

  function populateSheetFilter(sheets) {
    sheetFilter.innerHTML = '<option value="__all__">All sheets</option>';
    sheets.forEach(function (s) {
      var opt = document.createElement("option");
      opt.value = s.name;
      opt.textContent = s.name;
      sheetFilter.appendChild(opt);
    });
    sheetFilter.onchange = function () { renderGraph(workbookData); };
  }

  // ---- D3 dependency graph ----
  function renderGraph(data) {
    var container = document.getElementById("graph-container");
    var emptyMsg = document.getElementById("graph-empty");
    container.querySelectorAll("svg, .node-tooltip").forEach(function (el) { el.remove(); });
    if (graphSim) graphSim.stop();

    var selected = sheetFilter.value;
    var formulaSet = {};
    data.formulaCells.forEach(function (f) { formulaSet[f.fullAddr] = true; });
    var errorSet = {};
    data.issues.forEach(function (i) {
      if (i.severity === "critical") {
        var m = i.title.match(/[\w ]+![A-Z]+\d+/);
        if (m) errorSet[m[0]] = true;
      }
    });

    var fEdges = data.edges;
    if (selected !== "__all__") {
      fEdges = data.edges.filter(function (e) {
        return e.source.startsWith(selected + "!") || e.target.startsWith(selected + "!");
      });
    }
    if (fEdges.length > 600) fEdges = fEdges.slice(0, 600);

    var nodeSet = {};
    fEdges.forEach(function (e) { nodeSet[e.source] = true; nodeSet[e.target] = true; });
    var nodeIds = Object.keys(nodeSet);

    if (!nodeIds.length) { emptyMsg.classList.remove("hidden"); return; }
    emptyMsg.classList.add("hidden");

    var nodes = nodeIds.map(function (id) {
      return { id: id, isFormula: !!formulaSet[id], isError: !!errorSet[id], sheet: id.split("!")[0] };
    });
    var links = fEdges.map(function (e) { return { source: e.source, target: e.target }; });

    var degree = {};
    links.forEach(function (l) {
      degree[l.source] = (degree[l.source] || 0) + 1;
      degree[l.target] = (degree[l.target] || 0) + 1;
    });

    // Sheet color scale
    var sheetNames = data.sheets.map(function (s) { return s.name; });
    var sheetColors = {};
    var greens = ["#217346", "#2a9d5c", "#34c471", "#1a5c38"];
    sheetNames.forEach(function (n, i) { sheetColors[n] = greens[i % greens.length]; });

    var w = container.clientWidth;
    var h = container.clientHeight || 520;

    var svg = d3.select(container).append("svg")
      .attr("viewBox", [0, 0, w, h])
      .attr("preserveAspectRatio", "xMidYMid meet");
    var g = svg.append("g");
    var zoom = d3.zoom().scaleExtent([0.2, 6]).on("zoom", function (event) {
      g.attr("transform", event.transform);
    });
    svg.call(zoom);
    resetZoomBtn.onclick = function () {
      svg.transition().duration(400).call(zoom.transform, d3.zoomIdentity);
    };

    var tooltip = document.createElement("div");
    tooltip.className = "node-tooltip";
    container.appendChild(tooltip);

    var sim = d3.forceSimulation(nodes)
      .force("link", d3.forceLink(links).id(function (d) { return d.id; }).distance(45))
      .force("charge", d3.forceManyBody().strength(-70))
      .force("center", d3.forceCenter(w / 2, h / 2))
      .force("collision", d3.forceCollide().radius(10));
    graphSim = sim;

    var link = g.append("g").selectAll("line").data(links).join("line")
      .attr("class", "link-line").attr("stroke-width", 1);

    var node = g.append("g").selectAll("circle").data(nodes).join("circle")
      .attr("class", "node-circle")
      .attr("r", function (d) { return Math.min(3 + (degree[d.id] || 0) * 0.7, 14); })
      .attr("fill", function (d) {
        if (d.isError) return "#c0392b";
        if (d.isFormula) return sheetColors[d.sheet] || "#217346";
        return "#888";
      })
      .call(drag(sim));

    node.on("mouseenter", function (event, d) {
      var fi = data.formulaCells.find(function (f) { return f.fullAddr === d.id; });
      var html = "<strong>" + esc(d.id) + "</strong>";
      if (fi) html += "<br>=" + esc(fi.formula);
      var deps = links.filter(function (l) { return (l.target === d || l.target.id === d.id); }).length;
      var feeds = links.filter(function (l) { return (l.source === d || l.source.id === d.id); }).length;
      html += "<br>" + deps + " inputs, feeds " + feeds + " cells";
      tooltip.innerHTML = html;
      tooltip.classList.add("visible");
    }).on("mousemove", function (event) {
      var rect = container.getBoundingClientRect();
      tooltip.style.left = (event.clientX - rect.left + 14) + "px";
      tooltip.style.top = (event.clientY - rect.top - 12) + "px";
    }).on("mouseleave", function () {
      tooltip.classList.remove("visible");
    });

    var label = g.append("g").selectAll("text")
      .data(nodes.filter(function (d) { return (degree[d.id] || 0) >= 3; }))
      .join("text")
      .attr("class", "node-label")
      .text(function (d) { return d.id.split("!")[1]; })
      .attr("dy", -10);

    sim.on("tick", function () {
      link.attr("x1", function (d) { return d.source.x; })
        .attr("y1", function (d) { return d.source.y; })
        .attr("x2", function (d) { return d.target.x; })
        .attr("y2", function (d) { return d.target.y; });
      node.attr("cx", function (d) { return d.x; }).attr("cy", function (d) { return d.y; });
      label.attr("x", function (d) { return d.x; }).attr("y", function (d) { return d.y; });
    });
  }

  function drag(sim) {
    return d3.drag()
      .on("start", function (event, d) {
        if (!event.active) sim.alphaTarget(0.3).restart();
        d.fx = d.x; d.fy = d.y;
      })
      .on("drag", function (event, d) { d.fx = event.x; d.fy = event.y; })
      .on("end", function (event, d) {
        if (!event.active) sim.alphaTarget(0);
        d.fx = null; d.fy = null;
      });
  }

  // ---- Export ----
  function wireExport(data) {
    exportBtn.onclick = function () {
      var lines = [
        "SHORTCUT AUDIT REPORT",
        "Generated: " + new Date().toISOString().split("T")[0],
        "Health Score: " + data.score + "/100",
        "",
        "SUMMARY",
        "Sheets: " + data.sheets.length,
        "Formulas: " + data.sheets.reduce(function (s, sh) { return s + sh.formulaCount; }, 0),
        "Cells: " + data.sheets.reduce(function (s, sh) { return s + sh.cellCount; }, 0),
        "Cross-sheet links: " + data.crossSheetCount,
        "Issues: " + data.issues.length,
        "",
        "SHEETS",
      ];
      data.sheets.forEach(function (s) {
        lines.push("  " + s.name + ": " + s.formulaCount + " formulas, " + s.cellCount + " cells");
      });
      lines.push("");
      lines.push("FINDINGS");
      if (!data.issues.length) lines.push("  None.");
      data.issues.forEach(function (i) {
        lines.push("  [" + i.severity.toUpperCase() + "] " + i.title);
        lines.push("    " + i.detail);
      });
      lines.push("");
      lines.push("FUNCTIONS");
      Object.keys(data.functionCounts).sort(function (a, b) {
        return data.functionCounts[b] - data.functionCounts[a];
      }).forEach(function (fn) {
        lines.push("  " + fn + ": " + data.functionCounts[fn]);
      });

      var blob = new Blob([lines.join("\n")], { type: "text/plain" });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "shortcut-audit-report.txt";
      a.click();
    };
  }

  // ---- Helpers ----
  function esc(str) {
    var d = document.createElement("div");
    d.textContent = str;
    return d.innerHTML;
  }
  function fmtNum(n) {
    return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(n);
  }
})();
