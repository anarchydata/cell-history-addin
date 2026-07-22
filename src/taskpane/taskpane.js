/* global Office, Excel */

(function () {
  "use strict";

  var MAX_ENTRIES = 100;

  /** Committed locations, oldest first: { sheetId, sheetName, address, firstVisit, lastVisit, visits } */
  var entries = [];
  /** Index of the "current" location within entries (-1 = none yet). */
  var pointer = -1;

  /** Selection we're waiting on before committing it to history. */
  var pending = null;
  var dwellTimer = null;
  var dwellSeconds = 0.1;

  /**
   * While Date.now() < suppressUntil, ignore selection/activation events.
   * activate()+select() can fire multiple events; a one-shot boolean is not enough.
   */
  var suppressUntil = 0;

  var ui = {};

  var SETTINGS_KEY = "cellHistoryState";
  var saveTimer = null;
  var hoverTimer = null;
  var hoverRequestId = 0;

  // ---------------------------------------------------------------------------
  // Startup
  // ---------------------------------------------------------------------------

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      return;
    }

    // Ribbon Back / Forward (one step each).
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("ribbonGoBack", function (event) {
        goBack().finally(function () {
          event.completed();
        });
      });
      Office.actions.associate("ribbonGoForward", function (event) {
        goForward().finally(function () {
          event.completed();
        });
      });
    }

    bindUi();
    restoreState();
    registerExcelEvents()
      .then(function () {
        setStatus("Tracking selection. Stay on a cell " + dwellSeconds + "s to record it.");
      })
      .catch(function (err) {
        setStatus("Failed to start tracking: " + err.message);
      });

    // Keep the "time ago" labels fresh.
    setInterval(renderList, 10000);
  });

  function bindUi() {
    ui.backBtn = document.getElementById("back-btn");
    ui.forwardBtn = document.getElementById("forward-btn");
    ui.clearBtn = document.getElementById("clear-btn");
    ui.dwellInput = document.getElementById("dwell-input");
    ui.list = document.getElementById("history-list");
    ui.emptyState = document.getElementById("empty-state");
    ui.statusText = document.getElementById("status-text");

    ui.backBtn.addEventListener("click", function () { goBack(); });
    ui.forwardBtn.addEventListener("click", function () { goForward(); });
    ui.clearBtn.addEventListener("click", clearHistory);
    ui.dwellInput.addEventListener("change", function () {
      var v = parseFloat(ui.dwellInput.value);
      if (isNaN(v) || v < 0.1) { v = 0.1; }
      if (v > 60) { v = 60; }
      dwellSeconds = v;
      ui.dwellInput.value = String(v);
      persistState();
    });
  }

  // ---------------------------------------------------------------------------
  // Persistence (stored in the workbook via document settings)
  // ---------------------------------------------------------------------------

  function restoreState() {
    try {
      var saved = Office.context.document.settings.get(SETTINGS_KEY);
      if (saved && Array.isArray(saved.entries)) {
        entries = saved.entries;
        pointer = typeof saved.pointer === "number" ? saved.pointer : entries.length - 1;
        if (pointer >= entries.length) { pointer = entries.length - 1; }
        if (typeof saved.dwellSeconds === "number") {
          dwellSeconds = Math.max(0.1, Math.min(60, saved.dwellSeconds));
          ui.dwellInput.value = String(dwellSeconds);
        }
        renderList();
      }
    } catch (e) {
      // Corrupt/old state: start fresh.
    }
  }

  /** Debounced save; the state is persisted into the workbook file when the user saves it. */
  function persistState() {
    if (saveTimer) { clearTimeout(saveTimer); }
    saveTimer = setTimeout(function () {
      saveTimer = null;
      Office.context.document.settings.set(SETTINGS_KEY, {
        entries: entries,
        pointer: pointer,
        dwellSeconds: dwellSeconds
      });
      Office.context.document.settings.saveAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          setStatus("Could not save history: " + result.error.message);
        }
      });
    }, 2000);
  }

  function registerExcelEvents() {
    return Excel.run(function (context) {
      var sheets = context.workbook.worksheets;
      sheets.onSelectionChanged.add(onSelectionChanged);
      sheets.onActivated.add(onSheetActivated);
      return context.sync();
    });
  }

  // ---------------------------------------------------------------------------
  // Selection tracking with dwell timer
  // ---------------------------------------------------------------------------

  function onSelectionChanged(eventArgs) {
    return handleSelection(eventArgs.worksheetId, eventArgs.address);
  }

  /** Switching sheets doesn't fire onSelectionChanged, so capture the active cell on activation. */
  function onSheetActivated(eventArgs) {
    return Excel.run(function (context) {
      var cell = context.workbook.getActiveCell();
      cell.load("address");
      return context.sync().then(function () {
        // address comes back as "Sheet1!A1"; strip the sheet part.
        var addr = cell.address.split("!").pop();
        return handleSelection(eventArgs.worksheetId, addr);
      });
    }).catch(function () { /* no active cell (e.g. chart selected) */ });
  }

  function beginSuppressSelection() {
    // Hold through the Excel sync; endSuppressSelection extends a short grace period.
    suppressUntil = Date.now() + 10000;
  }

  function endSuppressSelection() {
    suppressUntil = Date.now() + 200;
  }

  function normalizeAddress(address) {
    if (!address) { return ""; }
    var raw = String(address).split("!").pop().replace(/\$/g, "").toUpperCase();
    return raw.split(",").map(function (part) {
      var p = part.trim();
      var ends = p.split(":");
      if (ends.length === 2 && ends[0] === ends[1]) { return ends[0]; }
      return p;
    }).join(",");
  }

  function sameLocation(a, sheetId, address) {
    return a && a.sheetId === sheetId && normalizeAddress(a.address) === normalizeAddress(address);
  }

  function handleSelection(worksheetId, address) {
    if (Date.now() < suppressUntil) {
      cancelPending();
      return Promise.resolve();
    }

    address = normalizeAddress(address);

    // Same place as the current entry: refresh its timestamp, don't start a new dwell.
    var current = pointer >= 0 ? entries[pointer] : null;
    if (sameLocation(current, worksheetId, address)) {
      current.lastVisit = Date.now();
      cancelPending();
      renderList();
      return Promise.resolve();
    }

    // Same as the already-pending candidate: keep the existing timer running.
    if (pending && sameLocation(pending, worksheetId, address)) {
      return Promise.resolve();
    }

    // New candidate: (re)start the dwell timer silently.
    cancelPending();
    pending = { sheetId: worksheetId, address: address, since: Date.now() };
    dwellTimer = setTimeout(commitPending, dwellSeconds * 1000);
    return Promise.resolve();
  }

  function cancelPending() {
    if (dwellTimer) { clearTimeout(dwellTimer); dwellTimer = null; }
    pending = null;
  }

  function commitPending() {
    if (!pending) { return; }
    var loc = pending;
    cancelPending();

    // Resolve the sheet name, then record.
    Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getItem(loc.sheetId);
      sheet.load("name");
      return context.sync().then(function () {
        recordEntry(loc.sheetId, sheet.name, loc.address);
      });
    }).catch(function (err) {
      setStatus("Could not record location: " + err.message);
    });
  }

  function recordEntry(sheetId, sheetName, address) {
    var now = Date.now();
    address = normalizeAddress(address);
    var current = pointer >= 0 ? entries[pointer] : null;

    if (sameLocation(current, sheetId, address)) {
      current.lastVisit = now;
      current.visits += 1;
    } else {
      // Drop any "forward" entries when branching from mid-history (browser-style).
      if (pointer >= 0 && pointer < entries.length - 1) {
        entries = entries.slice(0, pointer + 1);
      }
      entries.push({
        sheetId: sheetId,
        sheetName: sheetName,
        address: address,
        firstVisit: now,
        lastVisit: now,
        visits: 1
      });
      if (entries.length > MAX_ENTRIES) {
        entries.shift();
      }
      pointer = entries.length - 1;
    }
    renderList();
    persistState();
  }

  // ---------------------------------------------------------------------------
  // Navigation
  // ---------------------------------------------------------------------------

  function goBack() {
    return goBackBy(1);
  }

  function goForward() {
    return goForwardBy(1);
  }

  function goBackBy(steps) {
    if (steps < 1 || pointer < steps) { return Promise.resolve(); }
    var target = entries[pointer - steps];
    pointer -= steps;
    return navigateTo(target);
  }

  function goForwardBy(steps) {
    if (steps < 1 || pointer < 0 || pointer + steps >= entries.length) {
      return Promise.resolve();
    }
    var target = entries[pointer + steps];
    pointer += steps;
    return navigateTo(target);
  }

  function formatEntry(entry) {
    return entry.sheetName + "!" + entry.address;
  }

  function jumpTo(index) {
    if (index < 0 || index >= entries.length) { return Promise.resolve(); }
    pointer = index;
    return navigateTo(entries[index]);
  }

  function navigateTo(entry) {
    cancelPending();
    beginSuppressSelection();
    return Excel.run(function (context) {
      // Prefer the sheet ID; fall back to the name for restored entries whose ID went stale.
      var sheet = context.workbook.worksheets.getItemOrNullObject(entry.sheetId);
      return context.sync().then(function () {
        if (sheet.isNullObject) {
          sheet = context.workbook.worksheets.getItem(entry.sheetName);
        }
        sheet.activate();
        // Multi-area selections ("A1:B2,D4") need getRanges; single areas use getRange.
        var target = entry.address.indexOf(",") >= 0
          ? sheet.getRanges(entry.address)
          : sheet.getRange(entry.address);
        target.select();
        return context.sync();
      });
    }).then(function () {
      entry.lastVisit = Date.now();
      renderList();
      persistState();
    }).catch(function (err) {
      setStatus("Could not navigate (sheet may have been deleted): " + err.message);
    }).then(function () {
      endSuppressSelection();
    });
  }

  function clearHistory() {
    entries = [];
    pointer = -1;
    cancelPending();
    renderList();
    persistState();
    setStatus("History cleared.");
  }

  // ---------------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------------

  function renderList() {
    var canBack = pointer > 0;
    var canForward = pointer >= 0 && pointer < entries.length - 1;
    ui.backBtn.disabled = !canBack;
    ui.forwardBtn.disabled = !canForward;
    ui.emptyState.classList.toggle("hidden", entries.length > 0);
    ui.list.innerHTML = "";

    // Newest first.
    for (var i = entries.length - 1; i >= 0; i--) {
      ui.list.appendChild(buildItem(entries[i], i));
    }
    updateRibbonButtons(canBack, canForward);
  }

  function updateRibbonButtons(canBack, canForward) {
    if (!Office.ribbon || !Office.ribbon.requestUpdate) { return; }
    try {
      Office.ribbon.requestUpdate({
        tabs: [{
          id: "TabHome",
          groups: [{
            id: "CellHistory.Group",
            controls: [
              { id: "CellHistory.BackButton", enabled: !!canBack },
              { id: "CellHistory.ForwardButton", enabled: !!canForward }
            ]
          }]
        }]
      });
    } catch (e) {
      // Host may not support RibbonApi.
    }
  }

  function buildItem(entry, index) {
    var li = document.createElement("li");
    li.className = "history-item" + (index === pointer ? " current" : "");
    var jumpLabel =
      index < pointer
        ? "Jump back to " + entry.sheetName + "!" + entry.address
        : index > pointer
          ? "Jump forward to " + entry.sheetName + "!" + entry.address
          : "Current: " + entry.sheetName + "!" + entry.address;
    li.title = jumpLabel;

    var loc = document.createElement("div");
    loc.className = "item-location";
    loc.textContent = entry.sheetName + "!" + entry.address;

    var meta = document.createElement("div");
    meta.className = "item-meta";
    var visits = entry.visits > 1 ? " \u00b7 " + entry.visits + " visits" : "";
    meta.textContent = timeAgo(entry.lastVisit) + visits + (index === pointer ? " \u00b7 current" : "");

    li.appendChild(loc);
    li.appendChild(meta);
    li.addEventListener("click", function () { jumpTo(index); });
    li.addEventListener("mouseenter", function () {
      previewEntry(entry);
    });
    li.addEventListener("mouseleave", function () {
      restoreCurrentSelection();
    });
    return li;
  }

  /** Hover preview: select the cell without changing history. */
  function previewEntry(entry) {
    if (hoverTimer) { clearTimeout(hoverTimer); }
    var req = ++hoverRequestId;
    hoverTimer = setTimeout(function () {
      selectOnly(entry, req);
    }, 80);
  }

  function restoreCurrentSelection() {
    if (hoverTimer) { clearTimeout(hoverTimer); hoverTimer = null; }
    var req = ++hoverRequestId;
    if (pointer < 0 || !entries[pointer]) { return; }
    selectOnly(entries[pointer], req);
  }

  function selectOnly(entry, req) {
    beginSuppressSelection();
    return Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getItemOrNullObject(entry.sheetId);
      return context.sync().then(function () {
        if (req !== hoverRequestId) { return; }
        if (sheet.isNullObject) {
          sheet = context.workbook.worksheets.getItem(entry.sheetName);
        }
        sheet.activate();
        var target = entry.address.indexOf(",") >= 0
          ? sheet.getRanges(entry.address)
          : sheet.getRange(entry.address);
        target.select();
        return context.sync();
      });
    }).catch(function () { /* hover preview is best-effort */ })
      .then(function () {
        endSuppressSelection();
      });
  }

  function timeAgo(ts) {
    var s = Math.floor((Date.now() - ts) / 1000);
    if (s < 10) { return "just now"; }
    if (s < 60) { return s + "s ago"; }
    var m = Math.floor(s / 60);
    if (m < 60) { return m + "m ago"; }
    var h = Math.floor(m / 60);
    if (h < 24) { return h + "h " + (m % 60) + "m ago"; }
    return new Date(ts).toLocaleString();
  }

  function setStatus(text) {
    ui.statusText.textContent = text;
  }
})();
