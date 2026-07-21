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
  var countdownTimer = null;
  var dwellSeconds = 5;

  /** Set while we navigate programmatically, so that selection event doesn't restart the dwell flow. */
  var suppressNextEvent = false;

  var ui = {};

  var SETTINGS_KEY = "cellHistoryState";
  var saveTimer = null;

  // ---------------------------------------------------------------------------
  // Startup
  // ---------------------------------------------------------------------------

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      return;
    }

    // Ribbon "Back" button (ExecuteFunction on the shared runtime).
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("ribbonGoBack", function (event) {
        goBack().finally(function () {
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
    ui.pendingBox = document.getElementById("pending-indicator");
    ui.pendingText = document.getElementById("pending-text");
    ui.statusText = document.getElementById("status-text");

    ui.backBtn.addEventListener("click", function () { goBack(); });
    ui.forwardBtn.addEventListener("click", function () { goForward(); });
    ui.clearBtn.addEventListener("click", clearHistory);
    ui.dwellInput.addEventListener("change", function () {
      var v = parseInt(ui.dwellInput.value, 10);
      if (isNaN(v) || v < 1) { v = 1; }
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
          dwellSeconds = saved.dwellSeconds;
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

  function handleSelection(worksheetId, address) {
    if (suppressNextEvent) {
      suppressNextEvent = false;
      cancelPending();
      return Promise.resolve();
    }

    // Same place as the current entry: refresh its timestamp, don't start a new dwell.
    var current = pointer >= 0 ? entries[pointer] : null;
    if (current && current.sheetId === worksheetId && current.address === address) {
      current.lastVisit = Date.now();
      cancelPending();
      renderList();
      return Promise.resolve();
    }

    // Same as the already-pending candidate: keep the existing timer running.
    if (pending && pending.sheetId === worksheetId && pending.address === address) {
      return Promise.resolve();
    }

    // New candidate: (re)start the dwell countdown.
    cancelPending();
    pending = { sheetId: worksheetId, address: address, since: Date.now() };
    dwellTimer = setTimeout(commitPending, dwellSeconds * 1000);
    countdownTimer = setInterval(updatePendingIndicator, 250);
    updatePendingIndicator();
    return Promise.resolve();
  }

  function cancelPending() {
    if (dwellTimer) { clearTimeout(dwellTimer); dwellTimer = null; }
    if (countdownTimer) { clearInterval(countdownTimer); countdownTimer = null; }
    pending = null;
    ui.pendingBox.classList.add("hidden");
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
    var current = pointer >= 0 ? entries[pointer] : null;

    if (current && current.sheetId === sheetId && current.address === address) {
      current.lastVisit = now;
      current.visits += 1;
    } else {
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
    if (pointer <= 0) { return Promise.resolve(); }
    pointer -= 1;
    return navigateTo(entries[pointer]);
  }

  function goForward() {
    if (pointer < 0 || pointer >= entries.length - 1) { return Promise.resolve(); }
    pointer += 1;
    return navigateTo(entries[pointer]);
  }

  function jumpTo(index) {
    if (index < 0 || index >= entries.length) { return Promise.resolve(); }
    pointer = index;
    return navigateTo(entries[index]);
  }

  function navigateTo(entry) {
    cancelPending();
    suppressNextEvent = true;
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
      suppressNextEvent = false;
      setStatus("Could not navigate (sheet may have been deleted): " + err.message);
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
    ui.backBtn.disabled = pointer <= 0;
    ui.forwardBtn.disabled = pointer < 0 || pointer >= entries.length - 1;
    ui.emptyState.classList.toggle("hidden", entries.length > 0);
    ui.list.innerHTML = "";

    // Newest first.
    for (var i = entries.length - 1; i >= 0; i--) {
      ui.list.appendChild(buildItem(entries[i], i));
    }
  }

  function buildItem(entry, index) {
    var li = document.createElement("li");
    li.className = "history-item" + (index === pointer ? " current" : "");
    li.title = "Jump to " + entry.sheetName + "!" + entry.address;

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
    return li;
  }

  function updatePendingIndicator() {
    if (!pending) { return; }
    var remaining = Math.max(0, dwellSeconds - (Date.now() - pending.since) / 1000);
    ui.pendingText.textContent = "Recording " + pending.address + " in " + remaining.toFixed(1) + "s\u2026";
    ui.pendingBox.classList.remove("hidden");
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
