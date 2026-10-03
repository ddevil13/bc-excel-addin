/* ===================================================================
   taskpane.js
   Wires up the UI: Home / Wizard (Add Table) / Refresh / Manage views.
   =================================================================== */

let wizard = {
  connection: null,   // { tenantId, clientId, environmentType, environmentName, companyId, companyName }
  token: null,
  table: null,        // entry from BC_TABLES
  allFields: [],
  selectedFields: [],
  anchor: null,       // { sheetName, address }
  tableName: null
};

function $(id) { return document.getElementById(id); }

function setStatus(id, message, type) {
  const el = $(id);
  el.textContent = message || "";
  el.className = "status" + (type ? " " + type : "");
}

function showView(viewId) {
  ["home-view", "wizard-view", "refresh-view", "manage-view"].forEach(id => {
    $(id).style.display = id === viewId ? "" : "none";
  });
}

function showWizardStep(n) {
  [1, 2, 3].forEach(i => {
    $("wizard-step" + i).style.display = i === n ? "" : "none";
  });
  $("wizard-done").style.display = "none";
  document.querySelectorAll(".step-dot").forEach(dot => {
    dot.classList.toggle("active", Number(dot.dataset.step) === n);
  });
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) return;

  const params = new URLSearchParams(window.location.search);
  const mode = params.get("mode");

  populateSavedConnections();
  populateTableDropdown();

  if (mode === "wizard") {
    startWizard();
  } else if (mode === "refresh") {
    startRefreshView();
  } else {
    showView("home-view");
  }

  wireHomeButtons();
  wireWizardButtons();
  wireRefreshButtons();
  wireManageButtons();
  wireConflictModal();
});

/* ---------------------------- HOME ---------------------------- */

function wireHomeButtons() {
  $("home-add-table").onclick = startWizard;
  $("home-refresh").onclick = startRefreshView;
  $("home-manage").onclick = startManageView;
}

/* --------------------------- WIZARD ---------------------------- */

function startWizard() {
  wizard = { connection: null, token: null, table: null, allFields: [], selectedFields: [], anchor: null, tableName: null };
  showView("wizard-view");
  showWizardStep(1);
  $("step1-next").disabled = true;
  $("input-company").disabled = true;
  $("input-company").innerHTML = '<option value="">-- Sign in first --</option>';
  setStatus("signin-status", "");
}

function populateSavedConnections() {
  const list = getConnections();
  const sel = $("saved-connections");
  sel.innerHTML = '<option value="">-- New connection --</option>';
  list.forEach((c, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    opt.textContent = `${c.companyName || c.companyId} @ ${c.environmentName} (${c.environmentType})`;
    sel.appendChild(opt);
  });
}

function populateTableDropdown() {
  const sel = $("input-table");
  sel.innerHTML = '<option value="">-- Select a table --</option>';
  BC_TABLES.forEach(t => {
    const opt = document.createElement("option");
    opt.value = t.key;
    opt.textContent = t.label;
    sel.appendChild(opt);
  });
}

function wireWizardButtons() {
  $("saved-connections").onchange = () => {
    const idx = $("saved-connections").value;
    if (idx === "") return;
    const c = getConnections()[Number(idx)];
    $("input-tenant").value = c.tenantId;
    $("input-clientid").value = c.clientId;
    $("input-envtype").value = c.environmentType;
    $("input-envname").value = c.environmentName;
  };

  $("btn-signin").onclick = async () => {
    const tenantId = $("input-tenant").value.trim();
    const clientId = $("input-clientid").value.trim();
    const environmentType = $("input-envtype").value;
    const environmentName = $("input-envname").value.trim();

    if (!tenantId || !clientId || !environmentName) {
      setStatus("signin-status", "Please fill Tenant ID, Client ID and Environment Name.", "error");
      return;
    }
    setStatus("signin-status", "Signing in...", "busy");
    $("btn-signin").disabled = true;
    try {
      const { accessToken } = await bcSignIn(tenantId, clientId);
      wizard.connection = { tenantId, clientId, environmentType, environmentName };
      wizard.token = accessToken;

      setStatus("signin-status", "Signed in. Loading companies...", "busy");
      const companies = await bcListCompanies(wizard.connection, accessToken);
      const sel = $("input-company");
      sel.innerHTML = "";
      if (companies.length === 0) {
        sel.innerHTML = '<option value="">No companies found</option>';
      } else {
        companies.forEach(c => {
          const opt = document.createElement("option");
          opt.value = c.id;
          opt.dataset.name = c.name || c.displayName;
          opt.textContent = c.name || c.displayName;
          sel.appendChild(opt);
        });
      }
      sel.disabled = false;
      setStatus("signin-status", "Signed in successfully.", "ok");
      $("step1-next").disabled = false;
    } catch (err) {
      setStatus("signin-status", "Sign-in failed: " + err.message, "error");
    } finally {
      $("btn-signin").disabled = false;
    }
  };

  $("step1-cancel").onclick = () => showView("home-view");

  $("step1-next").onclick = () => {
    const sel = $("input-company");
    const opt = sel.options[sel.selectedIndex];
    if (!opt || !opt.value) {
      setStatus("signin-status", "Please select a company.", "error");
      return;
    }
    wizard.connection.companyId = opt.value;
    wizard.connection.companyName = opt.dataset.name || opt.textContent;
    showWizardStep(2);
  };

  $("step2-back").onclick = () => showWizardStep(1);

  $("input-table").onchange = async () => {
    const key = $("input-table").value;
    $("fields-list").innerHTML = "";
    $("step2-next").disabled = true;
    if (!key) return;
    wizard.table = BC_TABLES.find(t => t.key === key);

    setStatus("table-load-status", "Loading available fields...", "busy");
    try {
      const token = await bcGetToken(wizard.connection);
      wizard.token = token;
      const fields = await bcDiscoverFields(wizard.connection, token, wizard.connection.companyId, wizard.table.entitySet);
      wizard.allFields = fields;
      renderFieldsList(fields);
      setStatus("table-load-status", `${fields.length} fields found.`, "ok");
    } catch (err) {
      setStatus("table-load-status", "Could not load fields: " + err.message, "error");
    }
  };

  $("field-filter").oninput = () => {
    const q = $("field-filter").value.toLowerCase();
    document.querySelectorAll("#fields-list .field-row").forEach(row => {
      row.style.display = row.dataset.field.toLowerCase().includes(q) ? "" : "none";
    });
  };

  $("step2-next").onclick = () => {
    const checked = Array.from(document.querySelectorAll("#fields-list input[type=checkbox]:checked")).map(cb => cb.value);
    if (checked.length === 0) {
      setStatus("table-load-status", "Select at least one field.", "error");
      return;
    }
    wizard.selectedFields = checked;
    const suggested = "BC_" + wizard.table.label.replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
    $("input-tablename").value = suggested;
    $("anchor-status").textContent = "";
    wizard.anchor = null;
    showWizardStep(3);
  };

  $("step3-back").onclick = () => showWizardStep(2);

  $("btn-use-active-cell").onclick = async () => {
    try {
      const info = await getActiveCellInfo();
      wizard.anchor = { sheetName: info.sheetName, address: info.address };
      setStatus("anchor-status", `Selected: ${info.sheetName}!${info.address}`, "ok");
    } catch (err) {
      setStatus("anchor-status", "Could not read active cell: " + err.message, "error");
    }
  };

  $("btn-create-table").onclick = async () => {
    const tableName = $("input-tablename").value.trim();
    if (!tableName || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(tableName)) {
      setStatus("create-status", "Enter a valid table name (letters, numbers, underscore; must start with a letter).", "error");
      return;
    }
    if (!wizard.anchor) {
      setStatus("create-status", "Pick a placeholder cell first.", "error");
      return;
    }
    $("btn-create-table").disabled = true;
    setStatus("create-status", "Checking table name...", "busy");
    try {
      if (await tableNameExists(tableName)) {
        setStatus("create-status", "A table with this name already exists. Choose another name.", "error");
        $("btn-create-table").disabled = false;
        return;
      }

      setStatus("create-status", "Fetching data from Business Central...", "busy");
      const token = await bcGetToken(wizard.connection);
      const rows = await bcFetchAllRecords(wizard.connection, token, wizard.connection.companyId, wizard.table.entitySet, wizard.selectedFields);

      setStatus("create-status", `Writing ${rows.length} rows to Excel...`, "busy");
      await createPlaceholderTable(wizard.anchor.sheetName, wizard.anchor.address, tableName, wizard.selectedFields, rows);

      await upsertConnection({
        tenantId: wizard.connection.tenantId,
        clientId: wizard.connection.clientId,
        environmentType: wizard.connection.environmentType,
        environmentName: wizard.connection.environmentName,
        companyId: wizard.connection.companyId,
        companyName: wizard.connection.companyName
      });

      await upsertTable({
        tableName,
        label: wizard.table.label,
        entitySet: wizard.table.entitySet,
        apiVersion: wizard.table.apiVersion,
        fields: wizard.selectedFields,
        sheetName: wizard.anchor.sheetName,
        anchorAddress: wizard.anchor.address,
        tenantId: wizard.connection.tenantId,
        clientId: wizard.connection.clientId,
        environmentType: wizard.connection.environmentType,
        environmentName: wizard.connection.environmentName,
        companyId: wizard.connection.companyId,
        companyName: wizard.connection.companyName,
        rowCount: rows.length,
        lastRefreshed: new Date().toISOString()
      });

      wizard.tableName = tableName;
      $("done-message").textContent = `"${tableName}" created at ${wizard.anchor.sheetName}!${wizard.anchor.address} with ${rows.length} rows and ${wizard.selectedFields.length} columns.`;
      showWizardStep(0);
      $("wizard-done").style.display = "";
      populateSavedConnections();
    } catch (err) {
      setStatus("create-status", "Failed: " + err.message, "error");
    } finally {
      $("btn-create-table").disabled = false;
    }
  };

  $("btn-add-another").onclick = () => {
    // Keep the same connection/token, jump back to table selection.
    $("input-table").value = "";
    $("fields-list").innerHTML = "";
    $("step2-next").disabled = true;
    showWizardStep(2);
  };

  $("btn-wizard-home").onclick = () => showView("home-view");
}

function renderFieldsList(fields) {
  const container = $("fields-list");
  container.innerHTML = "";
  fields.forEach(f => {
    const row = document.createElement("label");
    row.className = "field-row";
    row.dataset.field = f;
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = f;
    cb.checked = true;
    cb.onchange = () => {
      const anyChecked = document.querySelectorAll("#fields-list input[type=checkbox]:checked").length > 0;
      $("step2-next").disabled = !anyChecked;
    };
    row.appendChild(cb);
    row.appendChild(document.createTextNode(" " + f));
    container.appendChild(row);
  });
  $("step2-next").disabled = fields.length === 0;
}

/* -------------------------- REFRESH ---------------------------- */

function startRefreshView() {
  showView("refresh-view");
  renderRefreshList();
  $("refresh-log").textContent = "";
}

function renderRefreshList() {
  const tables = getTables();
  const container = $("refresh-list");
  container.innerHTML = "";
  if (tables.length === 0) {
    container.innerHTML = '<p class="muted">No tables registered yet. Use "Add Table" first.</p>';
    return;
  }
  tables.forEach(t => {
    const row = document.createElement("div");
    row.className = "refresh-item";
    row.innerHTML = `
      <input type="checkbox" value="${t.id}" />
      <div>
        <div class="title">${t.tableName} <span class="muted">(${t.label})</span></div>
        <div class="meta">${t.sheetName}!${t.anchorAddress} &middot; ${t.rowCount || 0} rows &middot; last refreshed: ${t.lastRefreshed ? new Date(t.lastRefreshed).toLocaleString() : "never"}</div>
      </div>`;
    container.appendChild(row);
  });
}

function wireRefreshButtons() {
  $("btn-refresh-home").onclick = () => showView("home-view");

  $("btn-refresh-selected").onclick = async () => {
    const ids = Array.from(document.querySelectorAll("#refresh-list input[type=checkbox]:checked")).map(cb => cb.value);
    if (ids.length === 0) {
      logRefresh("Select at least one table to refresh.");
      return;
    }
    $("btn-refresh-selected").disabled = true;
    const allTables = getTables();
    for (const id of ids) {
      const tableDef = allTables.find(t => t.id === id);
      if (!tableDef) continue;
      await refreshOneTable(tableDef);
    }
    $("btn-refresh-selected").disabled = false;
    renderRefreshList();
  };
}

function logRefresh(msg) {
  const el = $("refresh-log");
  el.textContent += msg + "\n";
  el.scrollTop = el.scrollHeight;
}

async function refreshOneTable(tableDef) {
  logRefresh(`Refreshing "${tableDef.tableName}"...`);
  try {
    const connMeta = {
      tenantId: tableDef.tenantId,
      clientId: tableDef.clientId,
      environmentType: tableDef.environmentType,
      environmentName: tableDef.environmentName
    };
    const token = await bcGetToken(connMeta);
    const rows = await bcFetchAllRecords(connMeta, token, tableDef.companyId, tableDef.entitySet, tableDef.fields);

    // Compute the rectangle this table will occupy after refresh.
    const anchorInfo = await getRangeRowCol(tableDef.sheetName, tableDef.anchorAddress);
    const newRect = {
      sheetName: tableDef.sheetName,
      tableName: tableDef.tableName,
      startRow: anchorInfo.rowIndex,
      startCol: anchorInfo.columnIndex,
      endRow: anchorInfo.rowIndex + rows.length, // +1 header -1 zero-index = rows.length
      endCol: anchorInfo.columnIndex + tableDef.fields.length - 1
    };

    const others = getTables().filter(t => t.id !== tableDef.id && t.sheetName === tableDef.sheetName);
    let conflictResolved = true;
    for (const other of others) {
      const otherRect = await getCurrentTableRect(other.sheetName, other.tableName);
      if (!otherRect) continue; // table was deleted manually elsewhere; ignore
      if (rectsOverlap(newRect, otherRect)) {
        const outcome = await showConflictModal(tableDef, other, otherRect);
        if (outcome === "cancel") {
          logRefresh(`  Skipped "${tableDef.tableName}" (conflict with "${other.tableName}" not resolved).`);
          conflictResolved = false;
          break;
        }
        // outcome === "moved": the other table has been relocated; continue checking remaining others.
      }
    }
    if (!conflictResolved) return;

    await refreshPlaceholderTable(tableDef, rows);
    await upsertTable({ id: tableDef.id, rowCount: rows.length, lastRefreshed: new Date().toISOString() });
    logRefresh(`  Done: ${rows.length} rows written to "${tableDef.tableName}".`);
  } catch (err) {
    logRefresh(`  Error refreshing "${tableDef.tableName}": ${err.message}`);
  }
}

async function getRangeRowCol(sheetName, address) {
  let result = null;
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(address);
    range.load("rowIndex,columnIndex");
    await context.sync();
    result = { rowIndex: range.rowIndex, columnIndex: range.columnIndex };
  });
  return result;
}

/* ------------------------ CONFLICT MODAL ------------------------ */

let conflictResolveFn = null;
let conflictPendingAnchor = null;

function wireConflictModal() {
  $("conflict-pick-cell").onclick = async () => {
    try {
      const info = await getActiveCellInfo();
      conflictPendingAnchor = { sheetName: info.sheetName, address: info.address };
      setStatus("conflict-pick-status", `Will move to: ${info.sheetName}!${info.address}`, "ok");
      $("conflict-confirm-move").disabled = false;
    } catch (err) {
      setStatus("conflict-pick-status", "Could not read active cell: " + err.message, "error");
    }
  };

  $("conflict-confirm-move").onclick = async () => {
    if (conflictResolveFn) conflictResolveFn({ action: "moved", anchor: conflictPendingAnchor });
    $("conflict-modal").style.display = "none";
  };

  $("conflict-cancel").onclick = () => {
    if (conflictResolveFn) conflictResolveFn({ action: "cancel" });
    $("conflict-modal").style.display = "none";
  };
}

function showConflictModal(tableDef, otherTableDef, otherRect) {
  return new Promise(resolve => {
    conflictPendingAnchor = null;
    $("conflict-confirm-move").disabled = true;
    setStatus("conflict-pick-status", "");
    $("conflict-message").textContent =
      `Refreshing "${tableDef.tableName}" needs rows that overlap table "${otherTableDef.tableName}" ` +
      `currently at ${otherRect.sheetName}!${otherRect.address}. Click a free cell in the worksheet, ` +
      `then "Pick a new cell for the other table", or cancel this refresh.`;
    $("conflict-modal").style.display = "flex";

    conflictResolveFn = async (result) => {
      conflictResolveFn = null;
      if (result.action === "cancel") {
        resolve("cancel");
        return;
      }
      try {
        await moveTableToNewAnchor(otherTableDef.sheetName, otherTableDef.tableName, result.anchor.address);
        await upsertTable({ id: otherTableDef.id, sheetName: result.anchor.sheetName, anchorAddress: result.anchor.address });
        logRefresh(`  Moved "${otherTableDef.tableName}" to ${result.anchor.sheetName}!${result.anchor.address}.`);
        resolve("moved");
      } catch (err) {
        logRefresh(`  Failed to move "${otherTableDef.tableName}": ${err.message}`);
        resolve("cancel");
      }
    };
  });
}

/* -------------------------- MANAGE ------------------------------ */

function startManageView() {
  showView("manage-view");
  renderManageLists();
}

function renderManageLists() {
  const tables = getTables();
  const tDiv = $("manage-tables-list");
  tDiv.innerHTML = tables.length === 0 ? '<p class="muted">No tables registered.</p>' : "";
  tables.forEach(t => {
    const row = document.createElement("div");
    row.className = "manage-row";
    row.innerHTML = `<span>${t.tableName} - ${t.sheetName}!${t.anchorAddress}</span>
      <a href="#" class="remove-link" data-id="${t.id}">Remove</a>`;
    row.querySelector(".remove-link").onclick = async (e) => {
      e.preventDefault();
      if (!confirm(`Remove "${t.tableName}" and delete it from the sheet?`)) return;
      await deleteTableAndClear(t.sheetName, t.tableName);
      await removeTable(t.id);
      renderManageLists();
    };
    tDiv.appendChild(row);
  });

  const connections = getConnections();
  const cDiv = $("manage-connections-list");
  cDiv.innerHTML = connections.length === 0 ? '<p class="muted">No saved connections.</p>' : "";
  connections.forEach((c, idx) => {
    const row = document.createElement("div");
    row.className = "manage-row";
    row.innerHTML = `<span>${c.companyName} @ ${c.environmentName}</span>
      <a href="#" class="remove-link" data-idx="${idx}">Forget</a>`;
    row.querySelector(".remove-link").onclick = async (e) => {
      e.preventDefault();
      const list = getConnections().filter((_, i) => i !== idx);
      await saveSettings(SETTINGS_KEY_CONNECTIONS, list);
      renderManageLists();
      populateSavedConnections();
    };
    cDiv.appendChild(row);
  });
}

function wireManageButtons() {
  $("btn-manage-home").onclick = () => showView("home-view");
}
