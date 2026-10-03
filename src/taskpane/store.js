/* ===================================================================
   store.js
   Persists connection profiles and placeholder-table registrations
   inside the workbook itself (Office.context.document.settings), so
   the add-in remembers everything even after closing and reopening
   the file. No client secrets are ever stored here.
   =================================================================== */

function genId() {
  return "id_" + Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function loadSettings(key) {
  const raw = Office.context.document.settings.get(key);
  if (!raw) return [];
  try {
    return JSON.parse(raw);
  } catch (e) {
    return [];
  }
}

function saveSettings(key, value) {
  Office.context.document.settings.set(key, JSON.stringify(value));
  return new Promise((resolve, reject) => {
    Office.context.document.settings.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

function getConnections() {
  return loadSettings(SETTINGS_KEY_CONNECTIONS);
}

async function upsertConnection(conn) {
  const list = getConnections();
  const idx = list.findIndex(
    c => c.tenantId === conn.tenantId && c.clientId === conn.clientId &&
         c.environmentName === conn.environmentName && c.companyId === conn.companyId
  );
  if (idx >= 0) {
    list[idx] = { ...list[idx], ...conn };
  } else {
    conn.id = conn.id || genId();
    list.push(conn);
  }
  await saveSettings(SETTINGS_KEY_CONNECTIONS, list);
  return conn;
}

function getTables() {
  return loadSettings(SETTINGS_KEY_TABLES);
}

async function upsertTable(tableDef) {
  const list = getTables();
  const idx = list.findIndex(t => t.id === tableDef.id);
  if (idx >= 0) {
    list[idx] = { ...list[idx], ...tableDef };
  } else {
    tableDef.id = tableDef.id || genId();
    list.push(tableDef);
  }
  await saveSettings(SETTINGS_KEY_TABLES, list);
  return tableDef;
}

async function removeTable(tableId) {
  const list = getTables().filter(t => t.id !== tableId);
  await saveSettings(SETTINGS_KEY_TABLES, list);
}
