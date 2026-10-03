/* ===================================================================
   bcapi.js
   Calls to the Business Central Standard API v2.0, using a bearer
   token obtained via auth.js.
   =================================================================== */

function bcBaseUrl(connectionMeta) {
  // https://api.businesscentral.dynamics.com/v2.0/{tenantId}/{environmentName}/api/v2.0
  return `https://api.businesscentral.dynamics.com/v2.0/${connectionMeta.tenantId}/${encodeURIComponent(connectionMeta.environmentName)}/api/v2.0`;
}

async function bcFetch(url, token) {
  const resp = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json"
    }
  });
  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`BC API call failed (${resp.status} ${resp.statusText}): ${url}\n${text}`);
  }
  return resp.json();
}

/** Lists companies available in the environment. */
async function bcListCompanies(connectionMeta, token) {
  const url = `${bcBaseUrl(connectionMeta)}/companies`;
  const data = await bcFetch(url, token);
  return data.value || [];
}

/** Fetches one sample record from an entity set to discover field names. */
async function bcDiscoverFields(connectionMeta, token, companyId, entitySet) {
  const url = `${bcBaseUrl(connectionMeta)}/companies(${companyId})/${entitySet}?$top=1`;
  const data = await bcFetch(url, token);
  const sample = (data.value && data.value[0]) || null;
  if (!sample) {
    return [];
  }
  return Object.keys(sample).filter(k => !HIDDEN_FIELDS.has(k));
}

/**
 * Fetches ALL records for an entity set, following @odata.nextLink
 * pagination, restricted to the requested fields via $select.
 * A hard cap prevents runaway pulls from accidentally locking up Excel.
 */
async function bcFetchAllRecords(connectionMeta, token, companyId, entitySet, fields, maxRows = 20000) {
  const select = fields.join(",");
  let url = `${bcBaseUrl(connectionMeta)}/companies(${companyId})/${entitySet}?$select=${encodeURIComponent(select)}&$top=1000`;
  let rows = [];

  while (url) {
    const data = await bcFetch(url, token);
    rows = rows.concat(data.value || []);
    if (rows.length >= maxRows) {
      break;
    }
    url = data["@odata.nextLink"] || null;
  }
  return rows;
}
