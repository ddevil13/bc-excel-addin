/* ===================================================================
   config.js
   Static configuration: the 8 supported tables mapped to confirmed
   Business Central Standard API v2.0 entity sets.
   =================================================================== */

// NOTE on mapping:
// Business Central's Standard API does not expose one generic "Sales
// Header"/"Sales Line" entity across all document types - each document
// type (Quote/Order/Invoice/Credit Memo) has its own entity. "Sales
// Header"/"Sales Line" below are mapped to the Sales ORDER entities,
// which are the closest table-level equivalent. "Sales Inv. Header/Line"
// map directly and cleanly to the Sales Invoice entities.
const BC_TABLES = [
  {
    key: "customers",
    label: "Customer",
    entitySet: "customers",
    apiVersion: "v2.0"
  },
  {
    key: "vendors",
    label: "Vendor",
    entitySet: "vendors",
    apiVersion: "v2.0"
  },
  {
    key: "salesOrders",
    label: "Sales Header (Sales Order)",
    entitySet: "salesOrders",
    apiVersion: "v2.0"
  },
  {
    key: "salesOrderLines",
    label: "Sales Line (Sales Order Lines)",
    entitySet: "salesOrderLines",
    apiVersion: "v2.0"
  },
  {
    key: "customerLedgerEntries",
    label: "Customer Ledger Entry",
    entitySet: "customerLedgerEntries",
    apiVersion: "v2.0"
  },
  {
    key: "detailedCustomerLedgerEntries",
    label: "Detailed Customer Ledger Entry",
    entitySet: "detailedCustomerLedgerEntries",
    apiVersion: "v2.0"
  },
  {
    key: "salesInvoices",
    label: "Sales Inv. Header",
    entitySet: "salesInvoices",
    apiVersion: "v2.0"
  },
  {
    key: "salesInvoiceLines",
    label: "Sales Inv. Line",
    entitySet: "salesInvoiceLines",
    apiVersion: "v2.0"
  }
];

// Fields returned by the API that we never show/select (OData plumbing).
const HIDDEN_FIELDS = new Set(["@odata.etag", "@odata.context"]);

// Document-settings keys used to persist state inside the workbook.
const SETTINGS_KEY_CONNECTIONS = "bcConnector.connections";
const SETTINGS_KEY_TABLES = "bcConnector.tables";

/* ---- MSAL configuration is built at runtime once the user supplies
   Tenant ID + Client ID (see auth.js). We only keep the redirect URI
   fixed here - it MUST exactly match the URL this file is hosted at
   and MUST be registered on the Azure AD App Registration as a
   "Single-page application" (SPA) redirect URI. */
function getRedirectUri() {
  // Use the task pane URL without query string so sign-in works
  // whether the pane was opened in wizard or refresh mode.
  return window.location.origin + window.location.pathname;
}

const BC_API_SCOPE_TEMPLATE = "https://api.businesscentral.dynamics.com/user_impersonation";
