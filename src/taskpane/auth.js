/* ===================================================================
   auth.js
   Thin wrapper around MSAL.js for delegated (interactive) OAuth sign-in
   to Microsoft Entra ID, scoped to the Business Central API.
   No client secret is ever used or stored - this is a public client
   (SPA) flow with PKCE, which is the correct/secure pattern for code
   that runs inside a browser task pane.
   =================================================================== */

let msalApp = null;
let activeAccount = null;
let activeConnectionMeta = null; // { tenantId, clientId, environmentName, environmentType }

function buildMsalApp(tenantId, clientId) {
  const msalConfig = {
    auth: {
      clientId: clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri: getRedirectUri()
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false
    }
  };
  return new msal.PublicClientApplication(msalConfig);
}

/**
 * Interactive sign-in (popup). Returns the acquired account.
 */
async function bcSignIn(tenantId, clientId) {
  msalApp = buildMsalApp(tenantId, clientId);
  await msalApp.initialize();

  const loginRequest = { scopes: [BC_API_SCOPE_TEMPLATE] };
  const result = await msalApp.loginPopup(loginRequest);
  activeAccount = result.account;
  activeConnectionMeta = { tenantId, clientId };
  return { account: result.account, accessToken: result.accessToken };
}

/**
 * Gets a valid access token, refreshing silently when possible.
 * Falls back to an interactive popup if silent acquisition fails
 * (e.g., token expired and no refresh token cached).
 */
async function bcGetToken(connectionMeta) {
  if (!msalApp) {
    msalApp = buildMsalApp(connectionMeta.tenantId, connectionMeta.clientId);
    await msalApp.initialize();
  }
  const accounts = msalApp.getAllAccounts();
  const account = accounts.find(a => true) || activeAccount;
  const tokenRequest = { scopes: [BC_API_SCOPE_TEMPLATE], account };

  try {
    const result = await msalApp.acquireTokenSilent(tokenRequest);
    return result.accessToken;
  } catch (silentError) {
    const result = await msalApp.acquireTokenPopup(tokenRequest);
    activeAccount = result.account;
    return result.accessToken;
  }
}

function bcSignOut() {
  if (msalApp && activeAccount) {
    msalApp.logoutPopup({ account: activeAccount });
  }
  activeAccount = null;
}
