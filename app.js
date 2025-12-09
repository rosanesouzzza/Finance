// util simples para log
const logEl = document.getElementById("log");
const log = (msg) => {
  const time = new Date().toTimeString().slice(0,8);
  logEl.textContent += `\n[${time}] ${msg}`;
  logEl.scrollTop = logEl.scrollHeight;
};

// pega os campos da UI
const tenantIdEl   = document.getElementById("tenantId");
const clientIdEl   = document.getElementById("clientId");
const siteHostEl   = document.getElementById("siteHost");
const sitePathEl   = document.getElementById("sitePath");
const redirectEl   = document.getElementById("redirectUri");

const btnLogin = document.getElementById("btnLogin");
const btnTest  = document.getElementById("btnTest");
const btnLogout = document.getElementById("btnLogout");

// MSAL client (criado sob demanda para refletir mudanças de campos)
let msalApp = null;
let account = null;

function buildMsal() {
  const tenantId   = tenantIdEl.value.trim();
  const clientId   = clientIdEl.value.trim();
  const redirectUri= redirectEl.value.trim();

  const authority = `https://login.microsoftonline.com/${tenantId}`; // ⚠️ TENANT-SPECIFIC

  const msalConfig = {
    auth: {
      clientId,
      authority,            // single-tenant → NUNCA usar /common
      redirectUri,
      navigateToLoginRequestUrl: false
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false
    },
    system: {
      allowRedirectInIframe: false
    }
  };

  msalApp = new msal.PublicClientApplication(msalConfig);
}

async function ensureLogin() {
  if (!msalApp) buildMsal();

  // tenta silent primeiro
  const accounts = msalApp.getAllAccounts();
  if (accounts.length > 0) {
    account = accounts[0];
    log(`Conta ativa: ${account.username}`);
    return account;
  }

  log("Silent falhou, tentando loginPopup…");
  const loginRequest = {
    scopes: ["openid", "profile", "offline_access"]
  };

  const login = await msalApp.loginPopup(loginRequest);
  account = login.account;
  log(`Login ok: ${account.username}`);
  return account;
}

async function getSpoToken() {
  if (!account) await ensureLogin();

  const siteHost = siteHostEl.value.trim();

  // Para SharePoint REST use escopos delegados do recurso SPO:
  // Leia: AllSites.Read (ou AllSites.Write / AllSites.FullControl)
  const tokenRequest = {
    account,
    scopes: [`https://${siteHost}/AllSites.Read`]
  };

  const resp = await msalApp.acquireTokenSilent(tokenRequest)
               .catch(async (e) => {
                 log("Silent token SPO falhou, tentando popup…");
                 return msalApp.acquireTokenPopup(tokenRequest);
               });

  return resp.accessToken;
}

async function testLists() {
  try {
    await ensureLogin();
    const accessToken = await getSpoToken();

    const siteHost = siteHostEl.value.trim();
    const sitePath = sitePathEl.value.trim();

    const url = `https://${siteHost}${sitePath}/_api/web/lists?$select=Title,ItemCount&$top=5`;
    const r = await fetch(url, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json;odata=nometadata"
      }
    });

    if (!r.ok) {
      const text = await r.text();
      throw new Error(`HTTP ${r.status} – ${text}`);
    }

    const data = await r.json();
    log("Listas (top 5):");
    (data.value || []).forEach((l, i) => {
      log(`  ${i+1}. ${l.Title} (itens: ${l.ItemCount})`);
    });

  } catch (err) {
    log(`ERRO teste: ${err.message}`);
    console.error(err);
  }
}

async function doLogout() {
  if (!msalApp) buildMsal();
  const accounts = msalApp.getAllAccounts();
  if (accounts.length) {
    await msalApp.logoutPopup({ account: accounts[0] });
    log("Logout realizado.");
    account = null;
  }
}

// eventos
btnLogin.addEventListener("click", ensureLogin);
btnTest .addEventListener("click", testLists);
btnLogout.addEventListener("click", doLogout);

// boot
log("Pronto para conectar…");
