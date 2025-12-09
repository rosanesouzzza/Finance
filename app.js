// ===== Util =====
const $ = (id) => document.getElementById(id);
const log = (msg) => {
  const el = $("log");
  const now = new Date().toLocaleTimeString();
  el.textContent += `\n[${now}] ${msg}`;
  el.scrollTop = el.scrollHeight;
};

let msalApp = null;
let activeAccount = null;

// ===== Config inicial de MSAL (Popup) =====
function buildMsalConfig() {
  const clientId = $("clientId").value.trim();
  const redirectUri = $("redirectUri").value.trim();

  return {
    auth: {
      clientId,
      // Mesmo com popup, é saudável ter um redirectUri cadastrado como SPA.
      redirectUri,
      // se precisar: authority: `https://login.microsoftonline.com/${$("tenantId").value.trim()}`
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false
    },
    system: {
      loggerOptions: {
        loggerCallback: (level, message) => {
          if (/acquire token silent|CacheLookup|serverTelemetry/i.test(message)) return;
          log(`MSAL: ${message}`);
        },
        logLevel: msal.LogLevel.Warning
      }
    }
  };
}

function ensureMsalLoaded() {
  if (!window.msal || !window.msal.PublicClientApplication) {
    log("[ERRO] MSAL não disponível. Verifique vendor/msal-browser.min.js ou a ordem dos <script>.");
    alert("MSAL não carregou. Recarregue a página (Ctrl+F5).");
    return false;
  }
  return true;
}

function initMsal() {
  if (!ensureMsalLoaded()) return;

  const config = buildMsalConfig();
  msalApp = new msal.PublicClientApplication(config);

  // Restaura conta ativa, se houver
  const accounts = msalApp.getAllAccounts();
  if (accounts.length > 0) {
    activeAccount = accounts[0];
    msalApp.setActiveAccount(activeAccount);
    log(`Conta ativa: ${activeAccount.username}`);
  } else {
    log("deslogado");
  }
}

// ===== Login / Logout (POPUP) =====
const loginScopes = ["openid", "profile", "offline_access", "Sites.ReadWrite.All"];

async function loginPopup() {
  try {
    const res = await msalApp.loginPopup({
      scopes: loginScopes,
      prompt: "select_account"
    });
    activeAccount = res.account;
    msalApp.setActiveAccount(activeAccount);
    log(`Login ok: ${activeAccount?.username}`);
  } catch (e) {
    log(`ERRO login: ${e.errorCode || ""} ${e.message || e}`);
    alert("Falha no login (Popup). Veja o log.");
  }
}

async function logoutPopup() {
  try {
    const acc = msalApp.getActiveAccount();
    await msalApp.logoutPopup({ account: acc || undefined });
    activeAccount = null;
    log("Logout ok.");
  } catch (e) {
    log(`ERRO logout: ${e.message || e}`);
  }
}

// ===== Token (POPUP) =====
async function getTokenPopup(scopes) {
  if (!activeAccount) {
    throw new Error("Sem conta ativa. Faça login.");
  }
  try {
    const res = await msalApp.acquireTokenSilent({
      account: activeAccount,
      scopes
    });
    return res.accessToken;
  } catch (silentErr) {
    log("Silent falhou, tentando acquireTokenPopup…");
    const res = await msalApp.acquireTokenPopup({
      account: activeAccount,
      scopes
    });
    return res.accessToken;
  }
}

// ===== Teste: pegar listas do site via Graph =====
async function testListsTop5() {
  try {
    if (!activeAccount) {
      alert("Faça login primeiro.");
      return;
    }

    const hostname = $("siteHostname").value.trim();          // ex: alufran.sharepoint.com
    const sitePath = $("sitePath").value.trim();              // ex: /sites/DiretoriaAdministrativa9
    const token = await getTokenPopup(["Sites.ReadWrite.All"]);

    // 1) obtem siteId
    // GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{path}
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(hostname)}:/sites${encodeURI(sitePath)}`;
    log(`Graph: ${siteUrl}`);

    let r = await fetch(siteUrl, { headers: { Authorization: `Bearer ${token}` }});
    if (!r.ok) throw new Error(`Erro Graph site: ${r.status} ${await r.text()}`);
    const site = await r.json();

    // 2) listas top 5
    // GET /sites/{id}/lists?$top=5&$select=name,id
    const listsUrl = `https://graph.microsoft.com/v1.0/sites/${site.id}/lists?$top=5&$select=name,id`;
    r = await fetch(listsUrl, { headers: { Authorization: `Bearer ${token}` }});
    if (!r.ok) throw new Error(`Erro Graph lists: ${r.status} ${await r.text()}`);
    const data = await r.json();

    log(`Top 5 listas:`);
    (data.value || []).forEach((it, i) => log(`  ${i+1}) ${it.name} (${it.id})`));
  } catch (e) {
    log(`ERRO teste: ${e.message || e}`);
    alert("Falha no teste. Veja o log.");
  }
}

// ===== UI Bindings =====
window.addEventListener("DOMContentLoaded", () => {
  initMsal();

  $("btnLogin").addEventListener("click", loginPopup);
  $("btnLogout").addEventListener("click", logoutPopup);
  $("btnTest").addEventListener("click", testListsTop5);
});
