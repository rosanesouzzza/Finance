/* ==== util de log ======================================================= */
const $ = (s)=>document.querySelector(s);
const log = (m)=>{ const d=new Date().toTimeString().slice(0,8); $("#log").textContent += `[${d}] ${m}\n`; };

/* ==== preencher redirectUri e ganchos de UI ============================= */
(function bootstrapUI(){
  const redirect = location.origin + location.pathname.replace(/\/[^/]*$/, '') + '/redirect.html';
  $("#redirectUri").value = redirect;

  $("#btnLogin").addEventListener("click", () => loginRedirect());
  $("#btnLogout").addEventListener("click", () => logout());
  $("#btnTest").addEventListener("click", () => testLists());
  log("Pronto para conectar…");
})();

/* ==== MSAL config dinâmico ============================================= */
function getConfig() {
  const tenantId = $("#tenantId").value.trim();
  const clientId = $("#clientId").value.trim();
  const redirectUri = $("#redirectUri").value.trim();

  return {
    tenantId, clientId, redirectUri,
    msal: {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri
      },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
      system: { allowNativeBroker: false }
    },
    scopes: ["openid","profile","offline_access","Sites.ReadWrite.All"]
  };
}

/* ==== instancia única do MSAL ========================================== */
let pca;
function ensureMsal(){
  if (!window.msal) { log("[ERRO] MSAL não disponível. Verifique vendor/msal-browser.min.js e a ordem dos <script>."); throw new Error("msal not loaded"); }
  if (!pca) pca = new msal.PublicClientApplication(getConfig().msal);
  return pca;
}

/* ==== fluxo redirect: login e tratamento de retorno ===================== */
async function handleRedirect() {
  ensureMsal();
  try {
    await pca.handleRedirectPromise();
  } catch (e) {
    log(`[ERRO] handleRedirect: ${e.message || e}`);
  }
}
handleRedirect();

function loginRedirect(){
  const cfg = getConfig();
  ensureMsal();
  log("Abrindo popup/aba de login…");
  pca.loginRedirect({ scopes: cfg.scopes, prompt: "select_account" });
}

async function logout(){
  const cfg = getConfig();
  ensureMsal();
  const acc = (await pca.getTokenCache().getAllAccounts())[0];
  if (!acc) { log("Nenhuma sessão para encerrar."); return; }
  await pca.logoutRedirect({ account: acc, postLogoutRedirectUri: cfg.redirectUri });
}

/* ==== token Graph + chamadas =========================================== */
async function getGraphToken() {
  const cfg = getConfig();
  ensureMsal();

  let account = (await pca.getTokenCache().getAllAccounts())[0];
  if (!account) {
    log("Nenhuma sessão. Faça login.");
    throw new Error("no session");
  }

  try {
    const r = await pca.acquireTokenSilent({ account, scopes: cfg.scopes });
    return r.accessToken;
  } catch {
    // fallback
    await pca.acquireTokenRedirect({ account, scopes: cfg.scopes });
    return new Promise(()=>{}); // a execução retorna após o redirect
  }
}

/* ==== helpers Graph para montar o siteId e ler listas =================== */
async function getSiteId(accessToken, host, sitePath) {
  const trimmed = sitePath.replace(/^\/+|\/+$/g,''); // sem barras
  const url = `https://graph.microsoft.com/v1.0/sites/${host}:/${trimmed}`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` }});
  if (!res.ok) throw new Error(`Graph site lookup falhou: ${res.status}`);
  const json = await res.json();
  return json.id; // form: tenant,site,web ids
}

async function getTopLists(accessToken, siteId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$top=5&$select=id,displayName,webUrl&$orderby=createdDateTime desc`;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` }});
  if (!res.ok) throw new Error(`Graph lists falhou: ${res.status}`);
  const json = await res.json();
  return json.value || [];
}

/* ==== ação: testar acesso ============================================== */
async function testLists(){
  try{
    const host = $("#siteHost").value.trim();
    const path = $("#sitePath").value.trim();

    log("Adquirindo token (Graph)…");
    const token = await getGraphToken();

    log("Obtendo siteId…");
    const siteId = await getSiteId(token, host, path);
    log(`siteId: ${siteId}`);

    log("Lendo 5 listas…");
    const lists = await getTopLists(token, siteId);
    if (!lists.length) { log("Nenhuma lista encontrada."); return; }

    lists.forEach((l,i)=>log(`${i+1}. ${l.displayName} — ${l.webUrl}`));
    log("OK ✅");
  }catch(e){
    log(`[ERRO] ${e.message || e}`);
  }
}

