// ----- Configuração MSAL -----
const msalConfig = {
  auth: {
    clientId: '4de58585-d608-4d47-b7c7-c8ef972e3695',
    authority: 'https://login.microsoftonline.com/d80bf2e9-ace6-421a-86aa-38f9fed30153',
    redirectUri: 'https://rosanesouzzza.github.io/Finance/redirect.html',
    postLogoutRedirectUri: 'https://rosanesouzzza.github.io/Finance/'
  },
  cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
};
const loginRequest = { scopes: ['openid', 'profile', 'offline_access', 'Sites.ReadWrite.All'] };

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Processa o retorno do redirect (se for o caso) e depois toca a vida
msalInstance.handleRedirectPromise()
  .then((response) => {
    if (response) {
      // já temos token de ID; você pode guardar a account
      const acct = response.account;
      if (acct) msalInstance.setActiveAccount(acct);
    }
    onMsalReady();
  })
  .catch((err) => {
    log('ERRO handleRedirect: ' + (err && err.message));
    onMsalReady();
  });

// Função chamada quando o MSAL estiver pronto (com ou sem resposta)
function onMsalReady() {
  log('MSAL pronto.');
  enableUi();

  // Se recebeu postMessage da redirect.html, você pode reagir aqui também (opcional)
  window.addEventListener('message', (ev) => {
    if (ev && ev.data && ev.data.type === 'msal:auth:done') {
      log('Auth concluído (postMessage).');
      // aqui você pode, por ex., buscar o token silenciosamente
      ensureLogged().catch(()=>{ /* ignore */ });
    }
  });
}

// UI helpers
function log(msg){
  const box = document.querySelector('#log') || document.body;
  (box.value !== undefined) ? box.value += msg + '\n' : console.log(msg);
}
function enableUi(){
  document.querySelectorAll('button, a.btn').forEach(b=>b.disabled=false);
}

// Botão ENTRAR (redirect)
async function loginRedirect() {
  try {
    await msalInstance.loginRedirect(loginRequest);
  } catch (e) {
    log('ERRO loginRedirect: ' + e.message);
  }
}
document.getElementById('btnEntrarRedirect')?.addEventListener('click', loginRedirect);

// Garante session/account e tenta token silencioso
async function ensureLogged(){
  let account = msalInstance.getActiveAccount() || (msalInstance.getAllAccounts()[0] || null);
  if (!account) {
    // se quiser forçar login silencioso por ssoSilent, comente a próxima linha
    await msalInstance.loginRedirect(loginRequest);
    return;
  }
  const tokenResp = await msalInstance.acquireTokenSilent({
    ...loginRequest, account
  });
  log('Token OK (scopes): ' + tokenResp.scopes.join(', '));
  return tokenResp.accessToken;
}
