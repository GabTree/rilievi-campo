// Configurazione dell'app. L'unico valore da compilare è clientId,
// che si ottiene registrando l'app su Microsoft Entra (vedi ISTRUZIONI.md).
window.APP_CONFIG = {
  clientId: "bc3c6044-0421-47e1-a2b8-710c3e00f252",          // Microsoft (OneDrive): ID applicazione (client), es. "1a2b3c4d-...."
  googleClientId: "",    // Google Drive: ID client OAuth, es. "1234-abcd.apps.googleusercontent.com"
  authority: "https://login.microsoftonline.com/common",  // account personali e di lavoro
  scopes: ["User.Read", "Files.ReadWrite"],
  appVersion: "0.2.1"
};
