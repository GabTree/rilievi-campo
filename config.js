// Configurazione dell'app. L'unico valore da compilare è clientId,
// che si ottiene registrando l'app su Microsoft Entra (vedi ISTRUZIONI.md).
window.APP_CONFIG = {
  clientId: "",                                            // es. "1a2b3c4d-...."
  authority: "https://login.microsoftonline.com/common",  // account personali e di lavoro
  scopes: ["User.Read", "Files.ReadWrite"],
  appVersion: "0.1.0"
};
