// Configurazione dell'app. L'unico valore da compilare è clientId,
// che si ottiene registrando l'app su Microsoft Entra (vedi ISTRUZIONI.md).
window.APP_CONFIG = {
  clientId: "bc3c6044-0421-47e1-a2b8-710c3e00f252",          // Microsoft (OneDrive): ID applicazione (client), es. "1a2b3c4d-...."
  googleClientId: "",    // Google Drive: ID client OAuth, es. "1234-abcd.apps.googleusercontent.com"
  // Satellite: chiave API ArcGIS Location Platform (permesso Basemaps), scade il ~20/11/2027
  esriKey: "AAPTaIYu1gXH5q2uMAaVXqR94Zg..3fvMFmxmwZKL9824R38K7xDWbOoThgWq-mxfil9jeSVOqyMgCoFehQwmKQFXNh8K-wCdph-8xTQdhGON-5q1oZUNd1xHMEtcslqF_gcrG4pkTVYg9AXiQ7lH7YqxRZsGV7NgygG_704Qr5fJnA1ZpFiwe27VVItikzq-gPb5Gs_SmxKnKBwjZaOKqtG0UUAxAHI_iPdhy74xKn8-28hMxhnqnf_Z2n6IoED3_R0CIMS4IzvOR1iJAT1_vppcQJNi",
  authority: "https://login.microsoftonline.com/common",  // account personali e di lavoro
  scopes: ["User.Read", "Files.ReadWrite"],
  appVersion: "0.4.2"
};
