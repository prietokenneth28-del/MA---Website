// Configuración de MSAL para Microsoft Entra ID (Azure AD)
export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || "00000000-0000-0000-0000-000000000000",
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || "common"}`,
    redirectUri: import.meta.env.VITE_AZURE_REDIRECT_URI || window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  }
};

// Permisos requeridos para Graph API (lectura y escritura en OneDrive)
export const loginRequest = {
  scopes: ["User.Read", "Files.ReadWrite.All"]
};

// Comprueba si Azure Client ID está configurado o si estamos en Modo Simulación
export const isConfiguredAzure = () => {
  return (
    import.meta.env.VITE_AZURE_CLIENT_ID && 
    import.meta.env.VITE_AZURE_CLIENT_ID !== "your-azure-app-client-id-here" &&
    import.meta.env.VITE_AZURE_CLIENT_ID !== "00000000-0000-0000-0000-000000000000"
  );
};

