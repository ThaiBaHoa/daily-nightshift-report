export const msalConfig = {
    auth: {
        clientId: "YOUR_CLIENT_ID", // Cần thay thế bằng Client ID từ Azure Portal
        authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // Thay thế Tenant ID
        redirectUri: "http://localhost:3000"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

export const loginRequest = {
    scopes: [
        'Files.ReadWrite',
        'Files.ReadWrite.All',
        'Sites.ReadWrite.All'
    ]
};
