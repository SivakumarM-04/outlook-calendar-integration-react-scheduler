export const msalConfig = {
  auth: {
    clientId:"YOUR_CLIENT_ID",      
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:5173"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: ["User.Read", "Calendars.Read", "Calendars.ReadWrite"]
};