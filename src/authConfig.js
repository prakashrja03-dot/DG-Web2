export const msalConfig = {
  auth: {
    clientId: "fe3ae1e7-579c-4d52-a859-90782213d180",
    authority: "https://login.microsoftonline.com/87491d46-94d5-4ca3-b8fa-19aa35a0226c",
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
};

export const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All"]
};
