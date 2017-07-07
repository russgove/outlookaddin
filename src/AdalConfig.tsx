const adalConfig: adal.Config = {
  clientId: 'f8f8d2ad-7c9d-4aac-80eb-3f00a263c879',//this can only read greaph
  tenant: 'common',
  extraQueryParameter: 'nux=1',
  endpoints: {
    'https://graph.microsoft.com': 'https://graph.microsoft.com'
  },
  postLogoutRedirectUri: window.location.origin,
  cacheLocation: 'sessionStorage'
};

export default adalConfig;
