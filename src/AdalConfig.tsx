const adalConfig: adal.Config = {
  instance:'https://login.microsoftonline.com/',
  clientId: 'f8f8d2ad-7c9d-4aac-80eb-3f00a263c879',//this can only read greaph
  tenant: 'common',
  extraQueryParameter: 'nux=1',
  endpoints: {
    'https://rgove3.sharepoint.com': 'https://rgove3.sharepoint.com'
  },
  postLogoutRedirectUri: window.location.origin,
  cacheLocation: 'sessionStorage',
  popUp: false,
  navigateToLoginRequestUrl:false

};

export default adalConfig;
