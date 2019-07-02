export default {
  appId: process.env.REACT_APP_AD_CLIENT_ID || '',
  scopes: ['user.read', 'calendars.read', 'files.read']
};
