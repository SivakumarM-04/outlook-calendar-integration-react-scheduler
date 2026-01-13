// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const config = {
  appId: 'Your-App-ID',
  redirectUri: 'http://localhost:3000',
  scopes: [
    'user.read',
    'mailboxsettings.read',
    'calendars.readwrite'
  ]
};

export default config;
