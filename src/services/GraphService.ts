import { Client } from '@microsoft/microsoft-graph-client';
import config from '../Config';
async function getAuthenticatedClient() {
  const msal = window.msal as any;
  var accessToken = await msal.acquireTokenSilent({
    scopes: config.scopes
  });
  // Initialize Graph client
  const client = Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: done => {
      done(null, accessToken.accessToken);
    }
  });

  return client;
}

export async function getUserDetails() {
  const client = await getAuthenticatedClient();

  const user = await client.api('/me').get();
  return user;
}

export async function getEvents() {
  const client = await getAuthenticatedClient();

  const events = await client
    .api('/me/events')
    .select('subject,organizer,start,end')
    .orderby('createdDateTime DESC')
    .get();

  return events;
}

export async function getWorksheets(id: string) {
  const client = await getAuthenticatedClient();
  const worksheet = await client
    .api(`/me/drive/items/${id}/workbook/worksheets`)
    .select('name')
    .get();
  return worksheet;
}
export async function searchFiles(searchText: string) {
  const client = await getAuthenticatedClient();
  const results = await client
    .api(`/me/drive/root/search(q='${searchText}')?select=name,id,webUrl`)
    .get();
  return results;
}
export async function getTable(documentId: string, sheetName: string) {
  const client = await getAuthenticatedClient();
  const results = await client
    .api(`/me/drive/items/${documentId}/workbook/worksheets('${sheetName}')/usedRange`)
    .get();

  return results;
}
