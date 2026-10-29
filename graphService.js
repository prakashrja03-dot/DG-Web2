import { loginRequest } from "./authConfig";

export async function getAccessToken(instance, accounts) {
  const response = await instance.acquireTokenSilent({
    ...loginRequest,
    account: accounts[0]
  });
  return response.accessToken;
}

async function callGraph(token, url, method = "GET", body = null) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : null
  });
  if (!res.ok) throw new Error(`Graph error ${res.status}: ${await res.text()}`);
  return res.status === 204 ? null : res.json();
}

export async function getSiteId(token, hostname, sitePath) {
  const data = await callGraph(token, `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`);
  return data.id;
}

export async function getListId(token, siteId, listName) {
  const data = await callGraph(token, `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}`);
  return data.id;
}

export async function getListItems(token, siteId, listId) {
  const data = await callGraph(
    token,
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`
  );
  return data.value.map((item) => ({ id: item.id, ...item.fields }));
}

export async function updateListItem(token, siteId, listId, itemId, fields) {
  return callGraph(
    token,
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
    "PATCH",
    fields
  );
}
