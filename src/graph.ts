import { type IPublicClientApplication, type AccountInfo } from "@azure/msal-browser";

async function graphFetch(
  instance: IPublicClientApplication,
  account: AccountInfo,
  url: string,
  scopes: string[],
  init: RequestInit = {}
) {
  const token = await instance.acquireTokenSilent({ scopes, account });
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token.accessToken}`);
  return fetch(url, { ...init, headers });
}

// Get user mailbox settings (timezone)
export async function getMailboxTimeZone(
  instance: IPublicClientApplication,
  account: AccountInfo
) {
  const res = await graphFetch(
    instance,
    account,
    "https://graph.microsoft.com/v1.0/me/mailboxSettings?$select=timeZone",
    ["User.Read"]
  );
  if (!res.ok) {
    return "UTC";
  }
  const data = await res.json();
  return data?.timeZone || "UTC";
}

// List events for a range using /calendarView (handles pagination)
export async function listEventsForRange(
  instance: IPublicClientApplication,
  account: AccountInfo,
  startISO: string,
  endISO: string,
  windowsTz: string // "UTC" for read path, or a Windows TZ
) {
  const base = new URL("https://graph.microsoft.com/v1.0/me/calendarView");
  base.searchParams.set("startDateTime", startISO);
  base.searchParams.set("endDateTime", endISO);
  base.searchParams.set("$select", "id,subject,start,end,location,isAllDay,bodyPreview");
  base.searchParams.set("$orderby", "start/dateTime");
  base.searchParams.set("$top", "50");

  let url = base.toString();
  let all: any[] = [];

  while (url) {
    const res = await graphFetch(instance, account, url, ["Calendars.Read"], {
      headers: { Prefer: `outlook.timezone="${windowsTz}"` }
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "<no-body>");
      throw new Error(`Graph calendarView error ${res.status}: ${txt} -- URL: ${url}`);
    }

    const data = await res.json();
    all = all.concat(data.value || []);
    url = data["@odata.nextLink"] || "";
  }

  return all;
}

// Create event
export async function createEvent(
  instance: IPublicClientApplication,
  account: AccountInfo,
  subject: string,
  startISO: string,
  endISO: string,
  windowsTz: string,
  body?: string,
  location?: string
) {
  const payload: any = {
    subject,
    start: { dateTime: startISO, timeZone: windowsTz },
    end: { dateTime: endISO, timeZone: windowsTz }
  };
  if (body) payload.body = { contentType: "HTML", content: body };
  if (location) payload.location = { displayName: location };

  const res = await graphFetch(
    instance,
    account,
    "https://graph.microsoft.com/v1.0/me/events",
    ["Calendars.ReadWrite"],
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    }
  );
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

// Update event
export async function updateEvent(
  instance: IPublicClientApplication,
  account: AccountInfo,
  eventId: string,
  partial: any,
  ifMatch?: string
) {
  const headers: Record<string, string> = { "Content-Type": "application/json" };
  if (ifMatch) headers["If-Match"] = ifMatch;

  const res = await graphFetch(
    instance,
    account,
    `https://graph.microsoft.com/v1.0/me/events/${eventId}`,
    ["Calendars.ReadWrite"],
    {
      method: "PATCH",
      headers,
      body: JSON.stringify(partial)
    }
  );
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

// Delete event
export async function deleteEvent(
  instance: IPublicClientApplication,
  account: AccountInfo,
  eventId: string
) {
  const res = await graphFetch(
    instance,
    account,
    `https://graph.microsoft.com/v1.0/me/events/${eventId}`,
    ["Calendars.ReadWrite"],
    { method: "DELETE" }
  );
  if (!res.ok) throw new Error(await res.text());
}
