import { useRef, useState, useEffect, useCallback } from "react";
import {
  ScheduleComponent, Day, Week, WorkWeek, Month, Agenda, Inject,
  DragAndDrop,
  Resize
} from "@syncfusion/ej2-react-schedule";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import {
  listEventsForRange, createEvent, updateEvent, deleteEvent, getMailboxTimeZone
} from "./graph";
import { loginRequest } from "./authConfig";

type SfEvent = {
  Id: string;
  Subject: string;
  StartTime: Date;
  EndTime: Date;
  IsAllDay?: boolean;
  Location?: string;
  Description?: string;
  _etag?: string;
};

export default function App() {
  const { instance, accounts } = useMsal();
  const isAuth = useIsAuthenticated();
  const account = accounts?.[0];

  const scheduleRef = useRef<ScheduleComponent | null>(null);
  const [events, setEvents] = useState<SfEvent[]>([]);
  const [tz, setTz] = useState("UTC");

  // ---- Auth ----
  const signIn = async () => {
    try { await instance.loginRedirect(loginRequest); }
    catch { await instance.loginPopup(loginRequest).catch(console.error); }
  };
  const signOut = async () => {
    try { await instance.logoutRedirect({ account }); }
    catch { await instance.logoutPopup({ account, mainWindowRedirectUri: "/" }).catch(console.error); }
    finally { setEvents([]); }
  };

  // ---- TZ for create/update semantics ----
  useEffect(() => {
    if (isAuth && account) {
      getMailboxTimeZone(instance, account).then(setTz).catch(() => setTz("UTC"));
    }
  }, [isAuth, account, instance]);

  // ---- Helpers ----
  const pad = (n: number) => String(n).padStart(2, "0");
  const localDateToGraph = (d: Date) =>
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;

  const parseGraphDate = (dt?: string, zone?: string) => {
    if (!dt) return new Date();
    // When reading with Prefer: "UTC", Graph often returns "timeZone: UTC" without 'Z'
    const needsZ = zone && zone.toUpperCase() === "UTC" && !/[zZ]|[+\-]\d{2}:\d{2}$/.test(dt);
    return new Date(needsZ ? `${dt}Z` : dt);
  };

  const mapGraphToSf = (g: any): SfEvent => ({
    Id: g.id,
    Subject: g.subject || "(No subject)",
    StartTime: parseGraphDate(g.start?.dateTime, g.start?.timeZone),
    EndTime: parseGraphDate(g.end?.dateTime, g.end?.timeZone),
    IsAllDay: !!g.isAllDay,
    Location: g.location?.displayName,
    Description: g.bodyPreview,
    _etag: g["@odata.etag"]
  });

  // ---- Load visible range (UTC read path, Sunday-safe) ----
  const loadRange = useCallback(async () => {
    if (!scheduleRef.current || !account) return;

    const days = scheduleRef.current.getCurrentViewDates?.() || [];
    if (!days.length) return;

    const start = new Date(days[0]); start.setHours(0, 0, 0, 0);
    const last = new Date(days[days.length - 1]);
    const end = new Date(last); end.setDate(end.getDate() + 1); end.setHours(0, 0, 0, 0);

    try {
      const items = await listEventsForRange(
        instance, account, start.toISOString(), end.toISOString(), "UTC"
      );
      setEvents(items.map(mapGraphToSf));
    } catch (e) {
      console.error("Load failed", e);
    }
  }, [account, instance]);

  useEffect(() => { if (isAuth) loadRange(); }, [isAuth, tz, loadRange]);

  // ---- CRUD (simple, with rollback by reload on error) ----
  const onActionBegin = async (args: any) => {
    if (!isAuth) return;
    const e = Array.isArray(args.data) ? args.data[0] : args.data;
    if (!e) return;

    const finish = async (fn: () => Promise<void>) => {
      args.cancel = true;
      try { await fn(); }
      catch (err) { console.error(err); await loadRange(); }
    };

    if (args.requestType === "eventCreate") {
      await finish(async () => {
        const r = await createEvent(
          instance, account!,
          e.Subject, localDateToGraph(new Date(e.StartTime)),
          localDateToGraph(new Date(e.EndTime)), tz, e.Description, e.Location
        );
        setEvents(prev => [...prev, mapGraphToSf(r)]);
      });
    }

    if (args.requestType === "eventChange") {
      await finish(async () => {
        const r = await updateEvent(
          instance, account!, e.Id, {
            subject: e.Subject,
            start: { dateTime: localDateToGraph(new Date(e.StartTime)), timeZone: tz },
            end:   { dateTime: localDateToGraph(new Date(e.EndTime)),   timeZone: tz },
            location: e.Location ? { displayName: e.Location } : undefined
          }, e._etag
        );
        const mapped = mapGraphToSf(r);
        setEvents(prev => prev.map(x => x.Id === mapped.Id ? mapped : x));
      });
    }

    if (args.requestType === "eventRemove") {
      await finish(async () => {
        await deleteEvent(instance, account!, e.Id);
        setEvents(prev => prev.filter(x => x.Id !== e.Id));
      });
    }
  };

  return (
    <div style={{ padding: 16 }}>
      {!isAuth ? (
        <button onClick={signIn}>Sign in with Microsoft</button>
      ) : (
        <>
          <div style={{ marginBottom: 8, display: "flex" }}>
            <button onClick={signOut}>Sign out</button>
          </div>

          <ScheduleComponent
            ref={scheduleRef}
            height="650px"
            showWeekend={true}
            allowDragAndDrop={true}
            allowResizing={true}
            eventSettings={{
              dataSource: events,
              fields: {
                id: "Id", subject: "Subject",
                startTime: "StartTime", endTime: "EndTime",
                isAllDay: "IsAllDay", location: { name: "Location" },
                description: { name: "Description" }
              }
            }}
            actionBegin={onActionBegin}
            actionComplete={(a) => {
              if (["viewNavigate", "dateNavigate"].includes(a.requestType)) loadRange();
            }}
          >
            <Inject services={[Day, Week, WorkWeek, Month, Agenda,DragAndDrop,Resize]} />
          </ScheduleComponent>
        </>
      )}
    </div>
  );
}
