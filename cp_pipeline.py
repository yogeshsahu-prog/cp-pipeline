#!/usr/bin/env python3
"""
CP Pendency pipeline — Gmail -> parse .xlsb -> aggregate -> Drive agg.json
Runs headless on GitHub Actions. Authenticates as the user via a stored OAuth
refresh token (scopes: gmail.readonly + drive). Idempotent: skips if the newest
matching email hasn't changed since the last run.
"""
import os, io, json, base64, datetime as dt
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly",
          "https://www.googleapis.com/auth/drive",
          "https://www.googleapis.com/auth/spreadsheets"]

PROBLEM_PINS = {"System Deactivated", "Embargo", "COD OFF"}

def creds():
    return Credentials(
        None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        scopes=SCOPES,
    )

def latest_attachment(gmail, query):
    """Return (msg_id, filename, bytes) for the newest matching email whose
    attachment actually parses as an xlsb. We don't gate on filename — some
    messages expose it inconsistently — instead we verify by trying to open it."""
    print(f"DEBUG: query received = {query!r}")
    res = gmail.users().messages().list(userId="me", q=query, maxResults=10).execute()
    msgs = res.get("messages", [])
    print(f"DEBUG: {len(msgs)} messages matched")
    if not msgs:
        return None

    candidates = []
    for m in msgs:
        meta = gmail.users().messages().get(
            userId="me", id=m["id"], format="metadata",
            metadataHeaders=["Subject", "Date"]).execute()
        ts = int(meta.get("internalDate", 0))
        hdrs = {h["name"]: h["value"] for h in meta.get("payload", {}).get("headers", [])}
        print(f"DEBUG: id={m['id']} ts={ts} subject={hdrs.get('Subject','?')!r} date={hdrs.get('Date','?')!r}")
        candidates.append((ts, m["id"]))
    candidates.sort(key=lambda x: -x[0])   # newest first

    def walk(payload):
        if payload.get("body", {}).get("attachmentId") and payload.get("filename"):
            return payload["body"]["attachmentId"], payload["filename"]
        for p in payload.get("parts") or []:
            sub = walk(p)
            if sub:
                return sub
        return None

    for ts, msg_id in candidates:
        full = gmail.users().messages().get(userId="me", id=msg_id, format="full").execute()
        hit = walk(full["payload"])
        if hit:
            att_id, fn = hit
            att = gmail.users().messages().attachments().get(
                userId="me", messageId=msg_id, id=att_id).execute()
            data = base64.urlsafe_b64decode(att["data"])
            try:
                pd.ExcelFile(io.BytesIO(data), engine="pyxlsb")  # verify it's a real xlsb
            except Exception as e:
                print(f"DEBUG: id={msg_id} attachment {fn!r} isn't a valid xlsb ({e}), skipping")
                continue
            print(f"DEBUG: picked id={msg_id} file={fn!r}")
            return msg_id, fn, data
        else:
            print(f"DEBUG: id={msg_id} has no attachment at all, skipping")

    print("DEBUG: no candidate had a usable .xlsb attachment")
    return None

def read_state(drive, file_id):
    try:
        buf = io.BytesIO(drive.files().get_media(fileId=file_id).execute())
        return json.loads(buf.getvalue().decode() or "{}")
    except Exception:
        return {}

def write_json(drive, file_id, obj):
    body = io.BytesIO(json.dumps(obj, separators=(",", ":")).encode())
    media = MediaIoBaseUpload(body, mimetype="application/json", resumable=False)
    drive.files().update(fileId=file_id, media_body=media).execute()

REQUIRED_COLS = {"Bucket", "DC Hub Zone", "destination_state", "DC", "Status"}

def pick_data_sheet(xlsb_bytes):
    xl = pd.ExcelFile(io.BytesIO(xlsb_bytes), engine="pyxlsb")
    for name in xl.sheet_names:
        head = pd.read_excel(xl, sheet_name=name, engine="pyxlsb", nrows=0)
        cols = {str(c).strip() for c in head.columns}
        if REQUIRED_COLS.issubset(cols):
            return name
    raise ValueError(f"No data sheet with expected columns in {xl.sheet_names}")

def aggregate(xlsb_bytes, snapshot_label):
    sheet = pick_data_sheet(xlsb_bytes)
    df = pd.read_excel(io.BytesIO(xlsb_bytes), sheet_name=sheet, engine="pyxlsb")
    df.columns = [str(c).strip() for c in df.columns]
    op = df[df["Bucket"] == "Current Open"].copy()
    Z, S, H = "DC Hub Zone", "destination_state", "DC"
    for c in (Z, S, H):
        op[c] = op[c].astype(str).replace("nan", "Unknown").fillna("Unknown")
    op["cp"] = pd.to_numeric(op["customer_promise_date"], errors="coerce")
    today = pd.to_numeric(op["cp"], errors="coerce").min()

    def m(g):
        st = g["Status"].value_counts()
        cp = g["cp"]
        return dict(
            total=int(len(g)),
            rad=int(st.get("RAD", 0)),
            it=int(st.get("Intransit", 0)),
            mis=int(st.get("Misroute_IT", 0) + st.get("Investigate", 0)),
            today=int((cp <= today).sum()),
            tom=int((cp == today + 1).sum()),
            d2=int((cp >= today + 2).sum()),
            noplan=int((g["2167 Pins"] == "No Plan").sum()),
            badpin=int(g["Pincode"].isin(PROBLEM_PINS).sum()),
            cod=int((g["Order Type"] == "COD").sum()),
        )

    zones = []
    for z, gz in op.groupby(Z):
        zd = m(gz); zd["name"] = z; sts = []
        for s, gs in gz.groupby(S):
            sd = m(gs); sd["name"] = s
            sd["hubs"] = sorted([dict(m(gh), name=h) for h, gh in gs.groupby(H)],
                                key=lambda x: -x["total"])
            sts.append(sd)
        zd["states"] = sorted(sts, key=lambda x: -x["total"]); zones.append(zd)

    return dict(
        meta=dict(snapshot=snapshot_label, open=int(len(op)),
                  hubs=int(op[H].nunique()), states=int(op[S].nunique()),
                  zones=int(op[Z].nunique()),
                  updated=dt.datetime.now(dt.timezone.utc)
                          .astimezone(dt.timezone(dt.timedelta(hours=5, minutes=30)))
                          .strftime("%d %b, %I:%M %p IST")),
        overall=m(op),
        zones=sorted(zones, key=lambda x: -x["total"]),
    )

def append_trend(sheets, sheet_id, agg):
    o = agg["overall"]
    row = [[agg["meta"]["updated"], o["total"], o["rad"], o["it"] + o["mis"],
            o["today"], o["noplan"], o["badpin"]]]
    sheets.spreadsheets().values().append(
        spreadsheetId=sheet_id, range="Trend!A1",
        valueInputOption="USER_ENTERED", insertDataOption="INSERT_ROWS",
        body={"values": row}).execute()

def main():
    c = creds()
    gmail = build("gmail", "v1", credentials=c, cache_discovery=False)
    drive = build("drive", "v3", credentials=c, cache_discovery=False)

    hit = latest_attachment(gmail, os.environ["GMAIL_QUERY"])
    if not hit:
        print("No matching CP email found — nothing to do."); return
    msg_id, fn, data = hit

    state = read_state(drive, os.environ["DRIVE_STATE_FILE_ID"])
    if state.get("last_msg_id") == msg_id:
        print(f"Already processed {fn} ({msg_id}); skipping."); return

    label = fn.replace(".xlsb", "").replace("_", " ")
    agg = aggregate(data, label)
    write_json(drive, os.environ["DRIVE_AGG_FILE_ID"], agg)
    write_json(drive, os.environ["DRIVE_STATE_FILE_ID"],
               {"last_msg_id": msg_id, "file": fn, "open": agg["overall"]["total"]})

    if os.environ.get("SHEET_ID"):
        sheets = build("sheets", "v4", credentials=c, cache_discovery=False)
        try:
            append_trend(sheets, os.environ["SHEET_ID"], agg)
        except Exception as e:
            print("Trend append skipped:", e)

    print(f"Updated dashboard from {fn}: {agg['overall']['total']} open shipments.")

if __name__ == "__main__":
    main()
