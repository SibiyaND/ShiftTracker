# ShiftTrack — Staff Clock-In System

GPS-verified staff attendance tracking with role-based access and Excel export.

## Quick start

Open `index.html` in any modern browser. No build step, no dependencies to install — it's a single self-contained file.

## Deploy options

### Option A — Static hosting (simplest)
Upload `index.html` to any static host:
- **Netlify**: drag the file onto netlify.com/drop
- **Vercel**: `npx vercel --static index.html`
- **GitHub Pages**: push to a repo, enable Pages
- **Any web server**: copy to `/var/www/html/` or equivalent

### Option B — PWA (mobile install)
Add these two files alongside `index.html` to enable "Add to Home Screen":

**manifest.json**
```json
{
  "name": "ShiftTrack",
  "short_name": "ShiftTrack",
  "start_url": "/",
  "display": "standalone",
  "background_color": "#0a0a0f",
  "theme_color": "#6c63ff",
  "icons": [{"src": "icon-192.png","sizes": "192x192","type": "image/png"}]
}
```

Add to `<head>` in index.html:
```html
<link rel="manifest" href="manifest.json"/>
```

### Option C — Node.js server
```bash
npm init -y
npm install express
node -e "require('express')().use(require('express').static('.')).listen(3000)"
# Open http://localhost:3000
```

## Test staff numbers

| Number | Name | Status |
|---|---|---|
| EMP-00421 | Sipho Mokoena | 142.5h — near limit |
| EMP-00138 | Lerato Dube | 163h — at limit |
| EMP-00274 | Thabo Nkosi | 98h — on track |
| EMP-00512 | Aisha Patel | 55h — on track |
| EMP-00389 | Nomsa Khumalo | 161.5h — at limit |
| EMP-00601 | James Dlamini | 120h — on track |
| EMP-00712 | Zanele Motha | 88.5h — on track |
| EMP-00834 | Pieter van Wyk | 175h — near limit |

## Role access

| Feature | Staff | Manager | Admin |
|---|:---:|:---:|:---:|
| Clock in / out | ✅ | — | — |
| View own hours | ✅ | — | — |
| Live site dashboard | — | ✅ | ✅ |
| Export: Monthly timesheet | — | ✅ | ✅ |
| Export: GPS flagged shifts | — | ✅ | ✅ |
| Export: Hours limit warning | — | ✅ | ✅ |
| Export: Site summary | — | — | ✅ |
| Export: Full audit log | — | — | ✅ |
| Manage sites | — | — | ✅ |
| Manage staff | — | — | ✅ |

## Production checklist

Before going live, wire up these backend integrations:

- [ ] Replace `STAFF` object with API call to your staff database
- [ ] Store clock events to a database (PostgreSQL, Firebase, Supabase, etc.)
- [ ] Validate staff number server-side (not just in JS)
- [ ] Enforce 180hr limit server-side — never trust client-side only
- [ ] Replace mock GPS with `navigator.geolocation.getCurrentPosition()`
- [ ] Validate GPS server-side against site geofences
- [ ] Add authentication (JWT tokens, session cookies)
- [ ] Move role checks to server — never expose admin/manager data to staff via API
- [ ] Enable HTTPS (required for `navigator.geolocation` in production)

## GPS integration (browser)

Replace the mock GPS with real coordinates by adding this to the login flow:

```javascript
navigator.geolocation.getCurrentPosition(
  (pos) => {
    const { latitude, longitude, accuracy } = pos.coords;
    // Send to server for geofence validation
  },
  (err) => {
    // Handle denied / unavailable
  },
  { enableHighAccuracy: true, timeout: 10000 }
);
```

## Export files

All exports download directly as `.xlsx` files — compatible with Microsoft Excel and Google Sheets (File → Import).

| Export | Sheets included | Who |
|---|---|---|
| Monthly timesheet | Summary + Shift Log | Manager + Admin |
| GPS flagged shifts | Shift Log (filtered) | Manager + Admin |
| Hours limit warning | Summary (filtered) | Manager + Admin |
| Site activity summary | Site Summary | Admin only |
| Full audit log | Audit Log | Admin only |
| Full export | All sheets | Admin only |
