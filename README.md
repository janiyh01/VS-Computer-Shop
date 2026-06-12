# VS System

VS System is a local stock, billing, repair, barcode, due-payment, and report app.

## Run On This PC

```powershell
npm install
npm start
```

Open:

```text
http://localhost:3000
```

Default app login:

```text
Username: admin
Password: 1234
```

## Run On Phone Or Laptop On Same Wi-Fi

```powershell
npm run web:lan
```

Open the printed LAN URL from the other device, for example:

```text
http://192.168.8.183:3000
```

LAN mode keeps a browser password in front of the app for safety. The terminal prints that browser username/password. After that, use the normal app login.

## Desktop Electron Mode

```powershell
npm run desktop
```

## Deploy On Render With Turso

This project can run locally with SQLite and in production with Turso/libSQL.
When `TURSO_DATABASE_URL` is present, the server uses Turso automatically.

Render settings:

```text
Build command: npm install
Start command: npm start
```

Render environment variables:

```text
HOST=0.0.0.0
BROWSER_PASSWORD=1
WEB_USER=vs
WEB_PASSWORD=<choose a strong browser password>
TURSO_DATABASE_URL=<your Turso database URL>
TURSO_AUTH_TOKEN=<your Turso auth token>
```

Render also provides `PORT` automatically. Keep the browser password on for
deployed usage, then use the normal app login after the browser prompt.

The included `render.yaml` contains the same deployment settings.

## Deploy With Local SQLite

Use this only on a Node.js host that provides persistent disk storage.

```text
Build command: npm install
Start command: npm start
Environment:
  HOST=0.0.0.0
  PORT=<provided by host>
  DB_PATH=<persistent disk path>/yard.db
```

If the host does not provide persistent disk storage, SQLite data can reset when
the service restarts. Turso is recommended for Render.

## Main Features

- Product and stock management
- Auto barcode generation
- Barcode label printing
- Billing and invoice printing
- Paid amount and balance due tracking
- Customer profile and due list
- Repair jobs and repair bill printing
- Dashboard and reports
- Business preferences
- Backup and restore tools

## Data

Without Turso, the web server uses the same app-data database path as the
Electron app:

```text
%APPDATA%\VS Software\VS-System\yard.db
```

The database and runtime cache files are intentionally ignored by git.
