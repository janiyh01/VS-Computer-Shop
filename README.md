# VS System

VS System is a local stock, billing, repair, barcode, due-payment, and report app.

## Run On This PC

```powershell
npm install
npm run web
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
npm start
```

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

The web server uses the same app-data database path as the Electron app:

```text
%APPDATA%\VS Software\VS-System\yard.db
```

The database and runtime cache files are intentionally ignored by git.
