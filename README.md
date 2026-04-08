# Nubra Insti Excel Plugin

Minimal Excel Office add-in for institutional login only.

Current scope:
- Reuses the existing Nubra icon assets and Office add-in structure
- Supports Insti login via `POST /login-insti`
- Supports MPIN verification via `POST /verifypin`
- Stores `x-device-id`, `auth_token`, and `session_token` in browser storage for the add-in session

Run:

```powershell
npm install
npm run setup
npm run start
```
