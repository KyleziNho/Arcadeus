# Local Testing Instructions for Firebase Authentication

## Quick Start

1. **Start the local server:**
   ```bash
   python3 server.py
   ```

2. **Open your browser to:**
   ```
   http://localhost:8000/login.html
   ```

## Why a Local Server?

Firebase Authentication doesn't work with `file://` URLs due to security restrictions. By running a local HTTP server, you can test authentication features properly.

## Requirements

- Python 3 (comes pre-installed on macOS)
- Your Firebase configuration in `login.html`

## Alternative Methods

If you prefer not to use the provided server.py, you can use:

### Python's built-in server (without CORS headers):
```bash
python3 -m http.server 8000
```

### Node.js http-server (if installed):
```bash
npx http-server -p 8000
```

### VS Code Live Server Extension:
- Install the "Live Server" extension
- Right-click on `login.html`
- Select "Open with Live Server"

## Troubleshooting

1. **Port already in use:** Change the PORT variable in server.py to another number (e.g., 8080, 3000)

2. **Firebase errors:** Make sure your Firebase project is configured to allow localhost:8000 as an authorized domain:
   - Go to Firebase Console → Authentication → Settings → Authorized domains
   - Add `localhost` if not already present

3. **Python not found:** On macOS, use `python3` instead of `python`

## Security Note

This server is for development/testing only. Never use it in production.