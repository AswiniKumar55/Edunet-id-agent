# Deploy to Glitch — Step by Step

## Step 1 — Create a Glitch account
Go to https://glitch.com → Sign Up (free)

## Step 2 — Create a new project
- Click **New Project** → **glitch-hello-node**
- A new project opens in the Glitch editor

## Step 3 — Upload your files

### Option A — Upload via Glitch dashboard (recommended)
1. In the Glitch editor, click **Tools** (bottom left) → **Import / Export**
2. Click **Upload a file** and upload these files one by one:
   - `server.js`
   - `package.json`
   - `Edunet ids.xlsx`
   - `public/index.html`  ← create the `public` folder first

### Option B — Use Glitch Git import
1. Push your project to GitHub (without node_modules)
2. In Glitch: Tools → Import → Import from GitHub → paste your repo URL

## Step 4 — Set Environment Variables (IMPORTANT)
In Glitch editor → click `.env` file in the left panel → add:

```
PORT=3000
GMAIL_USER=aiaswinikumar@gmail.com
GMAIL_PASS=pcdodjtvnhlfgxmp
```

> Glitch `.env` is private — never shown to visitors ✅

## Step 5 — Fix package.json start script
Make sure `package.json` has:
```json
"scripts": {
  "start": "node server.js"
}
```

## Step 6 — Your app is live!
Glitch gives you a URL like:
`https://your-project-name.glitch.me`

Share this link with anyone — they can access the Edunet ID Agent from any browser!

## Notes
- Glitch free tier sleeps after 5 min of inactivity (wakes up in ~30 sec on next visit)
- To keep it always awake: use https://uptimerobot.com — add a monitor pinging your Glitch URL every 5 min (free)
- Your `state.json` (pool, history, employees) persists on Glitch's file system permanently
- Upgrade to Glitch Pro ($8/mo) to remove sleep and get custom domain
