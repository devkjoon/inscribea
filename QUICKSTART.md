# Quick Start Guide

Get your Inscribea Outlook add-in up and running in 5 minutes!

## Step 1: Install Dependencies

```bash
npm install
```

Or use the setup script:

```bash
./setup.sh
```

## Step 2: Get Your OpenAI API Key

1. Go to [OpenAI Platform](https://platform.openai.com/api-keys)
2. Sign up or log in
3. Create a new API key
4. Copy the key

## Step 3: Configure Environment

Create a `.env` file (or copy from `env.example`):

```bash
cp env.example .env
```

Edit `.env` and add your API key:

```
OPENAI_API_KEY=sk-your-actual-api-key-here
```

## Step 4: Set Up HTTPS (Required for Outlook)

Outlook add-ins require HTTPS. Choose one option:

### Option A: Self-Signed Certificate (Quickest for Development)

```bash
openssl req -x509 -newkey rsa:4096 -nodes -keyout key.pem -out cert.pem -days 365 -subj "/CN=localhost"
```

**Important**: You'll need to trust this certificate in your browser/OS:
- **macOS**: Open `cert.pem` in Keychain Access, double-click it, expand "Trust", and set to "Always Trust"
- **Windows**: Import `cert.pem` into Trusted Root Certification Authorities

### Option B: Use ngrok (Easiest, No Certificate Setup)

```bash
# Install ngrok
npm install -g ngrok

# Start ngrok (in a separate terminal)
ngrok http 3000
```

Copy the HTTPS URL (e.g., `https://abc123.ngrok.io`) and update `manifest.xml`:
- Replace all instances of `https://localhost:3000` with your ngrok URL

## Step 5: Start the Server

```bash
npm start
```

You should see:
```
✅ Server running on https://localhost:3000
📝 Make sure to configure OPENAI_API_KEY in your .env file
```

## Step 6: Install in Outlook

### Outlook Desktop (Windows/Mac)

1. Open Outlook
2. Go to **File** → **Manage Add-ins** (or **Get Add-ins**)
3. Click **My Add-ins** → **Add a Custom Add-in** → **Add from File**
4. Select `manifest.xml` from this project
5. Click **Install**

### Outlook on the Web

1. Go to [Outlook on the Web](https://outlook.office.com)
2. Click the gear icon → **View all Outlook settings**
3. Go to **Mail** → **Manage add-ins**
4. Click **+ Add a custom add-in** → **Add from file**
5. Upload `manifest.xml`

## Step 7: Use the Add-in

1. Open any email in Outlook
2. Look for the **Inscribea** button in the ribbon
3. Click it to open the task pane
4. Type a prompt like "Write a professional follow-up email"
5. Click **Generate Email**
6. Review and click **Insert into Email** or **Copy to Clipboard**

## Troubleshooting

### "Failed to generate email"
- ✅ Check your `.env` file has the correct `OPENAI_API_KEY`
- ✅ Verify the server is running
- ✅ Check you have OpenAI API credits

### Add-in won't load
- ✅ Verify HTTPS is working (check browser console)
- ✅ If using ngrok, make sure the URL in `manifest.xml` matches
- ✅ Check that the server is accessible from the URL in manifest

### SSL Certificate errors
- ✅ Trust the self-signed certificate (see Step 4)
- ✅ Or use ngrok instead (no certificate needed)

## Need Help?

Check the full [README.md](README.md) for detailed documentation.

