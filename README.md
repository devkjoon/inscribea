# Inscribea

A lightweight Outlook-integrated AI assistant that helps you write emails faster and more professionally. This add-in adds a dedicated task-pane inside Outlook where you can type a prompt, and the assistant will generate a complete email draft—reply, follow-up, clarification, escalation, or anything else you need—powered by the OpenAI API.

> 🚀 **New to this project?** Check out the [Quick Start Guide](QUICKSTART.md) to get up and running in 5 minutes!

## Features

- 🤖 **AI-Powered Email Generation**: Generate professional emails using ChatGPT
- 📧 **Email Context Awareness**: Automatically reads the current email to provide context-aware responses
- ✨ **One-Click Insert**: Insert generated emails directly into your Outlook compose window
- 📋 **Copy to Clipboard**: Easily copy generated emails for use elsewhere
- 🎨 **Modern UI**: Clean, intuitive interface that integrates seamlessly with Outlook

## Prerequisites

- Node.js (v14 or higher)
- npm or yarn
- Outlook (desktop or web)
- OpenAI API key ([Get one here](https://platform.openai.com/api-keys))

## Setup Instructions

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure Environment Variables

Create a `.env` file in the root directory:

```bash
cp .env.example .env
```

Edit `.env` and add your OpenAI API key:

```
OPENAI_API_KEY=your_openai_api_key_here
OPENAI_MODEL=gpt-4
PORT=3000
```

### 3. Start the Server

```bash
npm start
```

For development with auto-reload:

```bash
npm run dev
```

The server will start on `https://localhost:3000`

### 4. Install the Add-in in Outlook

#### For Outlook Desktop (Windows/Mac):

1. Open Outlook
2. Go to **File** → **Manage Add-ins** (or **Get Add-ins**)
3. Click **My Add-ins** → **Add a Custom Add-in** → **Add from File**
4. Select the `manifest.xml` file from this project
5. The add-in should appear in your Outlook ribbon

#### For Outlook on the Web:

1. Go to [Outlook on the Web](https://outlook.office.com)
2. Click the gear icon (Settings) → **View all Outlook settings**
3. Go to **Mail** → **Manage add-ins**
4. Click **+ Add a custom add-in** → **Add from file**
5. Upload the `manifest.xml` file

### 5. Trust the Localhost Certificate (Required for HTTPS)

Since Outlook requires HTTPS, you'll need to set up SSL for localhost. You have a few options:

**Option A: Use a self-signed certificate (Development only)**

1. Generate a self-signed certificate:
   ```bash
   openssl req -x509 -newkey rsa:4096 -nodes -keyout key.pem -out cert.pem -days 365 -subj "/CN=localhost"
   ```

2. Update `server.js` to use HTTPS (see HTTPS setup section below)

**Option B: Use ngrok or similar tunneling service**

1. Install ngrok: `npm install -g ngrok`
2. Run: `ngrok http 3000`
3. Update the URLs in `manifest.xml` to use the ngrok URL

**Option C: Use Office Add-in development tools**

Microsoft provides tools for local development. Consider using [Office Add-in development tools](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/develop-add-ins-visual-studio-code).

## Usage

1. Open an email in Outlook (read or compose)
2. Click the **Inscribea** button in the ribbon
3. The task pane will open on the right side
4. Optionally click **"Use Current Email Context"** to load the current email's details
5. Type your prompt (e.g., "Write a professional follow-up email")
6. Click **"Generate Email"**
7. Review the generated email
8. Click **"Insert into Email"** to add it to your compose window, or **"Copy to Clipboard"** to copy it

## Project Structure

```
inscribea/
├── manifest.xml          # Outlook add-in manifest
├── server.js             # Express server with OpenAI integration
├── package.json          # Node.js dependencies
├── .env.example          # Environment variables template
├── .gitignore            # Git ignore file
└── public/
    ├── taskpane.html     # Main UI HTML
    ├── taskpane.css      # Styles
    ├── taskpane.js       # Client-side JavaScript
    └── commands.html     # Command handlers
```

## API Endpoints

### POST `/api/generate-email`

Generates an email based on a prompt and optional email context.

**Request Body:**
```json
{
  "prompt": "Write a professional follow-up email",
  "emailContext": {
    "subject": "Meeting Request",
    "from": "sender@example.com",
    "body": "Original email body..."
  }
}
```

**Response:**
```json
{
  "email": "Generated email text...",
  "model": "gpt-4",
  "usage": {
    "prompt_tokens": 100,
    "completion_tokens": 200,
    "total_tokens": 300
  }
}
```

## Development

### Running in Development Mode

```bash
npm run dev
```

This uses `nodemon` to automatically restart the server when files change.

### Testing Locally

1. Make sure the server is running
2. Open Outlook and load the add-in
3. Test the email generation functionality

## Troubleshooting

### "Failed to generate email" Error

- Make sure your OpenAI API key is correctly set in `.env`
- Verify the server is running on the correct port
- Check that you have sufficient OpenAI API credits

### Add-in Not Loading

- Verify the URLs in `manifest.xml` match your server URL
- Check browser/Outlook console for errors
- Ensure HTTPS is properly configured (required for Outlook add-ins)

### CORS Errors

- The server includes CORS middleware, but if you see CORS errors, verify the `cors` package is installed and configured correctly

## Security Notes

- **Never commit your `.env` file** - it contains your API key
- Use environment variables for all sensitive configuration
- For production, use a proper SSL certificate (not self-signed)
- Consider implementing rate limiting for the API endpoint

## License

MIT License - see LICENSE file for details

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
