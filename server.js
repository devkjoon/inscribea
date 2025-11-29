const express = require('express');
const cors = require('cors');
const path = require('path');
const dotenv = require('dotenv');
const https = require('https');
const fs = require('fs');
const { OpenAI } = require('openai');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;
const USE_HTTPS = process.env.USE_HTTPS !== 'false'; // Default to true

// Initialize OpenAI client
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Generate email endpoint
app.post('/api/generate-email', async (req, res) => {
  try {
    const { prompt, emailContext } = req.body;

    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    if (!process.env.OPENAI_API_KEY) {
      return res.status(500).json({ error: 'OpenAI API key not configured' });
    }

    // Build the system message with context
    let systemMessage = `You are a professional email assistant. Generate well-written, professional emails based on the user's prompt.`;
    
    let userMessage = prompt;

    // If email context is provided, include it in the prompt
    if (emailContext) {
      systemMessage += `\n\nYou are responding to or composing an email. Here's the context:\n`;
      if (emailContext.subject) {
        systemMessage += `Subject: ${emailContext.subject}\n`;
      }
      if (emailContext.from) {
        systemMessage += `From: ${emailContext.from}\n`;
      }
      if (emailContext.body) {
        systemMessage += `Original Email Body:\n${emailContext.body}\n`;
      }
      systemMessage += `\nGenerate an appropriate email response or draft based on the user's prompt.`;
    }

    // Call OpenAI API
    const completion = await openai.chat.completions.create({
      model: process.env.OPENAI_MODEL || 'gpt-4',
      messages: [
        { role: 'system', content: systemMessage },
        { role: 'user', content: userMessage }
      ],
      temperature: 0.7,
      max_tokens: 1000
    });

    const generatedEmail = completion.choices[0].message.content;

    res.json({ 
      email: generatedEmail,
      model: completion.model,
      usage: completion.usage
    });

  } catch (error) {
    console.error('Error generating email:', error);
    res.status(500).json({ 
      error: 'Failed to generate email',
      message: error.message 
    });
  }
});

// Serve static files
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', req.path === '/' ? 'taskpane.html' : req.path));
});

// Start server with HTTPS if certificates exist, otherwise HTTP
if (USE_HTTPS) {
  const certPath = path.join(__dirname, 'cert.pem');
  const keyPath = path.join(__dirname, 'key.pem');
  
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    const httpsOptions = {
      cert: fs.readFileSync(certPath),
      key: fs.readFileSync(keyPath)
    };
    
    https.createServer(httpsOptions, app).listen(PORT, () => {
      console.log(`✅ Server running on https://localhost:${PORT}`);
      console.log(`📝 Make sure to configure OPENAI_API_KEY in your .env file`);
      console.log(`🔒 HTTPS enabled with SSL certificate`);
    });
  } else {
    console.warn('⚠️  HTTPS certificates not found. Running on HTTP.');
    console.warn('⚠️  Outlook requires HTTPS. Generate certificates with:');
    console.warn('   openssl req -x509 -newkey rsa:4096 -nodes -keyout key.pem -out cert.pem -days 365 -subj "/CN=localhost"');
    console.warn('   Or set USE_HTTPS=false in .env for development\n');
    
    app.listen(PORT, () => {
      console.log(`✅ Server running on http://localhost:${PORT}`);
      console.log(`📝 Make sure to configure OPENAI_API_KEY in your .env file`);
      console.log(`⚠️  Note: Outlook add-ins require HTTPS. Use ngrok or generate SSL certificates.`);
    });
  }
} else {
  app.listen(PORT, () => {
    console.log(`✅ Server running on http://localhost:${PORT}`);
    console.log(`📝 Make sure to configure OPENAI_API_KEY in your .env file`);
    console.log(`⚠️  Note: Outlook add-ins require HTTPS. Use ngrok or generate SSL certificates.`);
  });
}

