#!/bin/bash

# Setup script for Inscribea Outlook Add-in

echo "🚀 Setting up Inscribea Outlook Add-in..."
echo ""

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js first."
    exit 1
fi

echo "✅ Node.js version: $(node --version)"

# Install dependencies
echo ""
echo "📦 Installing dependencies..."
npm install

# Create .env file if it doesn't exist
if [ ! -f .env ]; then
    echo ""
    echo "📝 Creating .env file from template..."
    cp env.example .env
    echo "⚠️  Please edit .env and add your OPENAI_API_KEY"
else
    echo "✅ .env file already exists"
fi

# Check if SSL certificates exist
if [ ! -f cert.pem ] || [ ! -f key.pem ]; then
    echo ""
    echo "🔒 SSL certificates not found."
    echo "   Outlook add-ins require HTTPS. You have two options:"
    echo ""
    echo "   Option 1: Generate self-signed certificates (for development):"
    echo "   openssl req -x509 -newkey rsa:4096 -nodes -keyout key.pem -out cert.pem -days 365 -subj \"/CN=localhost\""
    echo ""
    echo "   Option 2: Use ngrok for tunneling:"
    echo "   npm install -g ngrok"
    echo "   ngrok http 3000"
    echo "   (Then update manifest.xml with the ngrok URL)"
    echo ""
else
    echo "✅ SSL certificates found"
fi

echo ""
echo "✨ Setup complete!"
echo ""
echo "Next steps:"
echo "1. Edit .env and add your OPENAI_API_KEY"
echo "2. Generate SSL certificates (see above) or use ngrok"
echo "3. Run 'npm start' to start the server"
echo "4. Install the add-in in Outlook using manifest.xml"

