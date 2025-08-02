#!/bin/bash
# Quick startup script for the PDF server

echo "🚀 Starting Email2PDF Server..."
echo "📦 Building and starting Docker container..."

# Change to the pdf-server directory
cd "$(dirname "$0")"

# Start with Docker Compose
docker-compose up -d

echo "⏳ Waiting for server to start..."
sleep 10

# Test if server is running
if curl -f http://localhost:5000/health > /dev/null 2>&1; then
    echo "✅ Server is running at http://localhost:5000"
    echo "🔍 Health check: http://localhost:5000/health"
    echo "📝 API endpoint: http://localhost:5000/convert"
    echo ""
    echo "To view logs: docker-compose logs -f"
    echo "To stop server: docker-compose down"
else
    echo "❌ Server failed to start. Check logs with: docker-compose logs"
fi
