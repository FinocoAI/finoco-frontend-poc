#!/bin/bash

# Excel AI Agent - Frontend Startup Script

echo "🚀 Starting Excel AI Agent Frontend..."
echo ""

cd frontend

echo "✅ Starting web server on port 3000..."
echo "Frontend will be available at http://localhost:3000"
echo ""
echo "Press Ctrl+C to stop the server"
echo ""

# Try different methods to start the server
if command -v npx &> /dev/null; then
    npx http-server -p 3000 --cors
elif command -v python3 &> /dev/null; then
    echo "Using Python's built-in server..."
    python3 -m http.server 3000
elif command -v python &> /dev/null; then
    echo "Using Python's built-in server..."
    python -m http.server 3000
else
    echo "❌ Error: No suitable web server found"
    echo "Please install Node.js or Python"
    exit 1
fi
