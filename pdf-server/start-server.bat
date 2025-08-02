@echo off
echo 🚀 Starting Email2PDF Server...
echo 📦 Building and starting Docker container...

REM Change to the pdf-server directory
cd /d "%~dp0"

REM Start with Docker Compose
docker-compose up -d

echo ⏳ Waiting for server to start...
timeout /t 10 /nobreak > nul

REM Test if server is running
curl -f http://localhost:5000/health >nul 2>&1
if %errorlevel% == 0 (
    echo ✅ Server is running at http://localhost:5000
    echo 🔍 Health check: http://localhost:5000/health
    echo 📝 API endpoint: http://localhost:5000/convert
    echo.
    echo To view logs: docker-compose logs -f
    echo To stop server: docker-compose down
) else (
    echo ❌ Server failed to start. Check logs with: docker-compose logs
)

pause
