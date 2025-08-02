# Use Python slim image for smaller size
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies for WeasyPrint and tools
RUN apt-get update && apt-get install -y \
    libpango-1.0-0 \
    libpangoft2-1.0-0 \
    libfontconfig1 \
    libcairo2 \
    libgdk-pixbuf2.0-0 \
    libglib2.0-0 \
    libgtk-3-0 \
    curl \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Set environment variables
# Set environment variables for Python optimization
ENV PYTHONUNBUFFERED=1 \
    GOTENBERG_URL=http://gotenberg:3000 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONOPTIMIZE=1

# Create cache directory
RUN mkdir -p /app/cache
ENV CACHE_DIR=/app/cache

# Copy requirements first for better Docker layer caching
COPY requirements.txt .

# Install Python dependencies with optimization flags
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY *.py ./

# Expose port
EXPOSE 5000

# Health check with faster interval
HEALTHCHECK --interval=20s --timeout=5s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:5000/health || exit 1

# Run the application with optimized settings
CMD ["python", "-O", "app.py"]
