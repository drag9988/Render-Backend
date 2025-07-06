# Use official Node.js LTS image
FROM node:18

# Install LibreOffice, Ghostscript, poppler-utils, qpdf, pdftk, and other tools for enhanced document conversion
RUN apt-get update && apt-get install -y \
    libreoffice \
    ghostscript \
    poppler-utils \
    qpdf \
    pdftk \
    curl \
    imagemagick \
    tesseract-ocr \
    tesseract-ocr-eng \
    python3 \
    python3-pip \
    pdfinfo \
    pdfimages \
    pdftotext \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install all dependencies (including devDependencies needed for build)
RUN npm install

# Copy rest of the app
COPY . .

# Copy and set permissions for start script
COPY start.sh /app/start.sh
RUN chmod +x /app/start.sh

# Make nest command executable
RUN chmod +x node_modules/.bin/nest

# Build the NestJS project using the script from package.json
RUN npm run build

# List the build output for debugging
RUN ls -la dist/

# Verify the build doesn't contain hardcoded port 3000
RUN grep -n "3000" dist/main.js || echo "✅ No hardcoded port 3000 found in build"

# Railway workaround: Expose port 3000 since Railway seems to expect it
EXPOSE 3000

# Set environment variables
ENV NODE_ENV=production
ENV TEMP_DIR=/tmp/pdf-converter

# Create temp directory
RUN mkdir -p /tmp/pdf-converter && chmod 777 /tmp/pdf-converter

# Add healthcheck
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
  CMD curl -f http://0.0.0.0:${PORT:-3000}/health || exit 1

# Set command to run the application directly
CMD ["npm", "start"]