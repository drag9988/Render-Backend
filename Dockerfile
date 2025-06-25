# Use official Node.js LTS image
FROM node:18

# Install LibreOffice, Ghostscript, and curl for document conversion, PDF compression, and healthcheck
RUN apt-get update && apt-get install -y \
    libreoffice \
    ghostscript \
    curl \
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

# Expose backend port (Railway will set the PORT env variable)
EXPOSE $PORT
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