# Use official Node.js LTS image
FROM node:18

# Install LibreOffice and Ghostscript for document conversion and PDF compression
RUN apt-get update && apt-get install -y \
    libreoffice \
    ghostscript \
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

# Make nest command executable
RUN chmod +x node_modules/.bin/nest

# Build the NestJS project using the script from package.json
RUN npm run build

# Expose backend port
EXPOSE 3000

# Run app in production mode
CMD ["node", "dist/main"]