# Use official Node.js LTS image
FROM node:18

# Set working directory
WORKDIR /app

# Copy package files and install dependencies
COPY package*.json ./
RUN npm install

# ✅ Install Nest CLI globally
RUN npm install -g @nestjs/cli

# Copy rest of the app
COPY . .

# Build the NestJS project
RUN nest build

# Expose backend port
EXPOSE 3000

# Run app in production mode
CMD ["node", "dist/main"]