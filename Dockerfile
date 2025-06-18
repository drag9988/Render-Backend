# Use official Node.js LTS image
FROM node:18

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install all dependencies (including devDependencies needed for build)
RUN npm install

# Copy rest of the app
COPY . .

# Build the NestJS project using the script from package.json
RUN npm run build

# Expose backend port
EXPOSE 3000

# Run app in production mode
CMD ["node", "dist/main"]