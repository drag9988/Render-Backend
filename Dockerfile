# Use official Node.js LTS image
FROM node:18

# Set working directory
WORKDIR /app

# Copy package files and install dependencies
COPY package*.json ./
RUN npm install

# Copy rest of the app
COPY . .

# Give permission to nest binary
RUN chmod +x ./node_modules/.bin/nest

# Build the NestJS project using npx
RUN npx nest build

# Expose backend port
EXPOSE 3000

# Run app in production mode
CMD ["node", "dist/main"]