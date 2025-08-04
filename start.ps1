# Set environment variables for development
$env:NODE_ENV="development"
$env:LOG_LEVEL="debug"
$env:ONLYOFFICE_URL="http://localhost:8081/"
$env:ONLYOFFICE_JWT_SECRET="your_jwt_secret"
$env:CONVERT_API_SECRET="your_convert_api_secret"
$env:PORT="3000"

# Start the application using ts-node for development
npm run start:dev
