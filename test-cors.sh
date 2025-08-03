#!/bin/bash
echo "🚀 Building and testing CORS fixes..."

# Build the project
echo "📦 Building TypeScript..."
npm run build

if [ $? -eq 0 ]; then
    echo "✅ Build successful!"
    
    # Start the server in background
    echo "🌐 Starting server..."
    npm start &
    SERVER_PID=$!
    
    # Wait for server to start
    sleep 5
    
    # Test CORS with curl
    echo "🧪 Testing CORS..."
    
    echo "Testing GET /cors-test..."
    curl -v -H "Origin: http://localhost:3000" http://localhost:3000/cors-test
    
    echo -e "\n\nTesting OPTIONS preflight..."
    curl -v -X OPTIONS \
         -H "Origin: http://localhost:3000" \
         -H "Access-Control-Request-Method: POST" \
         -H "Access-Control-Request-Headers: Content-Type" \
         http://localhost:3000/cors-test
    
    echo -e "\n\nTesting POST /cors-test..."
    curl -v -X POST \
         -H "Origin: http://localhost:3000" \
         -H "Content-Type: application/json" \
         -d '{"test": "data"}' \
         http://localhost:3000/cors-test
    
    # Stop the server
    kill $SERVER_PID
    echo -e "\n\n✅ CORS testing complete!"
    
else
    echo "❌ Build failed!"
fi
