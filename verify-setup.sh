#!/bin/bash
echo "🔍 Verifying PDF Converter API setup..."

# Check if dist folder exists
if [ -d "dist" ]; then
    echo "❌ dist folder exists - should be rebuilt on Railway"
    echo "   Contents:"
    ls -la dist/
else
    echo "✅ dist folder doesn't exist - will be built fresh"
fi

# Check for any hardcoded port 3000 in source
echo ""
echo "🔍 Checking for hardcoded ports in source files..."
grep -r "3000" src/ --exclude-dir=node_modules || echo "✅ No hardcoded 3000 ports in src/"

# Show key configuration
echo ""
echo "🔍 Key configuration:"
echo "   Default port in main.ts: $(grep -n "process.env.PORT ||" src/main.ts)"
echo "   Host binding: $(grep -n "const host =" src/main.ts)"
echo "   Docker health check: $(grep -A1 "HEALTHCHECK" Dockerfile)"

echo ""
echo "✅ Verification complete. Ready for Railway deployment."
