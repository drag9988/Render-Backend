const fs = require('fs');
const FormData = require('form-data');
const fetch = require('node-fetch');

async function testCompression() {
  console.log('🧪 Testing PDF compression endpoint...');
  
  try {
    // First test if server is running
    console.log('1. Testing server health...');
    const healthResponse = await fetch('http://localhost:3000/', {
      method: 'GET',
      timeout: 5000
    });
    
    if (healthResponse.ok) {
      const healthText = await healthResponse.text();
      console.log('✅ Server is running:', healthText);
    } else {
      console.log('❌ Server health check failed:', healthResponse.status);
      return;
    }
    
    // Create a minimal test PDF (just a few bytes to test the endpoint)
    console.log('2. Testing compress-pdf endpoint with small test file...');
    
    const form = new FormData();
    
    // Create a minimal PDF buffer for testing (this is just for endpoint testing)
    const testBuffer = Buffer.from('%PDF-1.4\n1 0 obj\n<<\n/Type /Catalog\n/Pages 2 0 R\n>>\nendobj\n2 0 obj\n<<\n/Type /Pages\n/Kids [3 0 R]\n/Count 1\n>>\nendobj\n3 0 obj\n<<\n/Type /Page\n/Parent 2 0 R\n/MediaBox [0 0 612 792]\n>>\nendobj\nxref\n0 4\n0000000000 65535 f \n0000000010 00000 n \n0000000079 00000 n \n0000000173 00000 n \ntrailer\n<<\n/Size 4\n/Root 1 0 R\n>>\nstartxref\n253\n%%EOF');
    
    form.append('file', testBuffer, {
      filename: 'test.pdf',
      contentType: 'application/pdf'
    });
    form.append('quality', 'moderate');
    
    console.log('📤 Sending request to compress-pdf...');
    
    const response = await fetch('http://localhost:3000/compress-pdf', {
      method: 'POST',
      body: form,
      timeout: 30000, // 30 second timeout
      headers: form.getHeaders()
    });
    
    console.log('📥 Response status:', response.status);
    console.log('📥 Response headers:', Object.fromEntries(response.headers.entries()));
    
    if (response.ok) {
      const buffer = await response.buffer();
      console.log('✅ Success! Received compressed PDF:', buffer.length, 'bytes');
    } else {
      const errorText = await response.text();
      console.log('❌ Error response:', errorText);
    }
    
  } catch (error) {
    console.log('💥 Test failed with error:', error.message);
    if (error.code === 'ECONNREFUSED') {
      console.log('🚨 Server is not running! Please start the server first.');
    }
  }
}

// Run the test
testCompression().catch(console.error);
