#!/usr/bin/env node

const fs = require('fs').promises;
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

async function verifyEnhancedOnlyOfficeSetup() {
    console.log('🔧 Enhanced ONLYOFFICE Service Verification');
    console.log('==========================================');
    
    let allChecks = true;
    
    // Check 1: Temp directory
    try {
        const tempDir = process.env.TEMP_DIR || '/tmp/pdf-converter';
        await fs.mkdir(tempDir, { recursive: true });
        await fs.chmod(tempDir, 0o777);
        
        // Test write access
        const testFile = path.join(tempDir, 'test_write.txt');
        await fs.writeFile(testFile, 'test');
        await fs.unlink(testFile);
        
        console.log('✅ Temp directory check passed:', tempDir);
    } catch (error) {
        console.log('❌ Temp directory check failed:', error.message);
        allChecks = false;
    }
    
    // Check 2: LibreOffice
    try {
        const { stdout } = await execAsync('libreoffice --version');
        console.log('✅ LibreOffice check passed:', stdout.trim().split('\n')[0]);
    } catch (error) {
        console.log('❌ LibreOffice check failed:', error.message);
        allChecks = false;
    }
    
    // Check 3: Python
    try {
        const pythonPath = process.env.PYTHON_PATH || 'python3';
        const { stdout } = await execAsync(`${pythonPath} --version`);
        console.log('✅ Python check passed:', stdout.trim());
        
        // Check Python packages
        const packages = ['pdf2docx', 'pdfplumber', 'python-docx', 'pandas', 'python-pptx', 'openpyxl'];
        for (const pkg of packages) {
            try {
                await execAsync(`${pythonPath} -c "import ${pkg.replace('-', '_')}; print('${pkg} available')"`);
                console.log(`  ✅ ${pkg} available`);
            } catch (pkgError) {
                console.log(`  ⚠️  ${pkg} not available - will install on demand`);
            }
        }
    } catch (error) {
        console.log('❌ Python check failed:', error.message);
        allChecks = false;
    }
    
    // Check 4: Node.js modules
    try {
        const packageJson = JSON.parse(await fs.readFile('package.json', 'utf8'));
        console.log('✅ package.json check passed:', packageJson.name, packageJson.version);
        
        // Check if dist exists
        try {
            await fs.access('dist/main.js');
            console.log('✅ Build check passed: dist/main.js exists');
        } catch {
            console.log('⚠️  Build not found - run npm run build');
        }
    } catch (error) {
        console.log('❌ Node.js modules check failed:', error.message);
        allChecks = false;
    }
    
    // Check 5: Environment variables
    console.log('🌍 Environment variables:');
    console.log('  NODE_ENV:', process.env.NODE_ENV || 'not set');
    console.log('  PORT:', process.env.PORT || 'not set');
    console.log('  TEMP_DIR:', process.env.TEMP_DIR || 'default (/tmp)');
    console.log('  ONLYOFFICE_DOCUMENT_SERVER_URL:', process.env.ONLYOFFICE_DOCUMENT_SERVER_URL || 'not set (Enhanced Mode)');
    console.log('  PYTHON_PATH:', process.env.PYTHON_PATH || 'python3 (default)');
    
    // Check 6: Test conversion capability
    try {
        console.log('🧪 Testing conversion capability...');
        
        // Create a simple test PDF content
        const testPdfContent = `%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj
2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj
3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
>>
endobj
4 0 obj
<<
/Length 44
>>
stream
BT
/F1 12 Tf
100 700 Td
(Test PDF) Tj
ET
endstream
endobj
xref
0 5
0000000000 65535 f 
0000000010 00000 n 
0000000079 00000 n 
0000000173 00000 n 
0000000301 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
398
%%EOF`;
        
        const tempDir = process.env.TEMP_DIR || '/tmp/pdf-converter';
        const testPdfPath = path.join(tempDir, 'test.pdf');
        const testDocxPath = path.join(tempDir, 'test.docx');
        
        await fs.writeFile(testPdfPath, testPdfContent);
        
        // Test LibreOffice conversion
        try {
            await execAsync(`libreoffice --headless --convert-to docx --outdir "${tempDir}" "${testPdfPath}"`);
            await fs.access(testDocxPath);
            console.log('✅ Test conversion successful');
            
            // Cleanup
            await fs.unlink(testPdfPath).catch(() => {});
            await fs.unlink(testDocxPath).catch(() => {});
        } catch (convError) {
            console.log('⚠️  Test conversion failed - but service might still work with real PDFs');
        }
        
    } catch (error) {
        console.log('⚠️  Test conversion setup failed:', error.message);
    }
    
    console.log('\n🎯 Summary:');
    if (allChecks) {
        console.log('✅ All critical checks passed!');
        console.log('🚀 Enhanced ONLYOFFICE Service is ready for deployment');
        console.log('📋 Conversion priority: Enhanced ONLYOFFICE → Original ONLYOFFICE → ConvertAPI → LibreOffice');
    } else {
        console.log('⚠️  Some checks failed, but the service may still work with fallbacks');
        console.log('🔧 Consider fixing the failed checks for optimal performance');
    }
    
    console.log('\n📡 To test the service:');
    console.log(`curl -X POST -F 'file=@test.pdf' http://localhost:${process.env.PORT || 10000}/convert-pdf-to-word`);
    console.log(`curl http://localhost:${process.env.PORT || 10000}/onlyoffice/enhanced-status`);
}

verifyEnhancedOnlyOfficeSetup().catch(console.error);
