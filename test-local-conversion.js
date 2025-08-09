#!/usr/bin/env node

/**
 * Local PDF to PowerPoint conversion test script
 * This script helps debug the Python conversion issues locally
 */

const fs = require('fs').promises;
const path = require('path');
const { exec } = require('child_process');
const { promisify } = require('util');

const execAsync = promisify(exec);

async function testPdfToPptConversion() {
    console.log('🚀 Starting local PDF to PowerPoint conversion test...\n');

    // Configuration
    const testPdfPath = './test.pdf'; // Put your test PDF here
    const outputDir = './test-output';
    const pythonPath = process.env.PYTHON_PATH || 'python3';
    
    try {
        // Create output directory
        await fs.mkdir(outputDir, { recursive: true });
        console.log('✅ Created output directory:', outputDir);

        // Check if test PDF exists
        try {
            await fs.access(testPdfPath);
            const stats = await fs.stat(testPdfPath);
            console.log(`📄 Found test PDF: ${testPdfPath} (${stats.size} bytes)`);
        } catch (error) {
            console.error('❌ Test PDF not found. Please place a PDF file named "test.pdf" in the current directory.');
            console.log('💡 You can use any PDF file you want to test with.');
            return;
        }

        // Test Python availability
        console.log('\n🐍 Testing Python environment...');
        try {
            const { stdout: pythonVersion } = await execAsync(`${pythonPath} --version`);
            console.log(`✅ Python version: ${pythonVersion.trim()}`);
        } catch (error) {
            console.error(`❌ Python not found at: ${pythonPath}`);
            console.log('💡 Try setting PYTHON_PATH environment variable or install Python3');
            return;
        }

        // Test if required Python packages are available
        console.log('\n📦 Testing Python packages...');
        const requiredPackages = ['pdf2docx', 'PyMuPDF', 'pdfplumber', 'python-pptx', 'openpyxl'];
        
        for (const pkg of requiredPackages) {
            try {
                const { stdout, stderr } = await execAsync(`${pythonPath} -c "import ${pkg.replace('-', '_').replace('PyMuPDF', 'fitz')}; print('${pkg} available')"`);
                console.log(`✅ ${pkg}: Available`);
            } catch (error) {
                console.log(`❌ ${pkg}: Not available - ${error.message.includes('ModuleNotFoundError') ? 'Not installed' : 'Error'}`);
            }
        }

        // Create a simple test Python script
        const testScript = `#!/usr/bin/env python3
import sys
import os

def main():
    print("🐍 Starting Python conversion test...")
    print(f"Python version: {sys.version}")
    print(f"Arguments: {sys.argv}")
    
    if len(sys.argv) != 4:
        print("❌ Usage: python script.py <input_pdf> <output_pptx> <format>")
        sys.exit(1)
    
    input_pdf = sys.argv[1]
    output_pptx = sys.argv[2]
    format_type = sys.argv[3]
    
    print(f"📍 Input: {input_pdf}")
    print(f"📍 Output: {output_pptx}")
    print(f"📍 Format: {format_type}")
    
    # Check if input exists
    if not os.path.exists(input_pdf):
        print(f"❌ Input file does not exist: {input_pdf}")
        sys.exit(1)
    
    file_size = os.path.getsize(input_pdf)
    print(f"📄 Input file size: {file_size} bytes")
    
    # Test imports
    print("\\n📦 Testing package imports...")
    
    try:
        import fitz  # PyMuPDF
        print("✅ PyMuPDF (fitz) imported successfully")
    except ImportError as e:
        print(f"❌ PyMuPDF import failed: {e}")
    
    try:
        from pptx import Presentation
        print("✅ python-pptx imported successfully")
    except ImportError as e:
        print(f"❌ python-pptx import failed: {e}")
    
    try:
        import pdfplumber
        print("✅ pdfplumber imported successfully")
    except ImportError as e:
        print(f"❌ pdfplumber import failed: {e}")
    
    try:
        from pdf2docx import Converter
        print("✅ pdf2docx imported successfully")
    except ImportError as e:
        print(f"❌ pdf2docx import failed: {e}")
    
    # Simple conversion attempt
    print("\\n🔄 Attempting simple conversion...")
    
    try:
        # Try using python-pptx to create a basic presentation
        from pptx import Presentation
        from pptx.util import Inches
        
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "Test Conversion"
        slide.shapes.placeholders[1].text = f"Converted from: {os.path.basename(input_pdf)}"
        
        prs.save(output_pptx)
        
        if os.path.exists(output_pptx):
            output_size = os.path.getsize(output_pptx)
            print(f"✅ Basic PowerPoint created successfully: {output_size} bytes")
        else:
            print("❌ Output file was not created")
            
    except Exception as e:
        print(f"❌ Basic conversion failed: {e}")
        import traceback
        traceback.print_exc()
    
    print("\\n✅ Python test script completed!")

if __name__ == "__main__":
    main()
`;

        const testScriptPath = path.join(outputDir, 'test_conversion.py');
        await fs.writeFile(testScriptPath, testScript);
        console.log('\n📜 Created test Python script:', testScriptPath);

        // Run the test script
        console.log('\n🚀 Running Python test script...');
        const testOutputPath = path.join(outputDir, 'test_output.pptx');
        const command = `${pythonPath} "${testScriptPath}" "${testPdfPath}" "${testOutputPath}" "pptx"`;
        
        console.log(`📝 Command: ${command}\n`);
        
        try {
            const { stdout, stderr } = await execAsync(command, { 
                timeout: 30000,
                maxBuffer: 1024 * 1024 * 5 
            });
            
            if (stdout) {
                console.log('📋 Python Output:');
                console.log(stdout);
            }
            
            if (stderr) {
                console.log('\n⚠️ Python Errors/Warnings:');
                console.log(stderr);
            }
            
        } catch (error) {
            console.error('\n❌ Python script execution failed:');
            console.error(error.message);
            
            if (error.stdout) {
                console.log('\n📋 Partial Output:');
                console.log(error.stdout);
            }
            
            if (error.stderr) {
                console.log('\n⚠️ Error Details:');
                console.log(error.stderr);
            }
        }

        console.log('\n🎯 Test Summary:');
        console.log('1. Check the output above for any missing Python packages');
        console.log('2. Install missing packages using: pip install <package-name>');
        console.log('3. If successful, the issue might be in the more complex Python script');
        console.log('4. Check the test-output directory for any generated files');
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
    }
}

// Run the test
testPdfToPptConversion();
