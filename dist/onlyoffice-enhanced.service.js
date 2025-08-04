"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var OnlyOfficeEnhancedService_1;
Object.defineProperty(exports, "__esModule", { value: true });
exports.OnlyOfficeEnhancedService = void 0;
const common_1 = require("@nestjs/common");
const axios_1 = require("axios");
const FormData = require("form-data");
const fs = require("fs/promises");
const path = require("path");
const child_process_1 = require("child_process");
const util_1 = require("util");
const file_validation_service_1 = require("./file-validation.service");
const execAsync = (0, util_1.promisify)(child_process_1.exec);
let OnlyOfficeEnhancedService = OnlyOfficeEnhancedService_1 = class OnlyOfficeEnhancedService {
    constructor(fileValidationService) {
        this.fileValidationService = fileValidationService;
        this.logger = new common_1.Logger(OnlyOfficeEnhancedService_1.name);
        this.documentServerUrl = process.env.ONLYOFFICE_DOCUMENT_SERVER_URL || '';
        this.timeout = parseInt(process.env.ONLYOFFICE_TIMEOUT || '120000', 10);
        this.jwtSecret = process.env.ONLYOFFICE_JWT_SECRET;
        this.pythonPath = process.env.PYTHON_PATH || 'python3';
        if (!this.documentServerUrl) {
            this.logger.warn('ONLYOFFICE_DOCUMENT_SERVER_URL not configured. Using enhanced LibreOffice + Python fallbacks.');
        }
        else {
            this.logger.log(`Enhanced ONLYOFFICE Document Server configured at: ${this.documentServerUrl}`);
        }
    }
    isAvailable() {
        return true;
    }
    async convertPdfToDocx(pdfBuffer, filename) {
        return this.convertPdf(pdfBuffer, filename, 'docx');
    }
    async convertPdfToXlsx(pdfBuffer, filename) {
        return this.convertPdf(pdfBuffer, filename, 'xlsx');
    }
    async convertPdfToPptx(pdfBuffer, filename) {
        return this.convertPdf(pdfBuffer, filename, 'pptx');
    }
    async convertPdf(pdfBuffer, filename, targetFormat) {
        if (!pdfBuffer || pdfBuffer.length === 0) {
            throw new common_1.BadRequestException('Invalid or empty PDF buffer provided');
        }
        const file = {
            fieldname: 'file',
            originalname: filename,
            encoding: '7bit',
            mimetype: 'application/pdf',
            buffer: pdfBuffer,
            size: pdfBuffer.length,
            stream: null,
            destination: '',
            filename: '',
            path: ''
        };
        const validation = this.fileValidationService.validateFile(file, 'pdf');
        if (!validation.isValid) {
            throw new common_1.BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
        }
        const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp/pdf-converter';
        await fs.mkdir(tempDir, { recursive: true });
        const timestamp = Date.now();
        const tempInputPath = path.join(tempDir, `${timestamp}_input.pdf`);
        const tempOutputPath = path.join(tempDir, `${timestamp}_output.${targetFormat}`);
        try {
            this.logger.log(`🚀 Starting PREMIUM PDF to ${targetFormat.toUpperCase()} conversion for: ${filename}`);
            this.logger.log(`📁 Using temporary directory: ${tempDir}`);
            await fs.writeFile(tempInputPath, pdfBuffer);
            this.logger.log(`📄 PDF file written to: ${tempInputPath} (${pdfBuffer.length} bytes)`);
            if (this.documentServerUrl) {
                try {
                    this.logger.log(`🥇 Attempting ONLYOFFICE Document Server conversion...`);
                    const result = await this.convertViaOnlyOfficeServer(tempInputPath, targetFormat, validation.sanitizedFilename);
                    if (result && result.length > 1000) {
                        this.logger.log(`✅ ONLYOFFICE Document Server: SUCCESS! Output: ${result.length} bytes`);
                        return result;
                    }
                }
                catch (onlyOfficeError) {
                    this.logger.warn(`❌ ONLYOFFICE Document Server failed: ${onlyOfficeError.message}`);
                }
            }
            else {
                this.logger.log(`⚠️ ONLYOFFICE Document Server not configured - using enhanced fallbacks`);
            }
            try {
                this.logger.log(`🥈 Attempting Premium Python conversion (pdf2docx, PyMuPDF)...`);
                const result = await this.convertViaPremiumPython(tempInputPath, tempOutputPath, targetFormat);
                if (result && result.length > 1000) {
                    this.logger.log(`✅ Premium Python conversion: SUCCESS! Output: ${result.length} bytes`);
                    return result;
                }
            }
            catch (pythonError) {
                this.logger.warn(`❌ Premium Python conversion failed: ${pythonError.message}`);
            }
            if (targetFormat === 'docx') {
                try {
                    this.logger.log(`🥉 Attempting Advanced LibreOffice with PDF import optimizations...`);
                    const result = await this.convertViaAdvancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
                    if (result && result.length > 1000) {
                        this.logger.log(`✅ Advanced LibreOffice conversion: SUCCESS! Output: ${result.length} bytes`);
                        return result;
                    }
                }
                catch (libreOfficeError) {
                    this.logger.warn(`❌ Advanced LibreOffice failed: ${libreOfficeError.message}`);
                }
            }
            try {
                this.logger.log(`🛡️ Attempting Fallback Python conversion methods...`);
                const result = await this.convertViaFallbackPython(tempInputPath, tempOutputPath, targetFormat);
                if (result && result.length > 1000) {
                    this.logger.log(`✅ Fallback Python conversion: SUCCESS! Output: ${result.length} bytes`);
                    return result;
                }
            }
            catch (fallbackError) {
                this.logger.warn(`❌ Fallback Python conversion failed: ${fallbackError.message}`);
            }
            throw new Error(`All premium conversion methods failed for PDF to ${targetFormat.toUpperCase()}. This may be due to:
      1. Complex PDF structure (scanned images, unusual fonts, complex layouts)
      2. Missing Python libraries (pdf2docx, PyMuPDF, pdfplumber)
      3. Corrupted or password-protected PDF
      4. ONLYOFFICE Document Server not configured
      
      Recommendation: Deploy ONLYOFFICE Document Server for best results.`);
        }
        finally {
            try {
                await fs.unlink(tempInputPath).catch(() => { });
                await fs.unlink(tempOutputPath).catch(() => { });
                const tempPattern = path.join(tempDir, `${timestamp}_*`);
                try {
                    const { stdout } = await execAsync(`ls ${tempPattern}`, { timeout: 5000 });
                    if (stdout) {
                        const files = stdout.trim().split('\n').filter(f => f.trim());
                        for (const file of files) {
                            await fs.unlink(file).catch(() => { });
                        }
                    }
                }
                catch (cleanupListError) {
                }
            }
            catch (cleanupError) {
                this.logger.warn(`Cleanup warning: ${cleanupError.message}`);
            }
        }
    }
    async convertViaOnlyOfficeServer(inputPath, targetFormat, filename) {
        var _a;
        try {
            this.logger.log(`🏢 ONLYOFFICE Document Server conversion starting...`);
            const fileBuffer = await fs.readFile(inputPath);
            this.logger.log(`📄 PDF file loaded: ${fileBuffer.length} bytes`);
            const formData = new FormData();
            formData.append('file', fileBuffer, {
                filename: filename,
                contentType: 'application/pdf'
            });
            let fileUrl;
            try {
                const uploadResponse = await axios_1.default.post(`${this.documentServerUrl}/upload`, formData, {
                    headers: Object.assign({}, formData.getHeaders()),
                    timeout: this.timeout / 2,
                    maxContentLength: 100 * 1024 * 1024,
                });
                if (uploadResponse.data && uploadResponse.data.url) {
                    fileUrl = uploadResponse.data.url;
                    this.logger.log(`📤 File uploaded to ONLYOFFICE server: ${fileUrl}`);
                }
                else {
                    throw new Error('No upload URL returned');
                }
            }
            catch (uploadError) {
                this.logger.warn(`📤 Direct upload failed: ${uploadError.message}, using alternative method`);
                const serverUrl = process.env.SERVER_URL || `http://localhost:${process.env.PORT || 10000}`;
                const tempEndpoint = `temp_${Date.now()}_${filename}`;
                fileUrl = `${serverUrl}/temp/${tempEndpoint}`;
            }
            const conversionRequest = Object.assign({ async: false, filetype: 'pdf', key: this.generateConversionKey(filename), outputtype: targetFormat, title: filename, url: fileUrl, thumbnail: {
                    aspect: 2,
                    first: true,
                    height: 100,
                    width: 100
                } }, (targetFormat === 'docx' && {
                region: 'US',
                delimiter: {
                    paragraph: true,
                    column: false
                }
            }));
            if (this.jwtSecret) {
                try {
                    conversionRequest['token'] = this.generateJWT(conversionRequest);
                    this.logger.log(`🔐 JWT token added for secure conversion`);
                }
                catch (jwtError) {
                    this.logger.warn(`🔐 JWT generation failed: ${jwtError.message}`);
                }
            }
            const conversionUrl = `${this.documentServerUrl}/ConvertService.ashx`;
            this.logger.log(`🔄 Sending conversion request to: ${conversionUrl}`);
            this.logger.log(`📋 Conversion parameters: PDF → ${targetFormat.toUpperCase()}`);
            const response = await axios_1.default.post(conversionUrl, conversionRequest, {
                timeout: this.timeout,
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                validateStatus: (status) => status < 500
            });
            if (response.data.error) {
                const errorMsg = typeof response.data.error === 'object'
                    ? JSON.stringify(response.data.error)
                    : response.data.error;
                throw new Error(`ONLYOFFICE conversion error: ${errorMsg}`);
            }
            if (!response.data.fileUrl) {
                throw new Error('ONLYOFFICE returned no download URL');
            }
            this.logger.log(`📥 Conversion completed, downloading from: ${response.data.fileUrl}`);
            let fileResponse;
            let retries = 3;
            while (retries > 0) {
                try {
                    fileResponse = await axios_1.default.get(response.data.fileUrl, {
                        responseType: 'arraybuffer',
                        timeout: this.timeout,
                        maxContentLength: 100 * 1024 * 1024,
                        headers: {
                            'Accept': '*/*'
                        }
                    });
                    break;
                }
                catch (downloadError) {
                    retries--;
                    if (retries === 0)
                        throw downloadError;
                    this.logger.warn(`📥 Download attempt failed, retrying... (${3 - retries}/3)`);
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            }
            const convertedBuffer = Buffer.from(fileResponse.data);
            if (convertedBuffer.length === 0) {
                throw new Error('Downloaded file is empty');
            }
            if (convertedBuffer.length < 100) {
                throw new Error(`Downloaded file too small: ${convertedBuffer.length} bytes`);
            }
            const fileSignature = convertedBuffer.subarray(0, 8).toString('hex');
            const expectedSignatures = {
                'docx': ['504b0304', '504b0506', '504b0708'],
                'xlsx': ['504b0304', '504b0506', '504b0708'],
                'pptx': ['504b0304', '504b0506', '504b0708']
            };
            const isValidFormat = (_a = expectedSignatures[targetFormat]) === null || _a === void 0 ? void 0 : _a.some(sig => fileSignature.toLowerCase().startsWith(sig));
            if (!isValidFormat) {
                this.logger.warn(`⚠️ File signature verification failed. Expected ${targetFormat}, got: ${fileSignature}`);
            }
            this.logger.log(`✅ ONLYOFFICE Document Server conversion successful: ${convertedBuffer.length} bytes`);
            return convertedBuffer;
        }
        catch (error) {
            this.logger.error(`❌ ONLYOFFICE Document Server conversion failed: ${error.message}`);
            if (error.response) {
                this.logger.error(`📊 Response status: ${error.response.status}`);
                this.logger.error(`📊 Response data: ${JSON.stringify(error.response.data)}`);
            }
            return null;
        }
    }
    async convertViaPremiumPython(inputPath, outputPath, targetFormat) {
        const premiumPythonScript = this.generatePremiumPythonScript();
        const scriptPath = inputPath.replace('.pdf', '_premium_convert.py');
        try {
            await fs.writeFile(scriptPath, premiumPythonScript);
            const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
            this.logger.log(`🐍 Executing Premium Python conversion: ${targetFormat.toUpperCase()}`);
            const { stdout, stderr } = await execAsync(command, {
                timeout: this.timeout,
                maxBuffer: 1024 * 1024 * 20
            });
            if (stderr && !stderr.includes('Warning') && !stderr.includes('INFO')) {
                this.logger.warn(`Premium Python stderr: ${stderr}`);
            }
            if (stdout) {
                this.logger.log(`Premium Python stdout: ${stdout}`);
            }
            const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
            if (!outputExists) {
                throw new Error('Premium Python conversion did not produce output file');
            }
            const result = await fs.readFile(outputPath);
            if (result.length < 100) {
                throw new Error(`Premium Python output file too small: ${result.length} bytes`);
            }
            return result;
        }
        catch (error) {
            throw new Error(`Premium Python conversion failed: ${error.message}`);
        }
        finally {
            await fs.unlink(scriptPath).catch(() => { });
        }
    }
    async convertViaAdvancedLibreOffice(inputPath, outputPath, targetFormat) {
        const advancedCommands = [
            `libreoffice --headless --writer --convert-to ${targetFormat}:"MS Word 2007 XML" --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            `libreoffice --headless --convert-to ${targetFormat} --infilter="impress_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            `libreoffice --headless --draw --convert-to ${targetFormat} --infilter="draw_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            `libreoffice --headless --writer --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`
        ];
        let lastError = '';
        for (let i = 0; i < advancedCommands.length; i++) {
            const command = advancedCommands[i];
            this.logger.log(`📚 Advanced LibreOffice attempt ${i + 1}/${advancedCommands.length}: ${targetFormat.toUpperCase()}`);
            try {
                const { stdout, stderr } = await execAsync(command, {
                    timeout: this.timeout,
                    maxBuffer: 1024 * 1024 * 15
                });
                if (stdout) {
                    this.logger.log(`LibreOffice output (${i + 1}): ${stdout}`);
                }
                if (stderr && !stderr.includes('Warning')) {
                    this.logger.warn(`LibreOffice stderr (${i + 1}): ${stderr}`);
                }
                const expectedOutputPath = path.join(path.dirname(outputPath), path.basename(inputPath, '.pdf') + '.' + targetFormat);
                try {
                    await fs.access(expectedOutputPath);
                    const result = await fs.readFile(expectedOutputPath);
                    if (result.length > 1000) {
                        this.logger.log(`✅ Advanced LibreOffice success on attempt ${i + 1}: ${result.length} bytes`);
                        await fs.unlink(expectedOutputPath).catch(() => { });
                        return result;
                    }
                    else {
                        throw new Error(`Generated file too small: ${result.length} bytes`);
                    }
                }
                catch (fileError) {
                    this.logger.warn(`Advanced LibreOffice output check failed (${i + 1}): ${fileError.message}`);
                    lastError = fileError.message;
                }
            }
            catch (execError) {
                this.logger.warn(`Advanced LibreOffice execution failed (${i + 1}): ${execError.message}`);
                lastError = execError.message;
            }
        }
        throw new Error(`Advanced LibreOffice conversion failed after ${advancedCommands.length} attempts. Last error: ${lastError}`);
    }
    async convertViaFallbackPython(inputPath, outputPath, targetFormat) {
        const fallbackPythonScript = this.generateFallbackPythonScript();
        const scriptPath = inputPath.replace('.pdf', '_fallback_convert.py');
        try {
            await fs.writeFile(scriptPath, fallbackPythonScript);
            const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
            this.logger.log(`🛡️ Executing Fallback Python conversion: ${targetFormat.toUpperCase()}`);
            const { stdout, stderr } = await execAsync(command, {
                timeout: this.timeout,
                maxBuffer: 1024 * 1024 * 15
            });
            if (stderr && !stderr.includes('Warning')) {
                this.logger.warn(`Fallback Python stderr: ${stderr}`);
            }
            if (stdout) {
                this.logger.log(`Fallback Python stdout: ${stdout}`);
            }
            const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
            if (!outputExists) {
                throw new Error('Fallback Python conversion did not produce output file');
            }
            const result = await fs.readFile(outputPath);
            if (result.length < 50) {
                throw new Error(`Fallback Python output file too small: ${result.length} bytes`);
            }
            return result;
        }
        catch (error) {
            throw new Error(`Fallback Python conversion failed: ${error.message}`);
        }
        finally {
            await fs.unlink(scriptPath).catch(() => { });
        }
    }
    generatePremiumPythonScript() {
        return `#!/usr/bin/env python3
"""
Premium PDF Converter - Best Quality PDF to Office Conversion
Uses the most advanced Python libraries for superior results
"""
import sys
import os
import subprocess
import io
import tempfile
from pathlib import Path

def install_package(package, extra=""):
    """Install Python package with enhanced error handling"""
    try:
        package_spec = f"{package}{extra}"
        print(f"📦 Installing {package_spec}...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", package_spec, 
            "--break-system-packages", "--upgrade", "--no-warn-script-location"
        ])
        print(f"✅ {package_spec} installed successfully.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Could not install {package_spec}: {e}")
        return False

def premium_convert_to_docx(input_path, output_path):
    """Premium PDF to DOCX conversion using multiple high-quality methods"""
    print(f"🚀 Starting PREMIUM PDF to DOCX conversion...")
    
    # Method 1: pdf2docx (Best for most PDFs)
    try:
        from pdf2docx import Converter
        print("📄 Using pdf2docx (Premium Method 1)...")
        
        cv = Converter(input_path)
        # Use enhanced settings for better quality
        cv.convert(output_path, start=0, end=None, pages=None)
        cv.close()
        
        # Verify output quality
        if os.path.exists(output_path) and os.path.getsize(output_path) > 5000:
            print(f"✅ pdf2docx conversion successful: {os.path.getsize(output_path)} bytes")
            return True
        else:
            print("⚠️ pdf2docx output seems small, trying alternative method...")
            
    except ImportError:
        print("📦 Installing pdf2docx...")
        if install_package('pdf2docx'):
            return premium_convert_to_docx(input_path, output_path)
        else:
            print("❌ Failed to install pdf2docx")
    except Exception as e:
        print(f"❌ pdf2docx method failed: {e}")
    
    # Method 2: PyMuPDF with enhanced text extraction
    try:
        import fitz  # PyMuPDF
        from docx import Document
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        print("📝 Using PyMuPDF + python-docx (Premium Method 2)...")
        
        pdf_doc = fitz.open(input_path)
        doc = Document()
        
        # Set document margins for better layout
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Extract text blocks with formatting
            blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]
            
            if page_num > 0:
                doc.add_page_break()
            
            # Add page header
            if page_num == 0:
                header = doc.add_paragraph(f"Converted from PDF - Page {page_num + 1}")
                header.alignment = WD_ALIGN_PARAGRAPH.CENTER
                header.style = doc.styles['Heading 3']
            
            for block in blocks:
                if 'lines' in block:
                    block_text = ""
                    font_size = 12
                    
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span['text'].strip()
                            if text:
                                line_text += text + " "
                                font_size = max(font_size, span.get('size', 12))
                        
                        if line_text.strip():
                            block_text += line_text.strip() + "\\n"
                    
                    if block_text.strip():
                        p = doc.add_paragraph(block_text.strip())
                        # Set font size based on extracted text
                        for run in p.runs:
                            run.font.size = Pt(min(font_size, 14))
        
        doc.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 2000:
            print(f"✅ PyMuPDF conversion successful: {os.path.getsize(output_path)} bytes")
            return True
            
    except ImportError:
        print("📦 Installing PyMuPDF and python-docx...")
        if install_package('PyMuPDF') and install_package('python-docx'):
            return premium_convert_to_docx(input_path, output_path)
    except Exception as e:
        print(f"❌ PyMuPDF method failed: {e}")
    
    return False

def premium_convert_to_xlsx(input_path, output_path):
    """Premium PDF to XLSX conversion with advanced table detection"""
    print(f"🚀 Starting PREMIUM PDF to XLSX conversion...")
    
    # Method 1: Camelot (Best for table extraction)
    try:
        import camelot
        import pandas as pd
        
        print("📊 Using Camelot (Premium Table Extraction)...")
        
        # Try both flavors for maximum compatibility
        tables = []
        try:
            tables = camelot.read_pdf(input_path, pages='all', flavor='lattice')
            print(f"📋 Camelot lattice found {tables.n} tables")
        except:
            pass
            
        if tables.n == 0:
            try:
                tables = camelot.read_pdf(input_path, pages='all', flavor='stream')
                print(f"📋 Camelot stream found {tables.n} tables")
            except:
                pass
        
        if tables.n > 0:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    sheet_name = f'Page_{table.page}_Table_{i+1}'
                    # Clean the data
                    df = table.df
                    # Remove empty rows and columns
                    df = df.dropna(how='all').dropna(axis=1, how='all')
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"✅ Camelot conversion successful: {tables.n} tables extracted")
            return True
            
    except ImportError:
        print("📦 Installing Camelot...")
        if install_package('camelot-py', '[cv]') and install_package('pandas') and install_package('openpyxl'):
            return premium_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ Camelot method failed: {e}")
    
    # Method 2: pdfplumber (Excellent for text-based tables)
    try:
        import pdfplumber
        import pandas as pd
        
        print("📄 Using pdfplumber (Premium Text Extraction)...")
        
        all_tables = []
        all_text_data = []
        
        with pdfplumber.open(input_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Extract tables
                page_tables = page.extract_tables()
                if page_tables:
                    for table_num, table in enumerate(page_tables):
                        if table and len(table) > 1:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            # Clean the data
                            df = df.dropna(how='all').dropna(axis=1, how='all')
                            if not df.empty:
                                df.name = f'Page_{page_num+1}_Table_{table_num+1}'
                                all_tables.append(df)
                
                # Extract text as structured data
                text = page.extract_text()
                if text:
                    lines = [line.strip() for line in text.split('\\n') if line.strip()]
                    if lines:
                        text_df = pd.DataFrame(lines, columns=[f'Page_{page_num+1}_Content'])
                        all_text_data.append(text_df)
        
        if all_tables or all_text_data:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Write tables
                for i, df in enumerate(all_tables):
                    sheet_name = getattr(df, 'name', f'Table_{i+1}')
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Write text data if no tables found
                if not all_tables:
                    for i, df in enumerate(all_text_data):
                        df.to_excel(writer, sheet_name=f'Text_Page_{i+1}', index=False)
            
            print(f"✅ pdfplumber conversion successful: {len(all_tables)} tables + text data")
            return True
            
    except ImportError:
        if install_package('pdfplumber') and install_package('pandas') and install_package('openpyxl'):
            return premium_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ pdfplumber method failed: {e}")
    
    return False

def premium_convert_to_pptx(input_path, output_path):
    """Premium PDF to PPTX conversion with high-quality image rendering"""
    print(f"🚀 Starting PREMIUM PDF to PPTX conversion...")
    
    try:
        import fitz  # PyMuPDF
        from pptx import Presentation
        from pptx.util import Inches
        import io
        
        print("🎨 Using PyMuPDF + python-pptx (Premium Method)...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        # Set slide dimensions based on PDF
        if len(pdf_doc) > 0:
            first_page = pdf_doc.load_page(0)
            page_rect = first_page.rect
            # Convert points to inches (72 points = 1 inch)
            slide_width = Inches(page_rect.width / 72)
            slide_height = Inches(page_rect.height / 72)
            prs.slide_width = int(slide_width)
            prs.slide_height = int(slide_height)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Create slide
            slide_layout = prs.slide_layouts[6]  # Blank layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Render page as high-quality image
            mat = fitz.Matrix(2.0, 2.0)  # 2x scaling for better quality
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_data = pix.tobytes("png")
            
            # Add image to slide
            image_stream = io.BytesIO(img_data)
            slide.shapes.add_picture(
                image_stream, 
                0, 0, 
                width=prs.slide_width, 
                height=prs.slide_height
            )
            
            print(f"📄 Processed slide {page_num + 1}/{len(pdf_doc)}")
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 10000:
            print(f"✅ Premium PPTX conversion successful: {os.path.getsize(output_path)} bytes")
            return True
            
    except ImportError:
        print("📦 Installing PyMuPDF and python-pptx...")
        if install_package('PyMuPDF') and install_package('python-pptx'):
            return premium_convert_to_pptx(input_path, output_path)
    except Exception as e:
        print(f"❌ Premium PPTX conversion failed: {e}")
    
    return False

def convert_pdf(input_path, output_path, format_type):
    """Main conversion function with premium quality methods"""
    print(f"🎯 Premium PDF Converter - Converting to {format_type.upper()}")
    
    if not os.path.exists(input_path):
        print(f"❌ Input file not found: {input_path}")
        return False
    
    success = False
    if format_type == 'docx':
        success = premium_convert_to_docx(input_path, output_path)
    elif format_type == 'xlsx':
        success = premium_convert_to_xlsx(input_path, output_path)
    elif format_type == 'pptx':
        success = premium_convert_to_pptx(input_path, output_path)
    else:
        print(f"❌ Unsupported format: {format_type}")
        return False
    
    if success and os.path.exists(output_path):
        file_size = os.path.getsize(output_path)
        print(f"🎉 PREMIUM CONVERSION COMPLETED! Output: {file_size} bytes")
        return True
    else:
        print(f"❌ Premium conversion failed for {format_type}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(f"Usage: python {sys.argv[0]} <input.pdf> <output.ext> <format>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    format_type = sys.argv[3].lower()
    
    success = convert_pdf(input_file, output_file, format_type)
    sys.exit(0 if success else 1)
`;
    }
    generateFallbackPythonScript() {
        return `#!/usr/bin/env python3
"""
Fallback PDF Converter - Basic but reliable conversion methods
"""
import sys
import os
import subprocess

def install_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package, "--break-system-packages"])
        return True
    except:
        return False

def basic_convert_to_docx(input_path, output_path):
    """Basic PDF to DOCX using simple text extraction"""
    try:
        import fitz
        from docx import Document
        
        doc = Document()
        pdf_doc = fitz.open(input_path)
        
        for page in pdf_doc:
            text = page.get_text()
            if text.strip():
                doc.add_paragraph(text)
        
        doc.save(output_path)
        pdf_doc.close()
        return True
        
    except ImportError:
        if install_package('PyMuPDF') and install_package('python-docx'):
            return basic_convert_to_docx(input_path, output_path)
    except Exception as e:
        print(f"Basic DOCX conversion failed: {e}")
    return False

def basic_convert_to_xlsx(input_path, output_path):
    """Basic PDF to XLSX using simple text extraction"""
    try:
        import fitz
        import pandas as pd
        
        text_data = []
        pdf_doc = fitz.open(input_path)
        
        for page_num, page in enumerate(pdf_doc):
            text = page.get_text()
            if text.strip():
                lines = text.split('\\n')
                for line in lines:
                    if line.strip():
                        text_data.append({'Page': page_num + 1, 'Content': line.strip()})
        
        df = pd.DataFrame(text_data)
        df.to_excel(output_path, index=False)
        pdf_doc.close()
        return True
        
    except ImportError:
        if install_package('PyMuPDF') and install_package('pandas') and install_package('openpyxl'):
            return basic_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"Basic XLSX conversion failed: {e}")
    return False

def basic_convert_to_pptx(input_path, output_path):
    """Basic PDF to PPTX conversion"""
    try:
        import fitz
        from pptx import Presentation
        import io
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        for page in pdf_doc:
            slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(slide_layout)
            
            pix = page.get_pixmap()
            img_data = pix.tobytes("png")
            image_stream = io.BytesIO(img_data)
            
            slide.shapes.add_picture(image_stream, 0, 0)
        
        prs.save(output_path)
        pdf_doc.close()
        return True
        
    except ImportError:
        if install_package('PyMuPDF') and install_package('python-pptx'):
            return basic_convert_to_pptx(input_path, output_path)
    except Exception as e:
        print(f"Basic PPTX conversion failed: {e}")
    return False

if __name__ == "__main__":
    if len(sys.argv) != 4:
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    format_type = sys.argv[3].lower()
    
    success = False
    if format_type == 'docx':
        success = basic_convert_to_docx(input_file, output_file)
    elif format_type == 'xlsx':
        success = basic_convert_to_xlsx(input_file, output_file)
    elif format_type == 'pptx':
        success = basic_convert_to_pptx(input_file, output_file)
    
    sys.exit(0 if success else 1)
`;
    }
    async convertViaEnhancedLibreOffice(inputPath, outputPath, targetFormat) {
        const commands = [
            targetFormat === 'docx'
                ? `libreoffice --headless --writer --convert-to docx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
                : targetFormat === 'xlsx'
                    ? `libreoffice --headless --calc --convert-to xlsx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
                    : `libreoffice --headless --draw --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            `libreoffice --headless --draw --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
            targetFormat === 'docx'
                ? `libreoffice --headless --writer --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`
                : `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`
        ];
        let lastError = '';
        for (let i = 0; i < commands.length; i++) {
            const command = commands[i];
            this.logger.log(`Enhanced LibreOffice attempt ${i + 1}: Converting to ${targetFormat}`);
            try {
                const { stdout, stderr } = await execAsync(command, {
                    timeout: this.timeout,
                    maxBuffer: 1024 * 1024 * 10
                });
                if (stdout) {
                    this.logger.log(`LibreOffice output: ${stdout}`);
                }
                if (stderr && !stderr.includes('Warning')) {
                    this.logger.warn(`LibreOffice stderr: ${stderr}`);
                }
                const expectedOutputPath = path.join(path.dirname(outputPath), path.basename(inputPath, '.pdf') + '.' + targetFormat);
                try {
                    await fs.access(expectedOutputPath);
                    const result = await fs.readFile(expectedOutputPath);
                    if (result.length > 100) {
                        this.logger.log(`✅ LibreOffice conversion successful: ${result.length} bytes`);
                        await fs.unlink(expectedOutputPath).catch(() => { });
                        return result;
                    }
                    else {
                        throw new Error(`Generated file too small: ${result.length} bytes`);
                    }
                }
                catch (fileError) {
                    this.logger.warn(`Output file check failed: ${fileError.message}`);
                    lastError = fileError.message;
                }
            }
            catch (execError) {
                this.logger.warn(`LibreOffice execution failed: ${execError.message}`);
                lastError = execError.message;
            }
        }
        throw new Error(`Enhanced LibreOffice conversion failed. Last error: ${lastError}`);
    }
    generateConversionKey(filename) {
        const timestamp = Date.now();
        const random = Math.random().toString(36).substring(2);
        return `${timestamp}_${random}_${filename.replace(/[^a-zA-Z0-9]/g, '_')}`;
    }
    generateJWT(payload) {
        if (!this.jwtSecret) {
            return '';
        }
        try {
            const jwt = require('jsonwebtoken');
            return jwt.sign(payload, this.jwtSecret, { expiresIn: '1h' });
        }
        catch (error) {
            this.logger.error(`JWT generation failed: ${error.message}`);
            return '';
        }
    }
    async healthCheck() {
        if (this.documentServerUrl) {
            try {
                const response = await axios_1.default.get(`${this.documentServerUrl}/healthcheck`, {
                    timeout: 5000
                });
                return response.status === 200;
            }
            catch (error) {
                try {
                    const response = await axios_1.default.get(`${this.documentServerUrl}/`, {
                        timeout: 5000
                    });
                    return response.status === 200;
                }
                catch (altError) {
                    this.logger.warn(`ONLYOFFICE server not available: ${error.message}`);
                }
            }
        }
        try {
            await execAsync('libreoffice --version', { timeout: 5000 });
            return true;
        }
        catch (error) {
            this.logger.error(`LibreOffice not available: ${error.message}`);
            return false;
        }
    }
    async getServerInfo() {
        const info = {
            onlyofficeServer: {
                available: !!this.documentServerUrl,
                url: this.documentServerUrl,
                jwtEnabled: !!this.jwtSecret
            },
            python: {
                available: false,
                path: this.pythonPath
            },
            libreoffice: {
                available: false,
                version: 'Unknown'
            }
        };
        try {
            const { stdout } = await execAsync(`${this.pythonPath} --version`, { timeout: 5000 });
            info.python.available = true;
            info.python['version'] = stdout.trim();
        }
        catch (error) {
            this.logger.warn(`Python not available: ${error.message}`);
        }
        try {
            const { stdout } = await execAsync('libreoffice --version', { timeout: 5000 });
            info.libreoffice.available = true;
            info.libreoffice.version = stdout.trim();
        }
        catch (error) {
            this.logger.warn(`LibreOffice not available: ${error.message}`);
        }
        if (this.documentServerUrl) {
            info.onlyofficeServer['healthy'] = await this.healthCheck();
        }
        return info;
    }
};
exports.OnlyOfficeEnhancedService = OnlyOfficeEnhancedService;
exports.OnlyOfficeEnhancedService = OnlyOfficeEnhancedService = OnlyOfficeEnhancedService_1 = __decorate([
    (0, common_1.Injectable)(),
    __metadata("design:paramtypes", [file_validation_service_1.FileValidationService])
], OnlyOfficeEnhancedService);
//# sourceMappingURL=onlyoffice-enhanced.service.js.map