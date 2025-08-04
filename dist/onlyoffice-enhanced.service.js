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
            this.logger.log(`Starting enhanced PDF to ${targetFormat.toUpperCase()} conversion for: ${filename}`);
            this.logger.log(`Temporary directory is: ${tempDir}`);
            await fs.writeFile(tempInputPath, pdfBuffer);
            this.logger.log(`Successfully wrote temporary PDF file to ${tempInputPath}`);
            if (this.documentServerUrl) {
                try {
                    const result = await this.convertViaOnlyOfficeServer(tempInputPath, targetFormat, validation.sanitizedFilename);
                    if (result) {
                        this.logger.log(`✅ ONLYOFFICE Document Server conversion successful: ${result.length} bytes`);
                        return result;
                    }
                }
                catch (onlyOfficeError) {
                    this.logger.warn(`❌ ONLYOFFICE Document Server failed: ${onlyOfficeError.message}`);
                }
            }
            try {
                const result = await this.convertViaPython(tempInputPath, tempOutputPath, targetFormat);
                if (result) {
                    this.logger.log(`✅ Python conversion successful: ${result.length} bytes`);
                    return result;
                }
            }
            catch (pythonError) {
                this.logger.warn(`❌ Python conversion failed: ${pythonError.message}`);
            }
            if (targetFormat === 'docx') {
                try {
                    const result = await this.convertViaEnhancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
                    this.logger.log(`✅ Enhanced LibreOffice conversion successful: ${result.length} bytes`);
                    return result;
                }
                catch (libreOfficeError) {
                    this.logger.error(`❌ Enhanced LibreOffice failed: ${libreOfficeError.message}`);
                    throw new Error(`All conversion methods failed. PDF to ${targetFormat.toUpperCase()} requires Python libraries (pdf2docx, PyMuPDF, pdfplumber) for best results. Last error: ${libreOfficeError.message}`);
                }
            }
            else {
                this.logger.warn(`⚠️ LibreOffice has limited support for PDF to ${targetFormat.toUpperCase()} conversions. Python libraries are strongly recommended.`);
                throw new Error(`PDF to ${targetFormat.toUpperCase()} conversion failed. This format requires specialized Python libraries (PyMuPDF, pdfplumber, camelot-py) which are not available. LibreOffice cannot reliably convert PDF to ${targetFormat.toUpperCase()}.`);
            }
        }
        finally {
            try {
                await fs.unlink(tempInputPath).catch(() => { });
                await fs.unlink(tempOutputPath).catch(() => { });
            }
            catch (cleanupError) {
                this.logger.warn(`Cleanup warning: ${cleanupError.message}`);
            }
        }
    }
    async convertViaOnlyOfficeServer(inputPath, targetFormat, filename) {
        try {
            const fileBuffer = await fs.readFile(inputPath);
            const serverUrl = process.env.SERVER_URL || 'http://localhost:10000';
            const tempUrl = `${serverUrl}/temp/${Date.now()}_${filename}`;
            const conversionRequest = {
                async: false,
                filetype: 'pdf',
                key: this.generateConversionKey(filename),
                outputtype: targetFormat,
                title: filename,
                url: tempUrl
            };
            if (this.jwtSecret) {
                conversionRequest['token'] = this.generateJWT(conversionRequest);
            }
            const response = await axios_1.default.post(`${this.documentServerUrl}/ConvertService.ashx`, conversionRequest, {
                timeout: this.timeout,
                headers: { 'Content-Type': 'application/json' }
            });
            if (response.data.error) {
                throw new Error(`ONLYOFFICE error: ${response.data.error}`);
            }
            if (!response.data.fileUrl) {
                throw new Error('No file URL returned');
            }
            const fileResponse = await axios_1.default.get(response.data.fileUrl, {
                responseType: 'arraybuffer',
                timeout: this.timeout
            });
            return Buffer.from(fileResponse.data);
        }
        catch (error) {
            this.logger.error(`ONLYOFFICE Document Server conversion failed: ${error.message}`);
            return null;
        }
    }
    async convertViaPython(inputPath, outputPath, targetFormat) {
        const pythonScript = this.generatePythonScript();
        const scriptPath = inputPath.replace('.pdf', '_convert.py');
        try {
            await fs.writeFile(scriptPath, pythonScript);
            const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
            this.logger.log(`Executing Python conversion: ${command}`);
            const { stdout, stderr } = await execAsync(command, { timeout: this.timeout });
            if (stderr && !stderr.includes('Warning')) {
                this.logger.warn(`Python stderr: ${stderr}`);
            }
            if (stdout) {
                this.logger.log(`Python stdout: ${stdout}`);
            }
            const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
            if (!outputExists) {
                throw new Error('Python conversion did not produce output file');
            }
            const result = await fs.readFile(outputPath);
            if (result.length < 100) {
                throw new Error(`Output file too small: ${result.length} bytes`);
            }
            return result;
        }
        catch (error) {
            throw new Error(`Python conversion failed: ${error.message}`);
        }
        finally {
            await fs.unlink(scriptPath).catch(() => { });
        }
    }
    generatePythonScript() {
        return `#!/usr/bin/env python3
import sys
import os
import subprocess
import io
from pathlib import Path

def install_package(package, extra=""):
    try:
        package_spec = f"{package}{extra}"
        print(f"Installing {package_spec}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_spec, "--break-system-packages"])
        print(f"{package_spec} installed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Could not install {package_spec}. Pip command failed: {e}")
        print(f"Please try installing the package manually, e.g., 'pip install {package_spec}' in a virtual environment.")
        raise

def convert_to_docx(input_path, output_path):
    try:
        from pdf2docx import Converter
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        print("✅ pdf2docx conversion successful")
        return True
    except ImportError:
        install_package('pdf2docx')
        from pdf2docx import Converter
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        print("✅ pdf2docx conversion successful after install")
        return True
    except Exception as e:
        print(f"❌ pdf2docx failed: {e}")
        return False

def convert_to_xlsx(input_path, output_path):
    # ...existing code for convert_to_xlsx...
    try:
        import pandas as pd
        import camelot
        tables = camelot.read_pdf(input_path, pages='all', flavor='stream')
        if tables.n > 0:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    table.df.to_excel(writer, sheet_name=f'Page_{table.page}_Table_{i+1}', index=False, header=True)
            print(f"✅ Camelot conversion successful, found {tables.n} tables.")
            return True
        else:
            print("Camelot did not find any tables. Falling back to pdfplumber.")
    except ImportError:
        print("Camelot not found, installing...")
        install_package('camelot-py', extra="[cv]")
        install_package('pandas')
        install_package('openpyxl')
        return convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ Camelot failed: {e}. Falling back to pdfplumber.")
    try:
        import pdfplumber
        import pandas as pd
        all_tables = []
        with pdfplumber.open(input_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            all_tables.append(df)
        if all_tables:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i, df in enumerate(all_tables):
                    df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
            print(f"✅ pdfplumber table extraction successful, found {len(all_tables)} tables.")
            return True
        else:
            print("pdfplumber did not find any tables, extracting raw text as last resort.")
            with pdfplumber.open(input_path) as pdf:
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if text:
                            lines = text.split('\n')
                            df = pd.DataFrame(lines, columns=['Content'])
                            df.to_excel(writer, sheet_name=f'Page_{i+1}_Text', index=False)
            print("✅ pdfplumber raw text extraction successful.")
            return True
    except ImportError:
        install_package('pdfplumber')
        install_package('pandas')
        install_package('openpyxl')
        return convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ pdfplumber failed: {e}")
        return False

def convert_to_pptx(input_path, output_path, overlay_text=False):
    try:
        import fitz  # PyMuPDF
        from pptx import Presentation
        from pptx.util import Emu, Pt
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        first_page = pdf_doc.load_page(0)
        prs.slide_width = int(first_page.rect.width * 12700)
        prs.slide_height = int(first_page.rect.height * 12700)
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            slide_layout = prs.slide_layouts[6]  # Blank slide layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Only add PDF page as background image, no text overlay to avoid doubling
            pix = page.get_pixmap(dpi=150)
            image_stream = io.BytesIO(pix.tobytes("png"))
            slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)
            
            # Note: Text overlay disabled by default to prevent doubling
            # The PDF image already contains all the text content
            # Only enable overlay_text if you need editable text (will cause duplication)
            if overlay_text:
                print(f"⚠️  Warning: Adding text overlay to slide {page_num + 1} - this may cause text duplication")
                blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]
                for block in blocks:
                    if 'lines' in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                text = span['text']
                                if not text.strip():
                                    continue
                                rect = span['bbox']
                                font_size = span['size']
                                left = int(rect[0] * 12700)
                                top = int(rect[1] * 12700)
                                width = int((rect[2] - rect[0]) * 12700)
                                height = int((rect[3] - rect[1]) * 12700)
                                if width > 0 and height > 0:
                                    txBox = slide.shapes.add_textbox(left, top, width, height)
                                    tf = txBox.text_frame
                                    tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom = 0, 0, 0, 0
                                    tf.word_wrap = False
                                    p = tf.paragraphs[0]
                                    run = p.add_run()
                                    run.text = text
                                    font = run.font
                                    font.size = Pt(int(font_size))
        
        prs.save(output_path)
        print("✅ PyMuPDF to PPTX conversion successful (image-based, no text doubling)")
        return True
    except ImportError:
        install_package('PyMuPDF')
        install_package('python-pptx')
        return convert_to_pptx(input_path, output_path, overlay_text)
    except Exception as e:
        print(f"❌ PyMuPDF to PPTX conversion failed: {e}")
        return False

def convert_pdf(input_path, output_path, format_type):
    success = False
    if format_type == 'docx':
        success = convert_to_docx(input_path, output_path)
    elif format_type == 'xlsx':
        success = convert_to_xlsx(input_path, output_path)
    elif format_type == 'pptx':
        # By default, do NOT overlay text to avoid doubling
        success = convert_to_pptx(input_path, output_path, overlay_text=False)
    return success

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(f"Usage: python {sys.argv[0]} <input.pdf> <output.ext> <format>")
        sys.exit(1)
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    format_type = sys.argv[3].lower()
    if not os.path.exists(input_file):
        print(f"❌ Input file not found: {input_file}")
        sys.exit(1)
    success = convert_pdf(input_file, output_file, format_type)
    if success:
        print(f"✅ Conversion completed: {output_file}")
        sys.exit(0)
    else:
        print(f"❌ Conversion failed for {input_file} to {format_type}")
        sys.exit(1)
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