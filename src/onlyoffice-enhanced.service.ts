import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import axios from 'axios';
import * as FormData from 'form-data';
import * as fs from 'fs/promises';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import { FileValidationService } from './file-validation.service';

const execAsync = promisify(exec);

/**
 * Enhanced ONLYOFFICE Service with Python integration for better performance
 * Supports both ONLYOFFICE Document Server and Python-based fallbacks
 */
@Injectable()
export class OnlyOfficeEnhancedService {
  private readonly logger = new Logger(OnlyOfficeEnhancedService.name);
  private readonly documentServerUrl: string;
  private readonly timeout: number;
  private readonly jwtSecret?: string;
  private readonly pythonPath: string;

  constructor(private readonly fileValidationService: FileValidationService) {
    this.documentServerUrl = process.env.ONLYOFFICE_DOCUMENT_SERVER_URL || '';
    this.timeout = parseInt(process.env.ONLYOFFICE_TIMEOUT || '120000', 10);
    this.jwtSecret = process.env.ONLYOFFICE_JWT_SECRET;
    this.pythonPath = process.env.PYTHON_PATH || 'python3';

    if (!this.documentServerUrl) {
      this.logger.warn('ONLYOFFICE_DOCUMENT_SERVER_URL not configured. Using enhanced LibreOffice + Python fallbacks.');
    } else {
      this.logger.log(`Enhanced ONLYOFFICE Document Server configured at: ${this.documentServerUrl}`);
    }
  }

  isAvailable(): boolean {
    // Always available - either via ONLYOFFICE server or enhanced LibreOffice
    return true;
  }

  /**
   * Convert PDF to DOCX with multiple conversion methods
   */
  async convertPdfToDocx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'docx');
  }

  /**
   * Convert PDF to XLSX with multiple conversion methods
   */
  async convertPdfToXlsx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'xlsx');
  }

  /**
   * Convert PDF to PPTX with multiple conversion methods
   */
  async convertPdfToPptx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'pptx');
  }

  /**
   * Enhanced PDF conversion with multiple fallback methods
   */
  private async convertPdf(pdfBuffer: Buffer, filename: string, targetFormat: string): Promise<Buffer> {
    if (!pdfBuffer || pdfBuffer.length === 0) {
      throw new BadRequestException('Invalid or empty PDF buffer provided');
    }

    // Validate file
    const file: Express.Multer.File = {
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
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp/pdf-converter';
    await fs.mkdir(tempDir, { recursive: true }); // Ensure temp directory exists
    const timestamp = Date.now();
    const tempInputPath = path.join(tempDir, `${timestamp}_input.pdf`);
    const tempOutputPath = path.join(tempDir, `${timestamp}_output.${targetFormat}`);

    try {
      this.logger.log(`Starting enhanced PDF to ${targetFormat.toUpperCase()} conversion for: ${filename}`);
      this.logger.log(`Temporary directory is: ${tempDir}`);

      // Write PDF to temp file
      await fs.writeFile(tempInputPath, pdfBuffer);
      this.logger.log(`Successfully wrote temporary PDF file to ${tempInputPath}`);

      // Method 1: ONLYOFFICE Document Server (if available)
      if (this.documentServerUrl) {
        try {
          const result = await this.convertViaOnlyOfficeServer(tempInputPath, targetFormat, validation.sanitizedFilename);
          if (result) {
            this.logger.log(`✅ ONLYOFFICE Document Server conversion successful: ${result.length} bytes`);
            return result;
          }
        } catch (onlyOfficeError) {
          this.logger.warn(`❌ ONLYOFFICE Document Server failed: ${onlyOfficeError.message}`);
        }
      }

      // Method 2: Python-based conversion (high quality)
      try {
        const result = await this.convertViaPython(tempInputPath, tempOutputPath, targetFormat);
        if (result) {
          this.logger.log(`✅ Python conversion successful: ${result.length} bytes`);
          return result;
        }
      } catch (pythonError) {
        this.logger.warn(`❌ Python conversion failed: ${pythonError.message}`);
      }

      // Method 3: Enhanced LibreOffice with specialized options (Note: Limited PDF to Office support)
      if (targetFormat === 'docx') {
        // LibreOffice can handle PDF to DOCX reasonably well
        try {
          const result = await this.convertViaEnhancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
          this.logger.log(`✅ Enhanced LibreOffice conversion successful: ${result.length} bytes`);
          return result;
        } catch (libreOfficeError) {
          this.logger.error(`❌ Enhanced LibreOffice failed: ${libreOfficeError.message}`);
          throw new Error(`All conversion methods failed. PDF to ${targetFormat.toUpperCase()} requires Python libraries (pdf2docx, PyMuPDF, pdfplumber) for best results. Last error: ${libreOfficeError.message}`);
        }
      } else {
        // For XLSX and PPTX, LibreOffice has very limited support
        this.logger.warn(`⚠️ LibreOffice has limited support for PDF to ${targetFormat.toUpperCase()} conversions. Python libraries are strongly recommended.`);
        throw new Error(`PDF to ${targetFormat.toUpperCase()} conversion failed. This format requires specialized Python libraries (PyMuPDF, pdfplumber, camelot-py) which are not available. LibreOffice cannot reliably convert PDF to ${targetFormat.toUpperCase()}.`);
      }

    } finally {
      // Cleanup temporary files
      try {
        await fs.unlink(tempInputPath).catch(() => {});
        await fs.unlink(tempOutputPath).catch(() => {});
      } catch (cleanupError) {
        this.logger.warn(`Cleanup warning: ${cleanupError.message}`);
      }
    }
  }

  /**
   * Convert via ONLYOFFICE Document Server API
   */
  private async convertViaOnlyOfficeServer(inputPath: string, targetFormat: string, filename: string): Promise<Buffer | null> {
    try {
      // Simple file serving approach for now
      const fileBuffer = await fs.readFile(inputPath);
      
      // Create a temporary file URL (in production, use proper file serving)
      const serverUrl = process.env.SERVER_URL || 'http://localhost:10000';
      const tempUrl = `${serverUrl}/temp/${Date.now()}_${filename}`;

      const conversionRequest = {
        async: false,
        filetype: 'pdf',
        key: this.generateConversionKey(filename),
        outputtype: targetFormat,
        title: filename,
        url: tempUrl // In production, implement proper file serving
      };

      if (this.jwtSecret) {
        conversionRequest['token'] = this.generateJWT(conversionRequest);
      }

      const response = await axios.post(`${this.documentServerUrl}/ConvertService.ashx`, conversionRequest, {
        timeout: this.timeout,
        headers: { 'Content-Type': 'application/json' }
      });

      if (response.data.error) {
        throw new Error(`ONLYOFFICE error: ${response.data.error}`);
      }

      if (!response.data.fileUrl) {
        throw new Error('No file URL returned');
      }

      const fileResponse = await axios.get(response.data.fileUrl, {
        responseType: 'arraybuffer',
        timeout: this.timeout
      });

      return Buffer.from(fileResponse.data);

    } catch (error) {
      this.logger.error(`ONLYOFFICE Document Server conversion failed: ${error.message}`);
      return null;
    }
  }

  /**
   * Convert using Python libraries (pdf2docx, pdfplumber, etc.)
   */
  private async convertViaPython(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer | null> {
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

      // Check if output file exists
      const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
      if (!outputExists) {
        throw new Error('Python conversion did not produce output file');
      }

      const result = await fs.readFile(outputPath);
      if (result.length < 100) {
        throw new Error(`Output file too small: ${result.length} bytes`);
      }

      return result;

    } catch (error) {
      throw new Error(`Python conversion failed: ${error.message}`);
    } finally {
      await fs.unlink(scriptPath).catch(() => {});
    }
  }

  /**
   * Generate Python script for different output formats
   */
  private generatePythonScript(): string {
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

  /**
   * Enhanced LibreOffice conversion with specialized options
   */
  private async convertViaEnhancedLibreOffice(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer> {
    const commands = [
      // Method 1: Import PDF as text and convert (for docx)
      targetFormat === 'docx' 
        ? `libreoffice --headless --writer --convert-to docx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
        : targetFormat === 'xlsx'
        ? `libreoffice --headless --calc --convert-to xlsx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
        : `libreoffice --headless --draw --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
      
      // Method 2: Standard conversion with PDF import filter
      `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
      
      // Method 3: Use Draw for all PDF conversions (most compatible)
      `libreoffice --headless --draw --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
      
      // Method 4: Force Writer import for text-based conversions
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
          maxBuffer: 1024 * 1024 * 10 // 10MB buffer
        });
        
        if (stdout) {
          this.logger.log(`LibreOffice output: ${stdout}`);
        }
        
        if (stderr && !stderr.includes('Warning')) {
          this.logger.warn(`LibreOffice stderr: ${stderr}`);
        }

        // Check for expected output file
        const expectedOutputPath = path.join(
          path.dirname(outputPath), 
          path.basename(inputPath, '.pdf') + '.' + targetFormat
        );

        try {
          await fs.access(expectedOutputPath);
          const result = await fs.readFile(expectedOutputPath);
          
          if (result.length > 100) {
            this.logger.log(`✅ LibreOffice conversion successful: ${result.length} bytes`);
            await fs.unlink(expectedOutputPath).catch(() => {});
            return result;
          } else {
            throw new Error(`Generated file too small: ${result.length} bytes`);
          }
        } catch (fileError) {
          this.logger.warn(`Output file check failed: ${fileError.message}`);
          lastError = fileError.message;
        }
        
      } catch (execError) {
        this.logger.warn(`LibreOffice execution failed: ${execError.message}`);
        lastError = execError.message;
      }
    }

    throw new Error(`Enhanced LibreOffice conversion failed. Last error: ${lastError}`);
  }

  /**
   * Generate a unique conversion key
   */
  private generateConversionKey(filename: string): string {
    const timestamp = Date.now();
    const random = Math.random().toString(36).substring(2);
    return `${timestamp}_${random}_${filename.replace(/[^a-zA-Z0-9]/g, '_')}`;
  }

  /**
   * Generate JWT token for ONLYOFFICE (if JWT is enabled)
   */
  private generateJWT(payload: any): string {
    if (!this.jwtSecret) {
      return '';
    }

    try {
      const jwt = require('jsonwebtoken');
      return jwt.sign(payload, this.jwtSecret, { expiresIn: '1h' });
    } catch (error) {
      this.logger.error(`JWT generation failed: ${error.message}`);
      return '';
    }
  }

  /**
   * Health check
   */
  async healthCheck(): Promise<boolean> {
    if (this.documentServerUrl) {
      try {
        const response = await axios.get(`${this.documentServerUrl}/healthcheck`, {
          timeout: 5000
        });
        return response.status === 200;
      } catch (error) {
        try {
          const response = await axios.get(`${this.documentServerUrl}/`, {
            timeout: 5000
          });
          return response.status === 200;
        } catch (altError) {
          this.logger.warn(`ONLYOFFICE server not available: ${error.message}`);
        }
      }
    }

    // Check if LibreOffice is available
    try {
      await execAsync('libreoffice --version', { timeout: 5000 });
      return true;
    } catch (error) {
      this.logger.error(`LibreOffice not available: ${error.message}`);
      return false;
    }
  }

  /**
   * Get server information
   */
  async getServerInfo(): Promise<any> {
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

    // Check Python availability
    try {
      const { stdout } = await execAsync(`${this.pythonPath} --version`, { timeout: 5000 });
      info.python.available = true;
      info.python['version'] = stdout.trim();
    } catch (error) {
      this.logger.warn(`Python not available: ${error.message}`);
    }

    // Check LibreOffice availability
    try {
      const { stdout } = await execAsync('libreoffice --version', { timeout: 5000 });
      info.libreoffice.available = true;
      info.libreoffice.version = stdout.trim();
    } catch (error) {
      this.logger.warn(`LibreOffice not available: ${error.message}`);
    }

    if (this.documentServerUrl) {
      info.onlyofficeServer['healthy'] = await this.healthCheck();
    }

    return info;
  }
}
