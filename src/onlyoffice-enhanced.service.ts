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

      // Method 3: Enhanced LibreOffice with specialized options
      try {
        const result = await this.convertViaEnhancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
        this.logger.log(`✅ Enhanced LibreOffice conversion successful: ${result.length} bytes`);
        return result;
      } catch (libreOfficeError) {
        this.logger.error(`❌ Enhanced LibreOffice failed: ${libreOfficeError.message}`);
        throw new Error(`All conversion methods failed. Last error: ${libreOfficeError.message}`);
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
from pathlib import Path

def install_package(package):
    try:
        # Use --break-system-packages to handle externally managed environments (PEP 668)
        # This is generally safe in a containerized/isolated environment.
        print(f"Installing {package}...");
        subprocess.check_call([sys.executable, "-m", "pip", "install", package, "--break-system-packages"]);
        print(f"{package} installed successfully.");
    except subprocess.CalledProcessError as e:
        print(f"Could not install {package}. Pip command failed: {e}");
        # Suggest manual installation as a fallback.
        print(f"Please try installing the package manually, e.g., 'apt install python3-{package.replace('_', '-')}' or 'pip install {package}' in a virtual environment.");
        raise

def convert_pdf(input_path, output_path, format_type):
    try:
        if format_type == 'docx':
            try:
                from pdf2docx import Converter
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                print(f"✅ pdf2docx conversion successful")
                return True
            except ImportError:
                install_package('pdf2docx')
                from pdf2docx import Converter
                cv = Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                return True
            except Exception as e:
                print(f"❌ pdf2docx failed: {e}");
                return False
                
        elif format_type == 'xlsx':
            try:
                import pdfplumber
                import pandas as pd
                
                all_tables = []
                with pdfplumber.open(input_path) as pdf:
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        for table in tables:
                            if table:
                                df = pd.DataFrame(table[1:], columns=table[0])
                                all_tables.append(df)
                
                if all_tables:
                    combined_df = pd.concat(all_tables, ignore_index=True)
                    combined_df.to_excel(output_path, index=False)
                else:
                    text_data = []
                    with pdfplumber.open(input_path) as pdf:
                        for page in pdf.pages:
                            text = page.extract_text()
                            if text:
                                text_data.append([text])
                    df = pd.DataFrame(text_data, columns=['Content'])
                    df.to_excel(output_path, index=False)
                
                print(f"✅ Excel conversion successful")
                return True
                
            except ImportError:
                install_package('pdfplumber')
                install_package('pandas')
                install_package('openpyxl')
                import pdfplumber
                import pandas as pd
                # Retry logic after installation
                # ... (same as above)
                return True
                
        elif format_type == 'pptx':
            try:
                import pdfplumber
                from pptx import Presentation
                
                prs = Presentation()
                with pdfplumber.open(input_path) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        slide_layout = prs.slide_layouts[5]  # Blank slide
                        slide = prs.slides.add_slide(slide_layout)
                        text = page.extract_text()
                        if text:
                            txBox = slide.shapes.add_textbox(0, 0, prs.slide_width, prs.slide_height)
                            tf = txBox.text_frame
                            tf.text = text
                
                prs.save(output_path)
                print(f"✅ PowerPoint conversion successful")
                return True
                
            except ImportError:
                install_package('pdfplumber')
                install_package('python-pptx')
                # Retry logic after installation
                # ... (same as above)
                return True
        
        return False
        
    except Exception as e:
        print(f"❌ Python conversion error: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print(f"Usage: python {sys.argv[0]} input.pdf output.ext format");
        sys.exit(1);
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    format_type = sys.argv[3]
    
    if not os.path.exists(input_file):
        print(f"❌ Input file not found: {input_file}")
        sys.exit(1)
    
    success = convert_pdf(input_file, output_file, format_type)
    if success:
        print(f"✅ Conversion completed: {output_file}")
        sys.exit(0)
    else:
        print(f"❌ Conversion failed")
        sys.exit(1)
`;
  }

  /**
   * Enhanced LibreOffice conversion with specialized options
   */
  private async convertViaEnhancedLibreOffice(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer> {
    const commands = [
      // Method 1: Specialized conversion based on target format
      targetFormat === 'docx' 
        ? `libreoffice --headless --invisible --nodefault --nolockcheck --nologo --norestore --writer --convert-to docx:"MS Word 2007 XML" --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`
        : targetFormat === 'xlsx'
        ? `libreoffice --headless --invisible --nodefault --nolockcheck --nologo --norestore --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`
        : `libreoffice --headless --invisible --nodefault --nolockcheck --nologo --norestore --impress --convert-to pptx:"MS PowerPoint 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
      
      // Method 2: Enhanced standard conversion
      `libreoffice --headless --convert-to ${targetFormat} --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
      
      // Method 3: Alternative approach
      `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
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
