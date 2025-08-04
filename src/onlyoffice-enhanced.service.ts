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
   * Enhanced PDF conversion with premium quality methods prioritized
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
      this.logger.log(`🚀 Starting PREMIUM PDF to ${targetFormat.toUpperCase()} conversion for: ${filename}`);
      this.logger.log(`📁 Using temporary directory: ${tempDir}`);

      // Write PDF to temp file
      await fs.writeFile(tempInputPath, pdfBuffer);
      this.logger.log(`📄 PDF file written to: ${tempInputPath} (${pdfBuffer.length} bytes)`);

      // PRIORITY 1: ONLYOFFICE Document Server (Best quality for office formats)
      if (this.documentServerUrl) {
        try {
          this.logger.log(`🥇 Attempting ONLYOFFICE Document Server conversion...`);
          const result = await this.convertViaOnlyOfficeServer(tempInputPath, targetFormat, validation.sanitizedFilename);
          if (result && result.length > 1000) { // Ensure meaningful output
            this.logger.log(`✅ ONLYOFFICE Document Server: SUCCESS! Output: ${result.length} bytes`);
            return result;
          }
        } catch (onlyOfficeError) {
          this.logger.warn(`❌ ONLYOFFICE Document Server failed: ${onlyOfficeError.message}`);
        }
      } else {
        this.logger.log(`⚠️ ONLYOFFICE Document Server not configured - using enhanced fallbacks`);
      }

      // PRIORITY 2: Premium Python Libraries (Excellent quality, especially for DOCX)
      try {
        this.logger.log(`🥈 Attempting Premium Python conversion (pdf2docx, PyMuPDF)...`);
        const result = await this.convertViaPremiumPython(tempInputPath, tempOutputPath, targetFormat);
        if (result && result.length > 1000) {
          this.logger.log(`✅ Premium Python conversion: SUCCESS! Output: ${result.length} bytes`);
          return result;
        }
      } catch (pythonError) {
        this.logger.warn(`❌ Premium Python conversion failed: ${pythonError.message}`);
      }

      // PRIORITY 3: Advanced LibreOffice with PDF import optimizations
      if (targetFormat === 'docx') {
        try {
          this.logger.log(`🥉 Attempting Advanced LibreOffice with PDF import optimizations...`);
          const result = await this.convertViaAdvancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
          if (result && result.length > 1000) {
            this.logger.log(`✅ Advanced LibreOffice conversion: SUCCESS! Output: ${result.length} bytes`);
            return result;
          }
        } catch (libreOfficeError) {
          this.logger.warn(`❌ Advanced LibreOffice failed: ${libreOfficeError.message}`);
        }
      }

      // PRIORITY 4: Fallback Python methods
      try {
        this.logger.log(`🛡️ Attempting Fallback Python conversion methods...`);
        const result = await this.convertViaFallbackPython(tempInputPath, tempOutputPath, targetFormat);
        if (result && result.length > 1000) {
          this.logger.log(`✅ Fallback Python conversion: SUCCESS! Output: ${result.length} bytes`);
          return result;
        }
      } catch (fallbackError) {
        this.logger.warn(`❌ Fallback Python conversion failed: ${fallbackError.message}`);
      }

      // If all methods failed
      throw new Error(`All premium conversion methods failed for PDF to ${targetFormat.toUpperCase()}. This may be due to:
      1. Complex PDF structure (scanned images, unusual fonts, complex layouts)
      2. Missing Python libraries (pdf2docx, PyMuPDF, pdfplumber)
      3. Corrupted or password-protected PDF
      4. ONLYOFFICE Document Server not configured
      
      Recommendation: Deploy ONLYOFFICE Document Server for best results.`);

    } finally {
      // Cleanup temporary files
      try {
        await fs.unlink(tempInputPath).catch(() => {});
        await fs.unlink(tempOutputPath).catch(() => {});
        // Also cleanup any additional temp files that might have been created
        const tempPattern = path.join(tempDir, `${timestamp}_*`);
        try {
          const { stdout } = await execAsync(`ls ${tempPattern}`, { timeout: 5000 });
          if (stdout) {
            const files = stdout.trim().split('\n').filter(f => f.trim());
            for (const file of files) {
              await fs.unlink(file).catch(() => {});
            }
          }
        } catch (cleanupListError) {
          // Ignore cleanup listing errors
        }
      } catch (cleanupError) {
        this.logger.warn(`Cleanup warning: ${cleanupError.message}`);
      }
    }
  }

  /**
   * Convert via ONLYOFFICE Document Server API with enhanced error handling
   */
  private async convertViaOnlyOfficeServer(inputPath: string, targetFormat: string, filename: string): Promise<Buffer | null> {
    try {
      this.logger.log(`🏢 ONLYOFFICE Document Server conversion starting...`);
      
      // Read the PDF file
      const fileBuffer = await fs.readFile(inputPath);
      this.logger.log(`📄 PDF file loaded: ${fileBuffer.length} bytes`);

      // Create form data for file upload
      const formData = new FormData();
      formData.append('file', fileBuffer, {
        filename: filename,
        contentType: 'application/pdf'
      });

      // Try to upload file to ONLYOFFICE server first
      let fileUrl: string;
      try {
        const uploadResponse = await axios.post(`${this.documentServerUrl}/upload`, formData, {
          headers: {
            ...formData.getHeaders(),
          },
          timeout: this.timeout / 2,
          maxContentLength: 100 * 1024 * 1024, // 100MB
        });

        if (uploadResponse.data && uploadResponse.data.url) {
          fileUrl = uploadResponse.data.url;
          this.logger.log(`📤 File uploaded to ONLYOFFICE server: ${fileUrl}`);
        } else {
          throw new Error('No upload URL returned');
        }
      } catch (uploadError) {
        this.logger.warn(`📤 Direct upload failed: ${uploadError.message}, using alternative method`);
        
        // Fallback: Create a temporary accessible URL
        const serverUrl = process.env.SERVER_URL || `http://localhost:${process.env.PORT || 10000}`;
        const tempEndpoint = `temp_${Date.now()}_${filename}`;
        fileUrl = `${serverUrl}/temp/${tempEndpoint}`;
        
        // Note: In production, you would need to implement a temporary file serving endpoint
        // For now, we'll proceed with the conversion request
      }

      // Prepare conversion request with enhanced parameters
      const conversionRequest = {
        async: false,
        filetype: 'pdf',
        key: this.generateConversionKey(filename),
        outputtype: targetFormat,
        title: filename,
        url: fileUrl,
        // Enhanced conversion options
        thumbnail: {
          aspect: 2,
          first: true,
          height: 100,
          width: 100
        },
        // Format-specific import options
        ...(targetFormat === 'docx' && {
          region: 'US',
          delimiter: {
            paragraph: true,
            column: false
          }
        }),
        // PowerPoint-specific options for better conversion
        ...(targetFormat === 'pptx' && {
          region: 'US',
          codePage: 65001, // UTF-8
          delimiter: {
            paragraph: true,
            column: true
          },
          // Enhanced PowerPoint conversion settings
          spreadsheetLayout: {
            orientation: 'landscape',
            fitToPage: true,
            gridLines: false
          },
          // Try to preserve text structure
          textSettings: {
            extractText: true,
            preserveFormatting: true,
            recognizeStructure: true
          }
        }),
        // Excel-specific options
        ...(targetFormat === 'xlsx' && {
          region: 'US',
          codePage: 65001,
          delimiter: {
            paragraph: false,
            column: true,
            row: true
          },
          spreadsheetLayout: {
            orientation: 'portrait',
            fitToPage: false,
            gridLines: true
          }
        })
      };

      // Add JWT token if configured
      if (this.jwtSecret) {
        try {
          conversionRequest['token'] = this.generateJWT(conversionRequest);
          this.logger.log(`🔐 JWT token added for secure conversion`);
        } catch (jwtError) {
          this.logger.warn(`🔐 JWT generation failed: ${jwtError.message}`);
        }
      }

      // Send conversion request
      const conversionUrl = `${this.documentServerUrl}/ConvertService.ashx`;
      this.logger.log(`🔄 Sending conversion request to: ${conversionUrl}`);
      this.logger.log(`📋 Conversion parameters: PDF → ${targetFormat.toUpperCase()}`);

      const response = await axios.post(conversionUrl, conversionRequest, {
        timeout: this.timeout,
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        validateStatus: (status) => status < 500 // Accept 4xx errors for handling
      });

      // Check for conversion errors
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

      // Download the converted file with retry logic
      let fileResponse;
      let retries = 3;
      
      while (retries > 0) {
        try {
          fileResponse = await axios.get(response.data.fileUrl, {
            responseType: 'arraybuffer',
            timeout: this.timeout,
            maxContentLength: 100 * 1024 * 1024, // 100MB
            headers: {
              'Accept': '*/*'
            }
          });
          break;
        } catch (downloadError) {
          retries--;
          if (retries === 0) throw downloadError;
          
          this.logger.warn(`📥 Download attempt failed, retrying... (${3 - retries}/3)`);
          await new Promise(resolve => setTimeout(resolve, 2000)); // Wait 2 seconds
        }
      }

      const convertedBuffer = Buffer.from(fileResponse.data);
      
      // Validate output
      if (convertedBuffer.length === 0) {
        throw new Error('Downloaded file is empty');
      }

      if (convertedBuffer.length < 100) {
        throw new Error(`Downloaded file too small: ${convertedBuffer.length} bytes`);
      }

      // Validate file type based on magic bytes
      const fileSignature = convertedBuffer.subarray(0, 8).toString('hex');
      const expectedSignatures = {
        'docx': ['504b0304', '504b0506', '504b0708'], // ZIP-based formats
        'xlsx': ['504b0304', '504b0506', '504b0708'],
        'pptx': ['504b0304', '504b0506', '504b0708']
      };

      const isValidFormat = expectedSignatures[targetFormat]?.some(sig => 
        fileSignature.toLowerCase().startsWith(sig)
      );

      if (!isValidFormat) {
        this.logger.warn(`⚠️ File signature verification failed. Expected ${targetFormat}, got: ${fileSignature}`);
        // Don't fail here as some servers might return valid files with different signatures
      }

      this.logger.log(`✅ ONLYOFFICE Document Server conversion successful: ${convertedBuffer.length} bytes`);
      return convertedBuffer;

    } catch (error) {
      this.logger.error(`❌ ONLYOFFICE Document Server conversion failed: ${error.message}`);
      
      // Log additional debugging information
      if (error.response) {
        this.logger.error(`📊 Response status: ${error.response.status}`);
        this.logger.error(`📊 Response data: ${JSON.stringify(error.response.data)}`);
      }
      
      return null;
    }
  }

  /**
   * Premium Python conversion using the best available libraries
   */
  private async convertViaPremiumPython(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer | null> {
    const premiumPythonScript = this.generatePremiumPythonScript();
    const scriptPath = inputPath.replace('.pdf', '_premium_convert.py');

    try {
      await fs.writeFile(scriptPath, premiumPythonScript);

      const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
      this.logger.log(`🐍 Executing Premium Python conversion: ${targetFormat.toUpperCase()}`);

      const { stdout, stderr } = await execAsync(command, { 
        timeout: this.timeout,
        maxBuffer: 1024 * 1024 * 20 // 20MB buffer for large outputs
      });

      if (stderr && !stderr.includes('Warning') && !stderr.includes('INFO')) {
        this.logger.warn(`Premium Python stderr: ${stderr}`);
      }

      if (stdout) {
        this.logger.log(`Premium Python stdout: ${stdout}`);
      }

      // Check if output file exists
      const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
      if (!outputExists) {
        throw new Error('Premium Python conversion did not produce output file');
      }

      const result = await fs.readFile(outputPath);
      if (result.length < 100) {
        throw new Error(`Premium Python output file too small: ${result.length} bytes`);
      }

      return result;

    } catch (error) {
      throw new Error(`Premium Python conversion failed: ${error.message}`);
    } finally {
      await fs.unlink(scriptPath).catch(() => {});
    }
  }

  /**
   * Advanced LibreOffice conversion with PDF import optimizations
   */
  private async convertViaAdvancedLibreOffice(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer> {
    const advancedCommands = [];
    
    if (targetFormat === 'pptx') {
      // Specialized PowerPoint conversion commands
      advancedCommands.push(
        // Method 1: Import as Draw document first, then convert to PowerPoint
        `libreoffice --headless --draw --convert-to pptx:"Impress MS PowerPoint 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 2: Use Impress directly with PDF import
        `libreoffice --headless --impress --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 3: Force PDF import filter for better text extraction
        `libreoffice --headless --convert-to pptx --infilter="impress_pdf_Import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 4: Draw import with explicit PDF handling
        `libreoffice --headless --draw --convert-to pptx --infilter="draw_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 5: Writer import first (for text extraction), then export
        `libreoffice --headless --writer --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
      );
    } else {
      // Original commands for other formats
      advancedCommands.push(
        // Method 1: Use Writer with PDF import and OCR-like text extraction
        `libreoffice --headless --writer --convert-to ${targetFormat}:"MS Word 2007 XML" --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 2: Advanced PDF import with text layer preservation
        `libreoffice --headless --convert-to ${targetFormat} --infilter="impress_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 3: Draw import for complex layouts, then export to target format
        `libreoffice --headless --draw --convert-to ${targetFormat} --infilter="draw_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 4: Writer import with enhanced text recovery
        `libreoffice --headless --writer --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 5: Standard conversion with maximum compatibility
        `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`
      );
    }

    let lastError = '';
    
    for (let i = 0; i < advancedCommands.length; i++) {
      const command = advancedCommands[i];
      this.logger.log(`📚 Advanced LibreOffice attempt ${i + 1}/${advancedCommands.length}: ${targetFormat.toUpperCase()}`);
      
      try {
        const { stdout, stderr } = await execAsync(command, { 
          timeout: this.timeout,
          maxBuffer: 1024 * 1024 * 15 // 15MB buffer
        });
        
        if (stdout) {
          this.logger.log(`LibreOffice output (${i + 1}): ${stdout}`);
        }
        
        if (stderr && !stderr.includes('Warning')) {
          this.logger.warn(`LibreOffice stderr (${i + 1}): ${stderr}`);
        }

        // Check for expected output file
        const expectedOutputPath = path.join(
          path.dirname(outputPath), 
          path.basename(inputPath, '.pdf') + '.' + targetFormat
        );

        try {
          await fs.access(expectedOutputPath);
          const result = await fs.readFile(expectedOutputPath);
          
          if (result.length > 1000) { // Require meaningful file size
            this.logger.log(`✅ Advanced LibreOffice success on attempt ${i + 1}: ${result.length} bytes`);
            await fs.unlink(expectedOutputPath).catch(() => {});
            return result;
          } else {
            throw new Error(`Generated file too small: ${result.length} bytes`);
          }
        } catch (fileError) {
          this.logger.warn(`Advanced LibreOffice output check failed (${i + 1}): ${fileError.message}`);
          lastError = fileError.message;
        }
        
      } catch (execError) {
        this.logger.warn(`Advanced LibreOffice execution failed (${i + 1}): ${execError.message}`);
        lastError = execError.message;
      }
    }

    throw new Error(`Advanced LibreOffice conversion failed after ${advancedCommands.length} attempts. Last error: ${lastError}`);
  }

  /**
   * Fallback Python conversion methods for when premium methods fail
   */
  private async convertViaFallbackPython(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer | null> {
    const fallbackPythonScript = this.generateFallbackPythonScript();
    const scriptPath = inputPath.replace('.pdf', '_fallback_convert.py');

    try {
      await fs.writeFile(scriptPath, fallbackPythonScript);

      const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
      this.logger.log(`🛡️ Executing Fallback Python conversion: ${targetFormat.toUpperCase()}`);

      const { stdout, stderr } = await execAsync(command, { 
        timeout: this.timeout,
        maxBuffer: 1024 * 1024 * 15 // 15MB buffer
      });

      if (stderr && !stderr.includes('Warning')) {
        this.logger.warn(`Fallback Python stderr: ${stderr}`);
      }

      if (stdout) {
        this.logger.log(`Fallback Python stdout: ${stdout}`);
      }

      // Check if output file exists
      const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
      if (!outputExists) {
        throw new Error('Fallback Python conversion did not produce output file');
      }

      const result = await fs.readFile(outputPath);
      if (result.length < 50) {
        throw new Error(`Fallback Python output file too small: ${result.length} bytes`);
      }

      return result;

    } catch (error) {
      throw new Error(`Fallback Python conversion failed: ${error.message}`);
    } finally {
      await fs.unlink(scriptPath).catch(() => {});
    }
  }

  /**
   * Generate Premium Python script with the best conversion libraries
   */
  private generatePremiumPythonScript(): string {
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
    """Enhanced PDF to PPTX conversion with text extraction and smart layout"""
    print(f"🚀 Starting ENHANCED PDF to PPTX conversion...")
    
    # Method 1: Advanced text and layout extraction with formatted slides
    try:
        import fitz  # PyMuPDF
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.enum.text import WD_ALIGN_PARAGRAPH
        from pptx.dml.color import RGBColor
        import io
        import re
        
        print("🎨 Using Enhanced PyMuPDF + python-pptx (Smart Layout Method)...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        # Set optimal slide dimensions (16:9 widescreen)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Extract text blocks with detailed formatting information
            blocks = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)["blocks"]
            
            # Create slide with title and content layout
            slide_layout = prs.slide_layouts[1]  # Title and Content layout
            slide = prs.slides.add_slide(slide_layout)
            
            # Initialize slide components
            title_box = slide.shapes.title
            content_box = slide.placeholders[1]
            
            # Extract and organize content
            title_text = ""
            content_lines = []
            large_text_items = []
            regular_text_items = []
            
            for block in blocks:
                if 'lines' in block:
                    block_text = ""
                    max_font_size = 0
                    block_color = None
                    
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span['text'].strip()
                            if text:
                                line_text += text + " "
                                font_size = span.get('size', 12)
                                max_font_size = max(max_font_size, font_size)
                                
                                # Extract color information
                                if 'color' in span and not block_color:
                                    color_int = span['color']
                                    block_color = {
                                        'r': (color_int >> 16) & 255,
                                        'g': (color_int >> 8) & 255,
                                        'b': color_int & 255
                                    }
                        
                        if line_text.strip():
                            block_text += line_text.strip() + "\\n"
                    
                    if block_text.strip():
                        text_item = {
                            'text': block_text.strip(),
                            'font_size': max_font_size,
                            'color': block_color,
                            'bbox': block.get('bbox', [0, 0, 0, 0])
                        }
                        
                        # Categorize text by size and position
                        if max_font_size >= 16:  # Title-like text
                            large_text_items.append(text_item)
                        else:  # Regular content
                            regular_text_items.append(text_item)
            
            # Set slide title from largest/first significant text
            if large_text_items:
                # Use the largest text or text at the top as title
                title_item = max(large_text_items, key=lambda x: x['font_size'])
                title_text = title_item['text'].split('\\n')[0]  # First line only
                title_box.text = title_text
                
                # Format title
                if title_box.text_frame.paragraphs:
                    title_para = title_box.text_frame.paragraphs[0]
                    title_para.font.size = Pt(min(title_item['font_size'] + 4, 28))
                    title_para.font.bold = True
                    if title_item['color']:
                        title_para.font.color.rgb = RGBColor(
                            title_item['color']['r'],
                            title_item['color']['g'],
                            title_item['color']['b']
                        )
            
            # Add content to slide
            if regular_text_items or large_text_items:
                content_text_frame = content_box.text_frame
                content_text_frame.clear()
                
                # Add content from regular text items
                all_content_items = regular_text_items + [item for item in large_text_items if item['text'] != title_text]
                
                for i, item in enumerate(all_content_items):
                    if i > 0:
                        content_text_frame.add_paragraph()
                    
                    para = content_text_frame.paragraphs[i] if i < len(content_text_frame.paragraphs) else content_text_frame.add_paragraph()
                    para.text = item['text']
                    para.font.size = Pt(min(max(item['font_size'], 10), 18))
                    
                    # Apply color if available
                    if item['color']:
                        para.font.color.rgb = RGBColor(
                            item['color']['r'],
                            item['color']['g'],
                            item['color']['b']
                        )
                    
                    # Add some spacing
                    para.space_after = Pt(6)
            
            # If no text was extracted, fall back to high-quality image
            if not title_text and not regular_text_items:
                print(f"📄 No text found on page {page_num + 1}, using high-quality image...")
                
                # Remove title and content placeholders
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
                
                # Render page as very high-quality image
                mat = fitz.Matrix(3.0, 3.0)  # 3x scaling for excellent quality
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                
                # Add image to slide with proper scaling
                image_stream = io.BytesIO(img_data)
                
                # Calculate proper image dimensions while maintaining aspect ratio
                page_rect = page.rect
                aspect_ratio = page_rect.width / page_rect.height
                slide_aspect = prs.slide_width / prs.slide_height
                
                if aspect_ratio > slide_aspect:
                    # Image is wider, fit to width
                    img_width = prs.slide_width
                    img_height = int(prs.slide_width / aspect_ratio)
                    left = 0
                    top = (prs.slide_height - img_height) // 2
                else:
                    # Image is taller, fit to height
                    img_height = prs.slide_height
                    img_width = int(prs.slide_height * aspect_ratio)
                    left = (prs.slide_width - img_width) // 2
                    top = 0
                
                slide.shapes.add_picture(
                    image_stream, 
                    left, top, 
                    width=img_width, 
                    height=img_height
                )
            
            print(f"📄 Enhanced slide {page_num + 1}/{len(pdf_doc)} completed")
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 10000:
            print(f"✅ Enhanced PPTX conversion successful: {os.path.getsize(output_path)} bytes")
            return True
            
    except ImportError:
        print("📦 Installing Enhanced PyMuPDF and python-pptx...")
        if install_package('PyMuPDF') and install_package('python-pptx'):
            return premium_convert_to_pptx(input_path, output_path)
    except Exception as e:
        print(f"❌ Enhanced PPTX conversion failed: {e}")
    
    # Method 2: Fallback to structured text extraction with better layout
    try:
        import fitz
        from pptx import Presentation
        from pptx.util import Inches, Pt
        import io
        
        print("📋 Using Structured Text Extraction Method...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Get text in a more structured way
            text_dict = page.get_text("dict")
            
            # Create slide
            slide_layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            
            # Extract all text
            all_text = page.get_text()
            lines = [line.strip() for line in all_text.split('\\n') if line.strip()]
            
            if lines:
                # Use first line as title
                slide.shapes.title.text = lines[0]
                
                # Use remaining lines as content
                if len(lines) > 1:
                    content_box = slide.placeholders[1]
                    text_frame = content_box.text_frame
                    text_frame.clear()
                    
                    for i, line in enumerate(lines[1:]):
                        if i > 0:
                            text_frame.add_paragraph()
                        para = text_frame.paragraphs[i] if i < len(text_frame.paragraphs) else text_frame.add_paragraph()
                        para.text = line
                        para.font.size = Pt(14)
            else:
                # No text, use image
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                mat = fitz.Matrix(2.5, 2.5)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                image_stream = io.BytesIO(img_data)
                slide.shapes.add_picture(image_stream, 0, 0, prs.slide_width, prs.slide_height)
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 10000:
            print(f"✅ Structured PPTX conversion successful: {os.path.getsize(output_path)} bytes")
            return True
            
    except Exception as e:
        print(f"❌ Structured PPTX conversion failed: {e}")
    
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

  /**
   * Generate Fallback Python script for basic conversion
   */
  private generateFallbackPythonScript(): string {
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
    """Enhanced basic PDF to PPTX conversion with text extraction"""
    try:
        import fitz
        from pptx import Presentation
        from pptx.util import Inches, Pt
        import io
        
        print("🎨 Enhanced Basic PPTX Conversion...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        # Set standard widescreen dimensions
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Try to extract text first
            text = page.get_text()
            text_lines = [line.strip() for line in text.split('\\n') if line.strip()]
            
            if text_lines and len(' '.join(text_lines)) > 50:  # Meaningful text content
                # Create slide with title and content layout
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                
                # Set title from first line
                title = text_lines[0]
                if len(title) > 60:  # Truncate long titles
                    title = title[:60] + "..."
                slide.shapes.title.text = title
                
                # Add content
                if len(text_lines) > 1:
                    content_box = slide.placeholders[1]
                    text_frame = content_box.text_frame
                    text_frame.clear()
                    
                    # Add up to 10 lines of content
                    content_lines = text_lines[1:11]  # Limit content
                    for i, line in enumerate(content_lines):
                        if i > 0:
                            text_frame.add_paragraph()
                        para = text_frame.paragraphs[i] if i < len(text_frame.paragraphs) else text_frame.add_paragraph()
                        para.text = line
                        para.font.size = Pt(14)
                        para.space_after = Pt(6)
                
                print(f"📄 Text slide {page_num + 1} created with {len(text_lines)} lines")
            else:
                # No meaningful text, use high-quality image
                slide_layout = prs.slide_layouts[6]  # Blank layout
                slide = prs.slides.add_slide(slide_layout)
                
                # Render as high-quality image
                mat = fitz.Matrix(2.5, 2.5)  # High resolution
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                image_stream = io.BytesIO(img_data)
                
                # Add image with proper scaling
                page_rect = page.rect
                aspect_ratio = page_rect.width / page_rect.height
                slide_aspect = float(prs.slide_width) / float(prs.slide_height)
                
                if aspect_ratio > slide_aspect:
                    # Fit to width
                    img_width = prs.slide_width
                    img_height = int(prs.slide_width / aspect_ratio)
                    left = 0
                    top = (prs.slide_height - img_height) // 2
                else:
                    # Fit to height
                    img_height = prs.slide_height
                    img_width = int(prs.slide_height * aspect_ratio)
                    left = (prs.slide_width - img_width) // 2
                    top = 0
                
                slide.shapes.add_picture(image_stream, left, top, img_width, img_height)
                print(f"📄 Image slide {page_num + 1} created")
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 5000:
            print(f"✅ Enhanced Basic PPTX conversion successful: {os.path.getsize(output_path)} bytes")
            return True
        
    except ImportError:
        if install_package('PyMuPDF') and install_package('python-pptx'):
            return basic_convert_to_pptx(input_path, output_path)
    except Exception as e:
        print(f"❌ Enhanced Basic PPTX conversion failed: {e}")
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
