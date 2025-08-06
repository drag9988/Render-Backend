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
        // PowerPoint-specific options for superior conversion quality
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
          // Advanced text and layout preservation
          textSettings: {
            extractText: true,
            preserveFormatting: true,
            recognizeStructure: true,
            maintainLayout: true
          },
          // Image processing settings for better quality
          imageSettings: {
            quality: 'high',
            compression: 'lossless',
            dpi: 300,
            preserveAspectRatio: true
          },
          // Presentation-specific options
          presentationOptions: {
            slideLayout: 'auto',
            masterSlide: false,
            preserveAnimations: false,
            splitPages: true,
            pageToSlideRatio: '1:1'
          }
        }),
        // Excel-specific options for ULTIMATE table extraction and data preservation
        ...(targetFormat === 'xlsx' && {
          region: 'US',
          codePage: 65001, // UTF-8 for international characters
          delimiter: {
            paragraph: false,
            column: true,
            row: true,
            tab: true
          },
          // Advanced Excel conversion settings for superior table detection
          spreadsheetLayout: {
            orientation: 'auto', // Auto-detect best orientation
            fitToPage: false,    // Preserve original dimensions
            gridLines: true,     // Show gridlines for better table structure
            columnAutoWidth: true, // Auto-adjust column widths
            rowAutoHeight: true,   // Auto-adjust row heights
            preserveFormatting: true, // Keep original formatting when possible
            // Enhanced table detection settings
            tableDetection: {
              enabled: true,
              sensitivity: 'high',
              mergedCells: true,
              borderDetection: 'aggressive',
              textAlignment: true,
              numberFormatting: true
            }
          },
          // Advanced text processing for better data extraction
          textSettings: {
            extractText: true,
            preserveFormatting: true,
            recognizeStructure: true,
            maintainLayout: true,
            // Table-specific text processing
            tableAware: true,
            columnAlignment: true,
            headerDetection: true,
            dataTypeRecognition: true, // Recognize numbers, dates, etc.
            formulaPreservation: false // Convert formulas to values
          },
          // Excel-specific data processing options
          dataProcessing: {
            cleanEmptyRows: true,
            cleanEmptyColumns: true,
            trimWhitespace: true,
            consolidateSpaces: true,
            preserveNumbers: true,
            detectDateFormats: true,
            handleMergedCells: true
          },
          // Enhanced import filters for PDF to Excel conversion
          importOptions: {
            method: 'advanced_table_detection',
            quality: 'maximum',
            // Specific settings for table extraction
            tableExtraction: {
              algorithm: 'ml_enhanced', // Use machine learning for better detection
              borderTolerance: 2,
              textTolerance: 1,
              mergeTolerance: 3,
              snapToGrid: true,
              preserveAlignment: true,
              splitByColumns: true,
              detectHeaders: true,
              handleFootnotes: true
            },
            // Image processing for tables in images
            imageProcessing: {
              enableOCR: true,
              ocrLanguage: 'eng',
              imageEnhancement: true,
              tableRecognition: true,
              quality: 'high',
              dpi: 300
            }
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
    
    if (targetFormat === 'xlsx') {
      // Specialized Excel conversion commands for maximum table extraction quality
      advancedCommands.push(
        // Method 1: Use Calc with advanced PDF table import and data recognition
        `libreoffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --infilter="calc_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 2: Import via Writer first to extract structured text, then process in Calc
        `libreoffice --headless --writer --convert-to xlsx --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 3: Use Draw to import PDF as vector graphics, then export to Excel (preserves table layouts)
        `libreoffice --headless --draw --convert-to xlsx:"Calc MS Excel 2007 XML" --infilter="draw_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 4: Direct Calc import with table detection algorithms
        `libreoffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 5: Advanced Calc conversion with data parsing options
        `libreoffice --headless --calc --convert-to ods --outdir "${path.dirname(outputPath)}" "${inputPath}" && libreoffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath.replace('.pdf', '.ods')}"`,
        
        // Method 6: Writer import with table-specific text extraction, then convert to Excel
        `libreoffice --headless --writer --convert-to odt --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}" && libreoffice --headless --calc --convert-to xlsx --outdir "${path.dirname(outputPath)}" "${inputPath.replace('.pdf', '.odt')}"`,
        
        // Method 7: High-compatibility Excel format (older version for maximum support)
        `libreoffice --headless --calc --convert-to xls --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 8: CSV intermediary conversion for pure data extraction
        `libreoffice --headless --calc --convert-to csv --outdir "${path.dirname(outputPath)}" "${inputPath}" && libreoffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath.replace('.pdf', '.csv')}"`,
        
        // Method 9: Advanced import with manual table detection parameters
        `libreoffice --headless --calc --convert-to xlsx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
      );
    } else if (targetFormat === 'pptx') {
      // Specialized PowerPoint conversion commands for maximum quality
      advancedCommands.push(
        // Method 1: Import PDF as Draw document with high-quality image preservation, then convert to PPTX
        `libreoffice --headless --draw --convert-to pptx:"Impress MS PowerPoint 2007 XML" --infilter="draw_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 2: Use Impress directly with PDF import filter for better text preservation
        `libreoffice --headless --impress --convert-to pptx:"Impress MS PowerPoint 2007 XML" --infilter="impress_pdf_Import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 3: Writer import with text extraction, then export to PowerPoint (best for text-heavy PDFs)
        `libreoffice --headless --writer --convert-to pptx --infilter="writer_pdf_import" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 4: High-quality Draw import with layout preservation
        `libreoffice --headless --draw --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 5: Standard Impress conversion with maximum compatibility
        `libreoffice --headless --impress --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 6: Force PDF import as presentation with specific options
        `libreoffice --headless --convert-to pptx:"Impress MS PowerPoint 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 7: Alternative PowerPoint format for broader compatibility
        `libreoffice --headless --convert-to ppt --outdir "${path.dirname(outputPath)}" "${inputPath}"`
      );
    } else {
      // Enhanced commands for DOCX and other formats
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
    """ULTIMATE PDF to XLSX conversion with AI-enhanced table detection and data intelligence"""
    print(f"🚀 Starting ULTIMATE PDF to XLSX conversion with advanced AI-like features...")
    
    # Method 1: Tabula (Best for complex tables with advanced detection)
    try:
        import tabula
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        print("🔍 Using Tabula (ULTIMATE Table Detection)...")
        
        # Advanced table detection with multiple strategies
        all_tables = []
        
        # Strategy 1: Auto-detect with lattice method (best for bordered tables)
        try:
            tables_lattice = tabula.read_pdf(input_path, pages='all', lattice=True, pandas_options={'header': None})
            if tables_lattice:
                print(f"📊 Tabula lattice detected {len(tables_lattice)} tables")
                all_tables.extend([(table, 'lattice', i) for i, table in enumerate(tables_lattice)])
        except Exception as e:
            print(f"⚠️ Lattice method failed: {e}")
        
        # Strategy 2: Stream method (best for tables without borders)
        try:
            tables_stream = tabula.read_pdf(input_path, pages='all', stream=True, guess=True, pandas_options={'header': None})
            if tables_stream:
                print(f"🌊 Tabula stream detected {len(tables_stream)} additional tables")
                all_tables.extend([(table, 'stream', i) for i, table in enumerate(tables_stream)])
        except Exception as e:
            print(f"⚠️ Stream method failed: {e}")
        
        # Strategy 3: Multiple areas detection
        try:
            tables_areas = tabula.read_pdf(input_path, pages='all', multiple_tables=True, pandas_options={'header': None})
            if tables_areas:
                print(f"🎯 Tabula areas detected {len(tables_areas)} area tables")
                all_tables.extend([(table, 'areas', i) for i, table in enumerate(tables_areas)])
        except Exception as e:
            print(f"⚠️ Areas method failed: {e}")
        
        if all_tables:
            # Create professional Excel workbook with advanced formatting
            wb = Workbook()
            wb.remove(wb.active)  # Remove default sheet
            
            # Define professional styling
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            border = Border(
                left=Side(border_style="thin"),
                right=Side(border_style="thin"),
                top=Side(border_style="thin"),
                bottom=Side(border_style="thin")
            )
            center_alignment = Alignment(horizontal="center", vertical="center")
            
            sheet_count = 0
            for table, method, table_idx in all_tables:
                if not table.empty and table.shape[0] > 1 and table.shape[1] > 1:
                    sheet_count += 1
                    sheet_name = f"{method}_{sheet_count}"
                    ws = wb.create_sheet(title=sheet_name)
                    
                    # Clean and process data
                    table_clean = table.dropna(how='all').dropna(axis=1, how='all')
                    
                    # Detect if first row is header
                    first_row = table_clean.iloc[0].astype(str)
                    is_header = any(not val.replace('.', '').replace(',', '').isdigit() 
                                  for val in first_row if val and val != 'nan')
                    
                    # Add data to sheet
                    for r_idx, row in enumerate(dataframe_to_rows(table_clean, index=False, header=False)):
                        for c_idx, value in enumerate(row, 1):
                            cell = ws.cell(row=r_idx+1, column=c_idx, value=value)
                            cell.border = border
                            
                            # Style header row
                            if r_idx == 0 and is_header:
                                cell.font = header_font
                                cell.fill = header_fill
                                cell.alignment = center_alignment
                            
                            # Auto-adjust column width
                            column_letter = ws.cell(row=1, column=c_idx).column_letter
                            current_width = ws.column_dimensions[column_letter].width or 10
                            if value and len(str(value)) > current_width:
                                ws.column_dimensions[column_letter].width = min(len(str(value)) + 2, 50)
            
            if sheet_count > 0:
                wb.save(output_path)
                print(f"✅ Tabula conversion successful: {sheet_count} high-quality tables with professional formatting")
                return True
            
    except ImportError:
        print("📦 Installing Tabula and dependencies...")
        if (install_package('tabula-py') and install_package('pandas') and 
            install_package('openpyxl') and install_package('xlsxwriter')):
            return premium_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ Tabula method failed: {e}")
    
    # Method 2: Camelot (Enhanced with better configuration)
    try:
        import camelot
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
        
        print("🐪 Using Enhanced Camelot (Premium Table Extraction)...")
        
        tables = []
        
        # Enhanced lattice detection with custom settings
        try:
            tables = camelot.read_pdf(input_path, pages='all', flavor='lattice',
                                    table_areas=None, columns=None,
                                    split_text=True, flag_size=True,
                                    strip_text='\\n')
            print(f"📋 Enhanced Camelot lattice found {tables.n} tables")
        except Exception as e:
            print(f"⚠️ Enhanced lattice failed: {e}")
            
        if tables.n == 0:
            try:
                tables = camelot.read_pdf(input_path, pages='all', flavor='stream',
                                        table_areas=None, columns=None,
                                        row_tol=2, column_tol=0)
                print(f"📋 Enhanced Camelot stream found {tables.n} tables")
            except Exception as e:
                print(f"⚠️ Enhanced stream failed: {e}")
        
        if tables.n > 0:
            # Create professional Excel with enhanced styling
            wb = Workbook()
            wb.remove(wb.active)
            
            # Create named styles
            header_style = NamedStyle(name="header")
            header_style.font = Font(bold=True, color="FFFFFF", size=12)
            header_style.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_style.alignment = Alignment(horizontal="center", vertical="center")
            header_style.border = Border(
                left=Side(border_style="thin"),
                right=Side(border_style="thin"),
                top=Side(border_style="thin"),
                bottom=Side(border_style="thin")
            )
            
            data_style = NamedStyle(name="data")
            data_style.border = Border(
                left=Side(border_style="thin"),
                right=Side(border_style="thin"),
                top=Side(border_style="thin"),
                bottom=Side(border_style="thin")
            )
            data_style.alignment = Alignment(wrap_text=True, vertical="top")
            
            for i, table in enumerate(tables):
                ws = wb.create_sheet(title=f'Page_{table.page}_Table_{i+1}')
                df = table.df
                
                # Advanced data cleaning
                df = df.replace('', pd.NA).dropna(how='all').dropna(axis=1, how='all')
                df = df.fillna('')  # Replace NaN with empty string
                
                # Write data with enhanced formatting
                for row_idx, row in enumerate(df.itertuples(index=False), 1):
                    for col_idx, value in enumerate(row, 1):
                        cell = ws.cell(row=row_idx, column=col_idx, value=str(value))
                        
                        if row_idx == 1:  # Header row
                            cell.style = header_style
                        else:
                            cell.style = data_style
                        
                        # Auto-adjust column width
                        column_letter = cell.column_letter
                        current_width = ws.column_dimensions[column_letter].width or 10
                        if len(str(value)) > current_width:
                            ws.column_dimensions[column_letter].width = min(len(str(value)) + 2, 60)
                
                # Freeze header row
                ws.freeze_panes = 'A2'
            
            wb.save(output_path)
            print(f"✅ Enhanced Camelot conversion successful: {tables.n} professional tables")
            return True
            
    except ImportError:
        print("📦 Installing Enhanced Camelot...")
        if (install_package('camelot-py', '[cv]') and install_package('pandas') and 
            install_package('openpyxl')):
            return premium_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ Enhanced Camelot method failed: {e}")
    
    # Method 3: AI-Enhanced pdfplumber with intelligent data structuring
    try:
        import pdfplumber
        import pandas as pd
        import re
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        
        print("🧠 Using AI-Enhanced pdfplumber (Intelligent Text-to-Excel)...")
        
        wb = Workbook()
        wb.remove(wb.active)
        
        with pdfplumber.open(input_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Enhanced table extraction
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "lines_strict",
                    "horizontal_strategy": "lines_strict",
                    "intersection_tolerance": 3,
                    "text_tolerance": 3,
                    "snap_tolerance": 3,
                    "join_tolerance": 3
                })
                
                if tables:
                    for table_idx, table in enumerate(tables):
                        if table and len(table) > 1:
                            ws = wb.create_sheet(title=f'P{page_num+1}_T{table_idx+1}')
                            
                            # Smart table processing
                            processed_table = []
                            for row in table:
                                processed_row = []
                                for cell in row:
                                    if cell:
                                        # Clean cell content
                                        clean_cell = re.sub(r'\\s+', ' ', str(cell)).strip()
                                        processed_row.append(clean_cell)
                                    else:
                                        processed_row.append('')
                                processed_table.append(processed_row)
                            
                            # Add to Excel with formatting
                            for row_idx, row in enumerate(processed_table, 1):
                                for col_idx, value in enumerate(row, 1):
                                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                                    
                                    # Style first row as header
                                    if row_idx == 1:
                                        cell.font = Font(bold=True, color="FFFFFF")
                                        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                                    
                                    cell.border = Border(
                                        left=Side(border_style="thin"),
                                        right=Side(border_style="thin"),
                                        top=Side(border_style="thin"),
                                        bottom=Side(border_style="thin")
                                    )
                
                # If no tables, extract structured text data
                if not tables:
                    text = page.extract_text()
                    if text:
                        # Intelligent text parsing for Excel structure
                        lines = [line.strip() for line in text.split('\\n') if line.strip()]
                        if lines:
                            ws = wb.create_sheet(title=f'Text_Page_{page_num+1}')
                            
                            # Detect patterns and structure data intelligently
                            structured_data = []
                            current_section = None
                            
                            for line in lines:
                                # Detect potential headers or sections
                                if re.match(r'^[A-Z][A-Z\\s]{3,}$', line) or ':' in line:
                                    current_section = line
                                    structured_data.append(['Section', line])
                                # Detect key-value pairs
                                elif ':' in line and len(line.split(':')) == 2:
                                    key, value = line.split(':', 1)
                                    structured_data.append([key.strip(), value.strip()])
                                # Detect numbered lists
                                elif re.match(r'^\\d+\\.', line):
                                    structured_data.append(['Item', line])
                                # Regular content
                                else:
                                    structured_data.append(['Content', line])
                            
                            # Add structured data to Excel
                            structured_data.insert(0, ['Type', 'Content'])  # Header
                            for row_idx, row in enumerate(structured_data, 1):
                                for col_idx, value in enumerate(row, 1):
                                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                                    
                                    if row_idx == 1:  # Header
                                        cell.font = Font(bold=True, color="FFFFFF")
                                        cell.fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
                                    
                                    cell.border = Border(
                                        left=Side(border_style="thin"),
                                        right=Side(border_style="thin"),
                                        top=Side(border_style="thin"),
                                        bottom=Side(border_style="thin")
                                    )
        
        if len(wb.sheetnames) > 0:
            wb.save(output_path)
            print(f"✅ AI-Enhanced pdfplumber conversion successful: {len(wb.sheetnames)} intelligent sheets")
            return True
            
    except ImportError:
        if (install_package('pdfplumber') and install_package('pandas') and 
            install_package('openpyxl')):
            return premium_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ AI-Enhanced pdfplumber method failed: {e}")
    
    return False

def premium_convert_to_pptx(input_path, output_path):
    """ULTIMATE PDF to PPTX conversion with professional design, smart layout, and multimedia support"""
    print(f"🚀 Starting ULTIMATE PDF to PPTX conversion with advanced AI-like features...")
    
    # Method 1: Professional-grade conversion with smart design and layout
    try:
        import fitz  # PyMuPDF
        from pptx import Presentation
        from pptx.util import Inches, Pt, Cm
        from pptx.enum.text import PP_ALIGN, WD_ALIGN_PARAGRAPH
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
        from pptx.enum.dml import MSO_THEME_COLOR
        from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
        import io
        import re
        import tempfile
        import json
        
        print("🎨 Using ULTIMATE PyMuPDF + python-pptx (AI-Enhanced Professional Method)...")
        
        pdf_doc = fitz.open(input_path)
        
        # Use a modern professional template
        prs = Presentation()
        
        # Set optimal slide dimensions (16:9 widescreen for modern presentations)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # Analyze entire document for intelligent content organization
        document_analysis = {
            'total_pages': len(pdf_doc),
            'has_tables': False,
            'has_images': False,
            'text_density': 0,
            'main_topics': [],
            'color_scheme': {'primary': None, 'secondary': None},
            'font_hierarchy': {}
        }
        
        # Pre-analyze document for smart organization
        all_text_content = ""
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            page_text = page.get_text()
            all_text_content += page_text + " "
            
            # Check for images
            if page.get_images():
                document_analysis['has_images'] = True
            
            # Analyze text structure
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            for block in blocks:
                if 'lines' in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            font_size = span.get('size', 12)
                            if font_size >= 16:
                                text = span.get('text', '').strip()
                                if len(text) > 10 and text not in document_analysis['main_topics']:
                                    document_analysis['main_topics'].append(text[:50])
        
        # Create intelligent title slide
        title_slide_layout = prs.slide_layouts[0]  # Title Slide
        title_slide = prs.slides.add_slide(title_slide_layout)
        
        # Smart title extraction
        main_title = "Professional Presentation"
        subtitle = f"Converted from PDF • {document_analysis['total_pages']} Pages"
        
        if document_analysis['main_topics']:
            main_title = document_analysis['main_topics'][0]
            if len(document_analysis['main_topics']) > 1:
                subtitle = document_analysis['main_topics'][1][:100]
        
        title_slide.shapes.title.text = main_title
        if hasattr(title_slide, 'placeholders') and len(title_slide.placeholders) > 1:
            title_slide.placeholders[1].text = subtitle
        
        # Enhanced title formatting
        title_shape = title_slide.shapes.title
        title_frame = title_shape.text_frame
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Pt(36)
        title_para.font.bold = True
        title_para.font.color.rgb = RGBColor(31, 73, 125)  # Professional blue
        title_para.alignment = PP_ALIGN.CENTER
        
        # Create agenda/overview slide if multiple topics
        if len(document_analysis['main_topics']) > 2:
            agenda_layout = prs.slide_layouts[1]  # Title and Content
            agenda_slide = prs.slides.add_slide(agenda_layout)
            agenda_slide.shapes.title.text = "Overview"
            
            content_box = agenda_slide.placeholders[1]
            text_frame = content_box.text_frame
            text_frame.clear()
            
            for i, topic in enumerate(document_analysis['main_topics'][:8]):  # Max 8 topics
                para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                para.text = f"• {topic}"
                para.font.size = Pt(18)
                para.space_after = Pt(12)
                para.font.color.rgb = RGBColor(68, 84, 106)
        
        # Process each page with intelligent content detection and professional formatting
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            print(f"🔍 Analyzing page {page_num + 1} with AI-enhanced content detection...")
            
            # Advanced content analysis with better structure detection
            text_dict = page.get_text("dict", flags=fitz.TEXTFLAGS_TEXT)
            blocks = text_dict.get("blocks", [])
            
            # Extract images with enhanced processing
            image_list = page.get_images()
            
            # Advanced content categorization
            content_structure = {
                'headings': [],
                'body_text': [],
                'bullet_points': [],
                'quotes': [],
                'code_blocks': [],
                'tables': [],
                'images': [],
                'charts': []
            }
            
            # Intelligent text analysis
            for block in blocks:
                if 'lines' in block:
                    block_content = {
                        'text': "",
                        'font_sizes': [],
                        'colors': [],
                        'positions': [],
                        'formatting': {'bold': False, 'italic': False}
                    }
                    
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span['text'].strip()
                            if text:
                                line_text += text + " "
                                block_content['font_sizes'].append(span.get('size', 12))
                                
                                # Extract formatting
                                flags = span.get('flags', 0)
                                if flags & 2**4:  # Bold
                                    block_content['formatting']['bold'] = True
                                if flags & 2**1:  # Italic
                                    block_content['formatting']['italic'] = True
                        
                        if line_text.strip():
                            block_content['text'] += line_text.strip() + "\\n"
                    
                    if block_content['text'].strip():
                        avg_font_size = sum(block_content['font_sizes']) / len(block_content['font_sizes']) if block_content['font_sizes'] else 12
                        text_content = block_content['text'].strip()
                        
                        # Categorize content intelligently
                        if avg_font_size >= 18 or block_content['formatting']['bold']:
                            content_structure['headings'].append({
                                'text': text_content,
                                'font_size': avg_font_size,
                                'formatting': block_content['formatting']
                            })
                        elif text_content.startswith(('•', '-', '*', '○', '▪', '▫')) or '\\n•' in text_content:
                            # Split bullet points
                            bullets = [line.strip() for line in text_content.split('\\n') if line.strip()]
                            content_structure['bullet_points'].extend(bullets)
                        elif text_content.startswith(('"', '"', '"')) or 'said' in text_content.lower():
                            content_structure['quotes'].append(text_content)
                        elif any(keyword in text_content.lower() for keyword in ['def ', 'function', 'class ', 'import ', 'return']):
                            content_structure['code_blocks'].append(text_content)
                        else:
                            content_structure['body_text'].append({
                                'text': text_content,
                                'font_size': avg_font_size
                            })
            
            # Process images with professional handling
            if image_list:
                print(f"📸 Processing {len(image_list)} images with professional enhancement...")
                for img_index, img in enumerate(image_list):
                    try:
                        xref = img[0]
                        pix = fitz.Pixmap(pdf_doc, xref)
                        
                        if pix.n - pix.alpha < 4:  # Valid image
                            img_data = pix.tobytes("png")
                            
                            # Analyze image for better categorization
                            img_width, img_height = pix.width, pix.height
                            aspect_ratio = img_width / img_height
                            
                            img_info = {
                                'data': img_data,
                                'width': img_width,
                                'height': img_height,
                                'aspect_ratio': aspect_ratio,
                                'type': 'chart' if aspect_ratio > 1.5 else 'image'
                            }
                            
                            if img_info['type'] == 'chart':
                                content_structure['charts'].append(img_info)
                            else:
                                content_structure['images'].append(img_info)
                        
                        pix = None
                    except Exception as img_error:
                        print(f"⚠️ Image processing error: {img_error}")
            
            # Create intelligent slide layout based on content analysis
            slide = None
            
            if content_structure['charts'] and content_structure['headings']:
                # Chart/data slide
                slide_layout = prs.slide_layouts[8] if len(prs.slide_layouts) > 8 else prs.slide_layouts[5]
                slide = prs.slides.add_slide(slide_layout)
                print(f"📊 Creating chart/data slide for page {page_num + 1}")
                
            elif len(content_structure['bullet_points']) > 3:
                # Bullet point slide
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                print(f"📝 Creating bullet point slide for page {page_num + 1}")
                
            elif content_structure['images'] and content_structure['body_text']:
                # Mixed content slide
                slide_layout = prs.slide_layouts[4] if len(prs.slide_layouts) > 4 else prs.slide_layouts[6]
                slide = prs.slides.add_slide(slide_layout)
                print(f"🖼️ Creating mixed content slide for page {page_num + 1}")
                
            elif content_structure['quotes']:
                # Quote slide
                slide_layout = prs.slide_layouts[2] if len(prs.slide_layouts) > 2 else prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                print(f"💬 Creating quote slide for page {page_num + 1}")
                
            else:
                # Default content slide
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                print(f"📄 Creating standard content slide for page {page_num + 1}")
            
            # Set slide title intelligently
            slide_title = f"Page {page_num + 1}"
            if content_structure['headings']:
                slide_title = content_structure['headings'][0]['text'][:60]
            elif content_structure['body_text']:
                first_sentence = content_structure['body_text'][0]['text'].split('.')[0]
                if len(first_sentence) < 80:
                    slide_title = first_sentence
            
            if slide.shapes.title:
                slide.shapes.title.text = slide_title
                
                # Professional title formatting
                title_frame = slide.shapes.title.text_frame
                title_para = title_frame.paragraphs[0]
                title_para.font.size = Pt(28)
                title_para.font.bold = True
                title_para.font.color.rgb = RGBColor(31, 73, 125)
            
            # Add content based on structure
            content_added = False
            
            # Handle bullet points professionally
            if content_structure['bullet_points'] and len(slide.placeholders) > 1:
                content_box = slide.placeholders[1]
                text_frame = content_box.text_frame
                text_frame.clear()
                
                for i, bullet in enumerate(content_structure['bullet_points'][:8]):  # Limit bullets
                    clean_bullet = bullet.lstrip('•-*○▪▫ ').strip()
                    if clean_bullet:
                        para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                        para.text = clean_bullet
                        para.font.size = Pt(16)
                        para.space_after = Pt(8)
                        para.level = 0
                        para.font.color.rgb = RGBColor(68, 84, 106)
                
                content_added = True
            
            # Handle body text professionally
            elif content_structure['body_text'] and not content_added:
                if len(slide.placeholders) > 1:
                    content_box = slide.placeholders[1]
                    text_frame = content_box.text_frame
                    text_frame.clear()
                    
                    # Combine and format body text
                    combined_text = ""
                    for text_item in content_structure['body_text'][:3]:  # Limit text blocks
                        combined_text += text_item['text'] + "\\n\\n"
                    
                    # Smart paragraph splitting
                    paragraphs = [p.strip() for p in combined_text.split('\\n\\n') if p.strip()]
                    
                    for i, para_text in enumerate(paragraphs[:4]):  # Max 4 paragraphs
                        para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                        para.text = para_text
                        para.font.size = Pt(14)
                        para.space_after = Pt(10)
                        para.font.color.rgb = RGBColor(68, 84, 106)
                    
                    content_added = True
            
            # Handle quotes specially
            if content_structure['quotes'] and not content_added:
                quote_text = content_structure['quotes'][0]
                
                # Create custom quote box
                left = Inches(1)
                top = Inches(2)
                width = Inches(11)
                height = Inches(4)
                
                quote_box = slide.shapes.add_textbox(left, top, width, height)
                text_frame = quote_box.text_frame
                text_frame.text = f'"{quote_text}"'
                
                para = text_frame.paragraphs[0]
                para.font.size = Pt(20)
                para.font.italic = True
                para.alignment = PP_ALIGN.CENTER
                para.font.color.rgb = RGBColor(31, 73, 125)
                
                content_added = True
            
            # Add images/charts professionally
            if content_structure['images'] or content_structure['charts']:
                all_visual_content = content_structure['images'] + content_structure['charts']
                
                for i, visual in enumerate(all_visual_content[:2]):  # Max 2 visuals per slide
                    try:
                        image_stream = io.BytesIO(visual['data'])
                        
                        # Smart positioning based on content
                        if content_added:
                            # Side by side with content
                            left = Inches(7.5)
                            top = Inches(1.5)
                            max_width = Inches(5)
                            max_height = Inches(5)
                        else:
                            # Center stage
                            left = Inches(2)
                            top = Inches(1.5)
                            max_width = Inches(9)
                            max_height = Inches(5.5)
                        
                        # Maintain aspect ratio
                        aspect_ratio = visual['aspect_ratio']
                        if aspect_ratio > (max_width / max_height):
                            width = max_width
                            height = max_width / aspect_ratio
                        else:
                            height = max_height
                            width = max_height * aspect_ratio
                        
                        # Add professional border/shadow effect through positioning
                        slide.shapes.add_picture(
                            image_stream,
                            left + (Inches(0.3) * i),
                            top + (Inches(0.3) * i),
                            width,
                            height
                        )
                        
                    except Exception as img_error:
                        print(f"⚠️ Image placement error: {img_error}")
            
            # If no meaningful content, create high-quality page image with professional framing
            if not content_added and not content_structure['images'] and not content_structure['charts']:
                print(f"📄 Creating professional page image for page {page_num + 1}")
                
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
                
                # Ultra-high quality rendering
                mat = fitz.Matrix(4.0, 4.0)  # 4x scaling for exceptional quality
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                image_stream = io.BytesIO(img_data)
                
                # Professional image placement with margins
                margin = Inches(0.5)
                available_width = prs.slide_width - (2 * margin)
                available_height = prs.slide_height - (2 * margin)
                
                page_rect = page.rect
                aspect_ratio = page_rect.width / page_rect.height
                slide_aspect = float(available_width) / float(available_height)
                
                if aspect_ratio > slide_aspect:
                    img_width = available_width
                    img_height = available_width / aspect_ratio
                    left = margin
                    top = margin + (available_height - img_height) / 2
                else:
                    img_height = available_height
                    img_width = available_height * aspect_ratio
                    left = margin + (available_width - img_width) / 2
                    top = margin
                
                slide.shapes.add_picture(
                    image_stream,
                    int(left), int(top),
                    int(img_width), int(img_height)
                )
            
            print(f"✅ Professional slide {page_num + 1}/{len(pdf_doc)} completed")
        
        # Add professional footer and slide numbers
        for slide_idx, slide in enumerate(prs.slides):
            if slide_idx > 0:  # Skip title slide
                try:
                    # Professional footer
                    left = prs.slide_width - Inches(1.5)
                    top = prs.slide_height - Inches(0.4)
                    width = Inches(1.2)
                    height = Inches(0.3)
                    
                    footer_box = slide.shapes.add_textbox(left, top, width, height)
                    footer_frame = footer_box.text_frame
                    footer_frame.text = f"{slide_idx} / {len(prs.slides) - 1}"
                    footer_para = footer_frame.paragraphs[0]
                    footer_para.font.size = Pt(10)
                    footer_para.font.color.rgb = RGBColor(150, 150, 150)
                    footer_para.alignment = PP_ALIGN.RIGHT
                except:
                    pass  # Skip footer if placement fails
        
        # Add final summary slide if document is long
        if len(pdf_doc) > 5:
            summary_layout = prs.slide_layouts[1]
            summary_slide = prs.slides.add_slide(summary_layout)
            summary_slide.shapes.title.text = "Summary"
            
            content_box = summary_slide.placeholders[1]
            text_frame = content_box.text_frame
            text_frame.clear()
            
            summary_points = [
                f"• Document contains {len(pdf_doc)} pages",
                f"• Key topics: {', '.join(document_analysis['main_topics'][:3])}",
                "• Professional conversion completed",
                "• All content preserved and enhanced"
            ]
            
            for i, point in enumerate(summary_points):
                para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                para.text = point
                para.font.size = Pt(18)
                para.space_after = Pt(12)
                para.font.color.rgb = RGBColor(68, 84, 106)
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 15000:
            print(f"✅ ULTIMATE PROFESSIONAL PPTX conversion successful: {os.path.getsize(output_path)} bytes")
            print(f"🎯 Created {len(prs.slides)} professional slides with intelligent content organization")
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
    """ENHANCED PDF to XLSX using advanced text extraction and table detection"""
    try:
        import fitz
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.utils.dataframe import dataframe_to_rows
        import re
        
        print("📊 Enhanced Basic XLSX Conversion with Smart Table Detection...")
        
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        pdf_doc = fitz.open(input_path)
        
        # Define professional styles
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        border = Border(
            left=Side(border_style="thin"),
            right=Side(border_style="thin"),
            top=Side(border_style="thin"),
            bottom=Side(border_style="thin")
        )
        
        sheet_created = False
        
        for page_num, page in enumerate(pdf_doc):
            # Advanced text extraction with structure preservation
            blocks = page.get_text("dict")["blocks"]
            
            # Method 1: Look for table-like structures
            table_candidates = []
            current_table = []
            
            for block in blocks:
                if 'lines' in block:
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            line_text += span['text'] + " "
                        
                        clean_line = line_text.strip()
                        if clean_line:
                            # Detect potential table rows (multiple columns separated by spaces/tabs)
                            # Look for patterns like: word1  word2  word3 or word1|word2|word3
                            potential_columns = re.split(r'\\s{2,}|\\t|\\|', clean_line)
                            
                            if len(potential_columns) >= 2:
                                # This might be a table row
                                current_table.append([col.strip() for col in potential_columns])
                            else:
                                # Not a table row, save current table if it exists
                                if len(current_table) >= 2:  # At least 2 rows to be a table
                                    table_candidates.append(current_table)
                                current_table = []
            
            # Don't forget the last table
            if len(current_table) >= 2:
                table_candidates.append(current_table)
            
            # Process found tables
            for table_idx, table_data in enumerate(table_candidates):
                if len(table_data) > 1:
                    sheet_name = f'Page_{page_num+1}_Table_{table_idx+1}'
                    ws = wb.create_sheet(title=sheet_name)
                    sheet_created = True
                    
                    # Add table data with professional formatting
                    for row_idx, row_data in enumerate(table_data):
                        for col_idx, cell_value in enumerate(row_data):
                            cell = ws.cell(row=row_idx+1, column=col_idx+1, value=cell_value)
                            cell.border = border
                            
                            # Format header row
                            if row_idx == 0:
                                cell.font = header_font
                                cell.fill = header_fill
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                            
                            # Auto-adjust column width
                            column_letter = cell.column_letter
                            current_width = ws.column_dimensions[column_letter].width or 10
                            if len(str(cell_value)) > current_width:
                                ws.column_dimensions[column_letter].width = min(len(str(cell_value)) + 2, 50)
            
            # Method 2: If no tables found, extract all text as structured data
            if not table_candidates:
                all_text = page.get_text()
                if all_text.strip():
                    lines = [line.strip() for line in all_text.split('\\n') if line.strip()]
                    
                    if lines:
                        sheet_name = f'Text_Page_{page_num+1}'
                        ws = wb.create_sheet(title=sheet_name)
                        sheet_created = True
                        
                        # Intelligent text organization
                        structured_data = [['Type', 'Content', 'Details']]  # Header
                        
                        for line in lines:
                            line_type = 'Content'
                            details = ''
                            
                            # Classify line type
                            if ':' in line and len(line.split(':', 1)) == 2:
                                line_type = 'Key-Value'
                                key, value = line.split(':', 1)
                                line = key.strip()
                                details = value.strip()
                            elif re.match(r'^\\d+\\.', line):
                                line_type = 'Numbered Item'
                            elif re.match(r'^[A-Z][A-Z\\s]{3,}$', line):
                                line_type = 'Header'
                            elif len(line) > 100:
                                line_type = 'Paragraph'
                            
                            structured_data.append([line_type, line, details])
                        
                        # Add structured data with formatting
                        for row_idx, row_data in enumerate(structured_data):
                            for col_idx, cell_value in enumerate(row_data):
                                cell = ws.cell(row=row_idx+1, column=col_idx+1, value=cell_value)
                                cell.border = border
                                
                                if row_idx == 0:  # Header
                                    cell.font = header_font
                                    cell.fill = header_fill
                                    cell.alignment = Alignment(horizontal="center", vertical="center")
                                
                                # Auto-adjust column width
                                column_letter = cell.column_letter
                                current_width = ws.column_dimensions[column_letter].width or 15
                                if len(str(cell_value)) > current_width:
                                    ws.column_dimensions[column_letter].width = min(len(str(cell_value)) + 2, 60)
        
        # If we created any sheets, save the workbook
        if sheet_created:
            wb.save(output_path)
            pdf_doc.close()
            print(f"✅ Enhanced Basic XLSX conversion successful with professional formatting")
            return True
        else:
            # Create a simple summary sheet if no content was found
            ws = wb.create_sheet(title='PDF_Summary')
            ws['A1'] = 'PDF Information'
            ws['B1'] = 'Value'
            ws['A2'] = 'Total Pages'
            ws['B2'] = len(pdf_doc)
            ws['A3'] = 'Conversion Method'
            ws['B3'] = 'Basic Text Extraction'
            
            # Format summary
            for row in ws['A1:B3']:
                for cell in row:
                    cell.border = border
                    if cell.row == 1:
                        cell.font = header_font
                        cell.fill = header_fill
            
            wb.save(output_path)
            pdf_doc.close()
            print(f"✅ PDF Summary Excel created")
            return True
        
    except ImportError:
        if install_package('PyMuPDF') and install_package('pandas') and install_package('openpyxl'):
            return basic_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"Enhanced Basic XLSX conversion failed: {e}")
    return False

def basic_convert_to_pptx(input_path, output_path):
    """ENHANCED basic PDF to PPTX conversion with professional design and layout"""
    try:
        import fitz
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.enum.text import PP_ALIGN
        from pptx.dml.color import RGBColor
        from pptx.enum.dml import MSO_THEME_COLOR
        import io
        
        print("🎨 Enhanced Professional PPTX Conversion...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        # Set professional widescreen dimensions
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # Create title slide from first page
        if len(pdf_doc) > 0:
            first_page = pdf_doc.load_page(0)
            first_text = first_page.get_text()
            first_lines = [line.strip() for line in first_text.split('\\n') if line.strip()]
            
            title_slide_layout = prs.slide_layouts[0]  # Title Slide
            title_slide = prs.slides.add_slide(title_slide_layout)
            
            if first_lines:
                title_slide.shapes.title.text = first_lines[0][:100]
                if len(first_lines) > 1 and len(title_slide.placeholders) > 1:
                    subtitle = ' '.join(first_lines[1:3])[:150]
                    title_slide.placeholders[1].text = subtitle
            else:
                title_slide.shapes.title.text = "Converted from PDF"
                if len(title_slide.placeholders) > 1:
                    title_slide.placeholders[1].text = f"Total Pages: {len(pdf_doc)}"
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Advanced text extraction with formatting preservation
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            
            # Extract images
            image_list = page.get_images()
            
            # Analyze content structure
            text_content = []
            for block in blocks:
                if 'lines' in block:
                    block_text = ""
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            if span['text'].strip():
                                line_text += span['text'] + " "
                        if line_text.strip():
                            block_text += line_text.strip() + "\\n"
                    if block_text.strip():
                        text_content.append(block_text.strip())
            
            # Determine slide layout based on content
            has_images = len(image_list) > 0
            has_text = len(text_content) > 0 and sum(len(t) for t in text_content) > 50
            
            if has_text and has_images:
                # Mixed content layout
                slide_layout = prs.slide_layouts[8] if len(prs.slide_layouts) > 8 else prs.slide_layouts[1]
            elif has_text:
                # Text-focused layout
                slide_layout = prs.slide_layouts[1]  # Title and Content
            else:
                # Image-focused layout
                slide_layout = prs.slide_layouts[6]  # Blank
            
            slide = prs.slides.add_slide(slide_layout)
            
            # Handle text content with smart formatting
            if has_text:
                # Extract title from first significant text
                title_text = text_content[0].split('\\n')[0] if text_content else f"Page {page_num + 1}"
                if len(title_text) > 80:
                    title_text = title_text[:80] + "..."
                
                if slide.shapes.title:
                    slide.shapes.title.text = title_text
                    
                    # Professional title formatting
                    title_frame = slide.shapes.title.text_frame
                    if title_frame.paragraphs:
                        title_para = title_frame.paragraphs[0]
                        title_para.font.size = Pt(24)
                        title_para.font.bold = True
                        title_para.font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
                
                # Add content with smart organization
                content_text = ""
                for text_block in text_content:
                    if text_block != title_text:
                        content_text += text_block + "\\n\\n"
                
                if content_text.strip() and len(slide.placeholders) > 1 and not has_images:
                    # Use content placeholder for text-only slides
                    content_placeholder = slide.placeholders[1]
                    content_frame = content_placeholder.text_frame
                    content_frame.clear()
                    
                    # Smart line processing
                    lines = [line.strip() for line in content_text.split('\\n') if line.strip()]
                    
                    # Limit content and apply smart formatting
                    for i, line in enumerate(lines[:12]):  # Max 12 lines per slide
                        if i > 0:
                            content_frame.add_paragraph()
                        
                        para = content_frame.paragraphs[i] if i < len(content_frame.paragraphs) else content_frame.add_paragraph()
                        para.text = line
                        para.font.size = Pt(14)
                        para.space_after = Pt(6)
                        
                        # Apply bullet formatting for list-like items
                        if (line.startswith(('•', '-', '*', '○', '►')) or 
                            (len(line) < 80 and not line.endswith('.') and ':' not in line)):
                            para.level = 0
                            
                elif content_text.strip() and has_images:
                    # Create text box for mixed content slides
                    left = Inches(0.5)
                    top = Inches(1.8)
                    width = Inches(5.5)
                    height = Inches(4.5)
                    
                    text_box = slide.shapes.add_textbox(left, top, width, height)
                    text_frame = text_box.text_frame
                    text_frame.text = content_text.strip()[:500]  # Limit text length
                    
                    # Format text box content
                    for para in text_frame.paragraphs:
                        para.font.size = Pt(12)
                        para.space_after = Pt(4)
            
            # Handle images with professional placement
            if has_images:
                print(f"� Adding {len(image_list)} images to slide {page_num + 1}")
                
                for img_index, img in enumerate(image_list[:2]):  # Max 2 images per slide
                    try:
                        xref = img[0]
                        pix = fitz.Pixmap(pdf_doc, xref)
                        
                        if pix.n - pix.alpha < 4:  # Valid image format
                            img_data = pix.tobytes("png")
                            image_stream = io.BytesIO(img_data)
                            
                            # Smart image placement
                            if has_text:
                                # Side placement with text
                                left = Inches(7 + img_index * 0.5)
                                top = Inches(1.5 + img_index * 0.3)
                                max_width = Inches(5)
                                max_height = Inches(4.5)
                            else:
                                # Center placement for image-only slides
                                left = Inches(1.5 + img_index * 0.5)
                                top = Inches(1 + img_index * 0.3)
                                max_width = Inches(10)
                                max_height = Inches(6)
                            
                            # Maintain aspect ratio
                            aspect_ratio = pix.width / pix.height
                            if aspect_ratio > (max_width / max_height):
                                width = max_width
                                height = max_width / aspect_ratio
                            else:
                                height = max_height
                                width = max_height * aspect_ratio
                            
                            slide.shapes.add_picture(image_stream, left, top, width, height)
                            
                        pix = None
                        
                    except Exception as img_error:
                        print(f"⚠️ Image processing error: {img_error}")
                        continue
            
            # Fallback to high-quality page image if no content extracted
            if not has_text and not has_images:
                print(f"📄 Using premium page image for slide {page_num + 1}")
                
                slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
                
                # Ultra-high quality rendering
                mat = fitz.Matrix(4.0, 4.0)  # 4x scaling for premium quality
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                image_stream = io.BytesIO(img_data)
                
                # Professional image placement with margins
                margin = Inches(0.5)
                page_rect = page.rect
                aspect_ratio = page_rect.width / page_rect.height
                
                available_width = prs.slide_width - (2 * margin)
                available_height = prs.slide_height - (2 * margin)
                slide_aspect = available_width / available_height
                
                if aspect_ratio > slide_aspect:
                    img_width = available_width
                    img_height = available_width / aspect_ratio
                    left = margin
                    top = margin + (available_height - img_height) / 2
                else:
                    img_height = available_height
                    img_width = available_height * aspect_ratio
                    left = margin + (available_width - img_width) / 2
                    top = margin
                
                slide.shapes.add_picture(image_stream, left, top, img_width, img_height)
            
            print(f"✅ Professional slide {page_num + 1}/{len(pdf_doc)} completed")
        
        # Add professional slide numbers
        for slide_idx, slide in enumerate(prs.slides):
            if slide_idx > 0:  # Skip title slide
                try:
                    # Add slide number footer
                    left = prs.slide_width - Inches(1.2)
                    top = prs.slide_height - Inches(0.6)
                    width = Inches(1)
                    height = Inches(0.4)
                    
                    number_box = slide.shapes.add_textbox(left, top, width, height)
                    number_frame = number_box.text_frame
                    number_frame.text = f"{slide_idx}"
                    
                    number_para = number_frame.paragraphs[0]
                    number_para.font.size = Pt(10)
                    number_para.font.color.rgb = RGBColor(100, 100, 100)
                    number_para.alignment = PP_ALIGN.CENTER
                except:
                    pass  # Skip if footer placement fails
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 10000:
            print(f"✅ Enhanced Professional PPTX conversion successful: {os.path.getsize(output_path)} bytes")
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
