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

      // PRIORITY 2: Premium Python Libraries (Excellent quality, especially for DOCX and XLSX)
      try {
        if (targetFormat === 'pptx') {
          this.logger.log(`🎨 Attempting Premium Python PPTX conversion (Simplified Method)...`);
        } else {
          this.logger.log(`🥈 Attempting Premium Python conversion (pdf2docx, PyMuPDF)...`);
        }
        const result = await this.convertViaPremiumPython(tempInputPath, tempOutputPath, targetFormat);
        if (result && result.length > 1000) {
          this.logger.log(`✅ Premium Python conversion: SUCCESS! Output: ${result.length} bytes`);
          return result;
        }
      } catch (pythonError) {
        this.logger.warn(`❌ Premium Python conversion failed: ${pythonError.message}`);
        if (targetFormat === 'pptx') {
          this.logger.error(`🔍 PPTX specific error details: ${pythonError.stack || pythonError.message}`);
        }
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
      this.logger.log(`🐍 Starting Premium Python conversion for ${targetFormat.toUpperCase()}`);
      this.logger.log(`📍 Input path: ${inputPath}`);
      this.logger.log(`📍 Output path: ${outputPath}`);
      this.logger.log(`📍 Script path: ${scriptPath}`);
      this.logger.log(`📍 Python path: ${this.pythonPath}`);
      
      // Check if input file exists
      const inputExists = await fs.access(inputPath).then(() => true).catch(() => false);
      if (!inputExists) {
        throw new Error(`Input PDF file does not exist: ${inputPath}`);
      }
      
      const inputStats = await fs.stat(inputPath);
      this.logger.log(`📄 Input file size: ${inputStats.size} bytes`);

      // Write Python script
      await fs.writeFile(scriptPath, premiumPythonScript);
      this.logger.log(`✅ Python script written successfully`);
      
      // Check if script was written correctly
      const scriptStats = await fs.stat(scriptPath);
      this.logger.log(`📜 Python script size: ${scriptStats.size} bytes`);

      const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
      this.logger.log(`🚀 Executing command: ${command}`);

      const startTime = Date.now();
      const { stdout, stderr } = await execAsync(command, { 
        timeout: this.timeout,
        maxBuffer: 1024 * 1024 * 30 // 30MB buffer for PPTX conversions
      });
      const endTime = Date.now();
      
      this.logger.log(`⏱️ Python execution time: ${endTime - startTime}ms`);

      // Enhanced logging for PPTX debugging
      if (targetFormat === 'pptx') {
        this.logger.log(`🔍 PPTX Debug - Full Python stdout: ${stdout}`);
        if (stderr) {
          this.logger.log(`🔍 PPTX Debug - Full Python stderr: ${stderr}`);
        }
      }

      // Log all output for debugging
      if (stderr) {
        this.logger.log(`🔍 Premium Python stderr (full): ${stderr}`);
        if (!stderr.includes('Warning') && !stderr.includes('INFO') && !stderr.includes('pip as the \'root\' user')) {
          this.logger.warn(`❗ Premium Python stderr (non-warning): ${stderr}`);
        }
      }

      if (stdout) {
        this.logger.log(`📝 Premium Python stdout (full): ${stdout}`);
      }

      // Check if output file exists
      this.logger.log(`🔍 Checking if output file exists: ${outputPath}`);
      const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
      if (!outputExists) {
        this.logger.error(`❌ Output file was not created: ${outputPath}`);
        
        // List files in the directory to see what was created
        try {
          const outputDir = path.dirname(outputPath);
          const files = await fs.readdir(outputDir);
          this.logger.log(`📁 Files in output directory: ${files.join(', ')}`);
        } catch (dirError) {
          this.logger.error(`❌ Could not list output directory: ${dirError.message}`);
        }
        
        throw new Error('Premium Python conversion did not produce output file');
      }

      const result = await fs.readFile(outputPath);
      this.logger.log(`📊 Output file size: ${result.length} bytes`);
      
      if (result.length < 100) {
        this.logger.error(`❌ Output file too small: ${result.length} bytes`);
        throw new Error(`Premium Python output file too small: ${result.length} bytes`);
      }

      this.logger.log(`✅ Premium Python conversion successful!`);
      return result;

    } catch (error) {
      this.logger.error(`❌ Premium Python conversion error: ${error.message}`);
      this.logger.error(`🔍 Full error details: ${JSON.stringify(error, null, 2)}`);
      throw new Error(`Premium Python conversion failed: ${error.message}`);
    } finally {
      // Clean up script file
      try {
        await fs.unlink(scriptPath);
        this.logger.log(`🧹 Cleaned up Python script: ${scriptPath}`);
      } catch (cleanupError) {
        this.logger.warn(`⚠️ Could not clean up script: ${cleanupError.message}`);
      }
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
      this.logger.log(`🛡️ Starting Fallback Python conversion for ${targetFormat.toUpperCase()}`);
      this.logger.log(`📍 Input path: ${inputPath}`);
      this.logger.log(`📍 Output path: ${outputPath}`);
      this.logger.log(`📍 Script path: ${scriptPath}`);

      await fs.writeFile(scriptPath, fallbackPythonScript);
      this.logger.log(`✅ Fallback Python script written successfully`);

      const command = `${this.pythonPath} "${scriptPath}" "${inputPath}" "${outputPath}" "${targetFormat}"`;
      this.logger.log(`� Executing command: ${command}`);

      const startTime = Date.now();
      const { stdout, stderr } = await execAsync(command, { 
        timeout: this.timeout,
        maxBuffer: 1024 * 1024 * 15 // 15MB buffer
      });
      const endTime = Date.now();
      
      this.logger.log(`⏱️ Fallback Python execution time: ${endTime - startTime}ms`);

      if (stderr) {
        this.logger.log(`🔍 Fallback Python stderr (full): ${stderr}`);
        if (!stderr.includes('Warning') && !stderr.includes('pip as the \'root\' user')) {
          this.logger.warn(`❗ Fallback Python stderr (non-warning): ${stderr}`);
        }
      }

      if (stdout) {
        this.logger.log(`📝 Fallback Python stdout (full): ${stdout}`);
      }

      // Check if output file exists
      this.logger.log(`🔍 Checking if output file exists: ${outputPath}`);
      const outputExists = await fs.access(outputPath).then(() => true).catch(() => false);
      if (!outputExists) {
        this.logger.error(`❌ Fallback output file was not created: ${outputPath}`);
        
        // List files in the directory to see what was created
        try {
          const outputDir = path.dirname(outputPath);
          const files = await fs.readdir(outputDir);
          this.logger.log(`📁 Files in output directory: ${files.join(', ')}`);
        } catch (dirError) {
          this.logger.error(`❌ Could not list output directory: ${dirError.message}`);
        }
        
        throw new Error('Fallback Python conversion did not produce output file');
      }

      const result = await fs.readFile(outputPath);
      this.logger.log(`📊 Fallback output file size: ${result.length} bytes`);
      
      if (result.length < 50) {
        this.logger.error(`❌ Fallback output file too small: ${result.length} bytes`);
        throw new Error(`Fallback Python output file too small: ${result.length} bytes`);
      }

      this.logger.log(`✅ Fallback Python conversion successful!`);
      return result;

    } catch (error) {
      this.logger.error(`❌ Fallback Python conversion error: ${error.message}`);
      this.logger.error(`🔍 Full error details: ${JSON.stringify(error, null, 2)}`);
      throw new Error(`Fallback Python conversion failed: ${error.message}`);
    } finally {
      try {
        await fs.unlink(scriptPath);
        this.logger.log(`🧹 Cleaned up fallback Python script: ${scriptPath}`);
      } catch (cleanupError) {
        this.logger.warn(`⚠️ Could not clean up fallback script: ${cleanupError.message}`);
      }
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
    """ULTIMATE PDF to XLSX conversion with AI-enhanced table detection - SINGLE SHEET VERSION"""
    print(f"🚀 Starting ULTIMATE PDF to XLSX conversion (Single Sheet Mode)...")
    
    # Method 1: ULTIMATE pdfplumber with intelligent data structuring (MOST RELIABLE)
    try:
        import pdfplumber
        import pandas as pd
        import re
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
        
        print("🧠 Using ULTIMATE pdfplumber (AI-Enhanced Single Sheet Mode)...")
        
        wb = Workbook()
        # Create a single sheet for all data
        ws = wb.active
        ws.title = "PDF_Data_Consolidated"
        
        # Professional styling
        header_style = NamedStyle(name="header")
        header_style.font = Font(bold=True, color="FFFFFF", size=12)
        header_style.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        header_style.alignment = Alignment(horizontal="center", vertical="center")
        header_style.border = Border(
            left=Side(border_style="medium", color="FFFFFF"),
            right=Side(border_style="medium", color="FFFFFF"),
            top=Side(border_style="medium", color="FFFFFF"),
            bottom=Side(border_style="medium", color="FFFFFF")
        )
        
        data_style = NamedStyle(name="data")
        data_style.border = Border(
            left=Side(border_style="thin"),
            right=Side(border_style="thin"),
            top=Side(border_style="thin"),
            bottom=Side(border_style="thin")
        )
        data_style.alignment = Alignment(wrap_text=True, vertical="center")
        
        page_separator_style = NamedStyle(name="page_separator")
        page_separator_style.font = Font(bold=True, color="FFFFFF", size=11)
        page_separator_style.fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")
        page_separator_style.alignment = Alignment(horizontal="center", vertical="center")
        
        current_row = 1
        max_cols_global = 0
        data_found = False
        
        with pdfplumber.open(input_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                print(f"🔍 Analyzing page {page_num + 1} with ULTIMATE table detection...")
                
                # Add page separator if not the first page
                if page_num > 0 and data_found:
                    # Add empty row for spacing
                    current_row += 1
                    # Add page separator row
                    ws.cell(row=current_row, column=1, value=f"--- PAGE {page_num + 1} ---")
                    ws.cell(row=current_row, column=1).style = page_separator_style
                    ws.merge_cells(f'A{current_row}:Z{current_row}')  # Merge across multiple columns
                    current_row += 1
                
                page_data_found = False
                
                # SUPER-ENHANCED table extraction with multiple detection methods
                tables_found = []
                
                # Method 1: Ultra-precise table detection
                try:
                    tables_precise = page.extract_tables(table_settings={
                        "vertical_strategy": "lines_strict",
                        "horizontal_strategy": "lines_strict",
                        "intersection_tolerance": 2,
                        "text_tolerance": 2,
                        "snap_tolerance": 2,
                        "join_tolerance": 2,
                        "edge_min_length": 3,
                        "min_words_vertical": 1,
                        "min_words_horizontal": 1
                    })
                    if tables_precise:
                        tables_found.extend(tables_precise)
                        print(f"📊 ULTIMATE precise detection found {len(tables_precise)} tables")
                except Exception as e:
                    print(f"⚠️ Precise method: {e}")
                
                # Method 2: Relaxed detection for complex layouts
                if not tables_found:
                    try:
                        tables_relaxed = page.extract_tables(table_settings={
                            "vertical_strategy": "text",
                            "horizontal_strategy": "text", 
                            "intersection_tolerance": 5,
                            "text_tolerance": 5,
                            "snap_tolerance": 5,
                            "join_tolerance": 5
                        })
                        if tables_relaxed:
                            tables_found.extend(tables_relaxed)
                            print(f"🎯 ULTIMATE smart detection found {len(tables_relaxed)} tables")
                    except Exception as e:
                        print(f"⚠️ Smart method: {e}")
                
                # Method 3: AI-like text pattern detection for data without clear tables
                if not tables_found:
                    try:
                        text = page.extract_text()
                        if text:
                            # ULTIMATE pattern-based table detection
                            detected_tables = detect_tabular_patterns(text)
                            if detected_tables:
                                tables_found.extend(detected_tables)
                                print(f"🤖 ULTIMATE AI pattern detection found {len(detected_tables)} data structures")
                    except Exception as e:
                        print(f"⚠️ AI pattern method: {e}")
                
                # Process all found tables and add to single sheet
                for table_idx, table in enumerate(tables_found):
                    if table and len(table) > 1:
                        page_data_found = True
                        data_found = True
                        
                        # Add table separator if multiple tables on same page
                        if table_idx > 0:
                            current_row += 1  # Add spacing
                        
                        # ULTIMATE table processing
                        processed_table = []
                        max_cols = 0
                        
                        for row_idx, row in enumerate(table):
                            processed_row = []
                            for cell in row:
                                if cell:
                                    # ULTIMATE cell cleaning
                                    clean_cell = re.sub(r'\\s+', ' ', str(cell)).strip()
                                    clean_cell = clean_cell.replace('\\n', ' ').replace('\\r', ' ')
                                    processed_row.append(clean_cell)
                                else:
                                    processed_row.append('')
                            processed_table.append(processed_row)
                            max_cols = max(max_cols, len(processed_row))
                        
                        # Update global max columns
                        max_cols_global = max(max_cols_global, max_cols)
                        
                        # Normalize all rows to have the same number of columns
                        for row in processed_table:
                            while len(row) < max_cols:
                                row.append('')
                        
                        # Add to single Excel sheet with ULTIMATE formatting
                        for row_idx, row in enumerate(processed_table):
                            for col_idx, value in enumerate(row, 1):
                                cell = ws.cell(row=current_row, column=col_idx, value=value)
                                
                                # ULTIMATE styling
                                if row_idx == 0 and table_idx == 0 and page_num == 0:  # Very first header
                                    cell.style = header_style
                                elif row_idx == 0:  # Other table headers
                                    cell.font = Font(bold=True, color="000000")
                                    cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
                                    cell.border = Border(
                                        left=Side(border_style="thin"),
                                        right=Side(border_style="thin"),
                                        top=Side(border_style="thin"),
                                        bottom=Side(border_style="thin")
                                    )
                                else:
                                    cell.style = data_style
                                
                                # ULTIMATE auto-sizing
                                column_letter = cell.column_letter
                                current_width = ws.column_dimensions[column_letter].width or 12
                                if value and len(str(value)) > current_width:
                                    ws.column_dimensions[column_letter].width = min(len(str(value)) + 3, 80)
                            
                            current_row += 1
                        
                        print(f"✅ ULTIMATE table processed: {len(processed_table)} rows × {max_cols} columns (Row {current_row - len(processed_table)} to {current_row - 1})")
                
                # If no tables found, extract structured text data
                if not page_data_found:
                    try:
                        text = page.extract_text()
                        if text and len(text.strip()) > 50:
                            structured_data = extract_structured_text_data(text, page_num)
                            if structured_data:
                                data_found = True
                                
                                # Add text data header if this is the first data
                                if current_row == 1:
                                    # Add header for text data
                                    header_row = ['Data Type', 'Content', 'Context']
                                    for col_idx, value in enumerate(header_row, 1):
                                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                                        cell.style = header_style
                                    current_row += 1
                                
                                # Add structured text data to single sheet
                                for row in structured_data[1:]:  # Skip the header from structured_data
                                    for col_idx, value in enumerate(row, 1):
                                        cell = ws.cell(row=current_row, column=col_idx, value=value)
                                        cell.style = data_style
                                        
                                        # Auto-sizing
                                        column_letter = cell.column_letter
                                        current_width = ws.column_dimensions[column_letter].width or 12
                                        if value and len(str(value)) > current_width:
                                            ws.column_dimensions[column_letter].width = min(len(str(value)) + 3, 80)
                                    current_row += 1
                                
                                print(f"📝 ULTIMATE text extraction: {len(structured_data) - 1} structured data rows added")
                    except Exception as e:
                        print(f"⚠️ Text extraction: {e}")
        
        # Apply final formatting to the single sheet
        if data_found:
            # Freeze first row if there's data
            ws.freeze_panes = 'A2'
            
            # Add auto-filter to the entire data range
            if current_row > 1 and max_cols_global > 0:
                ws.auto_filter.ref = f"A1:{ws.cell(row=current_row-1, column=max_cols_global).coordinate}"
            
            wb.save(output_path)
            print(f"✅ ULTIMATE pdfplumber conversion successful: Single consolidated sheet with {current_row-1} total rows")
            return True
            
    except ImportError:
        print("📦 Installing ULTIMATE pdfplumber dependencies...")
        if (install_package('pdfplumber') and install_package('pandas') and 
            install_package('openpyxl')):
            return premium_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"❌ ULTIMATE pdfplumber method failed: {e}")
    
    # Method 2: Advanced Tabula (If Java is available)
    try:
        import tabula
        import pandas as pd
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        print("🔍 Using Advanced Tabula (Premium Table Detection)...")
        
        # Try multiple Tabula strategies
        all_tables = []
        
        try:
            tables_lattice = tabula.read_pdf(input_path, pages='all', lattice=True, pandas_options={'header': None})
            if tables_lattice:
                print(f"📊 Tabula lattice detected {len(tables_lattice)} tables")
                all_tables.extend([(table, 'Lattice', i) for i, table in enumerate(tables_lattice)])
        except Exception as e:
            print(f"⚠️ Lattice method failed: {e}")
        
        try:
            tables_stream = tabula.read_pdf(input_path, pages='all', stream=True, guess=True, pandas_options={'header': None})
            if tables_stream:
                print(f"🌊 Tabula stream detected {len(tables_stream)} tables")
                all_tables.extend([(table, 'Stream', i) for i, table in enumerate(tables_stream)])
        except Exception as e:
            print(f"⚠️ Stream method failed: {e}")
        
        if all_tables:
            wb = Workbook()
            # Use single sheet for all Tabula data
            ws = wb.active
            ws.title = "PDF_Tabula_Data"
            
            current_row = 1
            
            for table_idx, (table, method, idx) in enumerate(all_tables):
                if not table.empty and table.shape[0] > 1:
                    # Add separator if not first table
                    if table_idx > 0:
                        current_row += 1  # Add spacing
                        ws.cell(row=current_row, column=1, value=f"--- {method} Table {idx+1} ---")
                        ws.cell(row=current_row, column=1).font = Font(bold=True)
                        current_row += 1
                    
                    # Clean and add data
                    table_clean = table.dropna(how='all').dropna(axis=1, how='all')
                    
                    for r_idx, row in enumerate(dataframe_to_rows(table_clean, index=False, header=False)):
                        for c_idx, value in enumerate(row, 1):
                            cell = ws.cell(row=current_row, column=c_idx, value=value)
                            if r_idx == 0:  # Header
                                cell.font = Font(bold=True, color="FFFFFF")
                                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        current_row += 1
            
            wb.save(output_path)
            print(f"✅ Advanced Tabula conversion successful: Single sheet with {len(all_tables)} tables consolidated")
            return True
            
    except ImportError:
        print("⚠️ Tabula not available (requires Java)")
    except Exception as e:
        print(f"❌ Tabula method failed: {e}")
    
    # Method 3: Enhanced pdfplumber fallback
    try:
        import pdfplumber
        import pandas as pd
        import re
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
        
        print("🧠 Using Enhanced pdfplumber fallback (Single Sheet Mode)...")
        
        wb = Workbook()
        # Use single sheet for fallback method too
        ws = wb.active
        ws.title = "PDF_Fallback_Data"
        
        current_row = 1
        data_found = False
        
        with pdfplumber.open(input_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Add page separator if not the first page
                if page_num > 0 and data_found:
                    current_row += 1
                    ws.cell(row=current_row, column=1, value=f"--- PAGE {page_num + 1} ---")
                    ws.cell(row=current_row, column=1).font = Font(bold=True, color="FF0000")
                    current_row += 1
                
                page_data_found = False
                
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
                            page_data_found = True
                            data_found = True
                            
                            # Add table separator if multiple tables
                            if table_idx > 0:
                                current_row += 1
                            
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
                            
                            # Add to single Excel sheet with formatting
                            for row in processed_table:
                                for col_idx, value in enumerate(row, 1):
                                    cell = ws.cell(row=current_row, column=col_idx, value=value)
                                    
                                    # Style first row as header
                                    if processed_table.index(row) == 0:
                                        cell.font = Font(bold=True, color="FFFFFF")
                                        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                                    
                                    cell.border = Border(
                                        left=Side(border_style="thin"),
                                        right=Side(border_style="thin"),
                                        top=Side(border_style="thin"),
                                        bottom=Side(border_style="thin")
                                    )
                                current_row += 1
                
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
        print(f"❌ Enhanced pdfplumber fallback failed: {e}")
    
    return False

def detect_tabular_patterns(text):
    """AI-like detection of tabular patterns in text"""
    detected_tables = []
    
    lines = [line.strip() for line in text.split('\\n') if line.strip()]
    current_table = []
    
    for line in lines:
        # Clean the line
        clean_line = re.sub(r'\\s+', ' ', line).strip()
        
        # Skip very short lines or section headers
        if len(clean_line) < 5 or re.match(r'^[A-Z\\s]+$', clean_line):
            # Save current table if it exists
            if len(current_table) >= 2:
                detected_tables.append(current_table)
            current_table = []
            continue
        
        # Look for patterns that suggest tabular data
        # Pattern 1: Multiple spaces or tabs separating values
        if re.search(r'\\s{3,}|\\t', clean_line):
            potential_columns = re.split(r'\\s{3,}|\\t', clean_line)
            if len(potential_columns) >= 2:
                current_table.append([col.strip() for col in potential_columns])
                continue
        
        # Pattern 2: Pipe or vertical bar separated
        if '|' in clean_line and clean_line.count('|') >= 2:
            potential_columns = clean_line.split('|')
            if len(potential_columns) >= 3:
                current_table.append([col.strip() for col in potential_columns if col.strip()])
                continue
        
        # Pattern 3: Comma separated (but be careful with sentences)
        if clean_line.count(',') >= 2 and len(clean_line) < 100:
            potential_columns = clean_line.split(',')
            if len(potential_columns) >= 3 and all(len(col.strip()) < 30 for col in potential_columns):
                current_table.append([col.strip() for col in potential_columns])
                continue
        
        # Pattern 4: Colon-separated key-value pairs
        if ':' in clean_line and clean_line.count(':') == 1:
            key, value = clean_line.split(':', 1)
            if len(key.strip()) < 50 and len(value.strip()) < 100:
                current_table.append([key.strip(), value.strip()])
                continue
        
        # Pattern 5: Numbers and text (financial data, etc.)
        number_pattern = r'\\b\\d+[,.]\\d+\\b|\\b\\d+\\b'
        if len(re.findall(number_pattern, clean_line)) >= 2:
            # This might be a data row
            parts = re.split(r'\\s{2,}', clean_line)
            if len(parts) >= 2:
                current_table.append(parts)
                continue
        
        # If we get here and we have a table, save it
        if len(current_table) >= 2:
            detected_tables.append(current_table)
            current_table = []
    
    # Don't forget the last table
    if len(current_table) >= 2:
        detected_tables.append(current_table)
    
    return detected_tables

def extract_structured_text_data(text, page_num):
    """Extract structured data from text for Excel"""
    lines = [line.strip() for line in text.split('\\n') if line.strip()]
    structured_data = [['Data Type', 'Content', 'Context']]  # Header
    
    current_section = f"Page {page_num + 1}"
    
    for line in lines:
        clean_line = line.strip()
        
        # Section headers (ALL CAPS or title case)
        if re.match(r'^[A-Z][A-Z\\s]{3,}$', clean_line) or clean_line.isupper():
            current_section = clean_line
            structured_data.append(['Section Header', clean_line, f'Page {page_num + 1}'])
        
        # Key-value pairs
        elif ':' in clean_line and clean_line.count(':') == 1:
            key, value = clean_line.split(':', 1)
            structured_data.append(['Key-Value', f'{key.strip()}: {value.strip()}', current_section])
        
        # Numbered items
        elif re.match(r'^\\d+[\\.\\)]', clean_line):
            structured_data.append(['Numbered Item', clean_line, current_section])
        
        # Bulleted items
        elif re.match(r'^[•▪▫○●◦‣⁃-]', clean_line):
            structured_data.append(['Bullet Point', clean_line, current_section])
        
        # Dates
        elif re.search(r'\\b\\d{1,2}[/-]\\d{1,2}[/-]\\d{2,4}\\b|\\b\\d{4}[/-]\\d{1,2}[/-]\\d{1,2}\\b', clean_line):
            structured_data.append(['Date Info', clean_line, current_section])
        
        # Numbers/Financial data
        elif re.search(r'\\$[\\d,]+\\.?\\d*|\\b\\d+[,.]\\d+\\b', clean_line):
            structured_data.append(['Numeric Data', clean_line, current_section])
        
        # Regular text (but only if substantial)
        elif len(clean_line) > 20:
            structured_data.append(['Text Content', clean_line[:100] + ('...' if len(clean_line) > 100 else ''), current_section])
    
    return structured_data if len(structured_data) > 1 else None
    
    return False

def premium_convert_to_pptx(input_path, output_path):
    """Enhanced PDF to PPTX conversion with editable text and structured content"""
    print(f"🚀 Starting ENHANCED PDF to PPTX conversion with editable text...")
    
    # Method 1: Enhanced conversion with editable text extraction
    try:
        import fitz  # PyMuPDF
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.enum.text import PP_ALIGN
        from pptx.dml.color import RGBColor
        import io
        import re
        
        print("📋 Using enhanced PyMuPDF + python-pptx with text extraction...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        # Use standard 16:9 slide size
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(5.625)
        
        # Create title slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "PDF Presentation"
        if len(title_slide.placeholders) > 1:
            title_slide.placeholders[1].text = f"Converted from PDF • {len(pdf_doc)} pages • Editable Content"
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            print(f"📄 Processing page {page_num + 1}/{len(pdf_doc)} with text extraction")
            
            # Extract text with structure information
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            
            # Organize content by structure
            slide_content = {
                'title': '',
                'headings': [],
                'paragraphs': [],
                'bullet_points': [],
                'tables': [],
                'has_images': len(page.get_images()) > 0
            }
            
            # Process text blocks to extract structured content
            for block in blocks:
                if "lines" in block:
                    block_text = ""
                    block_font_sizes = []
                    is_bold = False
                    
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            text = span.get("text", "").strip()
                            if text:
                                line_text += text + " "
                                font_size = span.get("size", 12)
                                block_font_sizes.append(font_size)
                                
                                # Check for bold formatting
                                flags = span.get("flags", 0)
                                if flags & 2**4:  # Bold flag
                                    is_bold = True
                        
                        if line_text.strip():
                            block_text += line_text.strip() + "\\n"
                    
                    if block_text.strip():
                        avg_font_size = sum(block_font_sizes) / len(block_font_sizes) if block_font_sizes else 12
                        clean_text = block_text.strip()
                        
                        # Categorize text based on formatting and content
                        if avg_font_size >= 18 or is_bold:
                            if not slide_content['title'] and len(clean_text) < 100:
                                slide_content['title'] = clean_text
                            else:
                                slide_content['headings'].append(clean_text)
                        elif clean_text.startswith(('•', '-', '*', '○', '▪', '▫', '1.', '2.', '3.', 'a.', 'b.', 'c.')):
                            # Process bullet points
                            bullet_lines = [line.strip() for line in clean_text.split('\\n') if line.strip()]
                            for bullet_line in bullet_lines:
                                if bullet_line.startswith(('•', '-', '*', '○', '▪', '▫')) or re.match(r'^[0-9a-zA-Z]\\.', bullet_line):
                                    slide_content['bullet_points'].append(bullet_line)
                                else:
                                    slide_content['paragraphs'].append(bullet_line)
                        else:
                            # Regular paragraph text
                            paragraphs = [p.strip() for p in clean_text.split('\\n\\n') if p.strip()]
                            slide_content['paragraphs'].extend(paragraphs)
            
            # Create slide layout based on content structure
            if slide_content['bullet_points']:
                # Use bullet point layout
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                print(f"📝 Creating bullet point slide for page {page_num + 1}")
            elif slide_content['headings'] and slide_content['paragraphs']:
                # Use title and content layout
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                print(f"📋 Creating structured content slide for page {page_num + 1}")
            elif slide_content['has_images'] and (slide_content['paragraphs'] or slide_content['headings']):
                # Use content with caption layout
                slide_layout = prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[1]
                slide = prs.slides.add_slide(slide_layout)
                print(f"🖼️ Creating image with text slide for page {page_num + 1}")
            else:
                # Default title and content
                slide_layout = prs.slide_layouts[1]  # Title and Content
                slide = prs.slides.add_slide(slide_layout)
                print(f"📄 Creating standard slide for page {page_num + 1}")
            
            # Set slide title
            slide_title = slide_content['title'] or f"Page {page_num + 1}"
            if slide_content['headings'] and not slide_content['title']:
                slide_title = slide_content['headings'][0][:60]
            
            slide.shapes.title.text = slide_title
            
            # Add content to slide
            content_added = False
            
            # Add bullet points if available
            if slide_content['bullet_points'] and len(slide.placeholders) > 1:
                content_placeholder = slide.placeholders[1]
                text_frame = content_placeholder.text_frame
                text_frame.clear()
                
                for i, bullet in enumerate(slide_content['bullet_points'][:8]):  # Limit to 8 bullets
                    # Clean bullet point text
                    clean_bullet = re.sub(r'^[•\\-\\*○▪▫0-9a-zA-Z]\\.?\\s*', '', bullet).strip()
                    if clean_bullet:
                        para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                        para.text = clean_bullet
                        para.level = 0
                        para.font.size = Pt(16)
                        para.space_after = Pt(6)
                
                content_added = True
                print(f"✅ Added {len(slide_content['bullet_points'])} bullet points to slide {page_num + 1}")
            
            # Add paragraphs if no bullets or additional content
            elif slide_content['paragraphs'] and len(slide.placeholders) > 1 and not content_added:
                content_placeholder = slide.placeholders[1]
                text_frame = content_placeholder.text_frame
                text_frame.clear()
                
                # Combine paragraphs intelligently
                combined_text = ""
                for i, para in enumerate(slide_content['paragraphs'][:4]):  # Limit to 4 paragraphs
                    if len(combined_text) + len(para) < 800:  # Prevent text overflow
                        combined_text += para + "\\n\\n"
                    else:
                        break
                
                if combined_text.strip():
                    paragraphs = [p.strip() for p in combined_text.strip().split('\\n\\n') if p.strip()]
                    for i, para_text in enumerate(paragraphs):
                        para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                        para.text = para_text
                        para.font.size = Pt(14)
                        para.space_after = Pt(8)
                        para.alignment = PP_ALIGN.LEFT
                
                content_added = True
                print(f"✅ Added {len(paragraphs)} paragraphs to slide {page_num + 1}")
            
            # Add headings as additional content
            if slide_content['headings'] and slide_content['title']:
                try:
                    # Add headings as formatted text
                    if len(slide.placeholders) > 1 and not content_added:
                        content_placeholder = slide.placeholders[1]
                        text_frame = content_placeholder.text_frame
                        text_frame.clear()
                        
                        for i, heading in enumerate(slide_content['headings'][:3]):  # Limit to 3 headings
                            para = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                            para.text = heading
                            para.font.size = Pt(18)
                            para.font.bold = True
                            para.space_after = Pt(10)
                            para.font.color.rgb = RGBColor(31, 73, 125)
                        
                        content_added = True
                        print(f"✅ Added {len(slide_content['headings'])} headings to slide {page_num + 1}")
                except Exception as heading_error:
                    print(f"⚠️ Could not add headings: {heading_error}")
            
            # If we have images and space, add them alongside text
            if slide_content['has_images'] and content_added:
                try:
                    image_list = page.get_images()
                    if image_list:
                        # Add first image as a smaller element
                        img = image_list[0]
                        xref = img[0]
                        pix = fitz.Pixmap(pdf_doc, xref)
                        
                        if pix.n - pix.alpha < 4:  # Valid image
                            img_data = pix.tobytes("png")
                            image_stream = io.BytesIO(img_data)
                            
                            # Position image on the right side
                            left = Inches(6.5)
                            top = Inches(1.5)
                            max_width = Inches(3)
                            max_height = Inches(3.5)
                            
                            slide.shapes.add_picture(image_stream, left, top, max_width, max_height)
                            print(f"✅ Added image to slide {page_num + 1}")
                        
                        pix = None
                except Exception as img_error:
                    print(f"⚠️ Could not add image to slide {page_num + 1}: {img_error}")
            
            # Fallback: if no editable content was added, add page as image (last resort)
            if not content_added:
                print(f"📄 No structured text found, adding page {page_num + 1} as image")
                try:
                    mat = fitz.Matrix(2.0, 2.0)  # 2x scaling for good quality
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    img_data = pix.tobytes("png")
                    image_stream = io.BytesIO(img_data)
                    
                    # Use blank layout for full page image
                    image_slide = prs.slides.add_slide(prs.slide_layouts[6])
                    
                    # Center the image
                    left = Inches(1)
                    top = Inches(0.5)
                    width = Inches(8)
                    height = Inches(4.5)
                    
                    image_slide.shapes.add_picture(image_stream, left, top, width, height)
                    print(f"⚠️ Added page {page_num + 1} as fallback image")
                    
                except Exception as fallback_error:
                    print(f"❌ Could not add fallback image for page {page_num + 1}: {fallback_error}")
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 5000:
            print(f"✅ ENHANCED PPTX conversion successful: {os.path.getsize(output_path)} bytes")
            print(f"🎯 Created presentation with editable text content")
            return True
            
    except ImportError as e:
        print(f"📦 Missing packages for enhanced PPTX conversion: {e}")
        print("📦 Installing PyMuPDF and python-pptx...")
        if install_package('PyMuPDF') and install_package('python-pptx'):
            return premium_convert_to_pptx(input_path, output_path)
    except Exception as e:
        print(f"❌ Enhanced PPTX conversion failed: {e}")
        import traceback
        traceback.print_exc()  # Print full stack trace for debugging
    
    # Method 2: Simple text extraction fallback
    try:
        import fitz
        from pptx import Presentation
        from pptx.util import Inches
        import io
        
        print("📋 Using Text-Only Fallback Method...")
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Create simple slide
            slide_layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            slide.shapes.title.text = f"Page {page_num + 1}"
            
            # Extract text and add as editable content
            text = page.get_text()
            if text.strip() and len(slide.placeholders) > 1:
                content_box = slide.placeholders[1]
                # Clean and format text
                lines = [line.strip() for line in text.split('\\n') if line.strip()]
                formatted_text = '\\n'.join(lines[:15])  # Limit to 15 lines
                content_box.text = formatted_text[:1000]  # Limit text length
                print(f"✅ Added editable text to page {page_num + 1}")
            else:
                # If no text, add a note
                if len(slide.placeholders) > 1:
                    content_box = slide.placeholders[1]
                    content_box.text = "No readable text found on this page"
        
        prs.save(output_path)
        pdf_doc.close()
        
        if os.path.exists(output_path) and os.path.getsize(output_path) > 3000:
            print(f"✅ Text-only PPTX fallback successful: {os.path.getsize(output_path)} bytes")
            return True
            
    except Exception as e:
        print(f"❌ Text-only PPTX fallback failed: {e}")
    
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
        from openpyxl import Workbook
        
        wb = Workbook()
        ws = wb.active
        ws.title = "PDF Content"
        
        pdf_doc = fitz.open(input_path)
        row = 1
        
        for page_num, page in enumerate(pdf_doc):
            # Add page header
            ws.cell(row=row, column=1, value=f"Page {page_num + 1}")
            row += 1
            
            # Extract text and add to cells
            text = page.get_text()
            lines = text.split('\\n')
            
            for line in lines:
                if line.strip():
                    ws.cell(row=row, column=1, value=line.strip())
                    row += 1
            
            row += 1  # Add spacing between pages
        
        wb.save(output_path)
        pdf_doc.close()
        return True
        
    except ImportError:
        if install_package('PyMuPDF') and install_package('openpyxl'):
            return basic_convert_to_xlsx(input_path, output_path)
    except Exception as e:
        print(f"Basic XLSX conversion failed: {e}")
    return False

def basic_convert_to_pptx(input_path, output_path):
    """Basic PDF to PPTX conversion"""
    try:
        import fitz
        from pptx import Presentation
        from pptx.util import Inches
        import io
        
        pdf_doc = fitz.open(input_path)
        prs = Presentation()
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc.load_page(page_num)
            
            # Create slide
            slide_layout = prs.slide_layouts[1]  # Title and Content
            slide = prs.slides.add_slide(slide_layout)
            
            # Set title
            title = slide.shapes.title
            title.text = f"Page {page_num + 1}"
            
            # Try to get page as image
            try:
                mat = fitz.Matrix(2.0, 2.0)  # 2x scaling for better quality
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img_data = pix.tobytes("png")
                image_stream = io.BytesIO(img_data)
                
                # Add image to slide
                left = Inches(1)
                top = Inches(1.5)
                width = Inches(8)
                height = Inches(5.5)
                
                slide.shapes.add_picture(image_stream, left, top, width, height)
                
            except Exception as img_error:
                # Fallback to text if image fails
                text = page.get_text()
                if text.strip():
                    content = slide.placeholders[1]
                    content.text = text[:500]  # Limit text length
        
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
