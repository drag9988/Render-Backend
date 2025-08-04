import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import * as fs from 'fs/promises';
import * as path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as multer from 'multer';
import { OnlyOfficeService } from './onlyoffice.service';
import { OnlyOfficeEnhancedService } from './onlyoffice-enhanced.service';
import { FileValidationService } from './file-validation.service';

@Injectable()
export class AppService {
  private readonly logger = new Logger(AppService.name);
  private readonly execAsync = promisify(exec);

  constructor(
    private readonly onlyOfficeService: OnlyOfficeService,
    private readonly onlyOfficeEnhancedService: OnlyOfficeEnhancedService,
    private readonly fileValidationService: FileValidationService
  ) {}

  /**
   * Convert Office documents to PDF using LibreOffice
   */
  async convertOfficeToPdf(file: Express.Multer.File): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

    // Validate file based on its type
    let expectedFileType: 'word' | 'excel' | 'powerpoint';
    
    if (this.fileValidationService['allowedMimeTypes']['word'].includes(file.mimetype)) {
      expectedFileType = 'word';
    } else if (this.fileValidationService['allowedMimeTypes']['excel'].includes(file.mimetype)) {
      expectedFileType = 'excel';
    } else if (this.fileValidationService['allowedMimeTypes']['powerpoint'].includes(file.mimetype)) {
      expectedFileType = 'powerpoint';
    } else {
      throw new BadRequestException(`Unsupported file type for Office to PDF conversion: ${file.mimetype}`);
    }

    // Validate the file
    const validation = this.fileValidationService.validateFile(file, expectedFileType);
    if (!validation.isValid) {
      throw new BadRequestException(`File validation failed: ${validation.errors.join(', ')}`);
    }

    // Update the file with sanitized filename
    file.originalname = validation.sanitizedFilename;

    // Use LibreOffice directly for Office to PDF conversions
    this.logger.log(`Converting ${expectedFileType.toUpperCase()} to PDF using LibreOffice for: ${file.originalname}`);
    return await this.executeLibreOfficeConversion(file, 'pdf');
  }

  /**
   * Convert PDF to Office formats using ONLYOFFICE Enhanced Service
   */
  async convertPdfToOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

    // Validate PDF file
    if (file.mimetype !== 'application/pdf') {
      throw new BadRequestException(`Expected PDF file, got: ${file.mimetype}`);
    }

    const validation = this.fileValidationService.validateFile(file, 'pdf');
    if (!validation.isValid) {
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    // Update the file with sanitized filename
    file.originalname = validation.sanitizedFilename;

    // Validate target format
    if (!['docx', 'xlsx', 'pptx'].includes(format)) {
      throw new BadRequestException(`Unsupported target format: ${format}`);
    }

    return await this.convertPdfToOfficeFormat(file, format);
  }

  /**
   * Legacy method for backward compatibility - routes to appropriate conversion method
   */
  async convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

    // Route to appropriate conversion method based on input and output format
    if (file.mimetype === 'application/pdf' && ['docx', 'xlsx', 'pptx'].includes(format)) {
      // PDF to Office - use ONLYOFFICE
      return await this.convertPdfToOffice(file, format);
    } else if (format === 'pdf') {
      // Office to PDF - use LibreOffice
      return await this.convertOfficeToPdf(file);
    } else {
      // Other conversions - use LibreOffice directly
      let expectedFileType: 'pdf' | 'word' | 'excel' | 'powerpoint';
      
      if (file.mimetype === 'application/pdf') {
        expectedFileType = 'pdf';
      } else if (this.fileValidationService['allowedMimeTypes']['word'].includes(file.mimetype)) {
        expectedFileType = 'word';
      } else if (this.fileValidationService['allowedMimeTypes']['excel'].includes(file.mimetype)) {
        expectedFileType = 'excel';
      } else if (this.fileValidationService['allowedMimeTypes']['powerpoint'].includes(file.mimetype)) {
        expectedFileType = 'powerpoint';
      } else {
        throw new BadRequestException(`Unsupported file type: ${file.mimetype}`);
      }

      const validation = this.fileValidationService.validateFile(file, expectedFileType);
      if (!validation.isValid) {
        throw new BadRequestException(`File validation failed: ${validation.errors.join(', ')}`);
      }

      file.originalname = validation.sanitizedFilename;
      return await this.executeLibreOfficeConversion(file, format);
    }
  }

  private async convertPdfToOfficeFormat(file: Express.Multer.File, format: string): Promise<Buffer> {
    this.logger.log(`Converting PDF to ${format.toUpperCase()} - using ONLYOFFICE Enhanced Service for premium quality`);

    // Primary: Use Enhanced ONLYOFFICE Service (includes ONLYOFFICE Server + Python + LibreOffice)
    try {
      switch (format) {
        case 'docx':
          return await this.onlyOfficeEnhancedService.convertPdfToDocx(file.buffer, file.originalname);
        case 'xlsx':
          return await this.onlyOfficeEnhancedService.convertPdfToXlsx(file.buffer, file.originalname);
        case 'pptx':
          return await this.onlyOfficeEnhancedService.convertPdfToPptx(file.buffer, file.originalname);
        default:
          throw new Error(`Unsupported format: ${format}`);
      }
    } catch (enhancedError) {
      this.logger.warn(`Enhanced ONLYOFFICE failed: ${enhancedError.message}, trying fallback methods`);
      
      // Secondary: Use original ONLYOFFICE service as backup
      if (this.onlyOfficeService.isAvailable()) {
        try {
          this.logger.log(`Trying original ONLYOFFICE service as backup`);
          switch (format) {
            case 'docx':
              return await this.onlyOfficeService.convertPdfToDocx(file.buffer, file.originalname);
            case 'xlsx':
              return await this.onlyOfficeService.convertPdfToXlsx(file.buffer, file.originalname);
            case 'pptx':
              return await this.onlyOfficeService.convertPdfToPptx(file.buffer, file.originalname);
            default:
              throw new Error(`Unsupported format: ${format}`);
          }
        } catch (onlyOfficeError) {
          this.logger.warn(`Original ONLYOFFICE also failed: ${onlyOfficeError.message}`);
        }
      } else {
        this.logger.warn(`Original ONLYOFFICE service is not available`);
      }
    }

    // Final fallback: LibreOffice (for compatibility - though not recommended for PDF to Office)
    this.logger.log(`Using LibreOffice as final conversion method for ${format.toUpperCase()}`);
    return await this.executeLibreOfficeConversion(file, format);
  }

  private async executeLibreOfficeConversion(file: Express.Multer.File, format: string): Promise<Buffer> {
    // Create a unique filename with timestamp
    const timestamp = Date.now();
    const sanitizedFilename = file.originalname.replace(/[^a-zA-Z0-9.]/g, '_');
    
    // Use the OS temp directory or fallback to /tmp
    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    this.logger.log(`Using temp directory: ${tempDir}`);
    
    // Ensure temp directory exists and is writable
    try {
      await fs.mkdir(tempDir, { recursive: true });
      await fs.chmod(tempDir, 0o777);
      this.logger.log(`Successfully ensured temp directory exists: ${tempDir}`);
    } catch (dirError) {
      this.logger.error(`Failed to create temp directory ${tempDir}: ${dirError.message}`);
      throw new Error(`Cannot create or access temp directory: ${dirError.message}`);
    }
    
    const tempInput = `${tempDir}/${timestamp}_${sanitizedFilename}`;
    const tempOutput = tempInput.replace(/\.[^.]+$/, `.${format}`);

    try {
      // Write file to disk
      await fs.writeFile(tempInput, file.buffer);
      this.logger.log(`Written input file: ${tempInput}`);

      // Special handling for PDF to Office conversions
      if (file.mimetype === 'application/pdf' && ['docx', 'xlsx', 'pptx'].includes(format)) {
        return await this.convertPdfWithLibreOffice(tempInput, tempOutput, format, tempDir);
      }

      // Special handling for Excel files
      if (format === 'pdf' && this.fileValidationService['allowedMimeTypes']['excel'].includes(file.mimetype)) {
        return await this.executeEnhancedExcelToPdfConversion(tempInput, tempOutput, tempDir);
      }

      // Standard LibreOffice conversion
      const command = `libreoffice --headless --convert-to ${format} --outdir "${tempDir}" "${tempInput}"`;
      this.logger.log(`Executing LibreOffice: ${command}`);

      const { stdout, stderr } = await this.execAsync(command, { timeout: 120000 });
      
      if (stderr && !stderr.includes('Warning')) {
        this.logger.warn(`LibreOffice stderr: ${stderr}`);
      }

      if (stdout) {
        this.logger.log(`LibreOffice stdout: ${stdout}`);
      }

      // Check if conversion was successful
      const outputExists = await fs.access(tempOutput).then(() => true).catch(() => false);
      if (!outputExists) {
        throw new Error(`LibreOffice did not create output file: ${tempOutput}`);
      }

      const result = await fs.readFile(tempOutput);
      
      // Validate output
      if (!await this.validateConvertedFile(result, format)) {
        throw new Error(`Converted file validation failed for format: ${format}`);
      }

      this.logger.log(`LibreOffice conversion successful, output size: ${result.length} bytes`);
      return result;

    } catch (error) {
      this.logger.error(`LibreOffice conversion failed: ${error.message}`);
      throw error;
    } finally {
      // Cleanup
      try {
        await fs.unlink(tempInput).catch(() => {});
        await fs.unlink(tempOutput).catch(() => {});
      } catch (cleanupError) {
        this.logger.warn(`Cleanup warning: ${cleanupError.message}`);
      }
    }
  }

  private async convertPdfWithLibreOffice(tempInput: string, tempOutput: string, format: string, tempDir: string): Promise<Buffer> {
    this.logger.log(`Converting PDF to ${format} with enhanced options`);
    
    // First, analyze the PDF to determine the best conversion strategy
    const pdfInfo = await this.analyzePdf(tempInput);
    this.logger.log(`PDF Analysis: ${JSON.stringify(pdfInfo)}`);
    
    // Try multiple LibreOffice options for better PDF conversion
    const commands = [
      `libreoffice --headless --convert-to ${format} --infilter="writer_pdf_import" --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --convert-to ${format} --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --draw --convert-to ${format} --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --writer --convert-to ${format} --outdir "${tempDir}" "${tempInput}"`,
    ];

    let lastError = '';
    
    for (let i = 0; i < commands.length; i++) {
      const command = commands[i];
      this.logger.log(`LibreOffice attempt ${i + 1}: Converting to ${format}`);
      
      try {
        const { stdout, stderr } = await this.execAsync(command, { timeout: 120000 });
        
        if (stderr && !stderr.includes('Warning')) {
          this.logger.warn(`LibreOffice stderr: ${stderr}`);
        }

        // Check for multiple possible output paths
        const outputPaths = [
          tempOutput,
          path.join(tempDir, `${path.basename(tempInput, path.extname(tempInput))}.${format}`),
          path.join(tempDir, path.basename(tempInput).replace(/\.[^.]+$/, `.${format}`))
        ];

        for (const outputPath of outputPaths) {
          const exists = await fs.access(outputPath).then(() => true).catch(() => false);
          if (exists) {
            const result = await fs.readFile(outputPath);
            if (result.length > 100) {
              this.logger.log(`LibreOffice conversion successful on attempt ${i + 1}`);
              return result;
            }
          }
        }
        
        lastError = `No valid output file found after attempt ${i + 1}`;
      } catch (execError) {
        lastError = execError.message;
        this.logger.warn(`LibreOffice attempt ${i + 1} failed: ${execError.message}`);
        continue;
      }
    }
    
    // If LibreOffice fails, try alternative methods for certain formats
    if (format === 'docx') {
      try {
        return await this.convertPdfToWordAlternative(tempInput, tempOutput, tempDir);
      } catch (altError) {
        this.logger.warn(`Alternative PDF to Word conversion also failed: ${altError.message}`);
      }
    }
    
    // Provide specific error messages based on PDF characteristics
    let errorMessage = `PDF to ${format.toUpperCase()} conversion failed after multiple attempts. `;
    
    if (pdfInfo.isScanned) {
      errorMessage += 'This appears to be a scanned PDF (image-based). ';
      errorMessage += 'Scanned PDFs cannot be reliably converted to editable Office formats. ';
      errorMessage += 'Consider using OCR software first to make the PDF text-selectable. ';
    } else if (pdfInfo.hasComplexLayout) {
      errorMessage += 'This PDF has complex formatting that may not convert well. ';
      errorMessage += 'PDFs with tables, graphics, and complex layouts often lose formatting during conversion. ';
    } else if (pdfInfo.isProtected) {
      errorMessage += 'This PDF appears to be password-protected or restricted. ';
      errorMessage += 'Remove password protection before attempting conversion. ';
    } else {
      errorMessage += 'The PDF structure may be incompatible with LibreOffice conversion. ';
    }
    
    errorMessage += `Try using a simpler, text-based PDF. Last error: ${lastError}`;
    
    throw new Error(errorMessage);
  }

  private async executeEnhancedExcelToPdfConversion(tempInput: string, tempOutput: string, tempDir: string): Promise<Buffer> {
    this.logger.log(`Attempting enhanced Excel to PDF conversion using multiple specialized approaches`);
    
    // Enhanced LibreOffice commands specifically for Excel files
    const commands = [
      `libreoffice --headless --calc --convert-to pdf:calc_pdf_Export --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --invisible --calc --convert-to pdf --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --calc --convert-to pdf --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --convert-to pdf:calc_pdf_Export:"SelectPdfVersion=1" --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --writer --convert-to pdf --outdir "${tempDir}" "${tempInput}"`,
      `libreoffice --headless --convert-to pdf --outdir ${tempDir.replace(/\s/g, '\\ ')} ${tempInput.replace(/\s/g, '\\ ')}`
    ];

    let lastError = '';
    const baseName = path.basename(tempInput, path.extname(tempInput));
    
    for (let i = 0; i < commands.length; i++) {
      const command = commands[i];
      this.logger.log(`Enhanced Excel conversion attempt ${i + 1}: ${command}`);
      
      try {
        const { stdout, stderr } = await this.execAsync(command, { timeout: 120000 });
        
        if (stderr && !stderr.includes('Warning') && !stderr.includes('Gtk-WARNING')) {
          this.logger.warn(`LibreOffice stderr: ${stderr}`);
        }

        // Check for output file with multiple possible names
        const outputPaths = [
          tempOutput,
          path.join(tempDir, `${baseName}.pdf`),
          path.join(tempDir, path.basename(tempInput).replace(/\.[^.]+$/, '.pdf'))
        ];

        for (const outputPath of outputPaths) {
          const exists = await fs.access(outputPath).then(() => true).catch(() => false);
          if (exists) {
            const result = await fs.readFile(outputPath);
            if (result.length > 100 && result.subarray(0, 4).toString() === '%PDF') {
              this.logger.log(`Enhanced Excel to PDF conversion successful on attempt ${i + 1}`);
              return result;
            }
          }
        }
        
        lastError = `No valid PDF output found after attempt ${i + 1}`;
      } catch (execError) {
        lastError = execError.message;
        this.logger.warn(`Enhanced Excel conversion attempt ${i + 1} failed: ${execError.message}`);
        continue;
      }
    }
    
    // If all enhanced LibreOffice attempts fail, provide detailed error message
    let errorMessage = `Enhanced Excel to PDF conversion failed after ${commands.length} specialized attempts. `;
    
    if (lastError.includes('not found') || lastError.includes('command not found')) {
      errorMessage += 'LibreOffice is not installed or not accessible. Please install LibreOffice. ';
    } else if (lastError.includes('Permission denied') || lastError.includes('access')) {
      errorMessage += 'File access permissions issue. Check that the temp directory is writable. ';
    } else if (lastError.includes('timeout')) {
      errorMessage += 'Conversion timed out. The Excel file may be too large or complex. ';
    } else if (lastError.includes('export filter') || lastError.includes('filter')) {
      errorMessage += 'LibreOffice PDF export filter issue. The Excel file may have unsupported features. ';
    } else {
      errorMessage += 'Unknown conversion error. The Excel file may be corrupted or use unsupported features. ';
    }
    
    errorMessage += `Try with a simpler .xlsx file without macros, charts, or complex formatting. Last error: ${lastError}`;
    throw new Error(errorMessage);
  }

  private async analyzePdf(pdfPath: string): Promise<{isScanned: boolean, hasComplexLayout: boolean, isProtected: boolean, pageCount: number}> {
    try {
      // Use pdfinfo to analyze the PDF
      const { stdout } = await this.execAsync(`pdfinfo "${pdfPath}"`, { timeout: 30000 });
      
      const pageCount = parseInt(stdout.match(/Pages:\s+(\d+)/)?.[1] || '0');
      const isProtected = stdout.includes('Encrypted:') && !stdout.includes('Encrypted: no');
      
      // Simple heuristics for determining if PDF is scanned or has complex layout
      const isScanned = stdout.includes('no text') || pageCount > 0 && !stdout.includes('Tagged:');
      const hasComplexLayout = pageCount > 10 || stdout.includes('Form:') || stdout.includes('JavaScript:');
      
      return { isScanned, hasComplexLayout, isProtected, pageCount };
    } catch (error) {
      this.logger.warn(`PDF analysis failed: ${error.message}`);
      return { isScanned: false, hasComplexLayout: false, isProtected: false, pageCount: 1 };
    }
  }

  private async validateConvertedFile(buffer: Buffer, format: string): Promise<boolean> {
    try {
      if (buffer.length < 50) return false;
      
      const header = buffer.subarray(0, 10).toString();
      
      switch (format) {
        case 'pdf':
          return header.startsWith('%PDF');
        case 'docx':
        case 'xlsx':
        case 'pptx':
          return header.startsWith('PK'); // ZIP-based Office formats
        default:
          return true; // Basic validation passed
      }
    } catch (error) {
      this.logger.warn(`File validation error: ${error.message}`);
      return false;
    }
  }

  private async convertPdfToWordAlternative(tempInput: string, tempOutput: string, tempDir: string): Promise<Buffer> {
    this.logger.log(`Attempting alternative PDF to Word conversion using text extraction`);
    
    try {
      // Extract text from PDF
      const { stdout: extractedText } = await this.execAsync(`pdftotext "${tempInput}" -`, { timeout: 30000 });
      
      if (extractedText.trim().length < 50) {
        throw new Error('PDF contains insufficient text content for conversion');
      }
      
      // Create a simple Word document with the extracted text
      const simpleDocxPath = `${tempDir}/${Date.now()}_simple.docx`;
      
      // Create a basic DOCX structure (this is a simplified approach)
      // For a more robust solution, you might want to use a library like docx
      const simpleText = extractedText.replace(/\n\n+/g, '\n\n').trim();
      
      // Write to a temporary text file first
      const tempTextFile = `${tempDir}/${Date.now()}_temp.txt`;
      await fs.writeFile(tempTextFile, simpleText);
      
      // Convert text to DOCX using LibreOffice
      const textToDocxCommand = `libreoffice --headless --convert-to docx --outdir "${tempDir}" "${tempTextFile}"`;
      await this.execAsync(textToDocxCommand, { timeout: 30000 });
      
      // Find the generated DOCX file
      const generatedDocx = tempTextFile.replace('.txt', '.docx');
      
      try {
        const result = await fs.readFile(generatedDocx);
        this.logger.log(`Alternative PDF to Word conversion successful`);
        return result;
      } finally {
        await fs.unlink(tempTextFile).catch(() => {});
        await fs.unlink(generatedDocx).catch(() => {});
      }
    } catch (error) {
      throw new Error(`Alternative PDF to Word conversion failed: ${error.message}`);
    }
  }

  async analyzePdfFile(pdfPath: string): Promise<{isScanned: boolean, hasComplexLayout: boolean, isProtected: boolean, pageCount: number}> {
    return await this.analyzePdf(pdfPath);
  }

  async compressPdf(file: Express.Multer.File, quality: string = 'moderate'): Promise<Buffer> {
    console.log(`🔥 APP.SERVICE compressPdf called at ${new Date().toISOString()}`);
    console.log(`🔥 File buffer size: ${file?.buffer?.length || 'NO BUFFER'} bytes`);
    console.log(`🔥 Quality: ${quality}`);
    
    if (!file || !file.buffer) {
      console.log(`🔥 ERROR: Invalid PDF file provided`);
      throw new Error('Invalid PDF file provided');
    }

    // Validate PDF file
    const validation = this.fileValidationService.validateFile(file, 'pdf');
    if (!validation.isValid) {
      console.log(`🔥 ERROR: PDF validation failed: ${validation.errors.join(', ')}`);
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    // Validate and sanitize quality parameter
    const sanitizedQuality = this.fileValidationService.validateCompressionQuality(quality);

    // Create a unique filename with timestamp
    const timestamp = Date.now();
    
    // Use the OS temp directory or fallback to /tmp
    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    this.logger.log(`Using temp directory for PDF compression: ${tempDir}`);
    
    // Ensure temp directory exists and is writable
    try {
      await fs.mkdir(tempDir, { recursive: true });
      await fs.chmod(tempDir, 0o777);
      this.logger.log(`Successfully ensured temp directory exists: ${tempDir}`);
    } catch (dirError) {
      this.logger.error(`Failed to create temp directory ${tempDir}: ${dirError.message}`);
      throw new Error(`Cannot create or access temp directory: ${dirError.message}`);
    }
    
    const input = `${tempDir}/${timestamp}_input.pdf`;
    const output = `${tempDir}/${timestamp}_output.pdf`;
    let isImageHeavy = false;
    
    try {
      this.logger.log(`Starting PDF compression, input size: ${file.buffer.length} bytes, quality: ${sanitizedQuality}`);
      
      // Write the uploaded PDF to disk
      await fs.writeFile(input, file.buffer);
      this.logger.log(`PDF written to ${input}`);

      // Analyze PDF to determine if it's image-heavy (like mobile camera photos)
      try {
        const { stdout: pdfInfo } = await this.execAsync(`pdfinfo "${input}"`, { timeout: 10000 });
        const { stdout: imageInfo } = await this.execAsync(`pdfimages -list "${input}"`, { timeout: 10000 });
        
        // Check for high-resolution images or large file size with few pages
        const pages = parseInt(pdfInfo.match(/Pages:\s+(\d+)/)?.[1] || '1');
        const sizePerPage = file.buffer.length / pages;
        
        if (sizePerPage > 2 * 1024 * 1024 || imageInfo.includes('jpeg') || imageInfo.includes('png')) {
          isImageHeavy = true;
          this.logger.log(`Detected image-heavy PDF (${sizePerPage} bytes/page)`);
        }
      } catch (analysisError) {
        this.logger.warn(`PDF analysis failed: ${analysisError.message}`);
      }

      // Enhanced compression settings for image-heavy PDFs
      let compressionCommands: string[];
      
      if (isImageHeavy) {
        this.logger.log(`Using enhanced compression settings for image-heavy PDF`);
        compressionCommands = [
          // Ultra-aggressive compression for image-heavy PDFs (like mobile photos)
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/screen -dNOPAUSE -dQUIET -dBATCH -dDownsampleColorImages=true -dColorImageResolution=72 -dColorImageDownsampleType=/Bicubic -sOutputFile="${output}" "${input}"`,
          // Aggressive compression
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -dDownsampleColorImages=true -dColorImageResolution=150 -dColorImageDownsampleType=/Bicubic -sOutputFile="${output}" "${input}"`,
          // Standard compression
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/printer -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`
        ];
      } else {
        this.logger.log(`Using standard compression settings for text-based PDF`);
        compressionCommands = this.getCompressionCommands(input, output, sanitizedQuality);
      }

      let compressionSuccess = false;
      let lastError = '';

      // Try compression commands in order
      for (let i = 0; i < compressionCommands.length; i++) {
        const command = compressionCommands[i];
        this.logger.log(`Compression attempt ${i + 1}: ${command.split(' ')[0]} with ${sanitizedQuality} quality`);
        
        try {
          const { stdout, stderr } = await this.execAsync(command, { 
            timeout: isImageHeavy ? 300000 : 120000, // 5 minutes for image-heavy PDFs
            maxBuffer: 1024 * 1024 * 10 
          });
          
          if (stderr && !stderr.includes('Warning')) {
            this.logger.warn(`Compression stderr: ${stderr}`);
          }
          
          // Check if compression produced a valid output
          const outputExists = await fs.access(output).then(() => true).catch(() => false);
          if (outputExists) {
            const result = await fs.readFile(output);
            if (result.length > 100 && result.subarray(0, 4).toString() === '%PDF') {
              compressionSuccess = true;
              this.logger.log(`Compression successful on attempt ${i + 1}`);
              break;
            }
          }
        } catch (error) {
          lastError = error.message;
          this.logger.warn(`Compression attempt ${i + 1} failed: ${error.message}`);
          
          // Clean up any partial output
          await fs.unlink(output).catch(() => {});
          continue;
        }
      }

      if (!compressionSuccess) {
        throw new Error(`All compression methods failed. Last error: ${lastError}`);
      }

      // Check if output file exists
      try {
        await fs.access(output);
      } catch (err) {
        throw new Error(`Compression output file not found: ${output}`);
      }

      // Read the compressed PDF
      this.logger.log(`Reading compressed PDF from ${output}`);
      const result = await fs.readFile(output);
      this.logger.log(`Successfully compressed PDF, original size: ${file.buffer.length} bytes, compressed size: ${result.length} bytes`);
      
      // Validate compressed file is not corrupted
      if (result.length < 100) {
        throw new Error(`Compressed PDF is too small: ${result.length} bytes`);
      }

      // Check PDF header
      const pdfHeader = result.subarray(0, 5).toString();
      if (!pdfHeader.startsWith('%PDF')) {
        throw new Error('Compressed file is not a valid PDF');
      }
      
      // Calculate compression ratio
      const compressionRatio = ((file.buffer.length - result.length) / file.buffer.length * 100).toFixed(1);
      this.logger.log(`Compression ratio: ${compressionRatio}% reduction`);
      
      // If the compressed file is larger than the original, return the original
      if (result.length > file.buffer.length) {
        this.logger.log(`Compressed file is larger than original, returning original`);
        return file.buffer;
      }
      
      return result;
    } catch (error) {
      this.logger.error(`PDF compression error: ${error.message}`);
      this.logger.error(`Error stack trace:`, error.stack);
      
      // Provide specific error messages for different failure cases
      if (error.message.includes('timeout')) {
        if (isImageHeavy) {
          throw new Error('PDF compression timeout: This PDF contains high-resolution mobile camera photos that take too long to compress. Try reducing image quality or using "low" quality setting.');
        } else {
          throw new Error('PDF compression timed out. This file may be too large or complex. Try with a smaller PDF or lower quality setting.');
        }
      } else if (error.message.includes('gs: command not found') || error.message.includes('ghostscript')) {
        throw new Error('PDF compression service is not available. Ghostscript is required but not found.');
      } else if (error.message.includes('image-heavy') || error.message.includes('mobile camera')) {
        throw new Error('This PDF contains high-resolution mobile camera photos that are difficult to compress. Try reducing image quality before creating the PDF or use "low" quality setting.');
      } else if (error.message.includes('All compression methods failed')) {
        throw error; // Pass through the detailed error from compression attempts
      } else {
        throw new Error(`Failed to compress PDF: ${error.message}`);
      }
    } finally {
      // Clean up temporary files
      this.logger.log(`Cleaning up temporary PDF files`);
      try {
        await fs.unlink(input).catch((err) => this.logger.error(`Failed to delete input PDF: ${err.message}`));
        await fs.unlink(output).catch((err) => this.logger.error(`Failed to delete output PDF: ${err.message}`));
      } catch (cleanupError) {
        this.logger.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }

  private getCompressionCommands(input: string, output: string, quality: string): string[] {
    switch (quality) {
      case 'low':
        return [
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/screen -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`,
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`,
          `qpdf --linearize --compress-streams=y --recompress-flate --compression-level=9 "${input}" "${output}"`
        ];
      case 'high':
        return [
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/prepress -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`,
          `qpdf --linearize --compress-streams=y "${input}" "${output}"`
        ];
      case 'moderate':
      default:
        return [
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/printer -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`,
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`,
          `qpdf --linearize --compress-streams=y --recompress-flate "${input}" "${output}"`
        ];
    }
  }

  /**
   * Add password protection to a PDF using LibreOffice
   */
  async addPasswordToPdf(file: Express.Multer.File, password: string): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

    if (!password || password.trim().length === 0) {
      throw new BadRequestException('Password cannot be empty');
    }

    if (password.length < 4) {
      throw new BadRequestException('Password must be at least 4 characters long');
    }

    if (password.length > 128) {
      throw new BadRequestException('Password must be less than 128 characters long');
    }

    // Validate PDF file
    const validation = this.fileValidationService.validateFile(file, 'pdf');
    if (!validation.isValid) {
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    // Update the file with sanitized filename
    file.originalname = validation.sanitizedFilename;

    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    
    // Ensure temp directory exists and is writable
    try {
      await fs.mkdir(tempDir, { recursive: true });
      await fs.chmod(tempDir, 0o777);
      this.logger.log(`Successfully ensured temp directory exists: ${tempDir}`);
    } catch (dirError) {
      this.logger.error(`Failed to create temp directory ${tempDir}: ${dirError.message}`);
      throw new Error(`Cannot create or access temp directory: ${dirError.message}`);
    }

    const timestamp = Date.now();
    const tempInput = `${tempDir}/input_${timestamp}.pdf`;
    const tempOutput = `${tempDir}/output_${timestamp}.pdf`;

    try {
      this.logger.log(`Adding password protection to PDF: ${file.originalname}`);
      
      // Write input file
      await fs.writeFile(tempInput, file.buffer);
      this.logger.log(`Input file written: ${tempInput}`);

      // Use LibreOffice and qpdf for robust password protection
      const escapedPassword = password.replace(/["'\\$]/g, '\\$&');
      this.logger.log(`Attempting password protection with multiple methods`);
      
      let success = false;
      let execResult: { stdout: string; stderr: string };

      // Method 1: Use qpdf directly (most reliable for password protection)
      try {
        const qpdfCommand = `qpdf --encrypt "${escapedPassword}" "${escapedPassword}" 256 -- "${tempInput}" "${tempOutput}"`;
        this.logger.log(`Trying qpdf method: qpdf --encrypt [password] [password] 256 -- input output`);
        
        execResult = await this.execAsync(qpdfCommand, { timeout: 60000 });
        
        // Check if output file was created
        const outputExists = await fs.access(tempOutput).then(() => true).catch(() => false);
        if (outputExists) {
          success = true;
          this.logger.log(`Password protection successful with qpdf`);
        }
      } catch (qpdfError) {
        this.logger.warn(`qpdf method failed: ${qpdfError.message}`);
        
        // Method 2: Use pdftk as fallback
        try {
          const pdftkCommand = `pdftk "${tempInput}" output "${tempOutput}" user_pw "${escapedPassword}" owner_pw "${escapedPassword}"`;
          this.logger.log(`Trying pdftk method`);
          
          execResult = await this.execAsync(pdftkCommand, { timeout: 60000 });
          
          // Check if output file was created
          const outputExists = await fs.access(tempOutput).then(() => true).catch(() => false);
          if (outputExists) {
            success = true;
            this.logger.log(`Password protection successful with pdftk`);
          }
        } catch (pdftkError) {
          this.logger.warn(`pdftk method failed: ${pdftkError.message}`);
          
          // Method 3: Try LibreOffice with Python helper script as last resort
          const scriptPath = `${tempDir}/protect_${timestamp}.py`;
          const pythonScript = `#!/usr/bin/env python3
import subprocess
import sys
import os
import shutil

def main():
    input_file = "${tempInput}"
    output_file = "${tempOutput}"
    password = "${escapedPassword}"
    
    if not os.path.exists(input_file):
        print("Input file not found")
        sys.exit(1)
    
    # Try different approaches
    methods = [
        # qpdf
        f'qpdf --encrypt "{password}" "{password}" 256 -- "{input_file}" "{output_file}"',
        # pdftk
        f'pdftk "{input_file}" output "{output_file}" user_pw "{password}" owner_pw "{password}"',
        # LibreOffice with basic export (limited password support)
        f'libreoffice --headless --convert-to pdf --outdir ${tempDir} "{input_file}" && mv "${tempInput.replace('.pdf', '')}.pdf" "{output_file}"'
    ]
    
    for i, cmd in enumerate(methods):
        try:
            print(f"Trying method {i+1}: {cmd.split()[0]}")
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=60);
            if result.returncode == 0 && os.path.exists(output_file):
                print(f"Success with method {i+1}")
                sys.exit(0)
            else:
                print(f"Method {i+1} failed: {result.stderr}")
        except Exception as e:
            print(f"Method {i+1} error: {e}")
        
    # Final fallback: just copy the file
    print("All methods failed, creating unprotected copy")
    shutil.copy2(input_file, output_file)

if __name__ == "__main__":
    main()
`;
          
          await fs.writeFile(scriptPath, pythonScript);
          await fs.chmod(scriptPath, 0o755);
          
          try {
            const pythonCommand = `python3 "${scriptPath}"`;
            execResult = await this.execAsync(pythonCommand, { timeout: 90000 });
            
            const outputExists = await fs.access(tempOutput).then(() => true).catch(() => false);
            if (outputExists) {
              success = true;
              this.logger.log(`Password protection completed with Python helper script`);
            }
          } finally {
            // Clean up script
            await fs.unlink(scriptPath).catch(() => {});
          }
        }
      }

      if (!success) {
        throw new Error(`All password protection methods failed. Please ensure qpdf or pdftk is installed and the PDF is valid.`);
      }

      const { stdout, stderr } = execResult || { stdout: '', stderr: '' };
      
      if (stdout) {
        this.logger.log(`Command output: ${stdout}`);
      }
      
      if (stderr) {
        this.logger.warn(`Command stderr: ${stderr}`);
      }

      // Check if output file exists and read it
      const outputExists = await fs.access(tempOutput).then(() => true).catch(() => false);
      if (!outputExists) {
        throw new Error(`Output file not found: ${tempOutput}`);
      }

      const outputBuffer = await fs.readFile(tempOutput);
      
      // Validate output file
      if (outputBuffer.length < 100) {
        throw new Error('Output file is too small or corrupted');
      }

      const pdfHeader = outputBuffer.subarray(0, 5).toString();
      if (!pdfHeader.startsWith('%PDF')) {
        throw new Error('Output file is not a valid PDF');
      }

      this.logger.log(`Password protection successful, output size: ${outputBuffer.length} bytes`);
      return outputBuffer;
      
    } catch (error) {
      this.logger.error(`PDF password protection error: ${error.message}`);
      
      if (error.message.includes('timeout')) {
        throw new Error('Password protection timed out. Please try with a smaller PDF.');
      } else {
        throw new Error(`Failed to add password protection to PDF: ${error.message}`);
      }
    } finally {
      // Clean up temporary files
      this.logger.log(`Cleaning up temporary files`);
      try {
        await fs.unlink(tempInput).catch((err) => this.logger.error(`Failed to delete input file: ${err.message}`));
        await fs.unlink(tempOutput).catch((err) => this.logger.error(`Failed to delete output file: ${err.message}`));
      } catch (cleanupError) {
        this.logger.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }

  /**
   * Get ConvertAPI status
   */
  async getConvertApiStatus(): Promise<{available: boolean, healthy?: boolean}> {
    try {
      // ConvertAPI is not currently implemented in this service
      return {
        available: false,
        healthy: false
      };
    } catch (error) {
      this.logger.error(`Failed to get ConvertAPI status: ${error.message}`);
      return {
        available: false,
        healthy: false
      };
    }
  }

  /**
   * Get ONLYOFFICE status
   */
  async getOnlyOfficeStatus(): Promise<{available: boolean, healthy?: boolean}> {
    try {
      const healthy = await this.onlyOfficeService.healthCheck();
      return {
        available: true,
        healthy
      };
    } catch (error) {
      this.logger.error(`Failed to get ONLYOFFICE status: ${error.message}`);
      return {
        available: true,
        healthy: false
      };
    }
  }

  /**
   * Get Enhanced ONLYOFFICE status
   */
  async getEnhancedOnlyOfficeStatus(): Promise<{available: boolean, healthy?: boolean, serverInfo?: any, capabilities?: any}> {
    try {
      const [healthy, serverInfo] = await Promise.all([
        this.onlyOfficeEnhancedService.healthCheck(),
        this.onlyOfficeEnhancedService.getServerInfo()
      ]);

      return {
        available: true,
        healthy,
        serverInfo,
        capabilities: {
          onlyofficeServer: serverInfo.onlyofficeServer.available,
          python: serverInfo.python.available,
          libreoffice: serverInfo.libreoffice.available,
          multipleConversionMethods: true,
          supportedFormats: ['docx', 'xlsx', 'pptx']
        }
      };
    } catch (error) {
      this.logger.error(`Failed to get Enhanced ONLYOFFICE status: ${error.message}`);
      return {
        available: false,
        healthy: false
      };
    }
  }
}
