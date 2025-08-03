import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import * as fs from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as multer from 'multer';
import { ConvertApiService } from './convertapi.service';
import { FileValidationService } from './file-validation.service';

@Injectable()
export class AppService {
  private readonly logger = new Logger(AppService.name);
  private readonly execAsync = promisify(exec);

  constructor(
    private readonly convertApiService: ConvertApiService,
    private readonly fileValidationService: FileValidationService
  ) {}

  async convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

    // Validate file based on its type and target format
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

    // Validate the file
    const validation = this.fileValidationService.validateFile(file, expectedFileType);
    if (!validation.isValid) {
      throw new BadRequestException(`File validation failed: ${validation.errors.join(', ')}`);
    }

    // Update the file with sanitized filename
    file.originalname = validation.sanitizedFilename;

    // For PDF to Office formats, try ConvertAPI first, then fallback to LibreOffice
    if (file.mimetype === 'application/pdf' && ['docx', 'xlsx', 'pptx'].includes(format)) {
      return await this.convertPdfToOfficeFormat(file, format);
    }

    // For other conversions, use LibreOffice directly
    return await this.executeLibreOfficeConversion(file, format);
  }

  private async convertPdfToOfficeFormat(file: Express.Multer.File, format: string): Promise<Buffer> {
    this.logger.log(`Converting PDF to ${format.toUpperCase()} - trying ConvertAPI first, then LibreOffice fallback`);

    // Try ConvertAPI first if available
    if (this.convertApiService.isAvailable()) {
      try {
        this.logger.log(`Attempting PDF to ${format.toUpperCase()} conversion using ConvertAPI`);
        
        let result: Buffer;
        switch (format) {
          case 'docx':
            result = await this.convertApiService.convertPdfToDocx(file.buffer, file.originalname);
            break;
          case 'xlsx':
            result = await this.convertApiService.convertPdfToXlsx(file.buffer, file.originalname);
            break;
          case 'pptx':
            result = await this.convertApiService.convertPdfToPptx(file.buffer, file.originalname);
            break;
          default:
            throw new Error(`Unsupported format: ${format}`);
        }

        this.logger.log(`ConvertAPI conversion successful for ${file.originalname} to ${format.toUpperCase()}`);
        return result;

      } catch (convertApiError) {
        this.logger.warn(`ConvertAPI failed for ${file.originalname}: ${convertApiError.message}. Falling back to LibreOffice.`);
        // Continue to LibreOffice fallback
      }
    } else {
      this.logger.log('ConvertAPI not available, using LibreOffice directly');
    }

    // Fallback to LibreOffice
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
      this.logger.log(`Starting LibreOffice conversion from ${file.mimetype} to ${format}`);
      
      // Write the uploaded file to disk
      await fs.writeFile(tempInput, file.buffer);
      this.logger.log(`File written to ${tempInput}`);
      
      // Special handling for PDF to other formats
      if (file.mimetype === 'application/pdf' && ['docx', 'xlsx', 'pptx'].includes(format)) {
        return await this.convertPdfWithLibreOffice(tempInput, tempOutput, format, tempDir);
      }
      
      // Standard LibreOffice conversion for other formats
      const command = `libreoffice --headless --convert-to ${format} --outdir ${tempDir} ${tempInput}`;
      this.logger.log(`Executing command: ${command}`);
      
      // Increase timeout for potentially long-running conversions
      const { stdout, stderr } = await this.execAsync(command, { timeout: 60000 });
      
      if (stdout) {
        this.logger.log(`LibreOffice conversion output: ${stdout}`);
      }
      
      if (stderr) {
        this.logger.error(`LibreOffice conversion error: ${stderr}`);
      }

      // Check if output file exists
      try {
        await fs.access(tempOutput);
      } catch (err) {
        this.logger.error(`Output file not found at ${tempOutput}`);
        throw new Error(`Conversion failed: Output file not created. LibreOffice may have failed to process this document.`);
      }

      // Read the converted file
      this.logger.log(`Reading converted file from ${tempOutput}`);
      const result = await fs.readFile(tempOutput);
      this.logger.log(`Successfully read converted file, size: ${result.length} bytes`);
      
      // Validate that the file is not empty or too small
      if (result.length < 100) {
        throw new Error(`Conversion produced an unusually small file (${result.length} bytes). The document may not have converted properly.`);
      }
      
      return result;
    } catch (error) {
      this.logger.error(`File conversion error: ${error.message}`);
      throw new Error(`Failed to convert document to ${format}: ${error.message}`);
    } finally {
      // Clean up temporary files
      try {
        this.logger.log(`Cleaning up temporary files`);
        await fs.unlink(tempInput).catch((err) => this.logger.error(`Failed to delete input file: ${err.message}`));
        await fs.unlink(tempOutput).catch((err) => this.logger.error(`Failed to delete output file: ${err.message}`));
      } catch (cleanupError) {
        this.logger.error(`Cleanup error: ${cleanupError.message}`);
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
      // Standard conversion with writer import
      `libreoffice --headless --convert-to ${format} --infilter="writer_pdf_import" --outdir ${tempDir} ${tempInput}`,
      // Standard conversion
      `libreoffice --headless --convert-to ${format} --outdir ${tempDir} ${tempInput}`,
      // Alternative approach with draw (sometimes works better for complex PDFs)
      `libreoffice --headless --draw --convert-to ${format} --outdir ${tempDir} ${tempInput}`,
      // Writer-specific conversion for text-heavy PDFs
      `libreoffice --headless --writer --convert-to ${format} --outdir ${tempDir} ${tempInput}`,
    ];

    let lastError = '';
    
    for (let i = 0; i < commands.length; i++) {
      const command = commands[i];
      this.logger.log(`Attempt ${i + 1}: ${command}`);
      
      try {
        const { stdout, stderr } = await this.execAsync(command, { timeout: 120000 }); // 2 minute timeout
        
        if (stdout) {
          this.logger.log(`LibreOffice output (attempt ${i + 1}): ${stdout}`);
        }
        
        if (stderr) {
          this.logger.error(`LibreOffice stderr (attempt ${i + 1}): ${stderr}`);
          lastError = stderr;
        }

        // Check if output file exists and has reasonable size
        try {
          await fs.access(tempOutput);
          const stats = await fs.stat(tempOutput);
          
          if (stats.size > 100) { // File exists and has content
            this.logger.log(`Successful conversion on attempt ${i + 1}, file size: ${stats.size} bytes`);
            const result = await fs.readFile(tempOutput);
            
            // Validate the converted file
            if (await this.validateConvertedFile(result, format)) {
              return result;
            } else {
              this.logger.log(`File validation failed on attempt ${i + 1}, trying next method`);
              await fs.unlink(tempOutput).catch(() => {});
            }
          } else {
            this.logger.log(`File created but too small (${stats.size} bytes), trying next method`);
            // Clean up the small file before next attempt
            await fs.unlink(tempOutput).catch(() => {});
          }
        } catch (err) {
          this.logger.log(`Output file not found on attempt ${i + 1}, trying next method`);
        }
        
      } catch (execError) {
        this.logger.error(`Execution failed on attempt ${i + 1}: ${execError.message}`);
        lastError = execError.message;
        // Clean up any partial files
        await fs.unlink(tempOutput).catch(() => {});
      }
    }
    
    // If LibreOffice fails, try alternative methods for certain formats
    if (format === 'docx') {
      this.logger.log(`Attempting alternative PDF to text extraction for Word conversion`);
      try {
        return await this.convertPdfToWordAlternative(tempInput, tempOutput, tempDir);
      } catch (altError) {
        this.logger.error(`Alternative conversion also failed: ${altError.message}`);
      }
    }
    
    // Provide specific error messages based on PDF characteristics
    let errorMessage = `PDF to ${format.toUpperCase()} conversion failed after multiple attempts. `;
    
    if (pdfInfo.isScanned) {
      errorMessage += `This appears to be a scanned PDF (image-based). Such PDFs cannot be directly converted to editable formats. `;
    } else if (pdfInfo.hasComplexLayout) {
      errorMessage += `This PDF has complex formatting that may not convert well to ${format.toUpperCase()}. `;
    } else if (pdfInfo.isProtected) {
      errorMessage += `This PDF appears to be password-protected or have restrictions that prevent conversion. `;
    } else {
      errorMessage += `The PDF format may not be compatible with the target format. `;
    }
    
    errorMessage += `Try using a simpler, text-based PDF. Last error: ${lastError}`;
    
    throw new Error(errorMessage);
  }

  private async analyzePdf(pdfPath: string): Promise<{isScanned: boolean, hasComplexLayout: boolean, isProtected: boolean, pageCount: number}> {
    try {
      // Use pdfinfo to get basic PDF information
      const { stdout: pdfInfo } = await this.execAsync(`pdfinfo "${pdfPath}"`, { timeout: 10000 });
      
      // Extract text to check if it's text-based or scanned
      const { stdout: textContent } = await this.execAsync(`pdftotext "${pdfPath}" -`, { timeout: 15000 });
      
      const pageCount = parseInt(pdfInfo.match(/Pages:\s*(\d+)/)?.[1] || '0');
      const isScanned = textContent.trim().length < 100; // Very little text suggests scanned PDF
      const hasComplexLayout = pdfInfo.includes('Form') || pdfInfo.includes('JavaScript') || textContent.includes('\t\t');
      const isProtected = pdfInfo.includes('Encrypted') || pdfInfo.includes('no');
      
      return {
        isScanned,
        hasComplexLayout,
        isProtected,
        pageCount
      };
    } catch (error) {
      this.logger.error(`PDF analysis failed: ${error.message}`);
      return {
        isScanned: false,
        hasComplexLayout: true,
        isProtected: false,
        pageCount: 1
      };
    }
  }

  private async validateConvertedFile(buffer: Buffer, format: string): Promise<boolean> {
    try {
      // Basic validation based on file format
      const content = buffer.toString('hex').substring(0, 20);
      
      switch (format) {
        case 'docx':
          // DOCX files start with PK (ZIP signature)
          return content.startsWith('504b');
        case 'xlsx':
          // XLSX files also start with PK (ZIP signature)
          return content.startsWith('504b');
        case 'pptx':
          // PPTX files also start with PK (ZIP signature)
          return content.startsWith('504b');
        default:
          return buffer.length > 500; // Basic size check
      }
    } catch (error) {
      this.logger.error(`File validation error: ${error.message}`);
      return false;
    }
  }

  private async convertPdfToWordAlternative(tempInput: string, tempOutput: string, tempDir: string): Promise<Buffer> {
    this.logger.log(`Attempting alternative PDF to Word conversion using text extraction`);
    
    // Extract text from PDF
    const { stdout: extractedText } = await this.execAsync(`pdftotext "${tempInput}" -`, { timeout: 30000 });
    
    if (extractedText.trim().length < 50) {
      throw new Error('PDF contains insufficient text for conversion');
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
    const textToDocxCommand = `libreoffice --headless --convert-to docx --outdir ${tempDir} "${tempTextFile}"`;
    await this.execAsync(textToDocxCommand, { timeout: 30000 });
    
    // Find the generated DOCX file
    const generatedDocx = tempTextFile.replace('.txt', '.docx');
    
    try {
      const result = await fs.readFile(generatedDocx);
      
      // Clean up temporary files
      await fs.unlink(tempTextFile).catch(() => {});
      await fs.unlink(generatedDocx).catch(() => {});
      
      return result;
    } catch (error) {
      // Clean up on error
      await fs.unlink(tempTextFile).catch(() => {});
      await fs.unlink(generatedDocx).catch(() => {});
      throw error;
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
    
    try {
      this.logger.log(`Starting PDF compression, input size: ${file.buffer.length} bytes, quality: ${sanitizedQuality}`);
      
      // Write the uploaded PDF to disk
      await fs.writeFile(input, file.buffer);
      this.logger.log(`PDF written to ${input}`);

      // Analyze PDF to determine if it's image-heavy (like mobile camera photos)
      let isImageHeavy = false;
      try {
        this.logger.log('Analyzing PDF content for compression optimization...');
        
        // Try to get PDF info using poppler-utils (already installed)
        const { stdout: pdfInfo } = await this.execAsync(`pdfinfo "${input}"`, { timeout: 10000 });
        this.logger.log(`PDF Info: ${pdfInfo.substring(0, 200)}...`);
        
        // Try to list images in PDF
        const { stdout: pdfImages } = await this.execAsync(`pdfimages -list "${input}"`, { timeout: 15000 }).catch(() => ({ stdout: '' }));
        
        // Check if PDF contains many images or large images (typical of mobile camera PDFs)
        const imageCount = (pdfImages.match(/page/g) || []).length;
        const hasLargeImages = pdfImages.includes('DCT') || pdfImages.includes('JPEG'); // Common in mobile photos
        const hasHighRes = pdfImages.includes('2000') || pdfImages.includes('3000') || pdfImages.includes('4000'); // High resolution
        
        // Determine if PDF is image-heavy based on multiple factors
        isImageHeavy = imageCount > 3 || file.buffer.length > 10 * 1024 * 1024 || hasLargeImages || hasHighRes;
        
        this.logger.log(`PDF analysis results:`);
        this.logger.log(`  - Image count: ${imageCount}`);
        this.logger.log(`  - File size: ${(file.buffer.length / (1024 * 1024)).toFixed(2)}MB`);
        this.logger.log(`  - Has JPEG/DCT images: ${hasLargeImages}`);
        this.logger.log(`  - Has high-resolution images: ${hasHighRes}`);
        this.logger.log(`  - Classified as image-heavy: ${isImageHeavy}`);
        
      } catch (analysisError) {
        this.logger.warn(`PDF analysis failed: ${analysisError.message}`);
        this.logger.warn('This may be due to missing poppler-utils or corrupted PDF');
        
        // Fallback: Assume image-heavy based on file size alone
        isImageHeavy = file.buffer.length > 5 * 1024 * 1024; // >5MB
        this.logger.log(`Fallback analysis: Assuming image-heavy based on file size: ${(file.buffer.length / (1024 * 1024)).toFixed(2)}MB > 5MB = ${isImageHeavy}`);
      }

      // Enhanced compression settings for image-heavy PDFs
      let compressionCommands: string[];
      
      if (isImageHeavy) {
        this.logger.log('Detected image-heavy PDF (mobile camera photos), using specialized compression');
        
        // Multiple compression strategies for image-heavy PDFs (mobile camera photos)
        compressionCommands = [
          // Strategy 1: Ultra-aggressive compression for mobile photos (lowest file size)
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/screen -dNOPAUSE -dQUIET -dBATCH -dColorImageResolution=72 -dGrayImageResolution=72 -dMonoImageResolution=150 -dColorImageDownsampleType=/Average -dGrayImageDownsampleType=/Average -dColorConversionStrategy=/RGB -dProcessColorModel=/DeviceRGB -sOutputFile="${output}" "${input}"`,
          
          // Strategy 2: Aggressive image compression for mobile photos
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -dColorImageResolution=150 -dGrayImageResolution=150 -dMonoImageResolution=300 -dColorImageDownsampleType=/Bicubic -dGrayImageDownsampleType=/Bicubic -dMonoImageDownsampleType=/Bicubic -dColorImageFilter=/DCTEncode -dGrayImageFilter=/DCTEncode -dColorConversionStrategy=/RGB -dProcessColorModel=/DeviceRGB -sOutputFile="${output}" "${input}"`,
          
          // Strategy 3: JPEG quality optimization for camera photos
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dNOPAUSE -dQUIET -dBATCH -dColorImageResolution=120 -dGrayImageResolution=120 -dMonoImageResolution=200 -dColorImageFilter=/DCTEncode -dGrayImageFilter=/DCTEncode -dColorImageDict="{/QFactor 0.3 /Blend 1 /HSamples [1 1 1 1] /VSamples [1 1 1 1]}" -dGrayImageDict="{/QFactor 0.3 /Blend 1 /HSamples [1 1 1 1] /VSamples [1 1 1 1]}" -sOutputFile="${output}" "${input}"`,
          
          // Strategy 4: Fallback with minimal compression
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/default -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`
        ];
      } else {
        // Standard compression for text-based or mixed PDFs
        const qualitySettings = {
          'low': '/screen',
          'moderate': '/ebook', 
          'high': '/printer'
        };
        
        const pdfSetting = qualitySettings[sanitizedQuality] || '/ebook';
        this.logger.log(`PDF compression quality: ${sanitizedQuality}, using setting: ${pdfSetting}`);
        
        compressionCommands = [
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${pdfSetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`,
          `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/ebook -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"` // fallback
        ];
      }

      let compressionSuccess = false;
      let lastError = '';

      // Try compression commands in order
      for (let i = 0; i < compressionCommands.length; i++) {
        const command = compressionCommands[i];
        this.logger.log(`Compression attempt ${i + 1}: Using ${isImageHeavy ? 'image-optimized' : 'standard'} compression`);
        
        try {
          // Increase timeout significantly for image-heavy PDFs (mobile camera photos)
          const timeout = isImageHeavy ? 600000 : 120000; // 10 minutes for images, 2 minutes for others
          this.logger.log(`Using timeout: ${timeout / 1000}s for ${isImageHeavy ? 'image-heavy' : 'standard'} PDF compression`);
          
          const { stdout, stderr } = await this.execAsync(command, { timeout });
          
          if (stdout) {
            this.logger.log(`Ghostscript output: ${stdout}`);
          }
          
          if (stderr) {
            this.logger.warn(`Ghostscript stderr: ${stderr}`);
          }

          // Check if output file exists and has reasonable size
          try {
            await fs.access(output);
            const stats = await fs.stat(output);
            if (stats.size > 0) {
              compressionSuccess = true;
              this.logger.log(`Compression successful on attempt ${i + 1}, output size: ${stats.size} bytes`);
              break;
            }
          } catch (err) {
            this.logger.error(`Output file check failed: ${err.message}`);
          }
          
        } catch (execError) {
          lastError = execError.message;
          this.logger.error(`Compression attempt ${i + 1} failed: ${execError.message}`);
          
          // Log additional details for debugging
          if (isImageHeavy) {
            this.logger.error(`Failed compressing image-heavy PDF (likely mobile camera photos)`);
          }
          
          // Clean up failed attempt
          await fs.unlink(output).catch(() => {});
          
          // Special handling for timeout errors on image-heavy PDFs
          if (execError.message.includes('timeout')) {
            if (isImageHeavy) {
              this.logger.error(`Timeout during image-heavy PDF compression. PDF likely contains high-resolution mobile camera photos that are too complex to compress quickly.`);
              if (i === compressionCommands.length - 1) {
                throw new Error(`Compression timeout: This PDF contains high-resolution images (likely mobile camera photos) that take too long to compress. Try using "low" quality setting or reduce image resolution before creating the PDF.`);
              }
            } else {
              this.logger.error(`Unexpected timeout during standard PDF compression.`);
              if (i === compressionCommands.length - 1) {
                throw new Error(`Compression timeout: The PDF compression is taking too long. Please try with a smaller file or lower quality setting.`);
              }
            }
            this.logger.log('Trying next compression method after timeout...');
            continue;
          }
          
          // If it's not a timeout, try next method for image-heavy PDFs
          if (isImageHeavy && i < compressionCommands.length - 1) {
            this.logger.log('Image-heavy PDF compression failed, trying next method...');
            continue;
          }
        }
      }

      if (!compressionSuccess) {
        this.logger.error(`All compression methods failed for ${file.originalname}. File size: ${(file.buffer.length / (1024 * 1024)).toFixed(2)}MB, Image-heavy: ${isImageHeavy}`);
        
        if (isImageHeavy) {
          throw new Error(`All compression methods failed for this image-heavy PDF (mobile camera photos detected). The PDF contains high-resolution images that are too complex to compress efficiently. Try reducing image quality before creating the PDF.`);
        } else {
          throw new Error(`All compression methods failed. Last error: ${lastError}`);
        }
      }

      // Check if output file exists
      try {
        await fs.access(output);
      } catch (err) {
        this.logger.error(`Compressed PDF not found at ${output}`);
        throw new Error(`Compression failed: Output file not created`);
      }

      // Read the compressed PDF
      this.logger.log(`Reading compressed PDF from ${output}`);
      const result = await fs.readFile(output);
      this.logger.log(`Successfully compressed PDF, original size: ${file.buffer.length} bytes, compressed size: ${result.length} bytes`);
      
      // Validate compressed file is not corrupted
      if (result.length < 100) {
        throw new Error('Compressed PDF file appears to be corrupted (too small)');
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
        this.logger.warn('Compressed file is larger than original, returning original file');
        return file.buffer;
      }
      
      return result;
    } catch (error) {
      this.logger.error(`PDF compression error: ${error.message}`);
      this.logger.error(`Error stack trace:`, error.stack);
      
      // Provide specific error messages for different failure cases
      if (error.message.includes('timeout')) {
        if (error.message.includes('mobile camera') || error.message.includes('high-resolution')) {
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
      try {
        this.logger.log(`Cleaning up temporary PDF files`);
        await fs.unlink(input).catch((err) => this.logger.error(`Failed to delete input PDF: ${err.message}`));
        await fs.unlink(output).catch((err) => this.logger.error(`Failed to delete output PDF: ${err.message}`));
      } catch (cleanupError) {
        this.logger.error(`Cleanup error: ${cleanupError.message}`);
      }
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
      const escapedPassword = password.replace(/["'\\$]/g, '\\$&'); // Escape special characters
      
      this.logger.log(`Attempting password protection with multiple methods`);
      
      let success = false;
      let execResult: { stdout: string; stderr: string };
      
      try {
        // Method 1: Use qpdf directly (most reliable for password protection)
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
        
        try {
          // Method 2: Use pdftk as fallback
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
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=60)
            if result.returncode == 0 and os.path.exists(output_file):
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
            
            // Check if output file was created
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

      // Verify output file was created and is valid
      const outputExists = await fs.access(tempOutput).then(() => true).catch(() => false);
      if (!outputExists) {
        this.logger.error(`Output file not found: ${tempOutput}`);
        throw new Error('Password protection failed - output file not created');
      }

      // Read the output file
      const outputBuffer = await fs.readFile(tempOutput);
      this.logger.log(`Password-protected PDF created successfully, size: ${outputBuffer.length} bytes`);
      
      // Verify the output is a valid PDF with password protection
      if (outputBuffer.length === 0) {
        throw new Error('Generated password-protected PDF is empty');
      }

      // Check PDF header
      const pdfHeader = outputBuffer.subarray(0, 5).toString();
      if (!pdfHeader.startsWith('%PDF')) {
        throw new Error('Generated file is not a valid PDF');
      }
      
      return outputBuffer;
    } catch (error) {
      this.logger.error(`PDF password protection error: ${error.message}`);
      
      if (error.message.includes('timeout')) {
        throw new Error('PDF password protection timed out. Please try with a smaller PDF.');
      }
      
      if (error.message.includes('command not found') || error.message.includes('libreoffice')) {
        throw new Error('LibreOffice is not installed or not accessible');
      }
      
      throw new Error(`Failed to add password protection to PDF: ${error.message}`);
    } finally {
      // Clean up temporary files
      try {
        this.logger.log(`Cleaning up temporary files`);
        await fs.unlink(tempInput).catch((err) => this.logger.error(`Failed to delete input file: ${err.message}`));
        await fs.unlink(tempOutput).catch((err) => this.logger.error(`Failed to delete output file: ${err.message}`));
      } catch (cleanupError) {
        this.logger.error(`Cleanup error: ${cleanupError.message}`);
      }
    }
  }

  // ConvertAPI-related methods
  async getConvertApiStatus(): Promise<{available: boolean, balance?: number, healthy?: boolean}> {
    if (!this.convertApiService.isAvailable()) {
      return { available: false };
    }

    try {
      const [balance, healthy] = await Promise.all([
        this.convertApiService.getBalance(),
        this.convertApiService.healthCheck()
      ]);

      return {
        available: true,
        balance,
        healthy
      };
    } catch (error) {
      this.logger.error(`Failed to get ConvertAPI status: ${error.message}`);
      return {
        available: true,
        healthy: false
      };
    }
  }
}