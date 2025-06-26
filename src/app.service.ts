import { Injectable, Logger } from '@nestjs/common';
import * as fs from 'fs/promises';
import { exec } from 'child_process';
import { promisify } from 'util';
import * as multer from 'multer';
import { ConvertApiService } from './convertapi.service';
import { SecurityService } from './security.service';

@Injectable()
export class AppService {
  private readonly logger = new Logger(AppService.name);
  private readonly execAsync = promisify(exec);

  constructor(
    private readonly convertApiService: ConvertApiService,
    private readonly securityService: SecurityService
  ) {}

  async convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer> {
    if (!file || !file.buffer) {
      throw new Error('Invalid file provided');
    }

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
    // Create a secure temporary filename
    const timestamp = Date.now();
    const secureFilename = this.securityService.generateSecureFilename(file.originalname, 'conversion');
    
    // Use the OS temp directory or fallback to /tmp
    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    this.logger.log(`Using temp directory: ${tempDir}`);
    
    const tempInput = `${tempDir}/${secureFilename}`;
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
      let command = `libreoffice --headless --convert-to ${format} --outdir "${tempDir}" "${tempInput}"`;
      this.logger.log(`Executing command: ${command}`);
      
      try {
        // Increase timeout for potentially long-running conversions
        const { stdout, stderr } = await this.execAsync(command, { timeout: 60000 });
        
        if (stdout) {
          this.logger.log(`LibreOffice conversion output: ${stdout}`);
        }
        
        if (stderr) {
          this.logger.error(`LibreOffice conversion error: ${stderr}`);
          // Don't throw immediately on stderr - LibreOffice often outputs warnings
        }
      } catch (execError) {
        this.logger.error(`LibreOffice command failed: ${execError.message}`);
        
        // Try alternative command with soffice
        try {
          command = `soffice --headless --convert-to ${format} --outdir "${tempDir}" "${tempInput}"`;
          this.logger.log(`Trying alternative command: ${command}`);
          
          const { stdout: altStdout, stderr: altStderr } = await this.execAsync(command, { timeout: 60000 });
          
          if (altStdout) {
            this.logger.log(`Alternative LibreOffice output: ${altStdout}`);
          }
          
          if (altStderr) {
            this.logger.error(`Alternative LibreOffice error: ${altStderr}`);
          }
        } catch (altError) {
          this.logger.error(`Alternative LibreOffice command also failed: ${altError.message}`);
          throw new Error(`LibreOffice is not available or not properly configured. Both 'libreoffice' and 'soffice' commands failed.`);
        }
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
    if (!file || !file.buffer) {
      throw new Error('Invalid PDF file provided');
    }

    // Create a unique filename with timestamp
    const timestamp = Date.now();
    
    // Use the OS temp directory or fallback to /tmp
    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    this.logger.log(`Using temp directory for PDF compression: ${tempDir}`);
    
    const input = `${tempDir}/${timestamp}_input.pdf`;
    const output = `${tempDir}/${timestamp}_output.pdf`;
    
    // Map quality settings to Ghostscript PDF settings
    // /screen - low quality, smaller size (72 dpi)
    // /ebook - medium quality, medium size (150 dpi)
    // /printer - high quality, larger size (300 dpi)
    const qualitySettings = {
      'low': '/screen',
      'moderate': '/ebook',
      'high': '/printer'
    };
    
    // Default to moderate if invalid quality is provided
    const pdfSetting = qualitySettings[quality] || '/ebook';
    this.logger.log(`PDF compression quality: ${quality}, using setting: ${pdfSetting}`);

    try {
      this.logger.log(`Starting PDF compression, input size: ${file.buffer.length} bytes`);
      
      // Write the uploaded PDF to disk
      await fs.writeFile(input, file.buffer);
      this.logger.log(`PDF written to ${input}`);

      // Execute Ghostscript for PDF compression
      let command = `gs -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${pdfSetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`;
      this.logger.log(`Executing command: ${command}`);
      
      try {
        const { stdout, stderr } = await this.execAsync(command, { timeout: 120000 }); // 2 minute timeout
        
        if (stdout) {
          this.logger.log(`Ghostscript output: ${stdout}`);
        }
        
        if (stderr) {
          this.logger.error(`Ghostscript compression error: ${stderr}`);
          // Don't throw immediately on stderr - check if file was created
        }
      } catch (execError) {
        this.logger.error(`Ghostscript command failed: ${execError.message}`);
        
        // Try alternative ghostscript command
        try {
          command = `ghostscript -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=${pdfSetting} -dNOPAUSE -dQUIET -dBATCH -sOutputFile="${output}" "${input}"`;
          this.logger.log(`Trying alternative ghostscript command: ${command}`);
          
          const { stdout: altStdout, stderr: altStderr } = await this.execAsync(command, { timeout: 120000 });
          
          if (altStdout) {
            this.logger.log(`Alternative Ghostscript output: ${altStdout}`);
          }
          
          if (altStderr) {
            this.logger.error(`Alternative Ghostscript error: ${altStderr}`);
          }
        } catch (altError) {
          this.logger.error(`Alternative Ghostscript command also failed: ${altError.message}`);
          throw new Error(`Ghostscript is not available or not properly configured. PDF compression failed.`);
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
      
      // If the compressed file is larger than the original, return the original
      if (result.length > file.buffer.length) {
        this.logger.log(`Compressed file is larger than original, returning original file`);
        return file.buffer;
      }
      
      return result;
    } catch (error) {
      this.logger.error(`PDF compression error: ${error.message}`);
      throw new Error(`Failed to compress PDF: ${error.message}`);
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