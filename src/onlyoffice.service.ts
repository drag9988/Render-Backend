import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import axios from 'axios';
import * as FormData from 'form-data';
import * as fs from 'fs/promises';
import * as path from 'path';
import { FileValidationService } from './file-validation.service';

export interface OnlyOfficeConfig {
  documentServerUrl: string;
  timeout?: number;
  jwtSecret?: string;
}

@Injectable()
export class OnlyOfficeService {
  private readonly logger = new Logger(OnlyOfficeService.name);
  private readonly documentServerUrl: string;
  private readonly timeout: number;
  private readonly jwtSecret?: string;

  constructor(private readonly fileValidationService: FileValidationService) {
    this.documentServerUrl = process.env.ONLYOFFICE_DOCUMENT_SERVER_URL || '';
    this.timeout = parseInt(process.env.ONLYOFFICE_TIMEOUT || '120000', 10);
    this.jwtSecret = process.env.ONLYOFFICE_JWT_SECRET;

    if (!this.documentServerUrl) {
      this.logger.warn('ONLYOFFICE_DOCUMENT_SERVER_URL not found. ONLYOFFICE integration will be disabled.');
    } else {
      this.logger.log(`ONLYOFFICE Document Server configured at: ${this.documentServerUrl}`);
    }
  }

  isAvailable(): boolean {
    return !!this.documentServerUrl;
  }

  /**
   * Convert PDF to DOCX using ONLYOFFICE Document Server
   */
  async convertPdfToDocx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'docx');
  }

  /**
   * Convert PDF to XLSX using ONLYOFFICE Document Server
   */
  async convertPdfToXlsx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'xlsx');
  }

  /**
   * Convert PDF to PPTX using ONLYOFFICE Document Server
   */
  async convertPdfToPptx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'pptx');
  }

  /**
   * Generic PDF conversion using ONLYOFFICE Document Server
   */
  private async convertPdf(pdfBuffer: Buffer, filename: string, targetFormat: string): Promise<Buffer> {
    if (!this.isAvailable()) {
      throw new Error('ONLYOFFICE Document Server is not available. Please set ONLYOFFICE_DOCUMENT_SERVER_URL environment variable.');
    }

    if (!pdfBuffer || pdfBuffer.length === 0) {
      throw new BadRequestException('Invalid or empty PDF buffer provided');
    }

    // Validate file using existing service
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

    const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
    const timestamp = Date.now();
    const tempInputPath = path.join(tempDir, `${timestamp}_input.pdf`);
    const tempOutputPath = path.join(tempDir, `${timestamp}_output.${targetFormat}`);

    try {
      this.logger.log(`Starting PDF to ${targetFormat.toUpperCase()} conversion using ONLYOFFICE for file: ${filename}`);

      // Write PDF to temp file
      await fs.writeFile(tempInputPath, pdfBuffer);
      this.logger.log(`PDF written to temporary file: ${tempInputPath}`);

      // Method 1: Try direct conversion API (if available)
      try {
        const convertedBuffer = await this.convertViaDirectAPI(tempInputPath, targetFormat, validation.sanitizedFilename);
        if (convertedBuffer) {
          this.logger.log(`Successfully converted PDF to ${targetFormat.toUpperCase()} via direct API. Output size: ${convertedBuffer.length} bytes`);
          return convertedBuffer;
        }
      } catch (directApiError) {
        this.logger.warn(`Direct API conversion failed: ${directApiError.message}. Trying LibreOffice with ONLYOFFICE integration.`);
      }

      // Method 2: Enhanced LibreOffice conversion with ONLYOFFICE-compatible settings
      const convertedBuffer = await this.convertViaEnhancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
      
      this.logger.log(`Successfully converted PDF to ${targetFormat.toUpperCase()} using enhanced LibreOffice. Output size: ${convertedBuffer.length} bytes`);
      return convertedBuffer;

    } catch (error) {
      this.logger.error(`ONLYOFFICE conversion failed for ${filename} to ${targetFormat}: ${error.message}`);
      throw new Error(`ONLYOFFICE conversion failed: ${error.message}`);
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
   * Convert via ONLYOFFICE Document Server direct API
   */
  private async convertViaDirectAPI(inputPath: string, targetFormat: string, filename: string): Promise<Buffer | null> {
    try {
      // Upload file to ONLYOFFICE Document Server
      const uploadUrl = await this.uploadFileToOnlyOffice(inputPath, filename);
      
      // Request conversion
      const conversionRequest = {
        async: false,
        filetype: 'pdf',
        key: this.generateConversionKey(filename),
        outputtype: targetFormat,
        title: filename,
        url: uploadUrl
      };

      // Add JWT if configured
      if (this.jwtSecret) {
        conversionRequest['token'] = this.generateJWT(conversionRequest);
      }

      const conversionUrl = `${this.documentServerUrl}/ConvertService.ashx`;
      
      this.logger.log(`Sending conversion request to ONLYOFFICE: ${conversionUrl}`);
      
      const response = await axios.post(conversionUrl, conversionRequest, {
        timeout: this.timeout,
        headers: {
          'Content-Type': 'application/json'
        }
      });

      if (response.data.error) {
        throw new Error(`ONLYOFFICE conversion error: ${response.data.error}`);
      }

      if (!response.data.fileUrl) {
        throw new Error('ONLYOFFICE returned no file URL');
      }

      // Download the converted file
      const fileResponse = await axios.get(response.data.fileUrl, {
        responseType: 'arraybuffer',
        timeout: this.timeout
      });

      const convertedBuffer = Buffer.from(fileResponse.data);
      
      if (convertedBuffer.length === 0) {
        throw new Error('Downloaded file is empty');
      }

      return convertedBuffer;

    } catch (error) {
      this.logger.warn(`Direct API conversion failed: ${error.message}`);
      return null;
    }
  }

  /**
   * Enhanced LibreOffice conversion with ONLYOFFICE-compatible settings
   */
  private async convertViaEnhancedLibreOffice(inputPath: string, outputPath: string, targetFormat: string): Promise<Buffer> {
    const { exec } = require('child_process');
    const { promisify } = require('util');
    const execAsync = promisify(exec);

    // Enhanced commands for better PDF to PowerPoint conversion
    const enhancedCommands = [];
    
    if (targetFormat === 'pptx') {
      enhancedCommands.push(
        // Method 1: Use Impress directly for PowerPoint conversion
        `libreoffice --headless --impress --convert-to pptx:"Impress MS PowerPoint 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 2: Draw import with PowerPoint export
        `libreoffice --headless --draw --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 3: Writer import then PowerPoint export (for text-heavy PDFs)
        `libreoffice --headless --writer --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Method 4: Standard PowerPoint conversion
        `libreoffice --headless --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`
      );
    } else {
      enhancedCommands.push(
        // Enhanced conversion for other formats
        `libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`,
        
        // Alternative methods based on target format
        targetFormat === 'docx' 
          ? `libreoffice --headless --writer --convert-to ${targetFormat}:"MS Word 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`
          : `libreoffice --headless --calc --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`
      );
    }

    let lastError = '';
    
    for (let i = 0; i < enhancedCommands.length; i++) {
      const command = enhancedCommands[i];
      this.logger.log(`Attempting enhanced LibreOffice conversion ${i + 1}/${enhancedCommands.length} for ${targetFormat}`);
      
      try {
        const { stdout, stderr } = await execAsync(command, { 
          timeout: this.timeout - 10000, // Leave some buffer time
          maxBuffer: 1024 * 1024 * 10 // 10MB
        });
        
        if (stdout) {
          this.logger.log(`LibreOffice output: ${stdout}`);
        }
        
        if (stderr && !stderr.includes('Warning')) {
          this.logger.warn(`LibreOffice stderr: ${stderr}`);
        }

        // Check for output file
        const expectedOutputPath = path.join(
          path.dirname(outputPath), 
          path.basename(inputPath, '.pdf') + '.' + targetFormat
        );

        try {
          await fs.access(expectedOutputPath);
          const result = await fs.readFile(expectedOutputPath);
          
          if (result.length > 500) { // Reasonable minimum size
            this.logger.log(`Enhanced LibreOffice conversion successful: ${result.length} bytes`);
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

    throw new Error(`Enhanced LibreOffice conversion failed after ${enhancedCommands.length} attempts. Last error: ${lastError}`);
  }

  /**
   * Upload file to ONLYOFFICE Document Server (if it supports file upload)
   */
  private async uploadFileToOnlyOffice(filePath: string, filename: string): Promise<string> {
    // Option 1: Use ONLYOFFICE's file upload if available
    try {
      const fileBuffer = await fs.readFile(filePath);
      const formData = new FormData();
      formData.append('file', fileBuffer, filename);

      const uploadResponse = await axios.post(`${this.documentServerUrl}/upload`, formData, {
        headers: formData.getHeaders(),
        timeout: this.timeout
      });

      if (uploadResponse.data && uploadResponse.data.url) {
        return uploadResponse.data.url;
      }
    } catch (uploadError) {
      this.logger.warn(`ONLYOFFICE upload failed: ${uploadError.message}`);
    }

    // Option 2: Create a temporary accessible URL via your own server
    // This would require implementing a temporary file serving endpoint
    const tempUrl = await this.createTempFileUrl(filePath, filename);
    return tempUrl;
  }

  /**
   * Create a temporary accessible URL for the file
   */
  private async createTempFileUrl(filePath: string, filename: string): Promise<string> {
    // For now, we'll use a simple approach
    // In production, you might want to use cloud storage or a proper file serving endpoint
    const serverUrl = process.env.SERVER_URL || 'http://localhost:10000';
    const tempEndpoint = `${serverUrl}/temp/${Date.now()}_${filename}`;
    
    // Note: You would need to implement the /temp endpoint in your controller
    // For this implementation, we'll assume the file is accessible at this URL
    return tempEndpoint;
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

    // Simple JWT implementation - in production, use a proper JWT library
    const jwt = require('jsonwebtoken');
    return jwt.sign(payload, this.jwtSecret, { expiresIn: '1h' });
  }

  /**
   * Health check for ONLYOFFICE Document Server
   */
  async healthCheck(): Promise<boolean> {
    if (!this.isAvailable()) {
      return false;
    }

    try {
      const response = await axios.get(`${this.documentServerUrl}/healthcheck`, {
        timeout: 5000
      });
      return response.status === 200;
    } catch (error) {
      // Try alternative health check endpoints
      try {
        const response = await axios.get(`${this.documentServerUrl}/`, {
          timeout: 5000
        });
        return response.status === 200;
      } catch (altError) {
        this.logger.error(`ONLYOFFICE health check failed: ${error.message}`);
        return false;
      }
    }
  }

  /**
   * Get ONLYOFFICE server information
   */
  async getServerInfo(): Promise<any> {
    if (!this.isAvailable()) {
      return { available: false, reason: 'Document server URL not configured' };
    }

    try {
      const healthy = await this.healthCheck();
      return {
        available: true,
        healthy,
        url: this.documentServerUrl,
        jwtEnabled: !!this.jwtSecret
      };
    } catch (error) {
      return {
        available: false,
        reason: error.message
      };
    }
  }
}
