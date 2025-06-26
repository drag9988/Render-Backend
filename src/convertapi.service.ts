import { Injectable, Logger, BadRequestException } from '@nestjs/common';
import axios from 'axios';
import * as FormData from 'form-data';
import { FileValidationService } from './file-validation.service';

export interface ConvertApiConfig {
  secret: string;
  baseUrl?: string;
  timeout?: number;
}

@Injectable()
export class ConvertApiService {
  private readonly logger = new Logger(ConvertApiService.name);
  private readonly secret: string;
  private readonly baseUrl: string;
  private readonly timeout: number;

  constructor(private readonly fileValidationService: FileValidationService) {
    this.secret = process.env.CONVERTAPI_SECRET;
    this.baseUrl = process.env.CONVERTAPI_BASE_URL || 'https://v2.convertapi.com';
    this.timeout = parseInt(process.env.CONVERTAPI_TIMEOUT || '60000', 10);

    if (!this.secret) {
      this.logger.warn('CONVERTAPI_SECRET not found. ConvertAPI integration will be disabled.');
    }
  }

  isAvailable(): boolean {
    return !!this.secret;
  }

  /**
   * Convert PDF to DOCX using ConvertAPI with validation
   */
  async convertPdfToDocx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
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

    // Validate file
    const validation = this.fileValidationService.validateFile(file, 'pdf');
    if (!validation.isValid) {
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    return this.convertPdf(pdfBuffer, validation.sanitizedFilename, 'docx');
  }

  /**
   * Convert PDF to XLSX using ConvertAPI with validation
   */
  async convertPdfToXlsx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
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

    // Validate file
    const validation = this.fileValidationService.validateFile(file, 'pdf');
    if (!validation.isValid) {
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    return this.convertPdf(pdfBuffer, validation.sanitizedFilename, 'xlsx');
  }

  /**
   * Convert PDF to PPTX using ConvertAPI with validation
   */
  async convertPdfToPptx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
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

    // Validate file
    const validation = this.fileValidationService.validateFile(file, 'pdf');
    if (!validation.isValid) {
      throw new BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
    }

    return this.convertPdf(pdfBuffer, validation.sanitizedFilename, 'pptx');
  }

  /**
   * Generic PDF conversion method with enhanced security
   */
  private async convertPdf(pdfBuffer: Buffer, filename: string, targetFormat: string): Promise<Buffer> {
    if (!this.isAvailable()) {
      throw new Error('ConvertAPI is not available. Please set CONVERTAPI_SECRET environment variable.');
    }

    // Additional buffer validation
    if (!pdfBuffer || pdfBuffer.length === 0) {
      throw new BadRequestException('Invalid or empty PDF buffer provided');
    }

    // Check for maximum file size (additional safety check)
    const maxSize = 50 * 1024 * 1024; // 50MB
    if (pdfBuffer.length > maxSize) {
      throw new BadRequestException(`File size ${pdfBuffer.length} bytes exceeds maximum allowed size of ${maxSize} bytes`);
    }

    try {
      this.logger.log(`Starting PDF to ${targetFormat.toUpperCase()} conversion using ConvertAPI for file: ${filename}`);

      // Sanitize filename for API
      const sanitizedFilename = this.sanitizeFilenameForApi(filename);

      // Create form data with validation
      const formData = new FormData();
      formData.append('File', pdfBuffer, {
        filename: sanitizedFilename,
        contentType: 'application/pdf'
      });

      // Set conversion parameters with input validation
      formData.append('StoreFile', 'true');
      
      // Format-specific parameters with validation
      if (targetFormat === 'docx') {
        formData.append('DocxVersion', '2019');
        formData.append('PageRange', '1-');
      } else if (targetFormat === 'xlsx') {
        formData.append('WorksheetName', 'Sheet1');
        formData.append('PageRange', '1-');
      } else if (targetFormat === 'pptx') {
        formData.append('PageRange', '1-');
      } else {
        throw new BadRequestException(`Unsupported target format: ${targetFormat}`);
      }

      // Make API request with enhanced error handling
      const response = await axios.post(
        `${this.baseUrl}/convert/pdf/to/${targetFormat}`,
        formData,
        {
          params: {
            Secret: this.secret
          },
          headers: {
            ...formData.getHeaders(),
          },
          timeout: this.timeout,
          responseType: 'json',
          maxContentLength: 100 * 1024 * 1024, // 100MB max response
          maxBodyLength: 60 * 1024 * 1024, // 60MB max request body
        }
      );

      // Enhanced response validation
      if (!response.data) {
        throw new Error('ConvertAPI returned empty response');
      }

      if (!response.data.Files || !Array.isArray(response.data.Files) || response.data.Files.length === 0) {
        throw new Error('ConvertAPI returned no files');
      }

      const fileInfo = response.data.Files[0];
      if (!fileInfo.Url) {
        throw new Error('ConvertAPI returned no file URL');
      }

      // Validate file URL
      if (!this.isValidUrl(fileInfo.Url)) {
        throw new Error('ConvertAPI returned invalid file URL');
      }

      this.logger.log(`ConvertAPI conversion successful. Downloading file from: ${fileInfo.Url}`);

      // Download the converted file with validation
      const fileResponse = await axios.get(fileInfo.Url, {
        responseType: 'arraybuffer',
        timeout: this.timeout,
        maxContentLength: 100 * 1024 * 1024, // 100MB max download
      });

      const convertedBuffer = Buffer.from(fileResponse.data);
      
      // Validate converted file size
      if (convertedBuffer.length === 0) {
        throw new Error('Downloaded file is empty');
      }

      if (convertedBuffer.length > 100 * 1024 * 1024) { // 100MB max
        throw new Error('Downloaded file size exceeds maximum limit');
      }

      this.logger.log(`Successfully converted PDF to ${targetFormat.toUpperCase()} using ConvertAPI. Output size: ${convertedBuffer.length} bytes`);

      return convertedBuffer;

    } catch (error) {
      this.logger.error(`ConvertAPI conversion failed for ${filename} to ${targetFormat}: ${error.message}`, error.stack);
      
      if (axios.isAxiosError(error)) {
        if (error.response) {
          this.logger.error(`ConvertAPI API Error: ${error.response.status} - ${JSON.stringify(error.response.data)}`);
        } else if (error.request) {
          this.logger.error('ConvertAPI request failed - no response received');
        }
      }
      
      throw new Error(`ConvertAPI conversion failed: ${error.message}`);
    }
  }

  /**
   * Get account balance/credits from ConvertAPI
   */
  async getBalance(): Promise<number> {
    if (!this.isAvailable()) {
      return 0;
    }

    try {
      const response = await axios.get(`${this.baseUrl}/user`, {
        params: {
          Secret: this.secret
        },
        timeout: 10000
      });

      return response.data?.SecondsLeft || 0;
    } catch (error) {
      this.logger.error(`Failed to get ConvertAPI balance: ${error.message}`);
      return 0;
    }
  }

  /**
   * Check if ConvertAPI service is healthy
   */
  async healthCheck(): Promise<boolean> {
    if (!this.isAvailable()) {
      return false;
    }

    try {
      const balance = await this.getBalance();
      return balance >= 0; // Even 0 balance means API is accessible
    } catch (error) {
      this.logger.error(`ConvertAPI health check failed: ${error.message}`);
      return false;
    }
  }

  /**
   * Sanitize filename for API usage
   */
  private sanitizeFilenameForApi(filename: string): string {
    if (!filename) {
      return 'document.pdf';
    }

    // Keep only alphanumeric, dots, dashes, and underscores
    const sanitized = filename
      .replace(/[^a-zA-Z0-9.\-_]/g, '_')
      .substring(0, 100); // Limit length for API

    return sanitized || 'document.pdf';
  }

  /**
   * Validate URL format
   */
  private isValidUrl(url: string): boolean {
    try {
      const urlObj = new URL(url);
      return urlObj.protocol === 'https:' && urlObj.hostname.includes('convertapi.com');
    } catch {
      return false;
    }
  }
}
