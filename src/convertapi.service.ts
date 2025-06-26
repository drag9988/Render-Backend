import { Injectable, Logger } from '@nestjs/common';
import axios from 'axios';
import * as FormData from 'form-data';

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

  constructor() {
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
   * Convert PDF to DOCX using ConvertAPI
   */
  async convertPdfToDocx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'docx');
  }

  /**
   * Convert PDF to XLSX using ConvertAPI
   */
  async convertPdfToXlsx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'xlsx');
  }

  /**
   * Convert PDF to PPTX using ConvertAPI
   */
  async convertPdfToPptx(pdfBuffer: Buffer, filename: string): Promise<Buffer> {
    return this.convertPdf(pdfBuffer, filename, 'pptx');
  }

  /**
   * Generic PDF conversion method
   */
  private async convertPdf(pdfBuffer: Buffer, filename: string, targetFormat: string): Promise<Buffer> {
    if (!this.isAvailable()) {
      throw new Error('ConvertAPI is not available. Please set CONVERTAPI_SECRET environment variable.');
    }

    try {
      this.logger.log(`Starting PDF to ${targetFormat.toUpperCase()} conversion using ConvertAPI for file: ${filename}`);

      // Create form data
      const formData = new FormData();
      formData.append('File', pdfBuffer, {
        filename: filename,
        contentType: 'application/pdf'
      });

      // Set conversion parameters
      formData.append('StoreFile', 'true');
      
      // Format-specific parameters
      if (targetFormat === 'docx') {
        formData.append('DocxVersion', '2019');
        formData.append('PageRange', '1-');
      } else if (targetFormat === 'xlsx') {
        formData.append('WorksheetName', 'Sheet1');
        formData.append('PageRange', '1-');
      } else if (targetFormat === 'pptx') {
        formData.append('PageRange', '1-');
      }

      // Make API request
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
          responseType: 'json'
        }
      );

      // Check if conversion was successful
      if (!response.data || !response.data.Files || response.data.Files.length === 0) {
        throw new Error('ConvertAPI returned no files');
      }

      const fileUrl = response.data.Files[0].Url;
      if (!fileUrl) {
        throw new Error('ConvertAPI returned no file URL');
      }

      this.logger.log(`ConvertAPI conversion successful. Downloading file from: ${fileUrl}`);

      // Download the converted file
      const fileResponse = await axios.get(fileUrl, {
        responseType: 'arraybuffer',
        timeout: this.timeout
      });

      const convertedBuffer = Buffer.from(fileResponse.data);
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
}
