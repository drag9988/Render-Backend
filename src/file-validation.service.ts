import { Injectable, BadRequestException } from '@nestjs/common';
import * as path from 'path';

export interface FileValidationResult {
  isValid: boolean;
  errors: string[];
  sanitizedFilename: string;
}

@Injectable()
export class FileValidationService {
  
  // Allowed MIME types for different conversions
  private readonly allowedMimeTypes = {
    pdf: ['application/pdf'],
    word: [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
      'application/msword', // .doc (legacy)
      'application/vnd.ms-word',
      'text/plain' // .txt files
    ],
    excel: [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel', // .xls (legacy)
      'application/excel',
      'application/x-excel',
      'text/csv' // .csv files
    ],
    powerpoint: [
      'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
      'application/vnd.ms-powerpoint', // .ppt (legacy)
      'application/mspowerpoint'
    ]
  };

  // Allowed file extensions
  private readonly allowedExtensions = {
    pdf: ['.pdf'],
    word: ['.docx', '.doc', '.txt'],
    excel: ['.xlsx', '.xls', '.csv'],
    powerpoint: ['.pptx', '.ppt']
  };

  // File size limits (in bytes) - 50MB max
  private readonly maxFileSize = 50 * 1024 * 1024; // 50MB

  /**
   * Validate PDF file
   */
  validatePdfFile(file: Express.Multer.File): FileValidationResult {
    const errors: string[] = [];
    
    // Check if file exists
    if (!file || !file.buffer) {
      errors.push('No file provided');
      return { isValid: false, errors, sanitizedFilename: '' };
    }

    // Check file size
    if (file.size > this.maxFileSize) {
      errors.push(`PDF file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
    }

    // Check MIME type
    if (!this.allowedMimeTypes.pdf.includes(file.mimetype)) {
      errors.push(`Invalid MIME type: ${file.mimetype}. Expected: ${this.allowedMimeTypes.pdf.join(', ')}`);
    }

    // Check file extension
    const ext = path.extname(file.originalname).toLowerCase();
    if (!this.allowedExtensions.pdf.includes(ext)) {
      errors.push(`Invalid file extension: ${ext}. Expected: ${this.allowedExtensions.pdf.join(', ')}`);
    }

    // Validate PDF header (magic bytes)
    if (file.buffer && file.buffer.length >= 4) {
      const header = file.buffer.slice(0, 4).toString('ascii');
      if (header !== '%PDF') {
        errors.push('Invalid PDF file format - missing PDF header');
      }
    }

    // Check for minimum file size (empty PDFs)
    if (file.size < 100) {
      errors.push('PDF file appears to be empty or corrupted');
    }

    const sanitizedFilename = this.sanitizeFilename(file.originalname);

    return {
      isValid: errors.length === 0,
      errors,
      sanitizedFilename
    };
  }

  /**
   * Validate Word file
   */
  validateWordFile(file: Express.Multer.File): FileValidationResult {
    const errors: string[] = [];
    
    if (!file || !file.buffer) {
      errors.push('No file provided');
      return { isValid: false, errors, sanitizedFilename: '' };
    }

    // Check file size
    if (file.size > this.maxFileSize) {
      errors.push(`Word file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
    }

    // Check MIME type
    if (!this.allowedMimeTypes.word.includes(file.mimetype)) {
      errors.push(`Invalid MIME type: ${file.mimetype}. Expected one of: ${this.allowedMimeTypes.word.join(', ')}`);
    }

    // Check file extension
    const ext = path.extname(file.originalname).toLowerCase();
    if (!this.allowedExtensions.word.includes(ext)) {
      errors.push(`Invalid file extension: ${ext}. Expected one of: ${this.allowedExtensions.word.join(', ')}`);
    }

    // Validate file header
    if (file.buffer && file.buffer.length >= 8) {
      const isValidWordFile = this.validateOfficeFileHeader(file.buffer, ext);
      if (!isValidWordFile) {
        errors.push(`Invalid ${ext} file format - file header does not match expected format`);
      }
    }

    // Check for minimum file size
    if (file.size < 1000) {
      errors.push('Word file appears to be empty or corrupted');
    }

    const sanitizedFilename = this.sanitizeFilename(file.originalname);

    return {
      isValid: errors.length === 0,
      errors,
      sanitizedFilename
    };
  }

  /**
   * Validate Excel file
   */
  validateExcelFile(file: Express.Multer.File): FileValidationResult {
    const errors: string[] = [];
    
    if (!file || !file.buffer) {
      errors.push('No file provided');
      return { isValid: false, errors, sanitizedFilename: '' };
    }

    // Check file size
    if (file.size > this.maxFileSize) {
      errors.push(`Excel file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
    }

    // Check MIME type
    if (!this.allowedMimeTypes.excel.includes(file.mimetype)) {
      errors.push(`Invalid MIME type: ${file.mimetype}. Expected one of: ${this.allowedMimeTypes.excel.join(', ')}`);
    }

    // Check file extension
    const ext = path.extname(file.originalname).toLowerCase();
    if (!this.allowedExtensions.excel.includes(ext)) {
      errors.push(`Invalid file extension: ${ext}. Expected one of: ${this.allowedExtensions.excel.join(', ')}`);
    }

    // Validate file header
    if (file.buffer && file.buffer.length >= 8) {
      const isValidExcelFile = this.validateOfficeFileHeader(file.buffer, ext);
      if (!isValidExcelFile) {
        errors.push(`Invalid ${ext} file format - file header does not match expected format`);
      }
    }

    // Check for minimum file size
    if (file.size < 1000) {
      errors.push('Excel file appears to be empty or corrupted');
    }

    const sanitizedFilename = this.sanitizeFilename(file.originalname);

    return {
      isValid: errors.length === 0,
      errors,
      sanitizedFilename
    };
  }

  /**
   * Validate PowerPoint file
   */
  validatePowerPointFile(file: Express.Multer.File): FileValidationResult {
    const errors: string[] = [];
    
    if (!file || !file.buffer) {
      errors.push('No file provided');
      return { isValid: false, errors, sanitizedFilename: '' };
    }

    // Check file size
    if (file.size > this.maxFileSize) {
      errors.push(`PowerPoint file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
    }

    // Check MIME type
    if (!this.allowedMimeTypes.powerpoint.includes(file.mimetype)) {
      errors.push(`Invalid MIME type: ${file.mimetype}. Expected one of: ${this.allowedMimeTypes.powerpoint.join(', ')}`);
    }

    // Check file extension
    const ext = path.extname(file.originalname).toLowerCase();
    if (!this.allowedExtensions.powerpoint.includes(ext)) {
      errors.push(`Invalid file extension: ${ext}. Expected one of: ${this.allowedExtensions.powerpoint.join(', ')}`);
    }

    // Validate file header
    if (file.buffer && file.buffer.length >= 8) {
      const isValidPptFile = this.validateOfficeFileHeader(file.buffer, ext);
      if (!isValidPptFile) {
        errors.push(`Invalid ${ext} file format - file header does not match expected format`);
      }
    }

    // Check for minimum file size
    if (file.size < 1000) {
      errors.push('PowerPoint file appears to be empty or corrupted');
    }

    const sanitizedFilename = this.sanitizeFilename(file.originalname);

    return {
      isValid: errors.length === 0,
      errors,
      sanitizedFilename
    };
  }

  /**
   * Sanitize filename to prevent path traversal and other security issues
   */
  private sanitizeFilename(filename: string): string {
    if (!filename) {
      return 'untitled_file';
    }

    // Remove path components
    const basename = path.basename(filename);
    
    // Remove or replace dangerous characters
    const sanitized = basename
      .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') // Replace dangerous chars with underscore
      .replace(/^\.+/, '') // Remove leading dots
      .replace(/\s+/g, '_') // Replace spaces with underscores
      .replace(/[^\w\-_.]/g, '_') // Keep only alphanumeric, dash, underscore, dot
      .substring(0, 255); // Limit length

    // Ensure filename is not empty after sanitization
    if (!sanitized || sanitized === '.' || sanitized === '..') {
      return 'sanitized_file';
    }

    return sanitized;
  }

  /**
   * Validate Office file headers (magic bytes)
   */
  private validateOfficeFileHeader(buffer: Buffer, extension: string): boolean {
    const header = buffer.slice(0, 8);
    
    switch (extension) {
      case '.docx':
      case '.xlsx':
      case '.pptx':
        // Modern Office files are ZIP-based (start with PK)
        return header[0] === 0x50 && header[1] === 0x4B;
      
      case '.doc':
      case '.xls':
      case '.ppt':
        // Legacy Office documents (OLE format)
        return header.slice(0, 4).toString('hex') === 'd0cf11e0';
      
      default:
        return false;
    }
  }

  /**
   * Check if file content appears to be malicious
   */
  private validateFileContent(file: Express.Multer.File): boolean {
    if (!file.buffer) {
      return false;
    }

    // Check for suspicious patterns in file content
    const content = file.buffer.toString('utf8', 0, Math.min(file.buffer.length, 1024));
    
    // Look for script tags, executable signatures, etc.
    const suspiciousPatterns = [
      /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
      /javascript:/gi,
      /vbscript:/gi,
      /on\w+\s*=/gi, // Event handlers
      /MZ/, // DOS/Windows executable header
      /\x7fELF/, // Linux executable header
    ];

    for (const pattern of suspiciousPatterns) {
      if (pattern.test(content)) {
        return false;
      }
    }

    return true;
  }

  /**
   * Comprehensive file validation based on type
   */
  validateFile(file: Express.Multer.File, expectedType: 'pdf' | 'word' | 'excel' | 'powerpoint'): FileValidationResult {
    // Basic security check
    if (!this.validateFileContent(file)) {
      return {
        isValid: false,
        errors: ['File contains potentially malicious content'],
        sanitizedFilename: this.sanitizeFilename(file.originalname)
      };
    }

    // Type-specific validation
    switch (expectedType) {
      case 'pdf':
        return this.validatePdfFile(file);
      case 'word':
        return this.validateWordFile(file);
      case 'excel':
        return this.validateExcelFile(file);
      case 'powerpoint':
        return this.validatePowerPointFile(file);
      default:
        return {
          isValid: false,
          errors: [`Unsupported file type: ${expectedType}`],
          sanitizedFilename: this.sanitizeFilename(file.originalname)
        };
    }
  }

  /**
   * Validate and sanitize compression quality parameter
   */
  validateCompressionQuality(quality: string): string {
    const allowedQualities = ['low', 'moderate', 'high'];
    const sanitizedQuality = quality?.toLowerCase()?.trim();
    
    if (!sanitizedQuality || !allowedQualities.includes(sanitizedQuality)) {
      return 'moderate'; // Default to moderate if invalid
    }
    
    return sanitizedQuality;
  }
}
