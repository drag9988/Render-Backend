import { Injectable, BadRequestException, Logger } from '@nestjs/common';
import * as path from 'path';
import * as crypto from 'crypto';

export interface FileValidationResult {
  isValid: boolean;
  sanitizedFilename: string;
  errors: string[];
  warnings: string[];
}

export interface SecurityConfig {
  maxFileSize: number;
  allowedMimeTypes: string[];
  allowedExtensions: string[];
  requireExtensionMatch: boolean;
}

@Injectable()
export class SecurityService {
  private readonly logger = new Logger(SecurityService.name);

  // Security configurations for different file types
  private readonly securityConfigs: Record<string, SecurityConfig> = {
    pdf: {
      maxFileSize: 50 * 1024 * 1024, // 50MB
      allowedMimeTypes: ['application/pdf'],
      allowedExtensions: ['.pdf'],
      requireExtensionMatch: true
    },
    office: {
      maxFileSize: 50 * 1024 * 1024, // 50MB
      allowedMimeTypes: [
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
        'application/msword', // .doc
        'application/vnd.ms-excel', // .xls
        'application/vnd.ms-powerpoint', // .ppt
        'text/plain', // .txt
        'application/rtf' // .rtf
      ],
      allowedExtensions: ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt', '.txt', '.rtf'],
      requireExtensionMatch: true
    }
  };

  /**
   * Sanitize filename to prevent path traversal and injection attacks
   */
  sanitizeFilename(filename: string): string {
    if (!filename || typeof filename !== 'string') {
      return 'untitled_file';
    }

    // Remove or replace dangerous characters
    let sanitized = filename
      .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') // Replace dangerous characters
      .replace(/\.\./g, '_') // Remove path traversal attempts
      .replace(/^\.+/, '') // Remove leading dots
      .replace(/\.+$/, '') // Remove trailing dots
      .replace(/\s+/g, '_') // Replace spaces with underscores
      .replace(/_+/g, '_') // Replace multiple underscores with single
      .substring(0, 255); // Limit length

    // Ensure filename has proper extension
    if (!sanitized.includes('.')) {
      sanitized += '.unknown';
    }

    // If filename becomes empty or just extension, generate a random name
    const nameWithoutExt = path.parse(sanitized).name;
    if (!nameWithoutExt || nameWithoutExt.length < 1) {
      const randomName = crypto.randomBytes(8).toString('hex');
      const ext = path.extname(sanitized) || '.unknown';
      sanitized = `file_${randomName}${ext}`;
    }

    return sanitized;
  }

  /**
   * Validate file against security rules
   */
  validateFile(file: Express.Multer.File, fileType: 'pdf' | 'office'): FileValidationResult {
    const result: FileValidationResult = {
      isValid: true,
      sanitizedFilename: this.sanitizeFilename(file.originalname),
      errors: [],
      warnings: []
    };

    const config = this.securityConfigs[fileType];
    
    // 1. File size validation
    if (file.size > config.maxFileSize) {
      result.isValid = false;
      result.errors.push(`File size ${Math.round(file.size / 1024 / 1024)}MB exceeds maximum allowed size of ${Math.round(config.maxFileSize / 1024 / 1024)}MB`);
    }

    // 2. MIME type validation
    if (!config.allowedMimeTypes.includes(file.mimetype)) {
      result.isValid = false;
      result.errors.push(`MIME type '${file.mimetype}' is not allowed. Allowed types: ${config.allowedMimeTypes.join(', ')}`);
    }

    // 3. File extension validation
    const fileExtension = path.extname(file.originalname).toLowerCase();
    if (!config.allowedExtensions.includes(fileExtension)) {
      result.isValid = false;
      result.errors.push(`File extension '${fileExtension}' is not allowed. Allowed extensions: ${config.allowedExtensions.join(', ')}`);
    }

    // 4. MIME type and extension consistency check
    if (config.requireExtensionMatch && result.isValid) {
      const mimeExtensionMatch = this.validateMimeExtensionConsistency(file.mimetype, fileExtension);
      if (!mimeExtensionMatch.isValid) {
        result.warnings.push(mimeExtensionMatch.message);
        // Don't fail validation, but log warning
        this.logger.warn(`MIME/Extension mismatch for file ${file.originalname}: ${mimeExtensionMatch.message}`);
      }
    }

    // 5. Buffer validation (basic checks)
    if (!file.buffer || file.buffer.length === 0) {
      result.isValid = false;
      result.errors.push('File buffer is empty or corrupted');
    }

    // 6. File signature validation (magic bytes)
    if (file.buffer && file.buffer.length > 0) {
      const signatureCheck = this.validateFileSignature(file.buffer, file.mimetype, fileExtension);
      if (!signatureCheck.isValid) {
        result.isValid = false;
        result.errors.push(signatureCheck.message);
      }
    }

    this.logger.log(`File validation for ${file.originalname}: ${result.isValid ? 'PASSED' : 'FAILED'}. Errors: ${result.errors.length}, Warnings: ${result.warnings.length}`);

    return result;
  }

  /**
   * Validate MIME type and file extension consistency
   */
  private validateMimeExtensionConsistency(mimeType: string, extension: string): { isValid: boolean; message: string } {
    const mimeExtensionMap: Record<string, string[]> = {
      'application/pdf': ['.pdf'],
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
      'application/msword': ['.doc'],
      'application/vnd.ms-excel': ['.xls'],
      'application/vnd.ms-powerpoint': ['.ppt'],
      'text/plain': ['.txt'],
      'application/rtf': ['.rtf']
    };

    const expectedExtensions = mimeExtensionMap[mimeType];
    if (expectedExtensions && !expectedExtensions.includes(extension)) {
      return {
        isValid: false,
        message: `MIME type '${mimeType}' does not match file extension '${extension}'. Expected: ${expectedExtensions.join(', ')}`
      };
    }

    return { isValid: true, message: 'MIME type and extension are consistent' };
  }

  /**
   * Validate file signature (magic bytes) to detect file type spoofing
   */
  private validateFileSignature(buffer: Buffer, mimeType: string, extension: string): { isValid: boolean; message: string } {
    if (buffer.length < 4) {
      return { isValid: false, message: 'File is too small to validate signature' };
    }

    const signature = buffer.toString('hex', 0, 8).toUpperCase();
    
    // PDF signature validation
    if (mimeType === 'application/pdf' || extension === '.pdf') {
      if (!buffer.toString('ascii', 0, 4).startsWith('%PDF')) {
        return { isValid: false, message: 'File does not have a valid PDF signature' };
      }
    }

    // Office documents (DOCX, XLSX, PPTX) - ZIP-based files
    if (mimeType.includes('openxmlformats') || ['.docx', '.xlsx', '.pptx'].includes(extension)) {
      // These files start with ZIP signature: PK (0x504B)
      if (!signature.startsWith('504B')) {
        return { isValid: false, message: 'Office document does not have a valid ZIP signature' };
      }
    }

    // Legacy Office documents
    if (mimeType.includes('msword') || mimeType.includes('ms-excel') || mimeType.includes('ms-powerpoint') || 
        ['.doc', '.xls', '.ppt'].includes(extension)) {
      // OLE2 signature: D0CF11E0A1B11AE1
      if (!signature.startsWith('D0CF11E0')) {
        return { isValid: false, message: 'Legacy Office document does not have a valid OLE2 signature' };
      }
    }

    return { isValid: true, message: 'File signature is valid' };
  }

  /**
   * Sanitize string input to prevent injection attacks
   */
  sanitizeString(input: string, maxLength: number = 255): string {
    if (!input || typeof input !== 'string') {
      return '';
    }

    return input
      .replace(/[<>'"&]/g, '') // Remove HTML/XML special characters
      .replace(/[;|&$`\\]/g, '') // Remove shell injection characters
      .replace(/\0/g, '') // Remove null bytes
      .trim()
      .substring(0, maxLength);
  }

  /**
   * Validate and sanitize quality parameter for PDF compression
   */
  validateQualityParameter(quality: string): string {
    const allowedQualities = ['low', 'moderate', 'high'];
    const sanitizedQuality = this.sanitizeString(quality, 20).toLowerCase();
    
    if (!allowedQualities.includes(sanitizedQuality)) {
      this.logger.warn(`Invalid quality parameter received: ${quality}. Using default 'moderate'.`);
      return 'moderate';
    }

    return sanitizedQuality;
  }

  /**
   * Generate secure temporary filename
   */
  generateSecureFilename(originalFilename: string, prefix: string = 'temp'): string {
    const sanitized = this.sanitizeFilename(originalFilename);
    const timestamp = Date.now();
    const random = crypto.randomBytes(8).toString('hex');
    const extension = path.extname(sanitized);
    const baseName = path.basename(sanitized, extension);
    
    return `${prefix}_${timestamp}_${random}_${baseName}${extension}`;
  }

  /**
   * Validate file buffer for potential malicious content
   */
  validateFileBuffer(buffer: Buffer): { isValid: boolean; message: string } {
    // Check for suspicious patterns
    const suspiciousPatterns = [
      /javascript:/gi,
      /<script/gi,
      /eval\(/gi,
      /document\.write/gi,
      /window\.location/gi,
      /\x00/g // Null bytes
    ];

    const bufferString = buffer.toString('ascii');
    
    for (const pattern of suspiciousPatterns) {
      if (pattern.test(bufferString)) {
        return {
          isValid: false,
          message: `Suspicious content detected in file: ${pattern.toString()}`
        };
      }
    }

    return { isValid: true, message: 'File buffer appears safe' };
  }

  /**
   * Rate limiting check for IP-based requests
   */
  validateRequestSource(req: any): { isValid: boolean; message: string } {
    // Get real IP address considering proxies
    const clientIp = req.headers['x-forwarded-for']?.split(',')[0] || 
                     req.headers['x-real-ip'] || 
                     req.connection?.remoteAddress || 
                     req.socket?.remoteAddress ||
                     req.ip;

    // Basic IP validation
    if (!clientIp || clientIp === '127.0.0.1' || clientIp === '::1') {
      return { isValid: true, message: 'Local request' };
    }

    // Check for private IP ranges (these are generally safe)
    const privateIPRegex = /^(10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)/;
    if (privateIPRegex.test(clientIp)) {
      return { isValid: true, message: 'Private IP range' };
    }

    return { isValid: true, message: `Request from IP: ${clientIp}` };
  }
}
