"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.FileValidationService = void 0;
const common_1 = require("@nestjs/common");
const path = require("path");
let FileValidationService = class FileValidationService {
    constructor() {
        this.allowedMimeTypes = {
            pdf: ['application/pdf'],
            word: [
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'application/msword',
                'application/vnd.ms-word',
                'text/plain'
            ],
            excel: [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/vnd.ms-excel',
                'application/excel',
                'application/x-excel',
                'text/csv'
            ],
            powerpoint: [
                'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'application/vnd.ms-powerpoint',
                'application/mspowerpoint'
            ]
        };
        this.allowedExtensions = {
            pdf: ['.pdf'],
            word: ['.docx', '.doc', '.txt'],
            excel: ['.xlsx', '.xls', '.csv'],
            powerpoint: ['.pptx', '.ppt']
        };
        this.maxFileSize = 50 * 1024 * 1024;
    }
    validatePdfFile(file) {
        const errors = [];
        if (!file || !file.buffer) {
            errors.push('No file provided');
            return { isValid: false, errors, sanitizedFilename: '' };
        }
        if (file.size > this.maxFileSize) {
            errors.push(`PDF file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
        }
        if (!this.allowedMimeTypes.pdf.includes(file.mimetype)) {
            errors.push(`Invalid MIME type: ${file.mimetype}. Expected: ${this.allowedMimeTypes.pdf.join(', ')}`);
        }
        const ext = path.extname(file.originalname).toLowerCase();
        if (!this.allowedExtensions.pdf.includes(ext)) {
            errors.push(`Invalid file extension: ${ext}. Expected: ${this.allowedExtensions.pdf.join(', ')}`);
        }
        if (file.buffer && file.buffer.length >= 4) {
            const header = file.buffer.slice(0, 4).toString('ascii');
            if (header !== '%PDF') {
                errors.push('Invalid PDF file format - missing PDF header');
            }
        }
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
    validateWordFile(file) {
        const errors = [];
        if (!file || !file.buffer) {
            errors.push('No file provided');
            return { isValid: false, errors, sanitizedFilename: '' };
        }
        if (file.size > this.maxFileSize) {
            errors.push(`Word file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
        }
        if (!this.allowedMimeTypes.word.includes(file.mimetype)) {
            errors.push(`Invalid MIME type: ${file.mimetype}. Expected one of: ${this.allowedMimeTypes.word.join(', ')}`);
        }
        const ext = path.extname(file.originalname).toLowerCase();
        if (!this.allowedExtensions.word.includes(ext)) {
            errors.push(`Invalid file extension: ${ext}. Expected one of: ${this.allowedExtensions.word.join(', ')}`);
        }
        if (file.buffer && file.buffer.length >= 8) {
            const isValidWordFile = this.validateOfficeFileHeader(file.buffer, ext);
            if (!isValidWordFile) {
                errors.push(`Invalid ${ext} file format - file header does not match expected format`);
            }
        }
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
    validateExcelFile(file) {
        const errors = [];
        if (!file || !file.buffer) {
            errors.push('No file provided');
            return { isValid: false, errors, sanitizedFilename: '' };
        }
        if (file.size > this.maxFileSize) {
            errors.push(`Excel file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
        }
        if (!this.allowedMimeTypes.excel.includes(file.mimetype)) {
            errors.push(`Invalid MIME type: ${file.mimetype}. Expected one of: ${this.allowedMimeTypes.excel.join(', ')}`);
        }
        const ext = path.extname(file.originalname).toLowerCase();
        if (!this.allowedExtensions.excel.includes(ext)) {
            errors.push(`Invalid file extension: ${ext}. Expected one of: ${this.allowedExtensions.excel.join(', ')}`);
        }
        if (file.buffer && file.buffer.length >= 8) {
            const isValidExcelFile = this.validateOfficeFileHeader(file.buffer, ext);
            if (!isValidExcelFile) {
                errors.push(`Invalid ${ext} file format - file header does not match expected format`);
            }
        }
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
    validatePowerPointFile(file) {
        const errors = [];
        if (!file || !file.buffer) {
            errors.push('No file provided');
            return { isValid: false, errors, sanitizedFilename: '' };
        }
        if (file.size > this.maxFileSize) {
            errors.push(`PowerPoint file size ${(file.size / (1024 * 1024)).toFixed(2)}MB exceeds maximum limit of ${this.maxFileSize / (1024 * 1024)}MB`);
        }
        if (!this.allowedMimeTypes.powerpoint.includes(file.mimetype)) {
            errors.push(`Invalid MIME type: ${file.mimetype}. Expected one of: ${this.allowedMimeTypes.powerpoint.join(', ')}`);
        }
        const ext = path.extname(file.originalname).toLowerCase();
        if (!this.allowedExtensions.powerpoint.includes(ext)) {
            errors.push(`Invalid file extension: ${ext}. Expected one of: ${this.allowedExtensions.powerpoint.join(', ')}`);
        }
        if (file.buffer && file.buffer.length >= 8) {
            const isValidPptFile = this.validateOfficeFileHeader(file.buffer, ext);
            if (!isValidPptFile) {
                errors.push(`Invalid ${ext} file format - file header does not match expected format`);
            }
        }
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
    sanitizeFilename(filename) {
        if (!filename) {
            return 'untitled_file';
        }
        const basename = path.basename(filename);
        const sanitized = basename
            .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
            .replace(/^\.+/, '')
            .replace(/\s+/g, '_')
            .replace(/[^\w\-_.]/g, '_')
            .substring(0, 255);
        if (!sanitized || sanitized === '.' || sanitized === '..') {
            return 'sanitized_file';
        }
        return sanitized;
    }
    validateOfficeFileHeader(buffer, extension) {
        const header = buffer.slice(0, 8);
        switch (extension) {
            case '.docx':
            case '.xlsx':
            case '.pptx':
                return header[0] === 0x50 && header[1] === 0x4B;
            case '.doc':
            case '.xls':
            case '.ppt':
                return header.slice(0, 4).toString('hex') === 'd0cf11e0';
            default:
                return false;
        }
    }
    validateFileContent(file) {
        if (!file.buffer) {
            return false;
        }
        const content = file.buffer.toString('utf8', 0, Math.min(file.buffer.length, 1024));
        const suspiciousPatterns = [
            /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi,
            /javascript:/gi,
            /vbscript:/gi,
            /on\w+\s*=/gi,
            /MZ/,
            /\x7fELF/,
        ];
        for (const pattern of suspiciousPatterns) {
            if (pattern.test(content)) {
                return false;
            }
        }
        return true;
    }
    validateFile(file, expectedType) {
        if (!this.validateFileContent(file)) {
            return {
                isValid: false,
                errors: ['File contains potentially malicious content'],
                sanitizedFilename: this.sanitizeFilename(file.originalname)
            };
        }
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
    validateCompressionQuality(quality) {
        var _a;
        const allowedQualities = ['low', 'moderate', 'high'];
        const sanitizedQuality = (_a = quality === null || quality === void 0 ? void 0 : quality.toLowerCase()) === null || _a === void 0 ? void 0 : _a.trim();
        if (!sanitizedQuality || !allowedQualities.includes(sanitizedQuality)) {
            return 'moderate';
        }
        return sanitizedQuality;
    }
};
exports.FileValidationService = FileValidationService;
exports.FileValidationService = FileValidationService = __decorate([
    (0, common_1.Injectable)()
], FileValidationService);
//# sourceMappingURL=file-validation.service.js.map