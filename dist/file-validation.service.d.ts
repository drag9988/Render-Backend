export interface FileValidationResult {
    isValid: boolean;
    errors: string[];
    sanitizedFilename: string;
}
export declare class FileValidationService {
    private readonly allowedMimeTypes;
    private readonly allowedExtensions;
    private readonly maxFileSize;
    validatePdfFile(file: Express.Multer.File): FileValidationResult;
    validateWordFile(file: Express.Multer.File): FileValidationResult;
    validateExcelFile(file: Express.Multer.File): FileValidationResult;
    validatePowerPointFile(file: Express.Multer.File): FileValidationResult;
    private sanitizeFilename;
    private validateOfficeFileHeader;
    private validateFileContent;
    validateFile(file: Express.Multer.File, expectedType: 'pdf' | 'word' | 'excel' | 'powerpoint'): FileValidationResult;
    validateCompressionQuality(quality: string): string;
}
