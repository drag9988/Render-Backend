import { FileValidationService } from './file-validation.service';
export interface OnlyOfficeConfig {
    documentServerUrl: string;
    timeout?: number;
    jwtSecret?: string;
}
export declare class OnlyOfficeService {
    private readonly fileValidationService;
    private readonly logger;
    private readonly documentServerUrl;
    private readonly timeout;
    private readonly jwtSecret?;
    constructor(fileValidationService: FileValidationService);
    isAvailable(): boolean;
    convertPdfToDocx(pdfBuffer: Buffer, filename: string): Promise<Buffer>;
    convertPdfToXlsx(pdfBuffer: Buffer, filename: string): Promise<Buffer>;
    convertPdfToPptx(pdfBuffer: Buffer, filename: string): Promise<Buffer>;
    private convertPdf;
    private convertViaDirectAPI;
    private convertViaEnhancedLibreOffice;
    private uploadFileToOnlyOffice;
    private createTempFileUrl;
    private generateConversionKey;
    private generateJWT;
    healthCheck(): Promise<boolean>;
    getServerInfo(): Promise<any>;
}
