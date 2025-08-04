import { FileValidationService } from './file-validation.service';
export declare class OnlyOfficeEnhancedService {
    private readonly fileValidationService;
    private readonly logger;
    private readonly documentServerUrl;
    private readonly timeout;
    private readonly jwtSecret?;
    private readonly pythonPath;
    constructor(fileValidationService: FileValidationService);
    isAvailable(): boolean;
    convertPdfToDocx(pdfBuffer: Buffer, filename: string): Promise<Buffer>;
    convertPdfToXlsx(pdfBuffer: Buffer, filename: string): Promise<Buffer>;
    convertPdfToPptx(pdfBuffer: Buffer, filename: string): Promise<Buffer>;
    private convertPdf;
    private convertViaOnlyOfficeServer;
    private convertViaPython;
    private generatePythonScript;
    private convertViaEnhancedLibreOffice;
    private generateConversionKey;
    private generateJWT;
    healthCheck(): Promise<boolean>;
    getServerInfo(): Promise<any>;
}
