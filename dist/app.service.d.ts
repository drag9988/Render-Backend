import { OnlyOfficeService } from './onlyoffice.service';
import { OnlyOfficeEnhancedService } from './onlyoffice-enhanced.service';
import { FileValidationService } from './file-validation.service';
export declare class AppService {
    private readonly onlyOfficeService;
    private readonly onlyOfficeEnhancedService;
    private readonly fileValidationService;
    private readonly logger;
    private readonly execAsync;
    constructor(onlyOfficeService: OnlyOfficeService, onlyOfficeEnhancedService: OnlyOfficeEnhancedService, fileValidationService: FileValidationService);
    convertOfficeToPdf(file: Express.Multer.File): Promise<Buffer>;
    convertPdfToOffice(file: Express.Multer.File, format: string): Promise<Buffer>;
    convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer>;
    private convertPdfToOfficeFormat;
    private executeLibreOfficeConversion;
    private convertPdfWithLibreOffice;
    private executeEnhancedExcelToPdfConversion;
    private analyzePdf;
    private validateConvertedFile;
    private convertPdfToWordAlternative;
    analyzePdfFile(pdfPath: string): Promise<{
        isScanned: boolean;
        hasComplexLayout: boolean;
        isProtected: boolean;
        pageCount: number;
    }>;
    compressPdf(file: Express.Multer.File, quality?: string): Promise<Buffer>;
    private getCompressionCommands;
    addPasswordToPdf(file: Express.Multer.File, password: string): Promise<Buffer>;
    getConvertApiStatus(): Promise<{
        available: boolean;
        healthy?: boolean;
    }>;
    getOnlyOfficeStatus(): Promise<{
        available: boolean;
        healthy?: boolean;
    }>;
    getEnhancedOnlyOfficeStatus(): Promise<{
        available: boolean;
        healthy?: boolean;
        serverInfo?: any;
        capabilities?: any;
    }>;
}
