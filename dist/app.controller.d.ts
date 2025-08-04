import { Response } from 'express';
import { AppService } from './app.service';
import { FileValidationService } from './file-validation.service';
export declare class AppController {
    private readonly appService;
    private readonly fileValidationService;
    constructor(appService: AppService, fileValidationService: FileValidationService);
    healthCheck(res: Response): Response<any, Record<string, any>>;
    health(res: Response): Response<any, Record<string, any>>;
    corsTest(res: Response): Response<any, Record<string, any>>;
    corsTestPost(res: Response, body: any): Response<any, Record<string, any>>;
    convertWordToPdf(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    convertExcelToPdf(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    convertPptToPdf(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    convertPdfToWord(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    convertPdfToExcel(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    convertPdfToPpt(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    analyzePdf(file: Express.Multer.File, res: Response): Promise<Response<any, Record<string, any>>>;
    private getConversionRecommendations;
    compressPdf(file: Express.Multer.File, quality: string, res: Response): Promise<Response<any, Record<string, any>>>;
    getConvertApiStatus(res: Response): Promise<Response<any, Record<string, any>>>;
    getOnlyOfficeStatus(res: Response): Promise<Response<any, Record<string, any>>>;
    getEnhancedOnlyOfficeStatus(res: Response): Promise<Response<any, Record<string, any>>>;
    addPasswordToPdf(file: Express.Multer.File, password: string, res: Response): Promise<Response<any, Record<string, any>>>;
}
