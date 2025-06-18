import { Express } from 'express';
export declare class AppService {
    convertLibreOffice(file: Express.Multer.File, format: string): Promise<Buffer>;
    compressPdf(file: Express.Multer.File): Promise<Buffer>;
}
