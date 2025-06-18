import { Response } from 'express';
export declare class AppController {
    private readonly appService;
    constructor(appService: AppService);
    convertDocToPdf(file: Multer.File, res: Response): Promise<void>;
    compressPdf(file: Multer.File, res: Response): Promise<void>;
}
