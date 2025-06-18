import { Response } from 'express';
import { AppService } from './app.service';
export declare class AppController {
    private readonly appService;
    constructor(appService: AppService);
    convertDocToPdf(file: Express.Multer.File, res: Response): Promise<void>;
    compressPdf(file: Express.Multer.File, res: Response): Promise<void>;
}
