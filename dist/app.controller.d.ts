import { Response } from 'express';
import { AppService } from './app.service';
import { File } from 'multer';
export declare class AppController {
    private readonly appService;
    constructor(appService: AppService);
    convertDocToPdf(file: File, res: Response): Promise<void>;
    compressPdf(file: File, res: Response): Promise<void>;
}
