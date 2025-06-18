import { File } from 'multer';
export declare class AppService {
    private readonly execAsync;
    convertLibreOffice(file: File, format: string): Promise<Buffer>;
    compressPdf(file: File): Promise<Buffer>;
}
