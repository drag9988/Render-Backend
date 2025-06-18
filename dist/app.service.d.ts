export declare class AppService {
    convertLibreOffice(file: Multer.File, format: string): Promise<Buffer>;
    compressPdf(file: Multer.File): Promise<Buffer>;
}
