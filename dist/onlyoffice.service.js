"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var OnlyOfficeService_1;
Object.defineProperty(exports, "__esModule", { value: true });
exports.OnlyOfficeService = void 0;
const common_1 = require("@nestjs/common");
const axios_1 = require("axios");
const FormData = require("form-data");
const fs = require("fs/promises");
const path = require("path");
const file_validation_service_1 = require("./file-validation.service");
let OnlyOfficeService = OnlyOfficeService_1 = class OnlyOfficeService {
    constructor(fileValidationService) {
        this.fileValidationService = fileValidationService;
        this.logger = new common_1.Logger(OnlyOfficeService_1.name);
        this.documentServerUrl = process.env.ONLYOFFICE_DOCUMENT_SERVER_URL || '';
        this.timeout = parseInt(process.env.ONLYOFFICE_TIMEOUT || '120000', 10);
        this.jwtSecret = process.env.ONLYOFFICE_JWT_SECRET;
        if (!this.documentServerUrl) {
            this.logger.warn('ONLYOFFICE_DOCUMENT_SERVER_URL not found. ONLYOFFICE integration will be disabled.');
        }
        else {
            this.logger.log(`ONLYOFFICE Document Server configured at: ${this.documentServerUrl}`);
        }
    }
    isAvailable() {
        return !!this.documentServerUrl;
    }
    async convertPdfToDocx(pdfBuffer, filename) {
        return this.convertPdf(pdfBuffer, filename, 'docx');
    }
    async convertPdfToXlsx(pdfBuffer, filename) {
        return this.convertPdf(pdfBuffer, filename, 'xlsx');
    }
    async convertPdfToPptx(pdfBuffer, filename) {
        return this.convertPdf(pdfBuffer, filename, 'pptx');
    }
    async convertPdf(pdfBuffer, filename, targetFormat) {
        if (!this.isAvailable()) {
            throw new Error('ONLYOFFICE Document Server is not available. Please set ONLYOFFICE_DOCUMENT_SERVER_URL environment variable.');
        }
        if (!pdfBuffer || pdfBuffer.length === 0) {
            throw new common_1.BadRequestException('Invalid or empty PDF buffer provided');
        }
        const file = {
            fieldname: 'file',
            originalname: filename,
            encoding: '7bit',
            mimetype: 'application/pdf',
            buffer: pdfBuffer,
            size: pdfBuffer.length,
            stream: null,
            destination: '',
            filename: '',
            path: ''
        };
        const validation = this.fileValidationService.validateFile(file, 'pdf');
        if (!validation.isValid) {
            throw new common_1.BadRequestException(`PDF validation failed: ${validation.errors.join(', ')}`);
        }
        const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
        const timestamp = Date.now();
        const tempInputPath = path.join(tempDir, `${timestamp}_input.pdf`);
        const tempOutputPath = path.join(tempDir, `${timestamp}_output.${targetFormat}`);
        try {
            this.logger.log(`Starting PDF to ${targetFormat.toUpperCase()} conversion using ONLYOFFICE for file: ${filename}`);
            await fs.writeFile(tempInputPath, pdfBuffer);
            this.logger.log(`PDF written to temporary file: ${tempInputPath}`);
            try {
                const convertedBuffer = await this.convertViaDirectAPI(tempInputPath, targetFormat, validation.sanitizedFilename);
                if (convertedBuffer) {
                    this.logger.log(`Successfully converted PDF to ${targetFormat.toUpperCase()} via direct API. Output size: ${convertedBuffer.length} bytes`);
                    return convertedBuffer;
                }
            }
            catch (directApiError) {
                this.logger.warn(`Direct API conversion failed: ${directApiError.message}. Trying LibreOffice with ONLYOFFICE integration.`);
            }
            const convertedBuffer = await this.convertViaEnhancedLibreOffice(tempInputPath, tempOutputPath, targetFormat);
            this.logger.log(`Successfully converted PDF to ${targetFormat.toUpperCase()} using enhanced LibreOffice. Output size: ${convertedBuffer.length} bytes`);
            return convertedBuffer;
        }
        catch (error) {
            this.logger.error(`ONLYOFFICE conversion failed for ${filename} to ${targetFormat}: ${error.message}`);
            throw new Error(`ONLYOFFICE conversion failed: ${error.message}`);
        }
        finally {
            try {
                await fs.unlink(tempInputPath).catch(() => { });
                await fs.unlink(tempOutputPath).catch(() => { });
            }
            catch (cleanupError) {
                this.logger.warn(`Cleanup warning: ${cleanupError.message}`);
            }
        }
    }
    async convertViaDirectAPI(inputPath, targetFormat, filename) {
        try {
            const uploadUrl = await this.uploadFileToOnlyOffice(inputPath, filename);
            const conversionRequest = {
                async: false,
                filetype: 'pdf',
                key: this.generateConversionKey(filename),
                outputtype: targetFormat,
                title: filename,
                url: uploadUrl
            };
            if (this.jwtSecret) {
                conversionRequest['token'] = this.generateJWT(conversionRequest);
            }
            const conversionUrl = `${this.documentServerUrl}/ConvertService.ashx`;
            this.logger.log(`Sending conversion request to ONLYOFFICE: ${conversionUrl}`);
            const response = await axios_1.default.post(conversionUrl, conversionRequest, {
                timeout: this.timeout,
                headers: {
                    'Content-Type': 'application/json'
                }
            });
            if (response.data.error) {
                throw new Error(`ONLYOFFICE conversion error: ${response.data.error}`);
            }
            if (!response.data.fileUrl) {
                throw new Error('ONLYOFFICE returned no file URL');
            }
            const fileResponse = await axios_1.default.get(response.data.fileUrl, {
                responseType: 'arraybuffer',
                timeout: this.timeout
            });
            const convertedBuffer = Buffer.from(fileResponse.data);
            if (convertedBuffer.length === 0) {
                throw new Error('Downloaded file is empty');
            }
            return convertedBuffer;
        }
        catch (error) {
            this.logger.warn(`Direct API conversion failed: ${error.message}`);
            return null;
        }
    }
    async convertViaEnhancedLibreOffice(inputPath, outputPath, targetFormat) {
        const { exec } = require('child_process');
        const { promisify } = require('util');
        const execAsync = promisify(exec);
        const enhancedCommands = [];
        if (targetFormat === 'pptx') {
            enhancedCommands.push(`libreoffice --headless --impress --convert-to pptx:"Impress MS PowerPoint 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`, `libreoffice --headless --draw --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`, `libreoffice --headless --writer --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`, `libreoffice --headless --convert-to pptx --outdir "${path.dirname(outputPath)}" "${inputPath}"`);
        }
        else {
            enhancedCommands.push(`libreoffice --headless --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`, targetFormat === 'docx'
                ? `libreoffice --headless --writer --convert-to ${targetFormat}:"MS Word 2007 XML" --outdir "${path.dirname(outputPath)}" "${inputPath}"`
                : `libreoffice --headless --calc --convert-to ${targetFormat} --outdir "${path.dirname(outputPath)}" "${inputPath}"`);
        }
        let lastError = '';
        for (let i = 0; i < enhancedCommands.length; i++) {
            const command = enhancedCommands[i];
            this.logger.log(`Attempting enhanced LibreOffice conversion ${i + 1}/${enhancedCommands.length} for ${targetFormat}`);
            try {
                const { stdout, stderr } = await execAsync(command, {
                    timeout: this.timeout - 10000,
                    maxBuffer: 1024 * 1024 * 10
                });
                if (stdout) {
                    this.logger.log(`LibreOffice output: ${stdout}`);
                }
                if (stderr && !stderr.includes('Warning')) {
                    this.logger.warn(`LibreOffice stderr: ${stderr}`);
                }
                const expectedOutputPath = path.join(path.dirname(outputPath), path.basename(inputPath, '.pdf') + '.' + targetFormat);
                try {
                    await fs.access(expectedOutputPath);
                    const result = await fs.readFile(expectedOutputPath);
                    if (result.length > 500) {
                        this.logger.log(`Enhanced LibreOffice conversion successful: ${result.length} bytes`);
                        await fs.unlink(expectedOutputPath).catch(() => { });
                        return result;
                    }
                    else {
                        throw new Error(`Generated file too small: ${result.length} bytes`);
                    }
                }
                catch (fileError) {
                    this.logger.warn(`Output file check failed: ${fileError.message}`);
                    lastError = fileError.message;
                }
            }
            catch (execError) {
                this.logger.warn(`LibreOffice execution failed: ${execError.message}`);
                lastError = execError.message;
            }
        }
        throw new Error(`Enhanced LibreOffice conversion failed after ${enhancedCommands.length} attempts. Last error: ${lastError}`);
    }
    async uploadFileToOnlyOffice(filePath, filename) {
        try {
            const fileBuffer = await fs.readFile(filePath);
            const formData = new FormData();
            formData.append('file', fileBuffer, filename);
            const uploadResponse = await axios_1.default.post(`${this.documentServerUrl}/upload`, formData, {
                headers: formData.getHeaders(),
                timeout: this.timeout
            });
            if (uploadResponse.data && uploadResponse.data.url) {
                return uploadResponse.data.url;
            }
        }
        catch (uploadError) {
            this.logger.warn(`ONLYOFFICE upload failed: ${uploadError.message}`);
        }
        const tempUrl = await this.createTempFileUrl(filePath, filename);
        return tempUrl;
    }
    async createTempFileUrl(filePath, filename) {
        const serverUrl = process.env.SERVER_URL || 'http://localhost:10000';
        const tempEndpoint = `${serverUrl}/temp/${Date.now()}_${filename}`;
        return tempEndpoint;
    }
    generateConversionKey(filename) {
        const timestamp = Date.now();
        const random = Math.random().toString(36).substring(2);
        return `${timestamp}_${random}_${filename.replace(/[^a-zA-Z0-9]/g, '_')}`;
    }
    generateJWT(payload) {
        if (!this.jwtSecret) {
            return '';
        }
        const jwt = require('jsonwebtoken');
        return jwt.sign(payload, this.jwtSecret, { expiresIn: '1h' });
    }
    async healthCheck() {
        if (!this.isAvailable()) {
            return false;
        }
        try {
            const response = await axios_1.default.get(`${this.documentServerUrl}/healthcheck`, {
                timeout: 5000
            });
            return response.status === 200;
        }
        catch (error) {
            try {
                const response = await axios_1.default.get(`${this.documentServerUrl}/`, {
                    timeout: 5000
                });
                return response.status === 200;
            }
            catch (altError) {
                this.logger.error(`ONLYOFFICE health check failed: ${error.message}`);
                return false;
            }
        }
    }
    async getServerInfo() {
        if (!this.isAvailable()) {
            return { available: false, reason: 'Document server URL not configured' };
        }
        try {
            const healthy = await this.healthCheck();
            return {
                available: true,
                healthy,
                url: this.documentServerUrl,
                jwtEnabled: !!this.jwtSecret
            };
        }
        catch (error) {
            return {
                available: false,
                reason: error.message
            };
        }
    }
};
exports.OnlyOfficeService = OnlyOfficeService;
exports.OnlyOfficeService = OnlyOfficeService = OnlyOfficeService_1 = __decorate([
    (0, common_1.Injectable)(),
    __metadata("design:paramtypes", [file_validation_service_1.FileValidationService])
], OnlyOfficeService);
//# sourceMappingURL=onlyoffice.service.js.map