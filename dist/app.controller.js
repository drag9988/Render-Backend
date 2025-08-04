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
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AppController = void 0;
const common_1 = require("@nestjs/common");
const platform_express_1 = require("@nestjs/platform-express");
const throttler_1 = require("@nestjs/throttler");
const app_service_1 = require("./app.service");
const file_validation_service_1 = require("./file-validation.service");
const os = require("os");
let AppController = class AppController {
    constructor(appService, fileValidationService) {
        this.appService = appService;
        this.fileValidationService = fileValidationService;
    }
    healthCheck(res) {
        try {
            const healthInfo = {
                status: 'ok',
                timestamp: new Date().toISOString(),
                uptime: process.uptime(),
                hostname: os.hostname(),
                network: Object.values(os.networkInterfaces())
                    .flat()
                    .filter(iface => iface && !iface.internal)
                    .map(iface => ({
                    address: iface.address,
                    family: iface.family,
                    netmask: iface.netmask,
                })),
                memory: {
                    total: os.totalmem(),
                    free: os.freemem(),
                },
                env: {
                    nodeEnv: process.env.NODE_ENV || 'development',
                    port: process.env.PORT || '3000',
                },
            };
            return res.status(200).json(healthInfo);
        }
        catch (error) {
            console.error('Health check error:', error.message);
            return res.status(500).json({ status: 'error', message: error.message });
        }
    }
    health(res) {
        return res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
    }
    corsTest(res) {
        res.header('Access-Control-Allow-Origin', '*');
        res.header('Access-Control-Allow-Methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
        res.header('Access-Control-Allow-Headers', 'Accept,Authorization,Content-Type,X-Requested-With,Range,Origin');
        return res.status(200).json({
            message: 'CORS test successful',
            timestamp: new Date().toISOString(),
            headers: {
                origin: res.req.get('Origin') || 'none',
                userAgent: res.req.get('User-Agent') || 'none',
                method: res.req.method
            }
        });
    }
    corsTestPost(res, body) {
        res.header('Access-Control-Allow-Origin', '*');
        res.header('Access-Control-Allow-Methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
        res.header('Access-Control-Allow-Headers', 'Accept,Authorization,Content-Type,X-Requested-With,Range,Origin');
        return res.status(200).json({
            message: 'CORS POST test successful',
            receivedBody: body,
            timestamp: new Date().toISOString()
        });
    }
    async convertWordToPdf(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            const validation = this.fileValidationService.validateFile(file, 'word');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'File validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            const output = await this.appService.convertOfficeToPdf(file);
            res.set({ 'Content-Type': 'application/pdf' });
            res.send(output);
        }
        catch (error) {
            console.error('Word to PDF conversion error:', error.message);
            if (error instanceof common_1.BadRequestException) {
                return res.status(400).json({ error: 'Invalid file', message: error.message });
            }
            res.status(500).json({ error: 'Failed to convert Word to PDF', message: error.message });
        }
    }
    async convertExcelToPdf(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            const validation = this.fileValidationService.validateFile(file, 'excel');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'File validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            const output = await this.appService.convertOfficeToPdf(file);
            res.set({ 'Content-Type': 'application/pdf' });
            res.send(output);
        }
        catch (error) {
            console.error('Excel to PDF conversion error:', error.message);
            if (error instanceof common_1.BadRequestException) {
                return res.status(400).json({ error: 'Invalid file', message: error.message });
            }
            res.status(500).json({ error: 'Failed to convert Excel to PDF', message: error.message });
        }
    }
    async convertPptToPdf(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            const validation = this.fileValidationService.validateFile(file, 'powerpoint');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'File validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            const output = await this.appService.convertOfficeToPdf(file);
            res.set({ 'Content-Type': 'application/pdf' });
            res.send(output);
        }
        catch (error) {
            console.error('PowerPoint to PDF conversion error:', error.message);
            if (error instanceof common_1.BadRequestException) {
                return res.status(400).json({ error: 'Invalid file', message: error.message });
            }
            res.status(500).json({ error: 'Failed to convert PowerPoint to PDF', message: error.message });
        }
    }
    async convertPdfToWord(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            const validation = this.fileValidationService.validateFile(file, 'pdf');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'PDF validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            console.log(`Converting PDF to Word: ${file.originalname}, size: ${file.size} bytes`);
            const output = await this.appService.convertPdfToOffice(file, 'docx');
            res.set({
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.docx')}"`
            });
            res.send(output);
        }
        catch (error) {
            console.error('PDF to Word conversion error:', error.message);
            if (error.message.includes('timeout')) {
                return res.status(408).json({
                    error: 'Conversion timeout',
                    message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.'
                });
            }
            else if (error.message.includes('complex formatting')) {
                return res.status(422).json({
                    error: 'Complex PDF format',
                    message: 'This PDF contains complex formatting that cannot be converted. Try using a text-based PDF instead of a scanned document.',
                    suggestions: [
                        'Ensure the PDF contains selectable text (not a scanned image)',
                        'Try with a simpler PDF with basic formatting',
                        'Consider using a PDF with fewer images and complex layouts'
                    ]
                });
            }
            else if (error.message.includes('scanned PDF')) {
                return res.status(422).json({
                    error: 'Scanned PDF detected',
                    message: 'This appears to be a scanned PDF (image-based) which cannot be converted to editable Word format.',
                    suggestions: [
                        'Use OCR software to convert the scanned PDF to text first',
                        'Try with a text-based PDF instead',
                        'Export the original document directly to Word if possible'
                    ]
                });
            }
            else if (error.message.includes('insufficient text')) {
                return res.status(422).json({
                    error: 'Insufficient text content',
                    message: 'This PDF does not contain enough text content for conversion.',
                    suggestions: [
                        'Verify the PDF contains readable text',
                        'Check if the PDF is password-protected',
                        'Try with a different PDF file'
                    ]
                });
            }
            else {
                return res.status(500).json({
                    error: 'Failed to convert PDF to Word',
                    message: error.message
                });
            }
        }
    }
    async convertPdfToExcel(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            const validation = this.fileValidationService.validateFile(file, 'pdf');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'PDF validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            console.log(`Converting PDF to Excel: ${file.originalname}, size: ${file.size} bytes`);
            const output = await this.appService.convertPdfToOffice(file, 'xlsx');
            res.set({
                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.xlsx')}"`
            });
            res.send(output);
        }
        catch (error) {
            console.error('PDF to Excel conversion error:', error.message);
            if (error.message.includes('timeout')) {
                return res.status(408).json({
                    error: 'Conversion timeout',
                    message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.'
                });
            }
            else if (error.message.includes('complex formatting')) {
                return res.status(422).json({
                    error: 'Complex PDF format',
                    message: 'This PDF may not contain tabular data suitable for Excel conversion.',
                    suggestions: [
                        'Ensure the PDF contains tables or structured data',
                        'Try with a PDF that has clear tabular layout',
                        'Consider manual copy-paste for complex data structures'
                    ]
                });
            }
            else if (error.message.includes('scanned PDF')) {
                return res.status(422).json({
                    error: 'Scanned PDF detected',
                    message: 'This appears to be a scanned PDF which cannot be converted to Excel format.',
                    suggestions: [
                        'Use OCR software first to make the PDF text-selectable',
                        'Try with a text-based PDF containing tables',
                        'Export data directly from the original source if possible'
                    ]
                });
            }
            else {
                return res.status(500).json({
                    error: 'Failed to convert PDF to Excel',
                    message: error.message
                });
            }
        }
    }
    async convertPdfToPpt(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            const validation = this.fileValidationService.validateFile(file, 'pdf');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'PDF validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            console.log(`Converting PDF to PowerPoint: ${file.originalname}, size: ${file.size} bytes`);
            const output = await this.appService.convertPdfToOffice(file, 'pptx');
            res.set({
                'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '.pptx')}"`
            });
            res.send(output);
        }
        catch (error) {
            console.error('PDF to PowerPoint conversion error:', error.message);
            if (error.message.includes('timeout')) {
                return res.status(408).json({
                    error: 'Conversion timeout',
                    message: 'The PDF conversion is taking too long. Please try with a smaller or simpler PDF.'
                });
            }
            else if (error.message.includes('complex formatting')) {
                return res.status(422).json({
                    error: 'Complex PDF format',
                    message: 'This PDF contains complex formatting that may not convert well to PowerPoint slides.',
                    suggestions: [
                        'Try with a PDF that has clear page-by-page content',
                        'Ensure the PDF has simple layout without complex graphics',
                        'Consider breaking down the PDF into smaller sections'
                    ]
                });
            }
            else if (error.message.includes('scanned PDF')) {
                return res.status(422).json({
                    error: 'Scanned PDF detected',
                    message: 'This appears to be a scanned PDF which cannot be converted to PowerPoint format.',
                    suggestions: [
                        'Use OCR software to make the PDF text-selectable first',
                        'Try with a text-based PDF',
                        'Export slides directly from the original presentation if possible'
                    ]
                });
            }
            else {
                return res.status(500).json({
                    error: 'Failed to convert PDF to PowerPoint',
                    message: error.message
                });
            }
        }
    }
    async analyzePdf(file, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No file uploaded' });
            }
            if (file.mimetype !== 'application/pdf') {
                return res.status(400).json({ error: 'Uploaded file is not a PDF' });
            }
            const timestamp = Date.now();
            const tempDir = process.env.TEMP_DIR || require('os').tmpdir() || '/tmp';
            try {
                await require('fs').promises.mkdir(tempDir, { recursive: true });
                await require('fs').promises.chmod(tempDir, 0o777);
            }
            catch (dirError) {
                console.error(`Failed to create temp directory ${tempDir}: ${dirError.message}`);
                return res.status(500).json({ error: 'Server configuration error', message: 'Cannot access temporary storage' });
            }
            const tempInput = `${tempDir}/${timestamp}_analysis.pdf`;
            await require('fs').promises.writeFile(tempInput, file.buffer);
            const analysis = await this.appService.analyzePdfFile(tempInput);
            await require('fs').promises.unlink(tempInput).catch(() => { });
            res.json({
                filename: file.originalname,
                size: file.size,
                analysis: analysis,
                recommendations: this.getConversionRecommendations(analysis)
            });
        }
        catch (error) {
            console.error('PDF analysis error:', error.message);
            res.status(500).json({ error: 'Failed to analyze PDF', message: error.message });
        }
    }
    getConversionRecommendations(analysis) {
        const recommendations = {
            canConvertToWord: true,
            canConvertToExcel: true,
            canConvertToPowerPoint: true,
            warnings: [],
            suggestions: []
        };
        if (analysis.isScanned) {
            recommendations.canConvertToWord = false;
            recommendations.canConvertToExcel = false;
            recommendations.canConvertToPowerPoint = false;
            recommendations.warnings.push('This appears to be a scanned PDF (image-based)');
            recommendations.suggestions.push('Use OCR software to make the PDF text-selectable first');
        }
        if (analysis.isProtected) {
            recommendations.canConvertToWord = false;
            recommendations.canConvertToExcel = false;
            recommendations.canConvertToPowerPoint = false;
            recommendations.warnings.push('This PDF appears to be password-protected or restricted');
            recommendations.suggestions.push('Remove password protection before conversion');
        }
        if (analysis.hasComplexLayout) {
            recommendations.warnings.push('Complex layout detected - conversion quality may vary');
            recommendations.suggestions.push('Simpler PDFs generally convert better');
        }
        if (analysis.pageCount > 50) {
            recommendations.warnings.push('Large PDF detected - conversion may take longer');
            recommendations.suggestions.push('Consider splitting into smaller files for faster processing');
        }
        return recommendations;
    }
    async compressPdf(file, quality = 'moderate', res) {
        console.log(`🔥 COMPRESS-PDF REQUEST RECEIVED at ${new Date().toISOString()}`);
        console.log(`🔥 Request headers:`, JSON.stringify(Object.keys(res.req.headers || {})));
        console.log(`🔥 File received:`, file ? `YES (${file.size} bytes)` : 'NO');
        console.log(`🔥 Quality parameter:`, quality);
        try {
            if (!file) {
                console.log(`🔥 ERROR: No file uploaded`);
                return res.status(400).json({ error: 'No PDF file uploaded' });
            }
            if (file.size > 50 * 1024 * 1024) {
                return res.status(400).json({
                    error: 'File too large',
                    message: 'PDF file size must be less than 50MB',
                    currentSize: `${(file.size / (1024 * 1024)).toFixed(2)}MB`,
                    maxSize: '50MB'
                });
            }
            const validation = this.fileValidationService.validateFile(file, 'pdf');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'PDF validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            console.log(`Starting PDF compression: ${file.originalname}, size: ${file.size} bytes, quality: ${quality}`);
            const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
            if (file.size > 10 * 1024 * 1024) {
                console.log(`Large PDF detected (${sizeMB}MB) - likely contains high-resolution images from mobile camera`);
                console.log(`Using extended timeout for image-heavy PDF compression...`);
            }
            console.log(`Compression quality setting: ${quality}`);
            const startTime = Date.now();
            const output = await this.appService.compressPdf(file, quality);
            const processingTime = (Date.now() - startTime) / 1000;
            const outputSizeMB = (output.length / (1024 * 1024)).toFixed(2);
            const compressionRatio = ((1 - output.length / file.size) * 100).toFixed(1);
            console.log(`PDF compression completed: ${file.originalname}`);
            console.log(`  Original: ${sizeMB}MB`);
            console.log(`  Compressed: ${outputSizeMB}MB`);
            console.log(`  Compression: ${compressionRatio}%`);
            console.log(`  Processing time: ${processingTime}s`);
            res.set({ 'Content-Type': 'application/pdf' });
            res.send(output);
        }
        catch (error) {
            console.error('PDF compression error:', error.message);
            console.error('Error stack:', error.stack);
            if (error.code === 'LIMIT_FILE_SIZE') {
                return res.status(400).json({
                    error: 'File too large',
                    message: 'PDF file size exceeds the 50MB limit',
                    maxSize: '50MB'
                });
            }
            if (error instanceof common_1.BadRequestException) {
                return res.status(400).json({ error: 'Invalid file', message: error.message });
            }
            if (error.message.includes('timeout') || error.message.includes('timed out')) {
                const isImageHeavyError = error.message.includes('mobile camera') || error.message.includes('high-resolution');
                return res.status(408).json({
                    error: 'Compression timeout',
                    message: isImageHeavyError
                        ? 'This PDF contains high-resolution images (likely mobile camera photos) that take too long to compress.'
                        : 'The PDF compression is taking too long. This often happens with large or complex PDFs.',
                    suggestions: [
                        'Try using "low" quality setting for faster compression',
                        'Reduce image resolution before creating the PDF',
                        'Try with a smaller PDF file (under 10MB for best results)',
                        'Consider compressing images separately before adding to PDF',
                        'Split large PDFs into smaller files'
                    ]
                });
            }
            else if (error.message.includes('image-heavy') || error.message.includes('high-resolution') || error.message.includes('mobile camera')) {
                return res.status(422).json({
                    error: 'High-resolution images detected',
                    message: 'This PDF contains high-resolution images (likely from mobile camera photos) that are difficult to compress.',
                    suggestions: [
                        'Use "low" quality setting for better compression of mobile photos',
                        'Reduce image quality/resolution before creating the PDF',
                        'Try compressing images to JPEG format before adding to PDF',
                        'Consider using image editing software to reduce file sizes first'
                    ]
                });
            }
            else if (error.message.includes('service is not available') || error.message.includes('command not found') || error.message.includes('ghostscript')) {
                return res.status(503).json({
                    error: 'Service unavailable',
                    message: 'PDF compression service is temporarily unavailable. The Ghostscript service may be down.'
                });
            }
            else if (error.message.includes('All compression methods failed')) {
                return res.status(422).json({
                    error: 'Compression failed',
                    message: 'This PDF could not be compressed using any available method. It may contain very complex content or be corrupted.',
                    suggestions: [
                        'Try with a different PDF file',
                        'Check if the PDF is password-protected',
                        'Try converting the PDF to images and back to PDF first',
                        'Ensure the PDF is not corrupted',
                        'Try with a smaller, simpler PDF'
                    ]
                });
            }
            else if (error.message.includes('poppler-utils') || error.message.includes('pdfinfo') || error.message.includes('pdfimages')) {
                return res.status(500).json({
                    error: 'PDF analysis failed',
                    message: 'Could not analyze PDF content, but compression may still work.',
                    details: 'PDF content analysis tools are not available',
                    suggestions: [
                        'Try compression anyway - it may still work',
                        'Use "low" quality setting',
                        'Try with a smaller PDF file'
                    ]
                });
            }
            res.status(500).json({
                error: 'Failed to compress PDF',
                message: 'An unexpected error occurred during compression.',
                details: error.message,
                suggestions: [
                    'Try with a different PDF file',
                    'Use "low" quality setting for mobile camera photos',
                    'Ensure the PDF is not corrupted or password-protected',
                    'Try with a smaller file (under 50MB)'
                ]
            });
        }
    }
    async getConvertApiStatus(res) {
        try {
            const status = await this.appService.getConvertApiStatus();
            return res.status(200).json({
                timestamp: new Date().toISOString(),
                convertapi: status
            });
        }
        catch (error) {
            console.error('ConvertAPI status check error:', error.message);
            return res.status(500).json({
                error: 'Failed to check ConvertAPI status',
                message: error.message
            });
        }
    }
    async getOnlyOfficeStatus(res) {
        try {
            const status = await this.appService.getOnlyOfficeStatus();
            return res.status(200).json({
                timestamp: new Date().toISOString(),
                onlyoffice: status
            });
        }
        catch (error) {
            console.error('ONLYOFFICE status check error:', error.message);
            return res.status(500).json({
                error: 'Failed to check ONLYOFFICE status',
                message: error.message
            });
        }
    }
    async getEnhancedOnlyOfficeStatus(res) {
        try {
            const status = await this.appService.getEnhancedOnlyOfficeStatus();
            return res.status(200).json({
                timestamp: new Date().toISOString(),
                service: 'ONLYOFFICE Enhanced Service',
                status: status
            });
        }
        catch (error) {
            console.error('Enhanced ONLYOFFICE status check error:', error.message);
            return res.status(500).json({
                error: 'Failed to check Enhanced ONLYOFFICE status',
                message: error.message
            });
        }
    }
    async addPasswordToPdf(file, password, res) {
        try {
            if (!file) {
                return res.status(400).json({ error: 'No PDF file uploaded' });
            }
            if (!password) {
                return res.status(400).json({ error: 'Password is required' });
            }
            const validation = this.fileValidationService.validateFile(file, 'pdf');
            if (!validation.isValid) {
                return res.status(400).json({
                    error: 'PDF validation failed',
                    details: validation.errors
                });
            }
            file.originalname = validation.sanitizedFilename;
            console.log(`Adding password protection to PDF: ${file.originalname}, size: ${file.size} bytes`);
            const output = await this.appService.addPasswordToPdf(file, password);
            res.set({
                'Content-Type': 'application/pdf',
                'Content-Disposition': `attachment; filename="${file.originalname.replace('.pdf', '_protected.pdf')}"`
            });
            res.send(output);
        }
        catch (error) {
            console.error('PDF password protection error:', error.message);
            if (error instanceof common_1.BadRequestException) {
                return res.status(400).json({ error: 'Invalid input', message: error.message });
            }
            if (error.message.includes('timeout')) {
                return res.status(408).json({
                    error: 'Password protection timeout',
                    message: 'The PDF password protection is taking too long. Please try with a smaller PDF.'
                });
            }
            else if (error.message.includes('LibreOffice')) {
                return res.status(503).json({
                    error: 'Service unavailable',
                    message: 'LibreOffice service is not available. Please try again later.'
                });
            }
            res.status(500).json({
                error: 'Failed to add password protection to PDF',
                message: error.message
            });
        }
    }
};
exports.AppController = AppController;
__decorate([
    (0, common_1.Get)(),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", void 0)
], AppController.prototype, "healthCheck", null);
__decorate([
    (0, common_1.Get)('health'),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", void 0)
], AppController.prototype, "health", null);
__decorate([
    (0, common_1.Get)('cors-test'),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", void 0)
], AppController.prototype, "corsTest", null);
__decorate([
    (0, common_1.Post)('cors-test'),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __param(1, (0, common_1.Body)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", void 0)
], AppController.prototype, "corsTestPost", null);
__decorate([
    (0, common_1.Post)('convert-word-to-pdf'),
    (0, throttler_1.Throttle)({ default: { limit: 20, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "convertWordToPdf", null);
__decorate([
    (0, common_1.Post)('convert-excel-to-pdf'),
    (0, throttler_1.Throttle)({ default: { limit: 20, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "convertExcelToPdf", null);
__decorate([
    (0, common_1.Post)('convert-ppt-to-pdf'),
    (0, throttler_1.Throttle)({ default: { limit: 20, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "convertPptToPdf", null);
__decorate([
    (0, common_1.Post)('convert-pdf-to-word'),
    (0, throttler_1.Throttle)({ default: { limit: 5, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "convertPdfToWord", null);
__decorate([
    (0, common_1.Post)('convert-pdf-to-excel'),
    (0, throttler_1.Throttle)({ default: { limit: 5, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "convertPdfToExcel", null);
__decorate([
    (0, common_1.Post)('convert-pdf-to-ppt'),
    (0, throttler_1.Throttle)({ default: { limit: 5, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "convertPdfToPpt", null);
__decorate([
    (0, common_1.Post)('analyze-pdf'),
    (0, throttler_1.Throttle)({ default: { limit: 10, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "analyzePdf", null);
__decorate([
    (0, common_1.Post)('compress-pdf'),
    (0, throttler_1.Throttle)({ default: { limit: 15, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file', {
        limits: {
            fileSize: 50 * 1024 * 1024,
        },
    })),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Query)('quality')),
    __param(2, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, String, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "compressPdf", null);
__decorate([
    (0, common_1.Get)('convertapi/status'),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "getConvertApiStatus", null);
__decorate([
    (0, common_1.Get)('onlyoffice/status'),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "getOnlyOfficeStatus", null);
__decorate([
    (0, common_1.Get)('onlyoffice/enhanced-status'),
    (0, throttler_1.SkipThrottle)(),
    __param(0, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "getEnhancedOnlyOfficeStatus", null);
__decorate([
    (0, common_1.Post)('add-password-to-pdf'),
    (0, throttler_1.Throttle)({ default: { limit: 10, ttl: 86400000 } }),
    (0, common_1.UseInterceptors)((0, platform_express_1.FileInterceptor)('file')),
    __param(0, (0, common_1.UploadedFile)()),
    __param(1, (0, common_1.Body)('password')),
    __param(2, (0, common_1.Res)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object, String, Object]),
    __metadata("design:returntype", Promise)
], AppController.prototype, "addPasswordToPdf", null);
exports.AppController = AppController = __decorate([
    (0, common_1.Controller)(),
    __metadata("design:paramtypes", [app_service_1.AppService,
        file_validation_service_1.FileValidationService])
], AppController);
//# sourceMappingURL=app.controller.js.map