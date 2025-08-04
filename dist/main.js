"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const core_1 = require("@nestjs/core");
const app_module_1 = require("./app.module");
const express_1 = require("express");
const express = require("express");
const os = require("os");
async function bootstrap() {
    try {
        console.log('Starting PDF Converter API...');
        const app = await core_1.NestFactory.create(app_module_1.AppModule);
        app.enableCors({
            origin: true,
            methods: ['GET', 'HEAD', 'PUT', 'PATCH', 'POST', 'DELETE', 'OPTIONS'],
            allowedHeaders: [
                'Accept',
                'Authorization',
                'Content-Type',
                'X-Requested-With',
                'Range',
                'Origin',
                'X-Forwarded-For',
                'X-Real-IP',
                'User-Agent',
                'Cache-Control'
            ],
            exposedHeaders: [
                'X-RateLimit-Limit',
                'X-RateLimit-Remaining',
                'X-RateLimit-Reset',
                'Retry-After',
                'Content-Length',
                'Content-Range',
                'Content-Disposition'
            ],
            credentials: true,
            preflightContinue: false,
            optionsSuccessStatus: 204,
        });
        app.use((req, res, next) => {
            res.header('Access-Control-Allow-Origin', '*');
            res.header('Access-Control-Allow-Methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
            res.header('Access-Control-Allow-Headers', 'Accept,Authorization,Content-Type,X-Requested-With,Range,Origin,X-Forwarded-For,X-Real-IP,User-Agent,Cache-Control');
            res.header('Access-Control-Expose-Headers', 'X-RateLimit-Limit,X-RateLimit-Remaining,X-RateLimit-Reset,Retry-After,Content-Length,Content-Range,Content-Disposition');
            if (req.method === 'OPTIONS') {
                console.log(`🔄 Handling preflight request for ${req.url}`);
                res.sendStatus(204);
                return;
            }
            next();
        });
        app.use((0, express_1.json)({ limit: '50mb' }));
        app.use(express.urlencoded({ extended: true, limit: '50mb' }));
        const port = process.env.PORT || 3000;
        app.use((req, res, next) => {
            const timestamp = new Date().toISOString();
            console.log(`🌐 [${timestamp}] ${req.method} ${req.url} - Body size: ${req.get('Content-Length') || 'unknown'} bytes`);
            if (req.url.includes('compress-pdf')) {
                console.log(`🔥 COMPRESS-PDF REQUEST DETECTED IN MIDDLEWARE`);
                console.log(`🔥 Content-Type: ${req.get('Content-Type') || 'none'}`);
                console.log(`🔥 User-Agent: ${req.get('User-Agent') || 'none'}`);
            }
            next();
        });
        const host = '0.0.0.0';
        await app.listen(port, host);
        if (process.env.PORT && process.env.PORT !== '3000') {
            try {
                const secondApp = await core_1.NestFactory.create(app_module_1.AppModule);
                secondApp.enableCors({
                    origin: true,
                    methods: ['GET', 'HEAD', 'PUT', 'PATCH', 'POST', 'DELETE', 'OPTIONS'],
                    allowedHeaders: [
                        'Accept',
                        'Authorization',
                        'Content-Type',
                        'X-Requested-With',
                        'Range',
                        'Origin',
                        'X-Forwarded-For',
                        'X-Real-IP',
                        'User-Agent',
                        'Cache-Control'
                    ],
                    exposedHeaders: [
                        'X-RateLimit-Limit',
                        'X-RateLimit-Remaining',
                        'X-RateLimit-Reset',
                        'Retry-After',
                        'Content-Length',
                        'Content-Range',
                        'Content-Disposition'
                    ],
                    credentials: true,
                    preflightContinue: false,
                    optionsSuccessStatus: 204,
                });
                secondApp.use((req, res, next) => {
                    res.header('Access-Control-Allow-Origin', '*');
                    res.header('Access-Control-Allow-Methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
                    res.header('Access-Control-Allow-Headers', 'Accept,Authorization,Content-Type,X-Requested-With,Range,Origin,X-Forwarded-For,X-Real-IP,User-Agent,Cache-Control');
                    res.header('Access-Control-Expose-Headers', 'X-RateLimit-Limit,X-RateLimit-Remaining,X-RateLimit-Reset,Retry-After,Content-Length,Content-Range,Content-Disposition');
                    if (req.method === 'OPTIONS') {
                        res.sendStatus(204);
                        return;
                    }
                    next();
                });
                secondApp.use((0, express_1.json)({ limit: '50mb' }));
                secondApp.use(express.urlencoded({ extended: true, limit: '50mb' }));
                await secondApp.listen(3000, host);
                console.log(`🔄 Backup server also listening on port 3000 for Railway compatibility`);
            }
            catch (err) {
                console.log(`ℹ️ Could not start backup server on port 3000: ${err.message}`);
            }
        }
        console.log(`✅ Server listening on http://${host}:${port}`);
        console.log(`✅ Server successfully started on port ${port}`);
        console.log(`🔍 Attempting to verify server is accessible...`);
        try {
            const http = require('http');
            const testReq = http.request({
                hostname: host === '0.0.0.0' ? '127.0.0.1' : host,
                port: port,
                path: '/health',
                method: 'GET',
                timeout: 5000
            }, (res) => {
                console.log(`✅ Local health check successful: ${res.statusCode}`);
            });
            testReq.on('error', (err) => {
                console.error(`❌ Local health check failed: ${err.message}`);
            });
            testReq.on('timeout', () => {
                console.error(`❌ Local health check timed out`);
            });
            testReq.end();
        }
        catch (err) {
            console.error(`❌ Failed to perform local health check: ${err.message}`);
        }
        console.log(`Environment PORT: ${process.env.PORT}, Using port: ${port}`);
        console.log(`Server running in ${process.env.NODE_ENV || 'development'} mode`);
        const networkInterfaces = os.networkInterfaces();
        console.log('Network interfaces:');
        Object.keys(networkInterfaces).forEach((interfaceName) => {
            const interfaces = networkInterfaces[interfaceName];
            if (interfaces) {
                interfaces.forEach((iface) => {
                    console.log(`  ${interfaceName}: ${iface.address} (${iface.family})`);
                });
            }
        });
        console.log('Environment variables:');
        console.log(`  NODE_ENV: ${process.env.NODE_ENV || 'not set'}`);
        console.log(`  PORT: ${process.env.PORT || 'not set'}`);
        console.log(`  RAILWAY_ENVIRONMENT: ${process.env.RAILWAY_ENVIRONMENT || 'not set'}`);
        console.log(`  RAILWAY_SERVICE_NAME: ${process.env.RAILWAY_SERVICE_NAME || 'not set'}`);
        console.log(`  TEMP_DIR: ${process.env.TEMP_DIR || 'not set'}`);
        console.log(`  Hostname: ${os.hostname()}`);
        console.log(`  Platform: ${os.platform()}`);
        console.log(`  Architecture: ${os.arch()}`);
        console.log(`  Total memory: ${Math.round(os.totalmem() / (1024 * 1024))} MB`);
        console.log(`  Free memory: ${Math.round(os.freemem() / (1024 * 1024))} MB`);
        console.log('🚀 PDF Converter API started successfully!');
        setInterval(() => {
            console.log(`📊 Server still running on ${host}:${port} - Memory: ${Math.round(process.memoryUsage().heapUsed / 1024 / 1024)}MB`);
        }, 30000);
        process.on('SIGTERM', async () => {
            console.log('SIGTERM received, shutting down gracefully...');
            await app.close();
            process.exit(0);
        });
        process.on('SIGINT', async () => {
            console.log('SIGINT received, shutting down gracefully...');
            await app.close();
            process.exit(0);
        });
        process.on('uncaughtException', (error) => {
            console.error('❌ Uncaught Exception:', error);
            process.exit(1);
        });
        process.on('unhandledRejection', (reason, promise) => {
            console.error('❌ Unhandled Rejection at:', promise, 'reason:', reason);
            process.exit(1);
        });
    }
    catch (error) {
        console.error('❌ Failed to start server:', error);
        process.exit(1);
    }
}
bootstrap();
//# sourceMappingURL=main.js.map