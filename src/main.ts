import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { json } from 'express';
import * as express from 'express';
import * as cors from 'cors';
import * as os from 'os';

async function bootstrap() {
  try {
    console.log('Starting PDF Converter API...');
    
    const app = await NestFactory.create(AppModule);
    
    // Configure CORS for all origins
    app.use(cors({
      origin: '*',
      methods: 'GET,HEAD,PUT,PATCH,POST,DELETE',
      preflightContinue: false,
      optionsSuccessStatus: 204,
      credentials: true,
    }));
    
    // Increase payload size limit for file uploads
    app.use(json({ limit: '50mb' }));
    app.use(express.urlencoded({ extended: true, limit: '50mb' }));
    
    // Use PORT from environment variables (for Railway) or default to 3000
    const port = process.env.PORT || 3000;
    
    // Listen on all interfaces (0.0.0.0) for better Railway compatibility
    await app.listen(port, '0.0.0.0');
    
    // Log the actual port being used
    const serverUrl = await app.getUrl();
    console.log(`✅ Server listening on ${serverUrl}`);
    console.log(`✅ Server successfully started on port ${port}`);
    console.log(`Environment PORT: ${process.env.PORT}, Using port: ${port}`);
    console.log(`Server running in ${process.env.NODE_ENV || 'development'} mode`);
    
    // Log network interfaces for debugging
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
    
    // Log environment variables
    console.log('Environment variables:');
    console.log(`  NODE_ENV: ${process.env.NODE_ENV || 'not set'}`);
    console.log(`  PORT: ${process.env.PORT || 'not set'}`);
    console.log(`  TEMP_DIR: ${process.env.TEMP_DIR || 'not set'}`);
    console.log(`  Hostname: ${os.hostname()}`);
    console.log(`  Platform: ${os.platform()}`);
    console.log(`  Architecture: ${os.arch()}`);
    console.log(`  Total memory: ${Math.round(os.totalmem() / (1024 * 1024))} MB`);
    console.log(`  Free memory: ${Math.round(os.freemem() / (1024 * 1024))} MB`);
    
    console.log('🚀 PDF Converter API started successfully!');
    
    // Graceful shutdown handling
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
    
  } catch (error) {
    console.error('❌ Failed to start server:', error);
    process.exit(1);
  }
}
bootstrap();