import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { json } from 'express';
import * as express from 'express';
import * as cors from 'cors';
import * as os from 'os';

async function bootstrap() {
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
  
  // Listen on IPv6 (::) which also supports IPv4 through IPv4-mapped IPv6 addresses
  await app.listen(port, '::');
  
  // Log the actual port being used
  const serverUrl = await app.getUrl();
  console.log(`Server listening on ${serverUrl}`);
  console.log(`Environment PORT: ${process.env.PORT}, Using port: ${port}`);
  console.log(`Server running in ${process.env.NODE_ENV || 'development'} mode`);
  
  // Log network interfaces for debugging
  const networkInterfaces = os.networkInterfaces();
  console.log('Network interfaces:');
  Object.keys(networkInterfaces).forEach((interfaceName) => {
    const interfaces = networkInterfaces[interfaceName];
    interfaces.forEach((iface) => {
      console.log(`  ${interfaceName}: ${iface.address} (${iface.family})`);
    });
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
}
bootstrap();