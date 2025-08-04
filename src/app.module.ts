import { Module } from '@nestjs/common';
import { ThrottlerModule, ThrottlerGuard } from '@nestjs/throttler';
import { APP_GUARD } from '@nestjs/core';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { OnlyOfficeService } from './onlyoffice.service';
import { OnlyOfficeEnhancedService } from './onlyoffice-enhanced.service';
import { FileValidationService } from './file-validation.service';

@Module({
  imports: [
    ThrottlerModule.forRoot([
      {
        name: 'short',
        ttl: 60000, // 1 minute
        limit: 15, // Increased limit for better UX with ONLYOFFICE
      },
      {
        name: 'daily-pdf-conversion',
        ttl: 86400000, // 24 hours (1 day)
        limit: 50, // Increased for ONLYOFFICE usage
      },
      {
        name: 'medium',
        ttl: 900000, // 15 minutes
        limit: 30, // Increased for better performance
      }
    ])
  ],
  controllers: [AppController],
  providers: [
    AppService,
    OnlyOfficeService,
    OnlyOfficeEnhancedService, // Add enhanced service
    FileValidationService,
    {
      provide: APP_GUARD,
      useClass: ThrottlerGuard,
    },
  ],
})
export class AppModule {}