import { Module } from '@nestjs/common';
import { ThrottlerModule, ThrottlerGuard } from '@nestjs/throttler';
import { APP_GUARD } from '@nestjs/core';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ConvertApiService } from './convertapi.service';
import { FileValidationService } from './file-validation.service';

@Module({
  imports: [
    ThrottlerModule.forRoot([
      {
        name: 'short',
        ttl: 60000, // 1 minute
        limit: 10, // 10 requests per minute for general endpoints
      },
      {
        name: 'daily-pdf-conversion',
        ttl: 86400000, // 24 hours (1 day)
        limit: 5, // 5 requests per day for PDF conversions
      },
      {
        name: 'medium',
        ttl: 900000, // 15 minutes
        limit: 20, // 20 requests per 15 minutes for other operations
      }
    ])
  ],
  controllers: [AppController],
  providers: [
    AppService,
    ConvertApiService,
    FileValidationService,
    {
      provide: APP_GUARD,
      useClass: ThrottlerGuard,
    },
  ],
})
export class AppModule {}