import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ConvertApiService } from './convertapi.service';

@Module({
  imports: [],
  controllers: [AppController],
  providers: [AppService, ConvertApiService],
})
export class AppModule {}