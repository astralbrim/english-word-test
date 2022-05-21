import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ApiController } from './api/api.controller';
import { ScrapeService } from './scrape/scrape.service';
import { ApiModule } from './api/api.module';

@Module({
  imports: [ApiModule],
  controllers: [AppController, ApiController],
  providers: [AppService, ScrapeService],
})
export class AppModule {}
