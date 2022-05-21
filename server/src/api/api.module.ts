import { Module } from '@nestjs/common';
import { ScrapeService } from '../scrape/scrape.service';
import { ApiController } from './api.controller';

@Module({
  imports: [],
  controllers: [ApiController],
  providers: [ScrapeService],
})
export class ApiModule {}
