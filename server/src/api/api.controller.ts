import { Controller, Get, Param } from '@nestjs/common';
import { ScrapeService } from '../scrape/scrape.service';
import { UnitsEntity } from '../scrape/word.entity';
import { UnitNamesEntity } from '../scrape/unit-name.entity';

@Controller('api')
export class ApiController {
  constructor(private readonly scrapeService: ScrapeService) {}
  @Get('/words/:grade')
  async getWords(@Param() param: { grade: string }): Promise<UnitsEntity> {
    return this.scrapeService.getWords(param.grade);
  }

  @Get('/units/:grade')
  async getUnits(@Param() param: { grade: string }): Promise<UnitNamesEntity> {
    return this.scrapeService.getUnitNames(param.grade);
  }
}
