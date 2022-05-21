export class WordEntity {
  constructor(public english: string, public japanese: string) {}
}

export class UnitEntity {
  constructor(public unitName: string) {
    this.words = [];
  }
  words: WordEntity[];

  addWord(word: WordEntity) {
    this.words.push(word);
  }
}

export class UnitsEntity {
  constructor() {
    this.units = [];
  }
  units: UnitEntity[];

  addUnit(unit: UnitEntity) {
    this.units.push(unit);
  }
}
