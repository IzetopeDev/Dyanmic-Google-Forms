const activeSpreadsheet = SpreadsheetApp.openById('placeholder'); // TODO: Replace ID

const sheets = {
  infraDash: activeSpreadsheet.getSheetByName('Infra Dashboard')!,
  jobArchive: activeSpreadsheet.getSheetByName('Job/Purchases Archive')!,
  safetyDash: activeSpreadsheet.getSheetByName('Safety Dashboard')!,
  settings: activeSpreadsheet.getSheetByName('Settings')!
};

const SITSettings = {
  refCells: [] as string[],
  formUrl: sheets.settings.getRange('B2').getValue().toString(),
  
  get list(): string[][] {
    return sheets.settings.getRange('A1:A').getValues().filter(
      (value: string[]) => value[0] !== '' && value[0] !== '--General--' && value[0] !== '--Developer--'
    );
  },

  getRefCells(enableLog = false): string[] {
    console.info(`SITSettings.getRefCells() called`);

    const output: string[] = [];
    let blankCounter = 0;

    for (let i = 0; blankCounter < 10; i++) {
      const currentRange = `B${i + 1}`;
      const currentValue = sheets.settings.getRange(currentRange).getValue();

      if (currentValue == '' || currentValue == null) {
        blankCounter++;
        continue;
      }

      output.push(currentRange);
    }

    if (enableLog) console.log(`SITSettings.refCells: ${output}`);
    this.refCells = output;
    return output;
  },

  setSetting(setting: string, newValue: any): void {
    const settingsList = this.list.map(item => item[0]);
    const index = settingsList.indexOf(setting);
    if (index === -1) return;
    sheets.settings.getRange(this.refCells[index]).setValue(newValue);
  }
};

class CustDate extends Date {
  abbvMonthArray: string[];

  constructor(dateStr?: string | Date) {
    const newDateStr = !dateStr ? new Date() : dateStr;
    super(newDateStr);
    this.abbvMonthArray = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  }

  toCustFormat(): string {
    return `${this.getDate()} ${this.abbvMonthArray[this.getMonth()]} ${this.getFullYear()}.`;
  }
}

function isDate(value: string): boolean {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const valueAsDate = new Date(value);
    return !isNaN(valueAsDate.getTime());
  }
  return false;
}

interface JobColumnEnum {
  JOB: string;
  STATUS: string;
  DATE_UPDATED: string;
  DETAILS: string;
  JOB_PURCHASE?: string;
  NOT_RAISED_SPOTTED?: string;
  AOR_DRAFTED?: string;
  AOR_APPROVED?: string;
  DR_RAISED?: string;
  JR_RAISED?: string;
  JRC_VERIFIED?: string;
  PO_RAISED?: string;
  JRR_VERIFIED?: string;
  COC_SIGNED?: string;
  GR_DONE?: string;
  CX_DONE?: string;
}

class UnlimitedTable {
  dataSets: string[][];
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  refCells: string[];
  column: JobColumnEnum;

  constructor(sheet: GoogleAppsScript.Spreadsheet.Sheet, refCells: string[], columnEnum: JobColumnEnum) {
    this.sheet = sheet;
    this.refCells = refCells;
    this.column = columnEnum;

    this.dataSets = this.refCells.map(refCell => {
      const rangeSelection = `${refCell}:${refCell[0]}`;
      const values = this.sheet.getRange(rangeSelection).getValues();
      const column: string[] = [];
      let blankCounter = 0;

      for (const valueRow of values) {
        const value = valueRow[0].toString();
        if (value === '') {
          blankCounter++;
          if (blankCounter > 10) break;
        } else {
          blankCounter = 0;
        }

        if (blankCounter <= 10) column.push(value);
      }

      return column;
    });
  }

  getFromCol(columnKey: string): string[] {
    const columnIndex = Object.values(this.column).indexOf(columnKey);
    if (columnIndex === -1) throw new Error(`Invalid column: ${columnKey}`);
    return this.dataSets[columnIndex];
  }

  updateSimTable(response: any[], dataStartIndex: number): void {
    const respAns = response.map((item) => {
      const answer = item.getResponse();
      return isDate(answer) ? new CustDate(answer).toCustFormat() : answer;
    });

    let jobIndex = -1;
    for (let i = 0; i < this.getFromCol(this.column.JOB).length; i++) {
      if (this.getFromCol(this.column.JOB)[i] === respAns[dataStartIndex - 1]) {
        jobIndex = i;
        break;
      }
    }

    if (respAns[3] === '') respAns[3] = respAns[2];
    if (respAns[5] === '') respAns[5] = new CustDate().toCustFormat();

    if (jobIndex === -1) {
      for (let i = dataStartIndex; i < respAns.length; i++) {
        this.dataSets[i - dataStartIndex].unshift(respAns[i]);
      }
      return;
    }

    let column = 0;
    for (let i = 0; i < respAns.length; i++) {
      if (i < dataStartIndex) continue;
      if (respAns[dataStartIndex] === "DELETE ENTRY") {
        this.dataSets[column].splice(jobIndex, 1);
        column++;
        continue;
      }

      if (respAns[i] === '') {
        column++;
        continue;
      }

      if (i === respAns.length - 1 && respAns[i] === 'REMOVE DEFECT DETAILS') {
        this.dataSets[this.dataSets.length - 1].splice(jobIndex, 1, '');
        column++;
        continue;
      }

      if (column > 3) return;

      this.dataSets[column].splice(jobIndex, 1, respAns[i]);
      column++;
    }
  }

  writeToSheet(): void {
    this.refCells.forEach((refCell, column) => {
      for (let row = 0; row < this.dataSets[column].length; row++) {
        this.sheet.getRange(`${refCell[0]}${Number(refCell[1]) + row}`).setValue(this.dataSets[column][row]);
      }
    });
  }

  sortTableByAlpahbet(): void {
    const unsortedNames = this.dataSets[0];
    const sortedNames = [...unsortedNames].sort((a, b) => -a.localeCompare(b));
    const sortIndexes = sortedNames.map(name => unsortedNames.indexOf(name));

    this.dataSets = this.dataSets.map(column =>
      sortIndexes.map(index => column[index])
    );
  }
}

const tables = {
  infraJobs: new UnlimitedTable(
    sheets.infraDash,
    ["A2", "B2", "C2", "D2"],
    {
      JOB: 'job',
      STATUS: 'status',
      DATE_UPDATED: 'date updated',
      DETAILS: 'details'
    }
  ),
  infraDefects: new UnlimitedTable(
    sheets.infraDash,
    ['F2', 'G2', 'H2', 'I2'],
    {
      JOB: 'job',
      STATUS: 'status',
      DATE_UPDATED: 'date updated',
      DETAILS: 'details'
    }
  ),
};
