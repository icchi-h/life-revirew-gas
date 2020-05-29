////////////////////////////////////////////
// SpreadSheet Lib for GAS

// Created at: 2020/03/10
// Last Updated: 2020/05/28
// Version: 0.2

// Update Description
// 2020/03/10: Initial commit
// 2020/05/28: クラス化
////////////////////////////////////////////

import Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
import Sheet = GoogleAppsScript.Spreadsheet.Sheet;

const DEFAULT_FILENAME = "SpreadSheet";
const DEFAULT_SHEET_NAME = "sheet1";
const DEFAULT_DATA_AREA_NUM_ROW = 1; // start: 1~
const DEFAULT_DATA_AREA_NUM_COLUMN = 1; // start: 1~
const DEFAULT_MAX_ROW_IDX = 100000; // spreadsheet limit: 200M cell

export class SpreadSheetService {
  readonly static DATE_FORMAT = "yyyy/MM/dd HH:mm:ss";

  templateSpreadsheetId: string | null;
  dirId: string;
  spreadsheetName: string;
  sheetName: string;
  dataAreaNumRow: number;
  dataAreaNumColumn: number;
  maxRowNum: number;
  // object
  spreadsheet: Spreadsheet;
  sheet: Sheet;

  constructor(
    dirId: string = DriveApp.getRootFolder().getId(),
    templateSpreadsheetId: string = null,
    spreadsheetName: string = DEFAULT_FILENAME,
    sheetName: string = DEFAULT_SHEET_NAME,
    dataAreaNumRow: number = DEFAULT_DATA_AREA_NUM_ROW,
    dataAreaNumColumn: number = DEFAULT_DATA_AREA_NUM_COLUMN,
    maxRowIdx: number = DEFAULT_MAX_ROW_IDX
  ) {
    this.dirId = dirId;
    this.templateSpreadsheetId = templateSpreadsheetId;
    this.spreadsheetName = spreadsheetName;
    this.sheetName = sheetName;
    this.dataAreaNumRow = dataAreaNumRow;
    this.dataAreaNumColumn = dataAreaNumColumn;
    this.maxRowNum = maxRowIdx;

    this.setup()
  }

  private setup() {
    // date
    const now = new Date();

    // get spreadsheet by name
    this.spreadsheet = this.getSpreadSheet();
    // if spreadsheet file doesn't exist, create new one
    if (!this.spreadsheet) {
      console.warn("Not found the target spreadsheet");
      this.spreadsheet = this.createSpreadSheet();
      console.log("Created a SpreadSheet:", this.spreadsheetName);
    }
    // get sheet
    this.sheet = this.spreadsheet.getSheetByName(this.sheetName);

    // check limit of spreadsheet (log rotation)
    if (!this.isFull()) {
      return;
    }

    console.warn("Log file size reached the limit");

    // backup filename
    const backupFilename = `${this.spreadsheetName}_${Utilities.formatDate(
      now,
      "Asia/Tokyo",
      "~yyyy-MM-DD"
    )}`;
    // move current log SS to old SS and create new one
    this.spreadsheet = this.logRotate(backupFilename);
    // overwrite with new log sheet
    this.sheet = this.spreadsheet.getSheetByName(this.sheetName);

    console.log("SUCCEEDED: Done log ration");
  }

  /**
   * @return Index Row Number (-1: Not found)
   */
  public getRowNumByValue(value: string, columnNum: number): number {
    // get date column's values (2D->1D)
    const columnValues = this.sheet
      .getRange(this.dataAreaNumRow, columnNum, this.sheet.getLastRow() || 1)
      .getValues()
      .map((elem) => elem[0]);
    return columnValues.indexOf(value) + 1;
  }

  public getSheet(): Sheet {
    return this.sheet;
  }

  public getValue(rowNum: number, columnNum: number): string | number {
    return this.sheet.getRange(rowNum, columnNum).getValue();
  }

  public setValues(
    values: any[][],
    rowNum: number,
    columnNum: number,
    rowCount: number,
    columnCount: number
  ) {
    this.sheet
      .getRange(rowNum, columnNum, rowCount, columnCount)
      .setValues(values);
  }

  public sortSheet(
    targetColumnNum: number,
    ascending: boolean = false,
    startRowNum = 1,
    startColumnNum = 1:
    rowCount = this.sheet.getLastRow(),
    columnCount = this.sheet.getLastColumn()
  ) {
    this.sheet.getRange(
      startRowNum,
      startColumnNum,
      rowCount,
      columnCount
    ).sort({column: targetColumnNum, ascending: ascending})
  }

  private getSpreadSheet(
    filename = this.spreadsheetName,
    dirId = this.dirId
  ): Spreadsheet {
    const dir = DriveApp.getFolderById(dirId);
    const ssItr = dir.getFilesByName(filename);

    // check to exit spreadsheet
    const isExist = ssItr.hasNext();
    if (!isExist) return null;

    return SpreadsheetApp.open(ssItr.next());
  }

  private createSpreadSheet(
    filename = this.spreadsheetName,
    outputDirId = this.dirId,
    templateId = this.templateSpreadsheetId,
    sheetName = DEFAULT_SHEET_NAME
  ): Spreadsheet {
    let ss: Spreadsheet = null;

    if (!templateId) {
      // create on root dir
      ss = SpreadsheetApp.create(filename);
      ss.getActiveSheet().setName(sheetName);

      // move created spreadsheet to output dir
      if (outputDirId) {
        // copy to output dir
        const createdSS = DriveApp.getFileById(ss.getId());
        const outputDir = DriveApp.getFolderById(outputDirId);
        outputDir.addFile(createdSS);
        ss = SpreadsheetApp.open(createdSS);

        // remove from root dir
        DriveApp.getRootFolder().removeFile(createdSS);
      }
    }
    // create from template
    else {
      const templateSS = DriveApp.getFileById(templateId);
      const outputDir = DriveApp.getFolderById(outputDirId);
      const copiedSS = templateSS.makeCopy(filename, outputDir);

      ss = SpreadsheetApp.open(copiedSS);
    }

    return ss;
  }

  private isFull(
    sheet = this.sheet,
    thresholdRowIdx = this.maxRowNum
  ): Boolean {
    // const columns: number = spreadsheet.getLastColumn();
    const lastRowNum: number = sheet.getLastRow();
    return lastRowNum >= thresholdRowIdx;
  }

  private logRotate(
    oldFilename: string,
    templateSpreadsheetId = this.templateSpreadsheetId,
    ss = this.spreadsheet,
    newFileOutputDirId = this.dirId,
    newFilename = this.spreadsheetName,
    newSheetName = this.sheetName,
    oldFileOutputDirId = this.dirId
  ): Spreadsheet {
    ss.rename(oldFilename);

    // move old file
    const outputDir = DriveApp.getFolderById(oldFileOutputDirId);
    const oldFile = DriveApp.getFileById(ss.getId());
    const currentDir = oldFile.getParents().next();
    // add
    outputDir.addFile(oldFile);
    // remove old file on current dir
    currentDir.removeFile(oldFile);

    // create new spreadsheet file
    const newSpreadsheet = this.createSpreadSheet(
      newFilename,
      newFileOutputDirId,
      templateSpreadsheetId,
      newSheetName
    );

    return newSpreadsheet;
  }
}
