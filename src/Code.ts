interface Sheet {
  getRange(row: number, column: number): Range
}
interface Range {
  setBackground(color: string);
  getColumn(): number;
  getRow(): number;
}

class Period {
  constructor(public startDate: Date, public finishDate: Date) { };
}

class AutoPaintGanttChart {
  constructor() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadSheet.getSheetByName(AutoPaintGanttChart.sheetName);

    this._sheet = sheet;
    this._firstDateCell = this._sheet.getRange(AutoPaintGanttChart.firstDateCellName);
    this._activeCell = this._sheet.getActiveCell();
  };

  static sheetName: string = 'ガントチャート';
  static startDateColumn: number = 6;
  static finishDateColumn: number = 7;
  static firstDateColumn: number = 10;
  static firstDateRow: number = 10;
  static firstDateCellName: string = 'K10';
  static projectDays: number = 180;

  private _sheet: any;
  private _firstDateCell: Range;
  private _activeCell: Range;

  isTargetCell(): boolean {
    const activeCell = this._sheet.getActiveCell();
    const activeColumn: number = activeCell.getColumn();
    return activeColumn === AutoPaintGanttChart.startDateColumn || activeColumn === AutoPaintGanttChart.finishDateColumn;
  };

  getPeriod(): Period {
    const startDate: Date = this._sheet.getRange(this._activeCell.getRow(), AutoPaintGanttChart.startDateColumn).getValue();
    const finishDate: Date = this._sheet.getRange(this._activeCell.getRow(), AutoPaintGanttChart.finishDateColumn).getValue();
    return new Period(startDate, finishDate);
  };

  clearPaint(): void {
    const range = this._sheet.getRange(this._activeCell.getRow(), AutoPaintGanttChart.firstDateColumn, 1, AutoPaintGanttChart.projectDays);
    range.setBackground(null);
  }

  paintPeriod(period: Period): void {
    const range = this.getChartPeriodRange(this._activeCell.getRow(), period.startDate, period.finishDate);
    if (range === null) return;
    range.setBackground("#0b5394");
    console.log('painted!!!')
  }

  getChartPeriodRange(row: number, startDate: Date, finishDate: Date): Range | null {
    let startDateCell = this.getChartDateCell(row, startDate);
    let finishDateCell = this.getChartDateCell(row, finishDate);


    if (startDateCell === null && finishDateCell === null) {
      return null;
    }

    if (startDateCell === null) {
      startDateCell = this._firstDateCell;
    }

    if (finishDateCell === null) {
      finishDateCell = this._sheet.getRange(
        row, this._firstDateCell.getColumn() + AutoPaintGanttChart.projectDays);
    }

    const startColumn: number = startDateCell.getColumn();
    const numColumns: number = finishDateCell.getColumn() - startDateCell.getColumn() + 1;
    return this._sheet.getRange(row, startColumn, 1, numColumns);
  }

  getChartDateCell(row: number, targetDate: Date): Range | null {
    console.log(targetDate)
    const firstDateCellColumn: number = this._firstDateCell.getColumn();
    for (let i: number = 0; i < AutoPaintGanttChart.projectDays; i++) {
      const column: number = firstDateCellColumn + i;
      const date: Date = this._sheet.getRange(AutoPaintGanttChart.firstDateRow, column).getValue();
      if (date.getTime() == targetDate.getTime()) {
        return this._sheet.getRange(row, column);
      }
    }

    return null;
  }
}

function autoPaint(): void {
  const gantt = new AutoPaintGanttChart();
  const period = gantt.getPeriod();
  if (gantt.isTargetCell()) {
    gantt.clearPaint();
    gantt.paintPeriod(period);
  }
}

// function autoPaintComplete(): void {};
// function paintToday(): void {};