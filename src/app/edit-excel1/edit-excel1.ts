import { Component, ViewChild } from '@angular/core';
import { HotTableModule, GridSettings, HotTableComponent } from '@handsontable/angular-wrapper';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-edit-excel1',
  standalone: true,
  imports: [HotTableModule],
  templateUrl: './edit-excel1.html',
  styleUrl: './edit-excel1.scss'
})
export class EditExcel1 {
  @ViewChild(HotTableComponent, { static: false }) hotTable!: HotTableComponent;

  tableData: any[][] = [];

    workbook:any;

    readonly gridSettings = <GridSettings>  {
      rowHeaders: true,
      colHeaders: true,

      //select остается в таблице при смене фокуса
      outsideClickDeselects: false,
  
      autoWrapRow: true,
      autoWrapCol: true,
  
      formula:true,
  
    }


  async onFileChange(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    this.workbook = new ExcelJS.Workbook();
    await this.workbook.xlsx.load(buffer);
    const sheet = this.workbook.worksheets[0];

    //console.log(sheet.getSheetValues())
    //возврат  sheet.getSheetValues() первая колонка и строка - пустые !!! 
    const rows = sheet.getSheetValues().slice(1).map((row: string | any[]) => Array.isArray(row) ? row.slice(1) : []);
    this.tableData = rows;
  }

  addRow() {
    
    let last = this.hotTable?.hotInstance?.getSelectedLast();

    if (last == null) return;

    const sheet = this.workbook.worksheets[0];

    const sourcePosition = last[0]+1; // строка, чьи стили копируем с учетом что в гриде начало с 1 она переходит вниз после вставки
    const newPosition = sourcePosition+1; // куда переходит старая  строка

    const emptyRow = sheet.getRow(sourcePosition);//она станет пустой после вставки
    const oldRow = sheet.getRow(newPosition);  //сюда сдвигается старая строка

    //console.log(oldRow)
    // Вставляем пустую строку
    sheet.spliceRows(sourcePosition, 0, []);
    //console.log(sourceRowIndex, oldRow, newRow)
    this.copyValueAndRowStyle( oldRow, emptyRow);
    this.shiftFormulasAfterInsert(sheet, sourcePosition);

    //console.log(sheet.getSheetValues())
    const rows = sheet.getSheetValues().slice(1).map((row: string | any[]) => Array.isArray(row) ? row.slice(1) : []);
    this.tableData = rows;
  }

  copyValueAndRowStyle(sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row) {
    targetRow.height = sourceRow.height;
    sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = targetRow.getCell(colNumber);
    targetCell.style = JSON.parse(JSON.stringify(cell.style)); // глубокая копия
    targetCell.numFmt = cell.numFmt;
    targetCell.value = cell.value;
  });
  }

  async saveOnDisk()  {
        const buffer = await this.workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), 'edited.xlsx');
  }
  
  shiftFormulasAfterInsert(sheet: ExcelJS.Worksheet, insertIndex: number, count = 1) {
  const rangeRegex = /\$?[A-Z]{1,3}\$?\d+/g; // все A1-ссылки
  sheet.eachRow({ includeEmpty: true }, row => {
    row.eachCell({ includeEmpty: true }, cell => {
      if (typeof cell.formula === "string") {
        const newFormula = cell.formula.replace(rangeRegex, ref => {
          const { col, row } = this.parseCellRef(ref);
          if (row >= insertIndex) {
            // Сдвигаем строки, но не трогаем колонки
            return ref.replace(/\d+/, (row + count).toString());
          }
          return ref;
        });
        if (newFormula !== cell.formula) {
          cell.value = { formula: newFormula };
        }
      }
    });
  });
  }

  parseCellRef(ref: string): { col: string; row: number } {
    const match = ref.match(/\$?([A-Z]{1,3})\$?(\d+)/);
    return match ? { col: match[1], row: parseInt(match[2], 10) } : { col: "", row: 0 };
  }

}
