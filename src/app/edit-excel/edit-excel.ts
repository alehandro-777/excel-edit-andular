import { Component, ViewChild } from "@angular/core";
import {
  GridSettings,
  HotTableComponent,
  HotTableModule, 
  
} from "@handsontable/angular-wrapper";
import { firstValueFrom } from 'rxjs';
import Handsontable from 'handsontable';

import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { HttpClient } from '@angular/common/http';
import { HyperFormula } from 'hyperformula';

import { CellChange, ChangeSource } from "handsontable/common";

import { textRenderer, dropdownRenderer, checkboxRenderer } from 'handsontable/renderers';

import dayjs from 'dayjs';
import 'dayjs/locale/uk' // import locale

 
@Component({
  selector: 'app-edit-excel',
  standalone: true,
  imports: [HotTableModule],
  templateUrl: './edit-excel.html',
  styleUrl: './edit-excel.scss'
})
export class EditExcel {

  constructor(private http: HttpClient, ) {

  }

  @ViewChild(HotTableComponent, { static: false }) hotTable!: HotTableComponent;

  selectedFile = "";
  data: any[] = [];
  data1: any[] = [];

  mergeCells: any[] = [];

  columns: any[] = [];
  styles: any[][] = [];

  hotSettings: any;

  hfInstance = HyperFormula.buildEmpty({
    licenseKey: 'gpl-v3'
  });

  //  grid settings !!!
  readonly gridSettings = <GridSettings>  {
    rowHeaders: true,
    colHeaders: true,
    //height: "auto",
    //rowHeights: 10,
    //manualRowResize: true,

    autoWrapRow: true,
    autoWrapCol: true,

    formula:true,

    formulas: {
      engine: this.hfInstance,
    },

    //select остается в таблице при смене фокуса
    outsideClickDeselects: false,

    //преобразования после редактирования ячейки
    //afterChange: (changes, source) => this.onAfterChange(changes, source),

  };
  //end  grid settings !!!



  argbToHex(argb: string): string {
    if (!argb || argb.length !== 8) return '#000000';

    // ARGB: AARRGGBB → выкидываем AA (альфа)
    const rgb = argb.substring(2); // "00FF00" из "FF00FF00"
    return `#${rgb}`;
  }
  
  argbToRgba(argb: string): string {
    if (!argb || argb.length !== 8) return 'rgba(0,0,0,1)';

    const a = parseInt(argb.substring(0, 2), 16) / 255;
    const r = parseInt(argb.substring(2, 4), 16);
    const g = parseInt(argb.substring(4, 6), 16);
    const b = parseInt(argb.substring(6, 8), 16);

    return `rgba(${r}, ${g}, ${b}, ${a.toFixed(2)})`;
  }

  async loadTemplateHTTP() {
   
    let buffer = await firstValueFrom(this.http.get('http://localhost:3000/download', { responseType: 'arraybuffer' }));

    await this.processBuffer(buffer);

    //console.log(this.styles)
  }

  async processBuffer(buffer: ArrayBuffer) {

    this.reset();

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[0];

    sheet.eachRow((row, rowNum) => {
      let rowData: any[] = [];
      const rowStyles: any[] = [];

      //1 чтение стилей
      row.eachCell((cell, colNum) => {
        // сохранить значение или формулу как строку, если есть
        if (cell.value && typeof cell.value === 'object' && cell.formula) {
          rowData.push('=' + cell.formula);
        } else {
          rowData.push(cell.value);
        }
        rowStyles.push({
          //цвет шрифта перенести не получится это глюк библиотеки !!!!
          font: cell.font,            //минимальный стиль ячейки
          fill: cell.fill,            //минимальный стиль ячейки
          border: cell.border,        //минимальный стиль ячейки
          alignment: cell.alignment,  //минимальный стиль ячейки
          style: cell.style           //минимальный стиль ячейки
        });
      });

      this.data1.push(rowData);
      this.styles.push(rowStyles);
      //console.log(rowData) //, rowStyles);
    });

    //2. Колонки динамически
    this.columns = Object.keys(this.data1[0]).map(key => ({ data: key }));
    this.hotTable.hotInstance?.updateSettings({ columns: this.columns });

    // 3. Считываем слияния
    this.mergeCells = [];
    sheet.model.merges.forEach((merge: any) => {
      const [tl, br] = merge.split(':'); // например "A1:B2"
      const start = sheet.getCell(tl);
      const end = sheet.getCell(br);
      this.mergeCells.push({
        row: +start.row - 1,
        col: +start.col - 1,
        rowspan: +end.row - +start.row + 1,
        colspan: +end.col - +start.col + 1,
      });
    });

    //console.log(this.data1)//, this.mergeCells, this.columns);

    for (let i = 0; i < this.data1.length; i++) {
      const row = this.data1[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j];

        this.setCell(cell, i, j);

      }      
    }

    this.hotTable.hotInstance?.updateSettings({ mergeCells: this.mergeCells });

    //console.log(this.data)

  }

  setCell(cell:any, row:number, col:number) {
        if (typeof cell === 'number') {
              //number
            this.hotTable?.hotInstance?.setCellMeta(row, col, "type",  'numeric');
            //this.hotTable?.hotInstance?.setCellMeta(row, col, "readOnly",  true);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorNumericRenderer.bind(this));
            //this.hotTable?.hotInstance?.setCellMeta(row, col, "numericFormat",  { pattern: '0\u202f0.00', culture: 'ru-RU' });//   не работал формат разделение тысяч пробелами
            this.hotTable?.hotInstance?.setDataAtCell(row, col, cell);
          return;
        }
        let numFmt = this.removeStrangeSym(this.styles[row][col].style.numFmt);

        if ((Object.prototype.toString.call(cell) === '[object Date]' || cell instanceof Date) && numFmt.includes(":")) {
          let format = this.OpenXMLTime(numFmt);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "type",  'time');
            this.hotTable?.hotInstance?.setCellMeta(row, col, "timeFormat",  format);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "correctFormat",  true);
            //this.hotTable?.hotInstance?.setCellMeta(0, 3, "allowInvalid",  false);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorTextRenderer.bind(this));
            this.hotTable?.hotInstance?.setDataAtCell(row, col, this.formatValueByNumFmt(cell, format));
            //console.log(format)
          return;
        }
        //если не время то дата
        if (Object.prototype.toString.call(cell) === '[object Date]' || cell instanceof Date) {
            let format = this.OpenXMLDate(numFmt);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "type",  'date');
            this.hotTable?.hotInstance?.setCellMeta(row, col, "dateFormat",  format);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "correctFormat",  true);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "defaultDate",  "01.12.2000");
            this.hotTable?.hotInstance?.setCellMeta(0, 3, "allowInvalid",  false);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorTextRenderer.bind(this));
            this.hotTable?.hotInstance?.setDataAtCell(row, col, this.formatValueByNumFmt(cell, format));
          return;
        }
        //  formula  
        if (typeof cell === 'string' && cell.startsWith("=")) {
          this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorNumericRenderer.bind(this));// !! без numericRenderer не форматирует формулы !
            this.hotTable?.hotInstance?.setDataAtCell(row, col, cell);
            //console.log(this.styles[row][col].style.numFmt)
          return;
        } 
        // JSON
        if (typeof cell === 'string' && cell.startsWith("{")) {

          let cellConfig = JSON.parse(cell);  // JSON with "elt":"sserwer" !!!

          if (cellConfig.type == "dropdown") {
            this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorDropRenderer.bind(this));// !! без numericRenderer не форматирует формулы !
              this.hotTable?.hotInstance?.setCellMeta(row, col, "type",  'dropdown');
              this.hotTable?.hotInstance?.setCellMeta(row, col, "source",  cellConfig.source);

          } else if (cellConfig.type == "checkbox") {
              this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorCheckRenderer.bind(this));// !! без numericRenderer не форматирует формулы !
              this.hotTable?.hotInstance?.setCellMeta(row, col, "type",  'checkbox');
              this.hotTable?.hotInstance?.setCellMeta(row, col, "label",  cellConfig.label);

          } else if (cellConfig.type == "numeric") {
            this.styles[row][col].style.numFmt = cellConfig.format; //set cell format
            this.styles[row][col].range = cellConfig.range;         //set input range
            this.hotTable?.hotInstance?.setCellMeta(row, col, "type",  'numeric');
            this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorNumericRenderer.bind(this));
            this.hotTable?.hotInstance?.setCellMeta(row, col, "validator",  this.customNumericValidator.bind(this));//
          }

          this.hotTable?.hotInstance?.setDataAtCell(row, col, cellConfig.cell);
            //console.log(cellConfig)
          return;
        }
        if (typeof cell === 'string' ) {
            this.hotTable?.hotInstance?.setCellMeta(row, col, "readOnly",  true);// затенение отключено в  renderer !!!
            this.hotTable?.hotInstance?.setDataAtCell(row, col, cell);
            this.hotTable?.hotInstance?.setCellMeta(row, col, "renderer",  this.colorTextRenderer.bind(this));
            //console.log(this.styles[row][col].style.numFmt)
          return;
        } 
  }

  onAfterChange(changes: CellChange[] | null, source: ChangeSource) {
    if (!changes) return;
    const hot = this.hotTable.hotInstance;

    if (source === 'edit') {

    if (!hot) return;

    console.log(changes)

    changes.forEach(([row, col, oldValue, newValue]) => {

        const numeric = parseFloat(newValue);

        if (!isNaN(numeric)) {          
          hot.setDataAtCell(row, +col, numeric, 'convert'); 
        } 

    });
  }
  }

  getCellName(row: number, col: number): string {
    const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    return letters[col] + (row + 1);
  }

  async saveExcel() {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');

    this.data.forEach((r, i) => {
      let row = Object.values(r);

      const excelRow = sheet.addRow(row.map((cell: any, j: number) => {
            const formula = this.hotTable?.hotInstance?.getSourceDataAtCell(i, j);
            const data = this.hotTable?.hotInstance?.getDataAtCell(i, j);

        if (typeof formula === 'string' && formula.startsWith('=')) {
          return { formula: formula.substring(1), result: data };
        }
        return data;
      }));

      row.forEach((cell: any, j: number) => {
        if (this.styles[i] && this.styles[i][j]) {
          const st = this.styles[i][j];
          excelRow.getCell(j + 1).font = st.font;
          excelRow.getCell(j + 1).fill = st.fill;
          excelRow.getCell(j + 1).border = st.border;
          excelRow.getCell(j + 1).alignment = st.alignment;
          excelRow.getCell(j + 1).style = st.style;
        }
      });
    });

    // Восстановление слияний
    this.mergeCells.forEach((merge) => {
      const tl = sheet.getCell(merge.row + 1, merge.col + 1).address;
      const br = sheet.getCell(merge.row + merge.rowspan, merge.col + merge.colspan).address;
      sheet.mergeCells(`${tl}:${br}`);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'edited.xlsx');
  }

  colorNumericRenderer(instance: Handsontable.Core, TD: HTMLTableCellElement, row: number, column: number, prop: string | number, 
    value: any, cellProperties: Handsontable.CellProperties): void  {
    // у грида  не работал формат разделение тысяч пробелами ???
    let value1 = this.formatValueByNumFmt(value , this.styles[row][column].style.numFmt);
    textRenderer.apply(this, [instance, TD, row, column, prop, value1, cellProperties]); 
    this.addTdStyles( TD, row, column, value );
  }
  colorTextRenderer(instance: Handsontable.Core, TD: HTMLTableCellElement, row: number, column: number, prop: string | number, 
    value: any, cellProperties: Handsontable.CellProperties): void  {
    textRenderer.apply(this, [instance, TD, row, column, prop, value, cellProperties]); 
    this.addTdStyles( TD, row, column, value );
  }
  colorCheckRenderer(instance: Handsontable.Core, TD: HTMLTableCellElement, row: number, column: number, prop: string | number, 
    value: any, cellProperties: Handsontable.CellProperties): void  {
    checkboxRenderer.apply(this, [instance, TD, row, column, prop, value, cellProperties]); 
    this.addTdStyles( TD, row, column, value );
  }
  colorDropRenderer(instance: Handsontable.Core, TD: HTMLTableCellElement, row: number, column: number, prop: string | number, 
    value: any, cellProperties: Handsontable.CellProperties): void  {
    dropdownRenderer.apply(this, [instance, TD, row, column, prop, value, cellProperties]); 
    this.addTdStyles( TD, row, column, value );
  }
//-----------------------------
  customNumericValidator(value: any, callback: (valid: boolean) => void) {
    // Разрешаем только числа 0–100
    const num = parseFloat(value);
    const isValid = !isNaN(num) && num >= 0 && num <= 100;

    callback(true);
  }
//-----------------------------
  addTdStyles( TD: HTMLTableCellElement, row: number, column: number, value:any ): void  {

        // Фон ячейки
    if ( this.styles[row][column]?.fill?.fgColor?.argb != undefined ) {
      TD.style.backgroundColor = this.argbToHex(this.styles[row][column].fill.fgColor.argb);
    } else {
      TD.style.backgroundColor = '';
    }

    //fontWeight
    if ( this.styles[row][column]?.font?.bold === true ) {
          TD.style.fontWeight  = 'bold';
    } else {
          TD.style.fontWeight = '';
    }

    //alignment
    if ( this.styles[row][column]?.alignment != undefined ) {
      TD.style.textAlign = this.styles[row][column].alignment.horizontal;
    } else {
      TD.style.textAlign = '';
    }

    TD.classList.remove('htDimmed'); // удаляем READonly затемнение !!!

    //-------------- условная раскраска min max задается в json
    let range = this.styles[row][column].range;

    if (range && !isNaN(value) && value > range.max) {
      TD.style.backgroundColor = '#dc3545'; // красная ячейка
      TD.style.color = 'white';
    }
    if (range && !isNaN(value) && value < range.min) {
      TD.style.backgroundColor = '#dc3545'; // красная ячейка
      TD.style.color = 'white';
    }    
  }
  formatValueByNumFmt(value: any, numFmt: string, locale = 'fr-FR'): string {
  
  //dayjs.locale('uk') // use locale

  if (value === null || value === undefined) return '';
  if (numFmt === null || numFmt === undefined) return value;

    // ======= 1. Проценты =======
    if (numFmt.includes('%') && typeof value === 'number') {
      const precision = (numFmt.split('.')[1]?.length ?? 0);
      return (value * 100).toLocaleString(locale, {
        minimumFractionDigits: precision,
        maximumFractionDigits: precision,
      }) + '%';
    }

    // ======= 2. Числовые форматы =======
    if (typeof value === 'number') {
      const parts = numFmt.split('.');
      const intPart = parts[0] || '';
      const fracPart = parts[1] || '';

      const useGrouping = intPart.includes(',') || intPart.includes(' '); // поддержка "тысячников"
      const minFraction = fracPart.length;
      const maxFraction = fracPart.length;

      return value.toLocaleString(locale, {
        useGrouping,
        minimumFractionDigits: minFraction,
        maximumFractionDigits: maxFraction,
      });
    }

    // ======= 3. Даты =======
    if (Object.prototype.toString.call(value) === '[object Date]' || value instanceof Date) {
      const date = dayjs(value);

      // OpenXML → DayJS
      const excelToDayjs = numFmt
        .replace(/\[\$\]/g, '')  //[$]dd.mm.yyyy;@ -  приходит из экселя такой формат возможно глюк библиотеки !!!
        .replace(/;@/g, '')
        .replace(/y/g, 'Y')
        .replace(/d/g, 'D')
        .replace(/m/g, 'M')
        .replace(/h/g, 'H')
        .replace(/AM\/PM/i, 'A');

      return date.format(excelToDayjs);
    }

    // ======= 4. Текстовые шаблоны =======
    if (typeof value === 'string' && numFmt.includes('@')) {
      return numFmt.replace('@', value);
    }

    return value.toString();
  }
  OpenXMLDate(numFmt: string)  {

      if (numFmt === null || numFmt === undefined) return "DD.MM.YY";

      return  numFmt
        .replace(/y/g, 'Y')
        .replace(/d/g, 'D')
        .replace(/m/g, 'M')
        .replace(/h/g, 'H')
        .replace(/AM\/PM/i, 'A');
  }
  OpenXMLTime(numFmt: string)  {

      if (numFmt === null || numFmt === undefined) return "";

      return  numFmt
        .replace(/h/g, 'H')
        .replace(/AM\/PM/i, 'A');
  }
  removeStrangeSym(numFmt: string)  {
      if (numFmt === null || numFmt === undefined) return "";

      return  numFmt
        .replace(/\[\$\]/g, '')  //[$]dd.mm.yyyy;@ -  приходит из экселя такой формат возможно глюк библиотеки !!!
        .replace(/;@/g, '');
  }
  addRow() {
    let last = this.hotTable?.hotInstance?.getSelectedLast();
    if (last == undefined) return;
    let i = last[0];

    //console.log(this.data1)
    this.styles.splice(i, 0, [...this.styles[i]]);
    this.data1.splice(i, 0, [...this.data1[i]]);
    //console.log(this.data1)
    this.hotTable?.hotInstance?.alter('insert_row_below', i, 1);
    
    const row = this.data1[i+1];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j];

        this.setCell(cell, i+1, j);

      } 

    //console.log(this.data)

    //this.hotTable?.hotInstance?.alter('remove_row', last[0], 1);
  }
  delRow() {
    let last = this.hotTable?.hotInstance?.getSelectedLast();
    if (last == undefined) return;
    let i = last[0];

    //console.log(this.data1)
    this.styles.splice(i, 1);
    this.data1.splice(i, 1);
    //console.log(this.data1) 
    this.hotTable?.hotInstance?.alter('remove_row', last[0], 1);
  }
  reset() {
    this.hotTable?.hotInstance?.updateSettings({data:[], columns:[]});

    this.hotTable?.hotInstance?.clear();
    this.styles = [];
    this.data1 = []; 
    this.data = [];

    //console.log(this.data)
  }

  selectedFileChange(e: any) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = async (e: any) => {
      const buffer = e.target.result;
        await this.processBuffer(buffer);

    };

    reader.readAsArrayBuffer(file);
  }

}

