import { AfterViewInit, Component, signal, ViewChild } from '@angular/core';
import { ColumnSettings, GridSettings,  HotTableComponent, HotTableModule,  } from '@handsontable/angular-wrapper';
import Handsontable from 'handsontable';
import { textRenderer } from 'handsontable/renderers';
import dayjs from 'dayjs';

@Component({
  selector: 'app-edit-excel2',
  standalone: true,
  imports: [HotTableModule],
  templateUrl: './edit-excel2.html',
  styleUrl: './edit-excel2.scss'
})
export class EditExcel2 {


  @ViewChild(HotTableComponent, { static: false }) hotTable!: HotTableComponent;

  readonly data = [];

  readonly apidata = [
    ['2017', 'Honda', 10, "2025-10-05", true, "10:30:00",],
    ['2018', 'Toyota', 20, "2025-10-05", true, "10:30:00",],
    ['2019', 'Nissan', 30, "2025-10-05", false, "10:30:00",],
  ];

  readonly columnsCfg: Handsontable.ColumnSettings[] = [ 
      {},
      {},
      {},
      {},
      {},
      {},
];

  readonly gridSettings: GridSettings = {

    height: 'auto',
    colHeaders: true,
    rowHeaders: true,
    autoWrapRow: true,
    autoWrapCol: true,
    //select остается в таблице при смене фокуса
    outsideClickDeselects: false,
  };


  loadData()  {
    this.hotTable.hotInstance?.updateSettings({ columns: this.columnsCfg });
    
    this.hotTable?.hotInstance?.setDataAtCell(0, 0, "ABCDEF");

    //-className -editor -renderer -type -source
    //this.hotTable?.hotInstance?.setCellMeta(0, 0, "readOnly",  true);

    this.hotTable?.hotInstance?.setDataAtCell(0, 1, "AAAAAA");
    this.hotTable?.hotInstance?.setCellMeta(0, 1, "type",  'select');
    this.hotTable?.hotInstance?.setCellMeta(0, 1, "selectOptions",  ['Kiaaaaaaaaaaaaaaaaa', 'Nissan', 'Toyota', 'Honda']);

    this.hotTable?.hotInstance?.setDataAtCell(1, 1, "red");
    this.hotTable?.hotInstance?.setCellMeta(1, 1, "type",  'dropdown');
    this.hotTable?.hotInstance?.setCellMeta(1, 1, "source",  ['yellow', 'red', 'orange', 'green', 'blue', 'gray', 'black', 'white']);

    //number
    this.hotTable?.hotInstance?.setCellMeta(1, 2, "type",  'numeric');
    this.hotTable?.hotInstance?.setCellMeta(1, 2, "renderer",  this.numericRenderer.bind(this));
    this.hotTable?.hotInstance?.setCellMeta(1, 2, "numericFormat",  { pattern: '0\u202f0.00', culture: 'ru-RU' });//https://numbrojs.com/languages.html
    this.hotTable?.hotInstance?.setDataAtCell(1, 2, 436.45);

    //checkbox
    this.hotTable?.hotInstance?.setDataAtCell(1, 4, "YES");
    this.hotTable?.hotInstance?.setCellMeta(1, 4, "type",  'checkbox');
    this.hotTable?.hotInstance?.setCellMeta(1, 4, "checkedTemplate",  'YES');
    this.hotTable?.hotInstance?.setCellMeta(1, 4, "uncheckedTemplate",  'NO');
    this.hotTable?.hotInstance?.setCellMeta(1, 4, "label",  { position: 'after', value: 'In black? '});
  
    //date
    this.hotTable?.hotInstance?.setCellMeta(0, 3, "type",  'date');
    this.hotTable?.hotInstance?.setCellMeta(0, 3, "dateFormat",  "DD.MM.YYYY");
    this.hotTable?.hotInstance?.setCellMeta(0, 3, "correctFormat",  true);
    this.hotTable?.hotInstance?.setCellMeta(0, 3, "defaultDate",  "01.12.2000");
    //this.hotTable?.hotInstance?.setCellMeta(0, 3, "allowInvalid",  false);
    this.hotTable?.hotInstance?.setDataAtCell(0, 3, "2025-12-01");

    //time
    this.hotTable?.hotInstance?.setCellMeta(0, 5, "type",  'time');
    this.hotTable?.hotInstance?.setCellMeta(0, 5, "timeFormat",  'HH:mm');
    this.hotTable?.hotInstance?.setCellMeta(0, 5, "correctFormat",  true);
    this.hotTable?.hotInstance?.setDataAtCell(0, 5, "08:00");

    //this.hotTable?.hotInstance?.render();
    //this.hotTable.hotInstance?.updateSettings({ mergeCells: this.mergeCells });

    //console.log(this.data);
  }

  numericRenderer(instance: Handsontable.Core, TD: HTMLTableCellElement, row: number, column: number, prop: string | number, 
    value: any, cellProperties: Handsontable.CellProperties): void  {

    let value1 = this.formatValueByNumFmt(value , "0,0.00");

    //console.log(value, typeof value)
/*
    if (isNaN(value)) {
      TD.style.backgroundColor = '#fff3cd'; // жёлтая ячейка
    } else if (value > 100) {
      TD.style.backgroundColor = '#dc3545'; // красная ячейка
      TD.style.color = 'white';
      TD.style.fontWeight = 'bold';
    }
*/
    textRenderer.apply(this, [instance, TD,row, column,prop, value1, cellProperties]); 
  }

  formatValueByNumFmt(value: any, numFmt: string, locale = 'fr-FR'): string {
  

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

  addRow()  {
    let last = this.hotTable?.hotInstance?.getSelectedLast();

    if (last == undefined) return;

    this.hotTable?.hotInstance?.alter('insert_row_below', last[0], 1);

    this.hotTable?.hotInstance?.render();
    console.log(this.data);
  }  

  deleteRow()  {
    let last = this.hotTable?.hotInstance?.getSelectedLast();

    if (last == undefined) return;

    this.hotTable?.hotInstance?.alter('remove_row', last[0], 1);

    this.hotTable?.hotInstance?.render();
    console.log(this.data);
  } 

}


