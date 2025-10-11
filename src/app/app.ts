import { Component, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { HttpClient } from '@angular/common/http';
import { registerAllModules } from "handsontable/registry";
import { EditExcel } from "./edit-excel/edit-excel";
import { EditExcel1 } from './edit-excel1/edit-excel1';
import { EditExcel2 } from './edit-excel2/edit-excel2';


registerAllModules();

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, EditExcel, EditExcel1, EditExcel2],
  templateUrl: './app.html',
  styleUrl: './app.scss'
})
export class App {
  protected readonly title = signal('editExcel-1');
}
