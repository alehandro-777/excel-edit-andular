import { Component, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';

import { registerAllModules } from "handsontable/registry";
import { EditExcel } from "./edit-excel/edit-excel";
import { EditExcel1 } from './edit-excel1/edit-excel1';
import { EditExcel2 } from './edit-excel2/edit-excel2';
import { HttpBusyService } from './http-busy.service';
import { ErrorService } from './error.service';
import { EditExcel3 } from './edit-excel3/edit-excel3';


registerAllModules();

@Component({
  selector: 'app-root',
  imports: [RouterOutlet, EditExcel3,],  // EditExcel1, EditExcel2],
  templateUrl: './app.html',
  styleUrl: './app.scss'
})
export class App {
  constructor(public busy: HttpBusyService, public errors: ErrorService) {}
  protected readonly title = signal('editExcel-1');
}
