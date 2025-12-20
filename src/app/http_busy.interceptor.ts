import { Injectable } from '@angular/core';
import {
  HttpInterceptor,
  HttpRequest,
  HttpHandler,
  HttpEvent
} from '@angular/common/http';
import { Observable, finalize } from 'rxjs';
import { HttpBusyService } from './http-busy.service';

@Injectable()
export class HttpBusyInterceptor implements HttpInterceptor {

  constructor(private busy: HttpBusyService) {}

  intercept(req: HttpRequest<any>, next: HttpHandler): Observable<HttpEvent<any>> {
    this.busy.start();

    return next.handle(req).pipe(
      finalize(() => this.busy.stop())
    );
  }
}
