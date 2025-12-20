import { Injectable } from '@angular/core';
import {
  HttpInterceptor,
  HttpRequest,
  HttpHandler,
  HttpEvent,
  HttpErrorResponse
} from '@angular/common/http';
import { Observable, catchError, throwError } from 'rxjs';
import { ErrorService } from './error.service';

@Injectable()
export class HttpErrorInterceptor implements HttpInterceptor {

  constructor(private errors: ErrorService) {}

  intercept(req: HttpRequest<any>, next: HttpHandler): Observable<HttpEvent<any>> {
    return next.handle(req).pipe(
      catchError((err: HttpErrorResponse) => {

        let message = 'Неизвестная ошибка';

        if (err.status === 0) {
          message = 'Нет соединения с сервером';
        } else if (err.status >= 500) {
          message = 'Ошибка сервера';
        } else if (err.status === 401) {
          message = 'Требуется авторизация';
        } else if (err.error?.message) {
          message = err.error.message;
        }

        this.errors.show({
          message,
          status: err.status
        });

        return throwError(() => err);
      })
    );
  }
}
