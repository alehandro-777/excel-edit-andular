import { Injectable, signal } from '@angular/core';

export interface AppError {
  message: string;
  status?: number;
}

@Injectable({ providedIn: 'root' })
export class ErrorService {
  readonly error = signal<AppError | null>(null);

  show(error: AppError) {
    this.error.set(error);
  }

  clear() {
    this.error.set(null);
  }
}
