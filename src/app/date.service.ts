import { Injectable } from '@angular/core';
import { format,
  parseISO,
  addDays,
  addHours,
  addMinutes,
  subDays,
  subHours,
  subMinutes,
  formatISO,
  isValid,
  differenceInHours,
  startOfDay
} from 'date-fns';

@Injectable({ providedIn: 'root' })
export class DateService {

  // ======================
  // Base helpers
  // ======================

  private parse(iso: string): Date {
    const date = parseISO(iso);
    if (!isValid(date)) {
      throw new Error(`Invalid ISO date: ${iso}`);
    }
    return date;
  }

  private toIso(date: Date): string {
    return formatISO(date);
  }

  // ======================
  // Add
  // ======================

  addDays(iso: string, days: number): string {
    return addDays(this.parse(iso), days).toISOString();
  }

  addHours(iso: string, hours: number): string {
    return addHours(this.parse(iso), hours).toISOString();
  }

  addMinutes(iso: string, minutes: number): string {
    return addMinutes(this.parse(iso), minutes).toISOString();
  }

  // ======================
  // Subtract
  // ======================

  subDays(iso: string, days: number): string {
    return subDays(this.parse(iso), days).toISOString();
  }

  subHours(iso: string, hours: number): string {
    return subHours(this.parse(iso), hours).toISOString();
  }

  subMinutes(iso: string, minutes: number): string {
    return subMinutes(this.parse(iso), minutes).toISOString();
  }

  diffInHours(a: string, b: string): number {
    return differenceInHours(this.parse(a), this.parse(b));
  }


  startOfDay(iso: string): string {
    return startOfDay(this.parse(iso)).toISOString();
  }
  
  now(): string {
    return formatISO(new Date());
  }
  
  //dd.MM.yyyy HH:mm
  format(dt: Date, fmt:string): string {
    return format(dt, fmt);
  }

}
