import { Injectable } from '@angular/core';
import {
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
    return this.toIso(addDays(this.parse(iso), days));
  }

  addHours(iso: string, hours: number): string {
    return this.toIso(addHours(this.parse(iso), hours));
  }

  addMinutes(iso: string, minutes: number): string {
    return this.toIso(addMinutes(this.parse(iso), minutes));
  }

  // ======================
  // Subtract
  // ======================

  subDays(iso: string, days: number): string {
    return this.toIso(subDays(this.parse(iso), days));
  }

  subHours(iso: string, hours: number): string {
    return this.toIso(subHours(this.parse(iso), hours));
  }

  subMinutes(iso: string, minutes: number): string {
    return this.toIso(subMinutes(this.parse(iso), minutes));
  }

  diffInHours(a: string, b: string): number {
    return differenceInHours(this.parse(a), this.parse(b));
  }


  startOfDay(iso: string): string {
    return this.toIso(startOfDay(this.parse(iso)));
  }
  
  now(): string {
    return formatISO(new Date());
  }

}
