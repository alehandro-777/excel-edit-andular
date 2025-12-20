import { Injectable, signal, computed } from '@angular/core';

@Injectable({ providedIn: 'root' })
export class HttpBusyService {

  private counter = signal(0);

  readonly busy = computed(() => this.counter() > 0);

  start() {
    this.counter.update(v => v + 1);
  }

  stop() {
    this.counter.update(v => Math.max(0, v - 1));
  }
}
