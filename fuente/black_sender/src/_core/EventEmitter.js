import { Subject, Subscription } from 'rxjs'

export class EventEmitter extends Subject {
  constructor () {
    super()
    this.subject = new Subject()
  }

  emit (type, payload) {
    super.next({ type, payload })
  }

  /**
   * Registers handlers for events emitted by this instance.
   * @param observerOrNext When supplied, a custom handler for emitted events, or an observer
   *     object
   * @param error When supplied, a custom handler for an error notification from this emitter.
   * @param complete When supplied, a custom handler for a completion notification from this
   *     emitter.
   */
  subscribe (observerOrNext, error, complete) {
    let nextFn = observerOrNext
    let errorFn = error || (() => null)
    let completeFn = complete

    if (observerOrNext && typeof observerOrNext === 'object') {
      const observer = observerOrNext
      nextFn = observer.next?.bind(observer)
      errorFn = observer.error?.bind(observer)
      completeFn = observer.complete?.bind(observer)
    }

    const sink = super.subscribe({ next: nextFn, error: errorFn, complete: completeFn })

    if (observerOrNext instanceof Subscription) {
      observerOrNext.add(sink)
    }

    return sink
  }
}
