import { WebhookPublisher } from '../_core'
import { performance } from 'perf_hooks'
import { Mutex } from 'async-mutex'

const MIN_REPLY_MS = 1000 // 1s
const THROTTLE_MS = 1000 // 1s

export class Replier {
  constructor (events, recinto, replicarCadaXMilisegundos) {
    this.events = events
    this.recinto = recinto
    this.publisher = new WebhookPublisher('replier')

    this.throttleMilliseconds = THROTTLE_MS
    this._replierMutex = new Mutex()
    this._replying = false
    this._replyAfterWait = false
    this._lastReplyTime = performance.now()
    this._paused = false
    this._unpauseInterval = null
    this.autoUnpauseMilliseconds = 7000

    // configurar RECINTO_WEBHOOK_REPLY_MS
    // 0 replica ante evento
    if (replicarCadaXMilisegundos) {
      this.replyOnEvent = false
      if (replicarCadaXMilisegundos < MIN_REPLY_MS) {
        console.log('Replier: replicarCadaXMilisegundos configurado en', replicarCadaXMilisegundos, 'milisegundos, minimo es', MIN_REPLY_MS)
        replicarCadaXMilisegundos = MIN_REPLY_MS
      }
      this.replyMilliseconds = replicarCadaXMilisegundos
      console.log('Replier: replica cada', this.replyMilliseconds, 'milisegundos')
    } else {
      this.replyOnEvent = true
      this.replyMilliseconds = 0
      console.log('Replier: replica solo ante cambios de Recinto')
    }
  }

  isPaused () {
    return this._paused
  }

  pause (autoUnpause = false) {
    if (this._paused) {
      if (autoUnpause && this._unpauseInterval === null) {
        console.log('Replier: auto unpause en', this.autoUnpauseMilliseconds, 'milisegundos')
        clearInterval(this._unpauseInterval)
        this._unpauseInterval = setInterval(() => {
          this.unpause()
        }, this.autoUnpauseMilliseconds)
      }
      return
    }
    this._paused = true
    console.log('Replier: paused')
    if (!autoUnpause) {
      return
    }
    console.log('Replier: auto unpause en', this.autoUnpauseMilliseconds, 'milisegundos')
    clearInterval(this._unpauseInterval)
    this._unpauseInterval = setInterval(() => {
      this.unpause()
    }, this.autoUnpauseMilliseconds)
  }

  unpause () {
    if (!this._paused) {
      return
    }
    clearInterval(this._unpauseInterval)
    this._unpauseInterval = null
    console.log('Replier: unpaused')
    this._paused = false
    this.reply()
  }

  to (url, apiKey) {
    this.publisher.subscribe(url, apiKey)
  }

  async reply () {
    if (this._paused) {
      return
    }
    const wait = await this._replierMutex.runExclusive(async () => {
      if (this._replying) {
        this._replyAfterWait = true
        return true
      }
      this._replying = true
      return false
    })

    // otra ejecucion se encuentra en paralelo
    if (wait) return

    // limitar la cantidad de replies
    const now = performance.now()
    const diff = now - this._lastReplyTime
    const throttle = diff < this.throttleMilliseconds
    if (throttle) {
      if (!this._throttling) {
        this._throttling = true
        setTimeout(() => {
          this._throttling = false
          this.reply()
        }, this.throttleMilliseconds)
      }
      this._replyAgain(false)
      return
    }

    // reply
    this._lastReplyTime = now
    this.publisher.publish('RECINTO_ESTADO', this.recinto.serializeToSync())
      .then(() => this._replyAgain())
      .catch(() => this._replyAgain())
  }

  async _replyAgain () {
    const replyAgain = await this._replierMutex.runExclusive(async () => {
      this._replying = false
      const replyAgain = this._replyAfterWait
      this._replyAfterWait = false
      return replyAgain
    })
    if (replyAgain) {
      this.reply()
    }
  }

  run () {
    if (this.replyOnEvent) {
      // Ante un evento se replica
      this.events.subscribe(e => {
        this.reply()
      })
    } else {
      // Replica cada x tiempo
      if (this.interval) return
      this.interval = setInterval(() => {
        this.reply()
      }, this.replyMilliseconds)
    }

    // reply one time on init
    this.reply()
  }
}
