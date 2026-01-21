import { http } from './http'

export class WebhookPublisher {
  constructor (name) {
    this.name = name
    this.webhooks = []
  }

  subscribe (url, apiKey) {
    this.webhooks.push({ url, apiKey })
    console.log(`publisher ${this.name}: webhook suscripto ${url}`)
  }

  publish (type, payload, callback) {
    if (!callback) {
      callback = res => {} // noop
    }
    const promises = []
    for (const w of this.webhooks) {
      const p = http.post(w.url, {
        type: type,
        payload: payload
      }, {
        timeout: 15000, // 15s
        headers: {
          'Content-Type': 'application/json',
          'X-API-KEY': w.apiKey
        }
      })
        .then(callback)
        .catch(err => {
          const message = err.response?.data?.error?.message ?? err.message
          console.log(`Error publishing to ${w.url}:`, message)
        })
      promises.push(p)
    }
    return Promise.all(promises)
  }
}
