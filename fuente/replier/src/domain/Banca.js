export class Banca {
  constructor (numero) {
    this.numero = numero
    this.bascula = false // true (sit_on) o false (sit_off)
    this.identificacion = null // null o cuil

    this._updateLastModified()
  }

  activarBascula () {
    if (this.bascula) {
      return false
    }
    this.bascula = true

    this._updateLastModified()
    return true
  }

  desactivarBascula () {
    if (!this.bascula) {
      return false
    }
    this.bascula = false

    this._updateLastModified()
    return true
  }

  identificar (cuil) {
    cuil = parseInt(cuil)
    if (this.identificacion === cuil) {
      return false
    }

    this.identificacion = cuil

    this._updateLastModified()
    return true
  }

  desidentificar () {
    if (!this.identificacion) {
      return false
    }
    const cuilPrevio = this.identificacion
    this.identificacion = null

    this._updateLastModified()
    return cuilPrevio
  }

  toJSON () {
    return {
      numero: this.numero,
      bascula: this.bascula,
      identificacion: this.identificacion
    }
  }

  _updateLastModified () {
    this.lastModified = new Date().toISOString()
  }
}
