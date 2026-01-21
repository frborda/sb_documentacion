export class ClientError extends Error {
  constructor (message, statusCode, metadata = []) {
    super(message)

    this.statusCode = statusCode ?? 417
    this.metadata = metadata

    // prevenir que esta clase se muestre en el stacktrace
    Error.captureStackTrace(this, this.constructor)
  }
}

export class ServerError extends Error {
  constructor (message, metadata = []) {
    super(message)

    this.statusCode = 500
    this.metadata = metadata

    // prevenir que esta clase se muestre en el stacktrace
    Error.captureStackTrace(this, this.constructor)
  }
}
