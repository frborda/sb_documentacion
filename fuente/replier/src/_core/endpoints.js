import config from '../config'
import { ClientError } from './errors'

export const endpoint = fn => (req, res, next) => {
  fn(req, res, next).catch(next)
}

export const requireServiceApiKey = (req, res, next) => {
  const headerApiKey = req.header('X-API-KEY')
  if (!headerApiKey || headerApiKey !== config.service.apiKey) {
    throw new ClientError('Autenticación inválida o vencida', 401)
  }
  next()
}
