import axios from 'axios'
import https from 'https'
import config from '../config'

// solo utilizamos certificados auto firmados en entorno local
const rejectUnauthorized = config.env !== 'local'

export const http = axios.create({
  httpsAgent: new https.Agent({
    rejectUnauthorized
  })
})
