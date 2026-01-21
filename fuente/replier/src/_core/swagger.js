import fs from 'fs'
import path from 'path'
import swaggerJsDoc from 'swagger-jsdoc'
import config from '../config'
import pkg from '../../package.json'

const srcPath = path.join(__dirname, '..')
const apiPath = path.join(srcPath, 'controllers')

const getControllerFiles = (p) => {
  const subpaths = fs.readdirSync(p).map(file => path.join(p, file))
  const files = []
  for (const sp of subpaths) {
    if (fs.lstatSync(sp).isDirectory()) {
      files.push(...getControllerFiles(sp))
    } else {
      files.push(sp)
    }
  }
  return files
}

const apiFiles = getControllerFiles(apiPath)

const serviceURL = new URL('/api', config.service.url)

export const swaggerSpec = swaggerJsDoc({
  definition: {
    // https://github.com/OAI/OpenAPI-Specification/blob/main/versions/3.0.3.md
    openapi: '3.0.3',
    info: {
      title: `${pkg.name}`,
      summary: pkg.description,
      version: pkg.version
    },
    servers: [
      {
        url: serviceURL.href,
        description: `Servicio [${config.env}]`
      }
    ],
    components: {
      securitySchemes: {
        ServiceApiKey: {
          type: 'apiKey',
          name: 'X-API-KEY',
          in: 'header'
        },
        WebhookApiKey: {
          type: 'apiKey',
          name: 'X-API-KEY',
          in: 'header'
        }
      }
    }
  },
  apis: apiFiles
})
