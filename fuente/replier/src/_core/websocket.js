import WebSocket from 'ws'
import queryString from 'query-string'
import { newID } from './id'

function serializeMessage (type, payload) {
  return JSON.stringify({ type, payload: payload ?? {} })
}

class WebsocketConnection {
  constructor (websocket, params) {
    this.id = newID()
    this.websocket = websocket
    this.params = params

    // if (this.params.rooms) {
    //   // unirse a distintos rooms
    // }
    // if (this.params.user) {
    //   // conexion no anonima
    // }

    // detectar aca el contexto: ip, device, os, timestamp, etc

    this.websocket.on('message', (msg) => {
      try {
        this.onMessage(JSON.parse(msg))
      } catch (err) {
        this.onMessage(msg)
      }
    })

    this.onMessage = (msg) => {
      if (msg.type === 'ping') {
        this.send('pong')
      }
    }
  }

  send (type, payload) {
    this.websocket.send(serializeMessage(type, payload))
  }

  rawSend (msg) {
    this.websocket.send(msg)
  }
}

// https://medium.com/hackernoon/implementing-a-websocket-server-with-node-js-d9b78ec5ffa8
class WebsocketServer {
  constructor (expressServer) {
    this.connections = new Map() // id => WebsocketConnection
    this.server = new WebSocket.Server({
      noServer: true,
      path: '/ws'
    })

    expressServer.on('upgrade', (request, socket, head) => {
      this.server.handleUpgrade(request, socket, head, (websocket) => {
        // nueva conexion
        const [, connQuery] = request?.url?.split('?')
        const conn = new WebsocketConnection(websocket, queryString.parse(connQuery))

        this.connections.set(conn.id, conn)
      })
    })

    // TODO ante desconexion, eliminar de this.connections
  }

  // enviar mensaje a todos las conexiones
  broadcast (type, payload) {
    const msg = serializeMessage(type, payload)
    for (const conn of this.connections.values()) {
      conn.rawSend(msg)
    }
  }
}

export function websocket (expressServer) {
  return new WebsocketServer(expressServer)
}
