import { endpoint } from '../_core'

export const hemiciclo = endpoint(async (req, res, next) => {
  const recinto = req.app.locals.recinto
  res.render('hemiciclo', {
    bancas: recinto.toJSON(),
    totalBasculasActivas: recinto.nroBasculasActivas,
    totalIdentificaciones: recinto.nroIdentificados
  })
})
