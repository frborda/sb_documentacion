// eslint-disable-next-line no-unused-vars
const HEMICICLO = new function () {
  const HEMICICLO = this
  const svgUrl = 'assets/img/hemiciclo.svg'

  HEMICICLO.dom = {
    bancas: new Array(258)
  }

  let _initialized = false

  HEMICICLO.init = (elementSelector) => {
    if (_initialized) {
      return Promise.resolve()
    }
    return new Promise((resolve, reject) => {
      HEMICICLO.dom.element = document.querySelector(elementSelector)
      if (!HEMICICLO.dom.element) {
        reject(new Error(`Elemento ${elementSelector} no encontrado`))
        return
      }
      HEMICICLO.dom.element.classList.add('hemiciclo--initializing')

      Promise.all([
        _fetchSVG()
      ]).then(() => {
        render()
        _initialized = true
        HEMICICLO.dom.element.classList.remove('hemiciclo--initializing')
        resolve()
      })
    })
  }

  const _fetchSVG = () => fetch(svgUrl)
    .then(response => response.text())
    .then(svgText => {
      HEMICICLO.dom.element.innerHTML = svgText
      HEMICICLO.dom.svg = HEMICICLO.dom.element.querySelector('svg')
      const bancas = HEMICICLO.dom.svg.querySelectorAll('foreignObject > [butaca]')
      bancas.forEach(el => {
        const bancaNro = el.getAttribute('butaca')
        HEMICICLO.dom.bancas[bancaNro] = el
      })
    })

  const render = () => {
    HEMICICLO.dom.bancas.forEach((bancaEl, bancaNro) => {
      // Agregar color y borde por defecto
      HEMICICLO.setearColorBanca(bancaNro, 'gray')
      HEMICICLO.setearBordeBanca(bancaNro, 'gray')

      // Enumerar banca
      const el = document.createElement('p')
      el.className = 'banca__numero'
      el.innerText = bancaNro
      bancaEl.append(el)
    })

    // Ocultar la banca extra
    HEMICICLO.dom.bancas[257].style.display = 'none'
  }

  HEMICICLO.setearColorBanca = (banca, color) => {
    const bancaSVG = HEMICICLO.dom.bancas[banca]
    if (!bancaSVG) {
      return
    }
    bancaSVG.style.backgroundColor = color
  }

  HEMICICLO.setearBordeBanca = (banca, color) => {
    const bancaSVG = HEMICICLO.dom.bancas[banca]
    if (!bancaSVG) {
      return
    }
    bancaSVG.style.border = `2px solid ${color}`
  }
}()
