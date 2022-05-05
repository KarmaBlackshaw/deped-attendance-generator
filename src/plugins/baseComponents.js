
// libs
import upperFirst from 'lodash/upperFirst'
import camelCase from 'lodash/camelCase'

// helpers
function pipe (initial, fns) {
  return fns.reduce((v, f) => f(v), initial)
}

export default {
  install: app => {
    const requireComponent = require.context('../components', true, /Base[A-Z]\w+\.(vue|js)$/)

    requireComponent.keys().forEach(fileName => {
      // Get component config
      const componentConfig = requireComponent(fileName)

      // Get PascalCase name of component
      const componentName = pipe(fileName, [
        val => val.split('/'),
        val => val.pop(),
        val => val.replace(/\.\w+$/, ''),
        val => camelCase(val),
        val => upperFirst(val)
      ])

      app.component(componentName, componentConfig.default)
    })
  }
}