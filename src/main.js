import { createApp } from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'

// styles
import '@/assets/scss/app/_@index.scss'

// libs
import upperFirst from 'lodash/upperFirst'
import camelCase from 'lodash/camelCase'

// instance
const app = createApp(App)

const requireComponent = require.context('./components', true, /Base[A-Z]\w+\.(vue|js)$/)

requireComponent.keys().forEach(fileName => {
  // Get component config
  const componentConfig = requireComponent(fileName)

  // Get PascalCase name of component
  const componentName = upperFirst(camelCase(fileName.split('/').pop().replace(/\.\w+$/, '')))

  app.component(componentName, componentConfig.default)
})

app.use(store)

app.use(router)

app.mount('#app')
