import { createApp } from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'

// styles
import '@/assets/scss/app/_@index.scss'

// plugins
import baseComponents from '@/plugins/baseComponents'

// instance
const app = createApp(App)

app.use(store)

app.use(baseComponents)

app.use(router)

app.mount('#app')
