import { onMounted, onUnmounted, reactive, computed } from 'vue'
import _debounce from 'lodash/debounce'

export default () => {
  const state = reactive({
    width: 0
  })

  const size = computed(() => {
    const width = state.width

    const _isMin = a => width >= a

    const screenSmMin = 640
    const screenMdMin = 768
    const screenLgMin = 1024
    const screenXlMin = 1280

    const breakpoints = {
      sm: _isMin(screenSmMin),
      md: _isMin(screenMdMin),
      lg: _isMin(screenLgMin),
      xl: _isMin(screenXlMin)
    }

    return breakpoints
  })

  const resizeCallback = _debounce(() => {
    state.width = window.innerWidth
  }, 10)

  onMounted(() => {
    state.width = window.innerWidth
    window.addEventListener('resize', resizeCallback)
  })

  onUnmounted(() => {
    window.removeEventListener('resize', resizeCallback)
  })

  return {
    size
  }
}
