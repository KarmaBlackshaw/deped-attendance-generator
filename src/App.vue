<template>
  <div class="app">
    <div class="app__background"></div>

    <div class="app__header">
      <TheNavigation />
    </div>

    <div class="app__body">
      <router-view />
    </div>
  </div>
</template>

<script>
import { provide } from 'vue'

// composables
import useBreakpoint from '@/composables/useBreakpoint'
import useTheme from '@/composables/useTheme'

// components
import TheNavigation from '@/components/the-navigation/TheNavigation'

export default {
  components: {
    TheNavigation
  },

  setup () {
    provide('breakpoint', useBreakpoint())
    provide('theme', useTheme())
  }
}
</script>

<style lang="scss">
@import '~assets/App';
</style>
