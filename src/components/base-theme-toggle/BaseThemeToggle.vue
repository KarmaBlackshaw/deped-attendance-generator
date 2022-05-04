<template>
  <div class="base-theme-toggle">
    <input
      id="checkbox"
      type="checkbox"
      class="checkbox"
      :checked="isDark"
      @change="toggleDark()"
    >

    <label
      for="checkbox"
      class="label"
    >
      <img
        src="./assets/images/moon.png"
        alt=""
      >

      <img
        src="./assets/images/sun.png"
        alt=""
      >

      <div class="ball"></div>
    </label>
  </div>
</template>

<script>
import { useDark, useToggle } from '@vueuse/core'

export default {
  name: "BaseThemeToggle",

  setup () {
    const isDark = useDark({
      attribute: 'data-theme',
      valueDark: 'dark',
      valueLight: 'light'
    })

    const toggleDark = useToggle(isDark)

    return {
      isDark,
      toggleDark
    }
  }
}
</script>

<style lang="scss" scoped>
@import './assets/scss/BaseThemeToggle.scss';
</style>