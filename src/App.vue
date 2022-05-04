<template>
  <div class="app">
    <div class="app__header">
      <div class="header-container">
        <base-theme-toggle class="app__themeToggle" />
      </div>
    </div>

    <div class="app__body">
      <router-view />
    </div>
  </div>
</template>

<style lang="scss">
@import '~assets/App';
</style>
