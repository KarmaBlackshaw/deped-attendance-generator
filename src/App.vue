<template>
  <div class="app">
    <div class="app__body">
      <router-view />
    </div>
  </div>
</template>

<script>
export default {
}
</script>
<style lang="scss">
.app {
  min-height: 100vh;
  height: 100%;
  padding: 100px;

  @include smAndBelow {
    padding: 10px;
  }

}
</style>
