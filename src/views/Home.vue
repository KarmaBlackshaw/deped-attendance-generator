<template>
  <div class="home">
    <div class="card">
      <div class="card__title">
        <span class="card__title--accent">DepEd</span> Attendance Generator
      </div>

      <div class="card__subtitle">
        <p>Lorem ipsum dolor, sit amet consectetur adipisicing elit. At accusamus iure incidunt! Rem, eveniet. Sunt sequi placeat soluta optio labore mollitia porro quidem praesentium veniam, cupiditate nihil hic saepe exercitationem?</p>

        <p>Lorem ipsum dolor, sit amet consectetur adipisicing elit. At accusamus iure incidunt! Rem, eveniet. Sunt sequi placeat soluta optio labore mollitia porro quidem praesentium veniam, cupiditate nihil hic saepe exercitationem?</p>
      </div>

      <div class="card__subtitle">
        <h5>Click the button below to get started.</h5>
      </div>

      <div class="card__subtitle btn-container">
        <button class="btn">
          Upload
        </button>
      </div>

      <div class="card__subtitle">
        Made with 💙 by <a
          href="https://jeash.tech"
          target="_blank"
        >Ernie Jeash</a>
      </div>
    </div>
  </div>
</template>

<script>
// @ is an alias to /src
import HelloWorld from '@/components/HelloWorld.vue'

export default {
  name: 'Home',
  components: {
    HelloWorld
  }
}
</script>

<style lang="scss" scoped>
.home {
  display: flex;
  justify-content: center;
}

.card {
  box-shadow: rgba(0, 0, 0, 0.05) 0px 6px 24px 0px, rgba(0, 0, 0, 0.08) 0px 0px 0px 1px;
  border-radius: 20px;
  padding-top: min(30px, 10%);
  padding-bottom: min(30px, 10%);
  padding-left: min(60px, 10%);
  padding-right: min(60px, 10%);
  width: 100%;
  max-width: 800px;

  @include smAndBelow {
    box-shadow: none;
  }

  .card__title {
    font-size: 2.5em;
    text-align: center;
    font-weight: bold;
    margin-bottom: 20px;
    color: $gray-600;

    @include smAndBelow {
      font-size: 2em;
    }

    .card__title--accent {
      color: $blue-500;
    }
  }

  .card__subtitle {
    line-height: 150%;
    color: $gray-500;
    text-align: center;

    p {
      margin-bottom: 20px;
    }

    h5 {
      font-weight: 600;
    }
  }
}

.btn-container {
  display: flex;
  justify-content: center;
  margin-top: 50px;
  margin-bottom: 50px;

  .btn {
    background-color: $blue-500;
    border: 0 solid #E5E7EB;
    box-sizing: border-box;
    color: #fff;
    display: flex;
    font-size: 1rem;
    font-weight: 700;
    justify-content: center;
    line-height: 1.75rem;
    padding: .75rem 1.65rem;
    position: relative;
    text-align: center;
    text-decoration: none #000000 solid;
    text-decoration-thickness: auto;
    width: 100%;
    max-width: 260px;
    position: relative;
    cursor: pointer;
    transform: rotate(-2deg);
    user-select: none;
    -webkit-user-select: none;
    touch-action: manipulation;

    @media (min-width: 768px) {
      padding: .75rem 3rem;
      font-size: 1.25rem;
    }

    &:focus {
      outline: 0;
    }

    &:after {
      content: '';
      position: absolute;
      border: 1px solid #000000;
      bottom: 4px;
      left: 4px;
      width: calc(100% - 1px);
      height: calc(100% - 1px);
      transition-duration: 200ms;
      transition-timing-function: ease;
      transition-property: bottom left;
    }

    &:hover:after {
      bottom: 2px;
      left: 2px;
    }

  }
}
</style>
