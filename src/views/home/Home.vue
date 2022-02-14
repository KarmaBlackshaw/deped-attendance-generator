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

      <div class="card__subtitle upload-container">
        <button
          class="btn upload__btn"
          @click="handleClickBrowseFiles"
        >
          Upload
        </button>

        <input
          id="input-file"
          ref="input-file"
          class="upload__input-file"
          type="file"
          accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
          @change="handleFileSelect"
        >
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
// libs
import * as XLSX from 'xlsx'

// composables
import useHelpers from '@/utils/useHelpers'

// helpers
const readFile = files => new Promise(resolve => {
  const reader = new FileReader()

  reader.onload = e => {
    const data = e.target.result
    const workbook = XLSX.read(data, { type: 'binary' })
    const wsname = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[wsname]
    const json = XLSX.utils.sheet_to_json(worksheet, { header: 2 })

    resolve(json)
  }

  reader.readAsBinaryString(files)
})

export default {
  name: 'Home',

  methods: {
    async handleClickBrowseFiles () {
      const { waitUntilElementIsLoaded } = useHelpers()

      const inputFile = await waitUntilElementIsLoaded('#input-file')

      inputFile.click()
    },

    async handleFileSelect (e) {
      const files = e.target.files

      const data = await readFile(files[0])
      e.target.value = null

      console.log(JSON.stringify(data))
    }
  }
}
</script>

<style lang="scss" scoped>
@import './styles/home';
</style>
