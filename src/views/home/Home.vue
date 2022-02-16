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
import XLSX from 'sheetjs-style'

// composables
import useHelpers from '@/utils/useHelpers'

// libs
import _toString from 'lodash/toString'

// helpers
const readFile = files => new Promise((resolve, reject) => {
  const reader = new FileReader()

  reader.onload = e => {
    const data = e.target.result
    const workbook = XLSX.read(data, { type: 'binary' })
    const wsname = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[wsname]
    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

    resolve(json)
  }

  reader.onerror = e => resolve(e)

  reader.readAsBinaryString(files)
})

export default {
  name: 'Home',

  created () {
    this.handleFileSelect()
  },

  methods: {
    async handleClickBrowseFiles () {
      const { waitUntilElementIsLoaded } = useHelpers()

      const inputFile = await waitUntilElementIsLoaded('#input-file')

      inputFile.click()
    },

    async handleFileSelect (e) {
      // const files = e.target.files

      // const data = await readFile(files[0])
      // e.target.value = null

      const workbook = XLSX.utils.book_new()

      const form = [
        [null, 'CIVIL SERVICE FORM NO. 48', null, null, null, null, null, null, null],
        [null, 'DAILY TIME RECORD', null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, 'LAST NAME, GIVEN NAME M.I.', null, null, null, null, null, null, null],
        [null, '(name)', null, null, null, null, null, null, null],
        [null, 'For the month of', null, null, 'MONTH 2021', null, null, null, null],
        [null, 'Official hours for arrival (Regular day)', null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null],
        [null, null, null, null, null, null, null, null, null]
      ]

      const worksheet = XLSX.utils.json_to_sheet(form, { skipHeader: true })

      /* add worksheet to workbook */
      XLSX.utils.book_append_sheet(workbook, worksheet, 'SheetJS')

      worksheet['!cols'] = (() => {
        const keys = Object.keys(form[0])
        const colStyles = keys.reduce((acc, curr) => {
          return { ...acc, [curr]: { width: 0 } }
        }, {})

        form.forEach(currData => {
          keys.forEach(key => {
            colStyles[key].width = Math.max(_toString(currData[key]).length, colStyles[key].width)
          })
        })

        for (const key in colStyles) {
          colStyles[key].width += 2
        }

        return Object.values(colStyles)
      })()

      worksheet['!rows'] = (() => {
        return [
          { hpx: 15 },
          { hpx: 30 },
          { hpx: 10 }
        ]
      })()

      worksheet['!merges'] = [
        { // CIVIL SERVICE FORM NO. 48
          s: { r: 0, c: 1 },
          e: { r: 0, c: 7 }
        },
        { // DAILY TIME RECORD
          s: { r: 1, c: 1 },
          e: { r: 1, c: 7 }
        },
        { // BLANK SPACE
          s: { r: 2, c: 1 },
          e: { r: 2, c: 7 }
        },
        { // NAME INPUT
          s: { r: 3, c: 1 },
          e: { r: 3, c: 7 }
        },
        { // NAME LABEL
          s: { r: 4, c: 1 },
          e: { r: 4, c: 7 }
        },
        { // FOR THE MONTH OF
          s: { r: 5, c: 1 },
          e: { r: 5, c: 3 }
        },
        { // FOR THE MONTH OF
          s: { r: 5, c: 4 },
          e: { r: 5, c: 7 }
        }
      ]

      // styles
      ;(() => {
        worksheet.B1.s = {
          font: {
            name: 'arial',
            sz: 10,
            bold: true
          },
          alignment: {
            vertical: 'center'
          }
        }

        worksheet.B2.s = {
          font: {
            name: 'arial',
            sz: 12,
            bold: true
          },
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          }
        }

        ;(() => {
          const cols = ['B4', 'C4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4']

          cols.forEach(col => {
            worksheet[col].s = {
              font: {
                name: 'arial',
                sz: 10,
                bold: true
              },
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              }
            }
          })
        })()
      })()

      XLSX.writeFile(workbook, 'test.xlsx')
    }
  }
}
</script>

<style lang="scss" scoped>
@import './styles/home';
</style>
