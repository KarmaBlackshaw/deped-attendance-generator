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

    excelCols ({ form }) {
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

      // Override
      colStyles[1].width = 4

      colStyles[2].width = 7
      colStyles[3].width = 7
      colStyles[4].width = 7
      colStyles[5].width = 7
      colStyles[6].width = 7
      colStyles[7].width = 7

      return Object.values(colStyles)
    },

    excelRows ({ form }) {
      const properties = {
        0: { hpx: 15 },
        1: { hpx: 30 },
        2: { hpx: 5 },
        3: { hpx: 15 },
        4: { hpx: 13 },
        5: { hpx: 15 },
        6: { hpx: 15 },
        7: { hpx: 15 },
        8: { hpx: 5 }
      }

      return Object.values(properties)
    },

    excelMerges ({ form }) {
      return [
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
        { // FOR THE MONTH OF DATE
          s: { r: 5, c: 4 },
          e: { r: 5, c: 7 }
        },
        { // OFFICIAL HOURS FOR ARRIVAL
          s: { r: 6, c: 1 },
          e: { r: 6, c: 5 }
        },
        { // OFFICIAL HOURS FOR ARRIVAL
          s: { r: 6, c: 6 },
          e: { r: 6, c: 7 }
        },
        { // AND DEPARTURE
          s: { r: 7, c: 1 },
          e: { r: 7, c: 5 }
        },
        { // AND DEPARTURE
          s: { r: 7, c: 6 },
          e: { r: 7, c: 7 }
        },
        { // BLANK SPACE
          s: { r: 8, c: 1 },
          e: { r: 8, c: 7 }
        },
        { // DAY
          s: { r: 9, c: 1 },
          e: { r: 10, c: 1 }
        },
        { // AM
          s: { r: 9, c: 2 },
          e: { r: 9, c: 3 }
        },
        { // PM
          s: { r: 9, c: 4 },
          e: { r: 9, c: 5 }
        },
        { // UNDETIME
          s: { r: 9, c: 6 },
          e: { r: 9, c: 7 }
        },
        { // i certify
          s: { r: daysInTheMonth + 12, c: 2 },
          e: { r: daysInTheMonth + 12, c: 7 }
        }
      ]
    },

    async handleFileSelect (e) {
      // const files = e.target.files

      // const data = await readFile(files[0])
      // e.target.value = null

      const workbook = XLSX.utils.book_new()

      const daysInTheMonth = 20
      const days = Array.from({ length: daysInTheMonth }, (_, i) => {
        return ['', i + 1, '', '', '', '', '', '', '']
      })

      const form = [
        ['', 'CIVIL SERVICE FORM NO. 48', '', '', '', '', '', '', ''],
        ['', 'DAILY TIME RECORD', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', 'LAST NAME, GIVEN NAME M.I.', '', '', '', '', '', '', ''],
        ['', '(name)', '', '', '', '', '', '', ''],
        ['', 'For the month of', '', '', 'MONTH 2021', '', '', '', ''],
        ['', 'Official hours for arrival (Regular day)', '', '', '', '', '', '', ''],
        ['', 'and departure (Saturdays)', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', 'Day', 'AM', '', 'PM', '', 'Undertime', '', ''],
        ['', '', 'Arrival', 'Departure', 'Arrival', 'Departure', 'Hours', 'Minutes', ''],
        ...days,
        ['', '', 'TOTAL', '', '', '', '', '', ''],
        ['', '', 'I CERTIFY on my honor that the above is a true and correct', '', '', '', '', '', ''],
        ['', 'report of the hours of work performed, record of which was made', '', '', '', '', '', '', ''],
        ['', 'daily at the time of arrival at and departure from office', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '']
      ]

      const worksheet = XLSX.utils.json_to_sheet(form, {
        skipHeader: true
      })

      /* add worksheet to workbook */
      XLSX.utils.book_append_sheet(workbook, worksheet, 'SheetJS')

      /**
       * COLUMNS
       */
      worksheet['!cols'] = this.excelCols({ form })

      /**
       * ROWS
       */
      worksheet['!rows'] = this.excelRows({ form })

      /**
       * MERGES
       */
      worksheet['!merges'] = this.excelMerges({ form })

      /**
       * STYLES
       */
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
              },
              border: {
                bottom: {
                  style: 'thin',
                  rgb: '#000'
                }
              }
            }
          })
        })()

        worksheet.B5.s = {
          font: {
            name: 'arial',
            sz: 8
          },
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          }
        }

        worksheet.B10.s = {
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          }
        }

        /**
         * AM|PM|Undertime
         */
        ;(() => {
          const cols = ['C10', 'E10', 'G10']

          cols.forEach(col => {
            worksheet[col].s = {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              }
            }
          })
        })()

        /**
         * Arival|Departure|Hours|minutes
         */
        ;(() => {
          const cols = ['C11', 'D11', 'E11', 'F11', 'G11', 'H11']

          cols.forEach(col => {
            worksheet[col].s = {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 9,
                italic: true
              }
            }
          })
        })()

        /**
         * Total
         */
        worksheet.C22.s = {
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          },
          font: {
            bold: true,
            sz: 9
          }
        }

        /**
         * I certify on my honor
         */
        // ;(() => {
        //   const cols = [
        //     'B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23',
        //     'B24', 'C24', 'D24', 'E24', 'F24', 'G24', 'H24',
        //     'B25', 'C25', 'D25', 'E25', 'F25', 'G25', 'H25'
        //   ]

        //   cols.forEach(col => {
        //     worksheet[col].s = {
        //       font: {
        //         sz: 10
        //       }
        //     }
        //   })
        // })()
      })()

      XLSX.writeFile(workbook, 'test.xlsx')
    }
  }
}
</script>

<style lang="scss" scoped>
@import './styles/home';
</style>
