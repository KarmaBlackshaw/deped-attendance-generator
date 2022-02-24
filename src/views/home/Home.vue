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
import _merge from 'lodash/merge'
import _flatten from 'lodash/flatten'

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

  setup () {
    const { waitUntilElementIsLoaded } = useHelpers()

    return {
      waitUntilElementIsLoaded
    }
  },

  created () {
    this.handleFileSelect()
  },

  methods: {
    async handleClickBrowseFiles () {
      const inputFile = await this.waitUntilElementIsLoaded('#input-file')

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

    excelMerges ({ form, daysInTheMonth }) {
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
        },
        { // report of the hours
          s: { r: daysInTheMonth + 13, c: 1 },
          e: { r: daysInTheMonth + 13, c: 7 }
        },
        { // daily at the time
          s: { r: daysInTheMonth + 14, c: 1 },
          e: { r: daysInTheMonth + 14, c: 7 }
        },
        { // Verified as to the prescribed office hours. underline
          s: { r: daysInTheMonth + 15, c: 1 },
          e: { r: daysInTheMonth + 15, c: 7 }
        },
        { // Verified as to the prescribed office hours
          s: { r: daysInTheMonth + 16, c: 1 },
          e: { r: daysInTheMonth + 16, c: 7 }
        },
        { // SCHOOL PRINCIPAL NAME
          s: { r: daysInTheMonth + 18, c: 3 },
          e: { r: daysInTheMonth + 18, c: 7 }
        },
        { // School Principal
          s: { r: daysInTheMonth + 19, c: 3 },
          e: { r: daysInTheMonth + 19, c: 7 }
        },
        { // In charge
          s: { r: daysInTheMonth + 20, c: 3 },
          e: { r: daysInTheMonth + 20, c: 7 }
        },
        { // (See instructions on back)
          s: { r: daysInTheMonth + 21, c: 1 },
          e: { r: daysInTheMonth + 21, c: 3 }
        },
        { // instructions
          s: { r: daysInTheMonth + 22, c: 1 },
          e: { r: daysInTheMonth + 22, c: 7 }
        },
        { // civil service form no 48
          s: { r: daysInTheMonth + 24, c: 1 },
          e: { r: daysInTheMonth + 27, c: 7 }
        },
        { // in the lieu
          s: { r: daysInTheMonth + 29, c: 1 },
          e: { r: daysInTheMonth + 37, c: 7 }
        },
        { // space
          s: { r: daysInTheMonth + 39, c: 1 },
          e: { r: daysInTheMonth + 42, c: 7 }
        },
        { // space
          s: { r: daysInTheMonth + 44, c: 1 },
          e: { r: daysInTheMonth + 46, c: 7 }
        },
        { // space
          s: { r: daysInTheMonth + 48, c: 1 },
          e: { r: daysInTheMonth + 58, c: 7 }
        },
        { // space
          s: { r: daysInTheMonth + 60, c: 1 },
          e: { r: daysInTheMonth + 64, c: 7 }
        }
      ]
    },

    async handleFileSelect (e) {
      // const files = e.target.files

      // const data = await readFile(files[0])
      // e.target.value = null

      const workbook = XLSX.utils.book_new()

      const daysInTheMonth = 10
      const days = Array.from({ length: daysInTheMonth }, (_, i) => {
        return ['', i + 1, '', '', '', '', '', '', '']
      })

      const emptySpaces = Array.from({ length: 100 }, () => {
        return ['', '', '', '', '', '', '', '', '']
      })

      const gap = num => daysInTheMonth + num

      const form = [
        ['', 'CIVIL SERVICE FORM NO. 48', '', '', '', '', '', '', ''],
        ['', 'DAILY TIME RECORD', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', 'LAST NAME, GIVEN NAME M.I.', '', '', '', '', '', '', ''],
        ['', '(Name)', '', '', '', '', '', '', ''],
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
        ['', 'Verified as to the prescribed office hours', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', 'SCHOOL PRINCIPAL NAME', '', '', '', '', ''],
        ['', '', '', 'School Principal', '', '', '', '', ''],
        ['', '', '', 'In-Charge', '', '', '', '', ''],
        ['', '(See instructions on back)', '', '', '', '', '', '', ''],
        ['', 'INSTRUCTIONS', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '              Civil Service Form No. 48, after completion should be filed in the records of the bureau or office which submits the monthly report on Civil Service Form No. 3 to the Bureau of Civil', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '              In the Lieu of the above, court interpreters and stenographers who accompany the judges of the court in First Instance will fill out the daily time reports on this form in triplicate, after which they should be approved by the judge with whom service has been rendered, or by an officer of the Department of Justice authorized to do so.The original should be forwarded promptly after the end of the month to the Bureau of Civil Service, thru the department of Justice; the duplicate to be kept in the Department of Justice; and the triplicate, in the office of the Clerk of Court where service were rendered.', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '              In the space provided for the purpose on the other side will be indicated the office hours the employee is required to observed, as for example. "Regular days, 3:00-12:00 and 1-4: Saturdays, 5:00-1:00."', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '              Attention is invited to paragraph 3, Civil Service Rule XV, Executive Order No. 5 series of 1909, which reads as follows:', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '              Each chief of bureau of office shall require a daily record of attendance of all officers and employees under him extitled to leave of a absence or vacation (including teachers) to be kept on the proper form and also a systematic office record showing for each day all absences from duty from any cause whatever. At the beginning of each month he shall report to the Commissioner on the proper form of all absences from any cause whatever, including the exact amount of undertime of each person for each day. Officers or employees must be included in the monthly report of changes and absences. Falsificationof time records will render the offending officer or employee liable to summary removal from the service anbd criminal prosecution."', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', '              (NOTE: - A record made from memory at sometime subsequent to the occurrence of an event is not reliable. Non-observance of office hours deprives the employee of the leave of privileges although he may have rendered overtime services. Where service rendered outside of the Office for the whole morning or afternoon, notation to that effect should be made.', '', '', '', '', '', '', ''],
        ...emptySpaces
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
      worksheet['!merges'] = this.excelMerges({ form, daysInTheMonth })

      /**
       * STYLES
       */
      const merge = (col, styles) => _merge(worksheet[col]?.s ?? {}, styles, { font: { name: 'Arial' } })

      ;(() => {
        /**
         * CIVIL SERVICE FORM No. 48
         */
        worksheet.B1.s = merge('B1', {
          font: {
            name: 'arial',
            sz: 9,
            bold: true
          },
          alignment: {
            vertical: 'center'
          }
        })

        /**
         * DAILY TIME RECORD
         */
        worksheet.B2.s = merge('B2', {
          font: {
            name: 'arial',
            sz: 12,
            bold: true
          },
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          }
        })

        /**
         * Fullname
         */
        ;(() => {
          const cols = ['B4', 'C4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4']

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
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
                  style: 'thin'
                }
              }
            })
          })
        })()

        /**
         * (name)
         */
        worksheet.B5.s = merge('B5', {
          font: {
            sz: 6
          },
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          }
        })

        /**
         * For the month of [static values]
         */
        ;(() => {
          const cols = [
            'B6', 'C6', 'D6', // For the month of
            'B7', 'C7', 'D7', 'E7', 'F7', // Official hours
            'B8', 'C8', 'D8', 'E8', 'F8' // Official hours
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'left'
              },
              font: {
                sz: 9
              }
            })
          })
        })()

        /**
         * For the month of [dynamic values]
         */
        ;(() => {
          const cols = [
            'E6', 'F6', 'G6', 'H6', // For the month of
            'G7', 'H7', // Official hours
            'G8', 'H8' // Official hours
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 8,
                bold: true
              },
              border: {
                bottom: {
                  style: 'thin'
                }
              }
            })
          })
        })()

        // Date table
        ;(() => {
          const dayCols = Array
            .from({ length: daysInTheMonth + 1 }, (_, i) => {
              return [`B${i + 11}`, `C${i + 11}`, `D${i + 11}`, `E${i + 11}`, `F${i + 11}`, `G${i + 11}`, `H${i + 11}`]
            })

          const cols = [
            'B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10',
            'B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11',
            ..._flatten(dayCols)
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              border: {
                bottom: {
                  style: 'thin'
                },
                top: {
                  style: 'thin'
                },
                left: {
                  style: 'thin'
                },
                right: {
                  style: 'thin'
                }
              }
            })
          })
        })()

        worksheet.B10.s = merge('B10', {
          alignment: {
            vertical: 'center',
            horizontal: 'center'
          },
          font: {
            sz: 7
          }
        })

        /**
         * AM|PM|Undertime
         */
        ;(() => {
          const cols = ['C10', 'E10', 'G10']

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: '8'
              }
            })
          })
        })()

        /**
         * Arival|Departure|Hours|minutes
         */
        ;(() => {
          const cols = ['C11', 'D11', 'E11', 'F11', 'G11', 'H11']

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 7
              }
            })
          })
        })()

        ;(() => {
          const cols = Array.from({ length: daysInTheMonth }, (_, i) => {
            return `B${12 + i}`
          })

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 10,
                italic: true
              }
            })
          })
        })()

        /**
         * Total
         */
        ;(() => {
          const cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H'].map(letter => {
            return `${letter}${gap(12)}`
          })

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                bold: true,
                sz: 9
              },
              border: {
                bottom: {
                  style: 'thin'
                }
              }
            })
          })
        })()

        /**
         * I certify on my honor
         */
        ;(() => {
          const cols = [
            `B${gap(13)}`, `C${gap(13)}`, `D${gap(13)}`, `E${gap(13)}`, `F${gap(13)}`, `G${gap(13)}`, `H${gap(13)}`,
            `B${gap(14)}`, `C${gap(14)}`, `D${gap(14)}`, `E${gap(14)}`, `F${gap(14)}`, `G${gap(14)}`, `H${gap(14)}`,
            `B${gap(15)}`, `C${gap(15)}`, `D${gap(15)}`, `E${gap(15)}`, `F${gap(15)}`, `G${gap(15)}`, `H${gap(15)}`
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              font: {
                sz: 8
              }
            })
          })
        })()

        /**
         * Verified as to the prescribed office hours
         */
        ;(() => {
          const cols = [
            `B${gap(17)}`, `C${gap(17)}`, `D${gap(17)}`, `E${gap(17)}`, `F${gap(17)}`, `G${gap(17)}`, `H${gap(17)}`
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 8
              },
              border: {
                top: {
                  style: 'thin'
                }
              }
            })
          })
        })()

        /**
         * School Principal Name
         */
        ;(() => {
          const cols = [
            `D${gap(19)}`, `E${gap(19)}`, `F${gap(19)}`, `G${gap(19)}`, `H${gap(19)}`
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                bold: true,
                sz: 9
              }
            })
          })
        })()

        /**
         * School Principal
         */
        ;(() => {
          const cols = [
            `D${gap(20)}`, `E${gap(20)}`, `F${gap(20)}`, `G${gap(20)}`, `H${gap(20)}`
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 8
              }
            })
          })
        })()

        /**
         * In-Charge
         */
        ;(() => {
          const cols = [
            `D${gap(21)}`, `E${gap(21)}`, `F${gap(21)}`, `G${gap(21)}`, `H${gap(21)}`
          ]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                sz: 8
              },
              border: {
                top: {
                  style: 'thin'
                }
              }
            })
          })
        })()

        /**
         * in-charge
         */
        ;(() => {
          const cols = [`B${gap(22)}`, `C${gap(22)}`, `D${gap(22)}`]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center'
              },
              font: {
                italic: true,
                sz: 8
              }
            })
          })
        })()

        /**
         * (See instructions on back)

         */
        ;(() => {
          const cols = [`B${gap(22)}`, `C${gap(22)}`, `D${gap(22)}`, `E${gap(22)}`, `F${gap(22)}`, `G${gap(22)}`, `H${gap(22)}`]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              border: {
                bottom: {
                  style: 'medium'
                }
              }
            })
          })
        })()

        /**
         * INSTRUCTIONS
         */
        ;(() => {
          const cols = [`B${gap(23)}`, `C${gap(23)}`, `D${gap(23)}`, `E${gap(23)}`, `F${gap(23)}`, `G${gap(23)}`, `H${gap(23)}`]

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'center'
              },
              font: {
                bold: true,
                sz: 12
              }
            })
          })
        })()

        /**
         * civil service form no. 48
         */
        ;(() => {
          const range = [25, 70]

          let i = range[0]
          const cols = []
          while (i <= range[1]) {
            const currCol = [
              `B${gap(i)}`, `C${gap(i)}`, `D${gap(i)}`, `E${gap(i)}`, `F${gap(i)}`, `G${gap(i)}`, `H${gap(i)}`
            ]
            cols.push(...currCol)
            i++
          }

          cols.forEach(col => {
            worksheet[col].s = merge(col, {
              alignment: {
                vertical: 'center',
                horizontal: 'left',
                wrapText: true
              },
              font: {
                sz: 10
              }
            })
          })
        })()

        // /**
        //  * civil service form no. 48
        //  */
        // ;(() => {
        //   const range = [30, 49]

        //   let i = range[0]
        //   const cols = []
        //   while (i <= range[1]) {
        //     const currCol = [
        //       `B${gap(i)}`, `C${gap(i)}`, `D${gap(i)}`, `E${gap(i)}`, `F${gap(i)}`, `G${gap(i)}`, `H${gap(i)}`
        //     ]
        //     cols.push(...currCol)
        //     i++
        //   }

        //   cols.forEach(col => {
        //     worksheet[col].s = merge(col, {
        //       alignment: {
        //         vertical: 'center',
        //         horizontal: 'left',
        //         wrapText: true
        //       },
        //       font: {
        //         sz: 10
        //       }
        //     })
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
