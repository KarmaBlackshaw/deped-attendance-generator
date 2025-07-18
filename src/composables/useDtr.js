
import moment from 'moment'
import _groupBy from 'lodash/groupBy'
import _get from 'lodash/get'
import _mapKeys from 'lodash/mapKeys'
import _trim from 'lodash/trim'
import _isNil from 'lodash/isNil'
import _toString from 'lodash/toString'
import _merge from 'lodash/merge'
import _flatten from 'lodash/flatten'

function pipe (initial, fns) {
  return fns.reduce((v, f) => f(v), initial)
}

export default () => {
  function getData (data) {
    const dictionary = {
      Department: 'department',
      Name: 'name',
      'No.': 'no',
      'Date/Time': 'datetime',
      Status: 'status',
      'Location ID': 'location_id',
      'ID Number': 'id',
      VerifyCode: 'verify_code',
      CardNo: 'card_no'
    }

    // helpers
    const generateDateTimelines = date => {
      const dateInstance = moment(date)

      const daysInMonth = dateInstance.daysInMonth()
      const year = dateInstance.format('YYYY')
      const month = dateInstance.format('MM')

      const toDateUsingDay = day => `${year}-${month}-${String(day).padStart(2, '0')}`

      return pipe(daysInMonth, [
        val => Array.from({ length: val }),
        val => val.map((_, idx) => toDateUsingDay(idx + 1)),
        val => val.reduce((acc, curr) => ({ ...acc, [curr]: {} }), {})
      ])
    }

    /**
        * Parse, trim, and sort data
        */
    const snakeCaseData = data
      .map(item => {
        const curr = _mapKeys(item, (_, key) => dictionary[key])

        const datetime = moment(_trim(curr.datetime))

        if (!datetime.isValid()) {
          throw new Error(`Invalid date format: ${curr.datetime}`)
        }

        curr.timestamp = {
          datetime: datetime.format('YYYY-MM-DD HH:mm:ss'),
          date: datetime.format('YYYY-MM-DD'),
          time: datetime.format('HH:mm:ss')
        }

        delete curr.datetime
        delete curr.no
        delete curr.id
        delete curr.card_no
        delete curr.location_id
        delete curr.verify_code

        return curr
      })
      .sort((x, y) => x.timestamp.datetime > y.timestamp.datetime ? 1 : -1)

    /**
      * Group data by date
      */
    const dateEntries = Object.entries(_groupBy(snakeCaseData, 'timestamp.date'))

    const attendanceByUserId = (() => {
      const timelines = {}

      dateEntries.forEach(([date, payload]) => {
        const timelineByUsers = _groupBy(payload, 'name')

        for (const userId in timelineByUsers) {
          const userTimeline = timelineByUsers[userId]

          const time = (() => {
            let morningTimeIn
            let morningTimeOut
            let afternoonTimeIn
            let afternoonTimeOut

            userTimeline.forEach(curr => {
              const status = curr.status

              if (status === 'C/In' && _isNil(morningTimeIn)) {
                morningTimeIn = curr
              }

              if (status === 'Out' && _isNil(morningTimeOut)) {
                morningTimeOut = curr
              }

              if (status === 'Out Back' && _isNil(afternoonTimeIn)) {
                afternoonTimeIn = curr
              }

              if (status === 'C/Out' && _isNil(afternoonTimeOut)) {
                afternoonTimeOut = curr
              }
            })

            return { morningTimeIn, morningTimeOut, afternoonTimeIn, afternoonTimeOut }
          })()

          if (!timelines[userId]) {
            timelines[userId] = {}

            const timeline = generateDateTimelines(date)
            for (const timelineDate in timeline) {
              timelines[userId][timelineDate] = { date: timelineDate }
            }
          }

          if (!timelines[userId][date]) {
            const timeline = generateDateTimelines(date)

            for (const timelineDate in timeline) {
              timelines[userId][timelineDate] = _get(timelines, [userId, timelineDate], { date: timelineDate })
            }
          }

          const timeData = time.morningTimeOut || time.morningTimeIn || time.afternoonTimeOut || time.afternoonTimeIn

          timelines[userId][date] = {
            date: date,
            department: timeData.department,
            morning: {
              time_in: _get(time.morningTimeIn, 'timestamp.time'),
              time_out: _get(time.morningTimeOut, 'timestamp.time')
            },
            afternoon: {
              time_in: _get(time.afternoonTimeIn, 'timestamp.time'),
              time_out: _get(time.afternoonTimeOut, 'timestamp.time')
            }
          }
        }
      })

      return timelines
    })()

    const attendanceList = []
    for (const userId in attendanceByUserId) {
      const currAttendance = attendanceByUserId[userId]

      const dates = Object.keys(currAttendance)

      const sortedDates = dates.sort((x, y) => (new Date(x) > new Date(y) ? 1 : -1))

      const attendance = {}
      sortedDates.forEach(curr => {
        const monthOf = moment(curr).format('MMMM YYYY')

        if (!attendance[monthOf]) {
          attendance[monthOf] = []
        }

        const item = currAttendance[curr]
        const dayNo = moment(curr).format('D')
        const amArrival = _get(item, 'morning.time_in', '')
        const amDeparture = _get(item, 'morning.time_out', '')
        const pmArrival = _get(item, 'afternoon.time_in', '')
        const pmDeparture = _get(item, 'afternoon.time_out', '')

        attendance[monthOf].push([
          '',
          dayNo,
          amArrival ? moment(amArrival, 'HH:mm:ss').format('hh:mm') : '',
          amDeparture ? moment(amDeparture, 'HH:mm:ss').format('hh:mm') : '',
          pmArrival ? moment(pmArrival, 'HH:mm:ss').format('hh:mm') : '',
          pmDeparture ? moment(pmDeparture, 'HH:mm:ss').format('hh:mm') : '',
          '',
          '',
          ''
        ])
      })

      attendanceList.push({
        user_id: userId,
        attendance
      })
    }

    return attendanceList
  }

  function excelCols ({ form }) {
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
  }

  function excelRows ({ form }) {
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
  }

  function excelMerges ({ form, daysInTheMonth }) {
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
      }
    ]
  }

  function excelStyles ({ worksheet, daysInTheMonth }) {
    const gap = num => daysInTheMonth + num
    const merge = (col, styles) => _merge(worksheet[col]?.s ?? {}, styles, { font: { name: 'Arial' } })

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
          font: {
            sz: 9
          },
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

    // Day
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
            sz: 8,
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
  }

  return {
    getData,
    excelCols,
    excelRows,
    excelMerges,
    excelStyles
  }
}
