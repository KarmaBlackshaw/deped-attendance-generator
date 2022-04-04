
import moment from 'moment'
import _groupBy from 'lodash/groupBy'
import _get from 'lodash/get'
import _mapKeys from 'lodash/mapKeys'
import _trim from 'lodash/trim'
import _isNil from 'lodash/isNil'

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
      const startOfMonth = moment(date).startOf('month')
      const endOfMonth = moment(date).endOf('month')
      const daysGap = endOfMonth.diff(startOfMonth, 'd')
      const timeline = {}

      for (let i = 0; i < daysGap; i++) {
        const currDate = moment(startOfMonth)
          .add(i, 'd')
          .format('YYYY-MM-DD')

        timeline[currDate] = {}
      }

      return timeline
    }

    /**
        * Parse, trim, and sort data
        */
    const snakeCaseData = data
      .map(item => {
        const curr = _mapKeys(item, (_, key) => dictionary[key])

        const datetime = moment(_trim(curr.datetime), 'M/D/YYYY H:mm:ss A')

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

  return {
    getData
  }
}
