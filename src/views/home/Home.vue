<script setup>
// libs
import { ref } from 'vue'
import XLSX from 'xlsx-js-style'
import JSZip from 'jszip'
import moment from 'moment'
import tryToCatch from 'try-to-catch'

// helpers
import { getElement } from '@/utils/helpers'

// composables
import useDtr from '@/composables/useDtr'
import useFile from '@/composables/useFile'

const {
  getData: getDtrData,
  excelCols,
  excelRows,
  excelMerges,
  excelStyles
} = useDtr()

const {
  saveBlobAs,
  readXlxs
} = useFile()

async function handleClickBrowseFiles () {
  const inputFile = await getElement('#input-file')

  inputFile.click()
}

const error = ref('')
async function handleFileSelect (e) {
  error.value = ''

  /**
    Helpers
    */
  const createEmptySpaces = length => Array.from({ length }, () => {
    return ['', '', '', '', '', '', '', '', '']
  })

  const files = e.target.files

  const data = await readXlxs(files[0])
  e.target.value = null

  const [err, dtrData] = await tryToCatch(() => getDtrData(data))

  if (err?.message) {
    error.value = err.message
    console.error('Error:', error.value)
    return
  }

  const zip = new JSZip()
  for (let i = 0; i < dtrData.length; i++) {
    const userData = dtrData[i]

    const workbook = XLSX.utils.book_new()

    for (const monthOf in userData.attendance) {
      const days = userData.attendance[monthOf]
      const daysInTheMonth = days.length

      const form = [
        ['', 'CIVIL SERVICE FORM NO. 48', '', '', '', '', '', '', ''],
        ['', 'DAILY TIME RECORD', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', ''],
        ['', userData.user_id, '', '', '', '', '', '', ''],
        ['', '(Name)', '', '', '', '', '', '', ''],
        ['', 'For the month of', '', '', monthOf, '', '', '', ''],
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
        ...createEmptySpaces(100)
      ]

      const worksheet = XLSX.utils.json_to_sheet(form, {
        skipHeader: true
      })

      const filename = userData.user_id

      /* add worksheet to workbook */
      XLSX.utils.book_append_sheet(workbook, worksheet, filename)

      /**
      * COLUMNS
      */
      worksheet['!cols'] = excelCols({ form })

      /**
      * ROWS
      */
      worksheet['!rows'] = excelRows({ form })

      /**
        * MERGES
        */
      worksheet['!merges'] = excelMerges({ form, daysInTheMonth })

      /**
        * STYLES
        */
      excelStyles({ worksheet, daysInTheMonth })

      const buffer = XLSX.write(workbook, {
        type: 'buffer'
      })

      zip.file(`${filename}.xlsx`, buffer)
    }
  }

  const blob = await zip.generateAsync({ type: "blob" })
  const fileName = `dtr-${moment().format('YYYY-MM-DD')}`

  saveBlobAs(blob, fileName)
}

</script>

<template>
  <div class="home">
    <div class="card">
      <div
        v-if="error"
        class="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded relative"
        role="alert"
      >
        <strong class="font-bold block">Holy smokes!</strong>

        <span class="block sm:inline">{{ error }}</span>
      </div>

      <div class="card__title">
        Excel to DTR Generator
      </div>

      <div class="card__subtitle">
        <p>This tool generates a <strong>DTR</strong> based on the excel file uploaded</p>
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

      <div class="card__subtitle documentation">
        <img
          src="./assets/images/excel-form2.png"
          alt=""
          class="documentation__img--excel-form"
        >

        <img
          src="./assets/svg/chevron-right.svg"
          alt=""
          class="documentation__img--chevron-right"
        >

        <img
          src="./assets/images/dtr-form2.png"
          alt=""
          class="documentation__img--dtr-form"
        >
      </div>

      <div
        class="
          flex gap-[5px] justify-center text-green-50
        "
      >
        <span>Made with</span>
        <img
          src="./assets/images/vue-icon.png"
          alt=""
          class="w-[20px] h-[20px]"
        >
        <span>by</span>
        <a
          href="https://jeash.tech"
          target="_blank"
          class="font-bold"
        >
          Ernie Jeash
        </a>
      </div>
    </div>
  </div>
</template>

<style lang="scss" scoped>
@import './styles/Home';
</style>

