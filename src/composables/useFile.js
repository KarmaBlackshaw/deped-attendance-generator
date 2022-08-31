import * as xlsx from 'xlsx'

export default () => {
  const readXlxs = files => new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = e => {
      const data = e.target.result
      const workbook = xlsx.read(data, { type: 'binary' })
      const wsname = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[wsname]
      const json = xlsx.utils.sheet_to_json(worksheet, { header: 2 })

      resolve(json)
    }

    reader.onerror = e => reject(e)

    reader.readAsBinaryString(files)
  })

  function saveBlobAs (blob, file_name) {
    if (typeof navigator.msSaveBlob == "function")
      return navigator.msSaveBlob(blob, file_name)

    const saver = document.createElementNS("http://www.w3.org/1999/xhtml", "a")
    const blobURL = saver.href = URL.createObjectURL(blob),
      body = document.body

    saver.download = file_name

    body.appendChild(saver)
    saver.dispatchEvent(new MouseEvent("click"))
    body.removeChild(saver)
    URL.revokeObjectURL(blobURL)
  }

  return {
    readXlxs,
    saveBlobAs
  }
}