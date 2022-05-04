module.exports = {
  css: {
    loaderOptions: {
      sass: {
        additionalData: `
          @import "@/styles/import.scss";
        `
      }
    }
  }
}
