module.exports = {
  publicPath: '/deped-attendance-generator/',
  assetsDir: '/deped-attendance-generator/',
  css: {
    loaderOptions: {
      sass: {
        prependData: `
          @import "@/styles/import.scss";
        `
      }
    }
  }
}
