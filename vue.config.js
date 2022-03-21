module.exports = {
  publicPath: process.env.NODE_ENV === 'production' ? '/deped-attendance-generator/' : '/',
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
