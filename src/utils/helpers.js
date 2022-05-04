export const getElement = async selector => {
  while (!document.querySelector(selector)) {
    // eslint-disable-next-line no-await-in-loop
    await new Promise(resolve => requestAnimationFrame(resolve))
  }

  return document.querySelector(selector)
}
