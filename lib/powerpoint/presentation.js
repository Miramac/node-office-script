var fs = require('fs')
var application = require('./application')
var Slide = require('./slide')
var Shape = require('./shape')

function Presentation (presentationPath) {
  presentationPath = presentationPath || null
  var nativePresentation
  var presentation = this
  if (fs.existsSync(presentationPath)) {
    nativePresentation = application.open(presentationPath, true)
  } else {
    nativePresentation = application.fetch(presentationPath, true)
  }

  presentation.slides = function (selector) {
    var slides = []
    var i
    var nativeSlides = nativePresentation.slides(selector, true)
    for (i = 0; i < nativeSlides.length; i++) {
      slides.push(new Slide(nativeSlides[i]))
    }
    return slides
  }

  presentation.addSlide = function (options) {
    return new Slide(nativePresentation.addSlide(options, true))
  }

  presentation.textReplace = function (find, replace) {
    if (typeof find === 'string' && typeof replace === 'string') {
      nativePresentation.textReplace({'find': find, 'replace': replace}, true)
    } else if (typeof find === 'object') {
      replace = (typeof replace === 'function') ? replace : true
      nativePresentation.textReplace({'batch': find}, replace)
    }
    return presentation
  }

  presentation.nativeSlides = function (selector) {
    return nativePresentation.slides(selector, true)
  }

  presentation.shapes = function (selector, slides) {
    var shapes = []
    var i
    slides = (typeof slides === 'string') ? presentation.slides(slides) : slides
    slides = (typeof slides === 'undefined') ? presentation.slides() : slides
    slides = (slides.length) ? slides : [slides]

    for (i = 0; i < slides.length; i++) {
      shapes = shapes.concat(slides[i].shapes(selector))
    }
    return shapes
  }

  presentation.attr = function (input, cb) {
    cb = cb || true
    return nativePresentation.attr(input, cb)
  }

  presentation.dispose = function () {
    return nativePresentation.dispose(null, true)
  }

  presentation.quit = function (cb) {
    cb = cb || true
    nativePresentation.close(null, true)
    application.quit(null, cb)
  }
  presentation.close = function (cb) {
    cb = cb || true
    nativePresentation.close(null, cb)
  }

  presentation.save = function (cb) {
    cb = cb || true
    nativePresentation.save(null, cb)
  }

  presentation.saveAs = function (input, cb) {
    cb = cb || true
    nativePresentation.saveAs(input, cb)
  }

  presentation.saveAsCopy = function (input, cb) {
    cb = cb || true
    nativePresentation.saveAs(input, cb)
  }

  // * Attr shortcuts
  presentation.path = function () {
    return presentation.attr('Path', true)
  }

  presentation.name = function () {
    return presentation.attr('Name', true)
  }

  presentation.fullName = function () {
    return presentation.attr('FullName', true)
  }

  presentation.builtinProp = function (prop, value) {
    if (typeof prop !== 'string') {
      return nativePresentation.properties(null, true).getAllBuiltinProperties(null, true)
    } else if (typeof value !== 'string') {
      return nativePresentation.properties(null, true).getBuiltinProperty(prop, true)
    } else {
      return nativePresentation.properties(null, true).setBuiltinProperty({'prop': prop, 'value': value}, true)
    }
  }

  presentation.customProp = function (prop, value) {
    if (typeof prop !== 'string') {
      return nativePresentation.properties(null, true).getAllCustomProperties(null, true)
    } else if (typeof value !== 'string') {
      return nativePresentation.properties(null, true).getCustomProperty(prop, true)
    } else {
      return nativePresentation.properties(null, true).setCustomProperty({'prop': prop, 'value': value}, true)
    }
  }

  presentation.tag = {
    get: function (name) {
      return nativePresentation.tags(null, true).get(name, true)
    },
    set: function (name, value) {
      nativePresentation.tags(null, true).set({name: name, value: value}, true)
      return presentation
    },
    remove: function (name) {
      nativePresentation.tags(null, true).remove(name, true)
      return presentation
    }
  }

  presentation.tags = nativePresentation.tags(null, true).all(null, true)

  presentation.getType = function () {
    return nativePresentation.getType(null, true)
  }

  presentation.getSelectedShape = function () {
    return new Shape(nativePresentation.getSelectedShape(null, true))
  }
  presentation.selectedShape = presentation.getSelectedShape

  presentation.getActiveSlide = function () {
    return new Slide(nativePresentation.getActiveSlide(null, true))
  }
  presentation.activeSlide = presentation.getActiveSlide

  presentation.slideHeight = function () {
    return presentation.attr('SlideHeight', true)
  }

  presentation.slideWidth = function () {
    return presentation.attr('SlideWidth', true)
  }

  presentation.pasteSlide = function (index) {
    index = index || 1
    if (typeof index !== 'number') throw new Error('presentation.pasteSlide(index) : Index must be a number!  ')
    index = (index === -1) ? presentation.slides().length + 1 : index
    index = (index < 1) ? 1 : index
    return new Slide(nativePresentation.pasteSlide(index, true))
  }
}

module.exports = Presentation
