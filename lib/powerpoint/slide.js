var _ = require('lodash')
var Shape = require('./shape')

function Slide (nativeSlide) {
  var slide = this

  slide.attr = function (name, value) {
    if (typeof value !== 'undefined') {
      return new Slide(nativeSlide.attr({name: name, value: value}, true))
    }
    return nativeSlide.attr(name, true)
  }

  slide.dispose = function () {
    return nativeSlide.dispose(null, true)
  }

  // inject attr
  _.assign(slide, require('./slide.attr'))

  slide.remove = function () {
    return nativeSlide.remove(null, true)
  }

  slide.duplicate = function () {
    return new Slide(nativeSlide.duplicate(null, true))
  }

  slide.copy = function () {
    return new Slide(nativeSlide.copy(null, true))
  }

  slide.select = function () {
    return new Slide(nativeSlide.select(null, true))
  }

  slide.textReplace = function (find, replace) {
    if (typeof find === 'string' && typeof replace === 'string') {
      nativeSlide.textReplace({'find': find, 'replace': replace}, true)
    } else if (typeof find === 'object') {
      replace = (typeof replace === 'function') ? replace : true
      nativeSlide.textReplace({'batch': find}, replace)
    }
    return slide
  }

  slide.addTextbox = function (options) {
    options = options || {}
    options.left = (typeof options.left !== 'undefined') ? options.left : 0
    options.top = (typeof options.top !== 'undefined') ? options.top : 0
    options.height = (typeof options.height !== 'undefined') ? options.height : 100
    options.width = (typeof options.width !== 'undefined') ? options.width : 100

    return new Shape(nativeSlide.addTextbox(options, true))
  }

  slide.addPicture = function (path, options) {
    if (typeof path !== 'string') {
      throw new Error('Missing path!')
    }
    options = options || {}
    options.left = (typeof options.left !== 'undefined') ? options.left : 0
    options.top = (typeof options.top !== 'undefined') ? options.top : 0
    options.path = path
    return new Shape(nativeSlide.addPicture(options, true))
  }

  slide.shapes = function (selector) {
    var shapes = []
    var i
    var nativeShapes = nativeSlide.shapes(selector, true)
    for (i = 0; i < nativeShapes.length; i++) {
      shapes.push(new Shape(nativeShapes[i]))
    }
    return shapes
  }

  slide.tag = {
    get: function (name) {
      return nativeSlide.tags(null, true).get(name, true)
    },
    set: function (name, value) {
      nativeSlide.tags(null, true).set({name: name, value: value}, true)
      return slide
    },
    remove: function (name) {
      nativeSlide.tags(null, true).remove(name, true)
      return slide
    }
  }

  slide.tags = nativeSlide.tags(null, true).all(null, true)
}

module.exports = Slide
