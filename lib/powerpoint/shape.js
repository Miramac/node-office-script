var _ = require('lodash')
var Paragraph = require('./paragraph')
var Character = require('./character')

function Shape (nativeShape, isTableCell) {
  isTableCell = isTableCell || false
  var shape = this

  shape.attr = function (name, value) {
    if (typeof value !== 'undefined') {
      return new Shape(nativeShape.attr({name: name, value: value}, true), isTableCell)
    }
    return nativeShape.attr(name, true)
  }

  shape.dispose = function () {
    return nativeShape.dispose(null, true)
  }

  // inject attr
  _.assign(shape, require('./shape.attr'))

  shape.has = function (objectName) {
    return nativeShape.hasObject(objectName, true)
  }

  shape.paragraph = function (start, length) {
    start = start || -1
    length = length || -1
    if (shape.text() === null) {
      return null
    }
    return new Paragraph(nativeShape.paragraph({'start': start, 'length': length}, true))
  }

  shape.p = shape.paragraph

  shape.character = function (start, length) {
    start = start || -1
    length = length || -1
    if (shape.text() === null) {
      return null
    }
    return new Character(nativeShape.character({'start': start, 'length': length}, true))
  }

  shape.c = shape.character
  shape.char = shape.character

  shape.textReplace = function (findString, replaceString) {
    return new Shape(nativeShape.textReplace({'find': findString, 'replace': replaceString}, true), isTableCell)
  }

  shape.getType = function () {
    return nativeShape.getType(null, true)
  }

  if (!isTableCell) {
    shape.remove = function () {
      return nativeShape.remove(null, true)
    }

    shape.duplicate = function () {
      return new Shape(nativeShape.duplicate(null, true))
    }

    shape.addLine = function (text, pos) {
      if (typeof pos !== 'number') {
        shape.paragraph(shape.paragraph().count() + 1, -1).text(text)
        return shape
      } else {
        shape.paragraph(pos, -1).text(text)
        return shape
      }
    }

    shape.removeLine = function (pos) {
      if (typeof pos !== 'number') {
        shape.paragraph(shape.paragraph().count(), -1).remove()
        return shape
      } else {
        shape.paragraph(pos, -1).remove()
        return shape
      }
    }

    shape.exportAs = function (options) {
      if (typeof options === 'string') {
        var path = options
        options = {'path': path}
        return nativeShape.exportAs(options)
      } else if (typeof options === 'object') {
        return nativeShape.exportAs(options)
      }
    }

    shape.zIndex = function (cmd) {
      if (typeof cmd === 'string') {
        return nativeShape.setZindex(cmd, true)
      } else if (typeof cmd === 'number') {
        var command
        var index = nativeShape.getZindex(null, true)
        if (index < cmd) {
          command = 'forward'
        } else if (index > cmd) {
          command = 'back'
        }
        while (index !== cmd) {
          nativeShape.setZindex(command, true)
          index = nativeShape.getZindex(null, true)
        }
        return new Shape(nativeShape)
      } else {
        return nativeShape.getZindex(null, true)
      }
    }

    shape.z = shape.zIndex

    shape.tag = {
      get: function (name) {
        return nativeShape.tags(null, true).get(name, true)
      },
      set: function (name, value) {
        nativeShape.tags(null, true).set({name: name, value: value}, true)
        return shape
      },
      remove: function (name) {
        nativeShape.tags(null, true).remove(name, true)
        return shape
      }
    }

    shape.tags = nativeShape.tags(null, true).all(null, true)

    shape.table = function () {
      var nativeTable = shape.attr('Table')
      var table = []
      var rowIndex = 0
      var colIndex = 0
      var row

      for (rowIndex = 0; rowIndex < nativeTable.length; rowIndex++) {
        row = []
        for (colIndex = 0; colIndex < nativeTable[rowIndex].length; colIndex++) {
          var cell = new Shape(nativeTable[rowIndex][colIndex], true)
          row.push(cell)
        }
        table.push(row)
      }
      return table
    }
  }
}

module.exports = Shape
