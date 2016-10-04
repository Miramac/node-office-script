var _ = require('lodash')

function Character (nativeCharacter) {
  var character = this
  var native = {
    character: nativeCharacter,
    font: nativeCharacter.font(null, true)
  }
  character.attr = function (name, value, target) {
    target = target || 'character'
    if (typeof value !== 'undefined') {
      return new Character(native[target].attr({name: name, value: value}, true))
    }
    return native[target].attr(name, true)
  }

  // inject attr
  _.assign(character, require('./character.attr'))

  character.remove = function () {
    return nativeCharacter.remove(null, true)
  }

  character._font = function () {
    return native.font
  }
}

module.exports = Character
