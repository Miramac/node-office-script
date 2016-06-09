var _ = require('lodash')

function Paragraph (nativeParagraph) {
  var paragraph = this
  var native = {
    paragraph: nativeParagraph,
    format: nativeParagraph.format(null, true),
    font: nativeParagraph.font(null, true)
  }
  paragraph.attr = function (name, value, target) {
    target = target || 'paragraph'
    if (typeof value !== 'undefined') {
      return new Paragraph(native[target].attr({name: name, value: value}, true))
    }
    return native[target].attr(name, true)
  }

  // inject attr
  _.assign(paragraph, require('./paragraph.attr'))

  paragraph.remove = function () {
    return nativeParagraph.remove(null, true)
  }

  paragraph._format = function () {
    return native.format
  }

  paragraph._font = function () {
    return native.font
  }

  paragraph.copyFont = function (srcParagraph) {
    paragraph.fontName(srcParagraph.fontName())
    paragraph.fontSize(srcParagraph.fontSize())
    paragraph.fontColor(srcParagraph.fontColor())
    paragraph.fontItalic(srcParagraph.fontItalic())
    paragraph.fontBold(srcParagraph.fontBold())
  }

  paragraph.copyFormat = function (srcParagraph) {
    paragraph.align(srcParagraph.align())
    paragraph.indent(srcParagraph.indent())
    paragraph.bulletCharacter(srcParagraph.bulletCharacter())
    paragraph.bulletFontName(srcParagraph.bulletFontName())
    paragraph.bulletFontBold(srcParagraph.bulletFontBold())
    paragraph.bulletFontSize(srcParagraph.bulletFontSize())
    paragraph.bulletFontColor(srcParagraph.bulletFontColor())
    paragraph.bulletVisible(srcParagraph.bulletVisible())
    paragraph.bulletRelativeSize(srcParagraph.bulletRelativeSize())
    paragraph.firstLineIndent(srcParagraph.firstLineIndent())
    paragraph.leftIndent(srcParagraph.leftIndent())
    paragraph.lineRuleBefore(srcParagraph.lineRuleBefore())
    paragraph.hangingPunctuation(srcParagraph.hangingPunctuation())
    paragraph.spaceBefore(srcParagraph.spaceBefore())
    paragraph.spaceAfter(srcParagraph.spaceAfter())
    paragraph.spaceWithin(srcParagraph.spaceWithin())
    return paragraph
  }

  paragraph.copyStyle = function (srcParagraph) {
    paragraph.copyFont(srcParagraph)
    paragraph.copyFormat(srcParagraph)
    return paragraph
  }
}

module.exports = Paragraph
