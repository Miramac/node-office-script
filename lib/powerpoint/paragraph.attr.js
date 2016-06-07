var attributes = {
  text: function (text) {
    return this.attr('Text', text)
  },

  count: function () {
    return this.attr('Count')
  },

  // Font properties
  fontName: function (fontName) {
    return this.attr('Name', fontName, 'font')
  },

  fontSize: function (fontName) {
    return this.attr('Size', fontName, 'font')
  },

  fontColor: function (fontName) {
    return this.attr('Color', fontName, 'font')
  },

  fontItalic: function (fontName) {
    return this.attr('Italic', fontName, 'font')
  },

  fontBold: function (fontName) {
    return this.attr('Bold', fontName, 'font')
  },

  // Format properties
  align: function (align) {
    return this.attr('Alignment', align, 'format')
  },

  indent: function (indent) {
    return this.attr('IndentLevel', indent, 'format')
  },

  bulletCharacter: function (bulletCharacter) {
    return this.attr('BulletCharacter', bulletCharacter, 'format')
  },

  bulletFontName: function (bulletFontName) {
    return this.attr('BulletFontName', bulletFontName, 'format')
  },

  bulletFontBold: function (bulletFontBold) {
    return this.attr('BulletFontBold', bulletFontBold, 'format')
  },

  bulletFontSize: function (bulletFontSize) {
    return this.attr('BulletFontSize', bulletFontSize, 'format')
  },

  bulletFontColor: function (bulletFontColor) {
    return this.attr('BulletFontColor', bulletFontColor, 'format')
  },

  bulletVisible: function (bulletVisible) {
    return this.attr('BulletVisible', bulletVisible, 'format')
  },

  bulletRelativeSize: function (bulletRelativeSize) {
    return this.attr('BulletRelativeSize', bulletRelativeSize, 'format')
  },

  firstLineIndent: function (firstLineIndent) {
    return this.attr('FirstLineIndent', firstLineIndent, 'format')
  },

  leftIndent: function (leftIndent) {
    return this.attr('LeftIndent', leftIndent, 'format')
  },

  lineRuleBefore: function (lineRuleBefore) {
    return this.attr('LineRuleBefore', lineRuleBefore, 'format')
  },

  lineRuleAfter: function (lineRuleAfter) {
    return this.attr('LineRuleAfter', lineRuleAfter, 'format')
  },

  hangingPunctuation: function (hangingPunctuation) {
    return this.attr('HangingPunctuation', hangingPunctuation, 'format')
  },

  spaceBefore: function (spaceBefore) {
    return this.attr('SpaceBefore', spaceBefore, 'format')
  },

  spaceAfter: function (spaceAfter) {
    return this.attr('SpaceAfter', spaceAfter, 'format')
  },

  spaceWithin: function (spaceWithin) {
    return this.attr('SpaceWithin', spaceWithin, 'format')
  }
}

module.exports = attributes
