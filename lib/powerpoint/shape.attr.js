var attributes = {
  /**
  * Setzt den Namen eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Namen des PowerPoint-Objekts zurück, wenn der Parameter 'name' nicht definiert ist.
  * @method name
  * @param {String} name
  * @chainable
  *
  * @example
  * Liest den Namen des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeName'.
  * @example
  *     var shapeName = $shapes('selector').name();
  *
  * @example
  * Ändert den Namen des PowerPoint-Objekts in 'Textbox_1337'.
  * @example
  *     $shapes('selector').name('Textbox_1337')
  */
  name: function (name) {
    return this.attr('Name', name)
  },

  /**
  * Setzt den Text eines PowerPoint-Objekts auf den übergebenen Wert oder gib den Text des PowerPoint-Objekts zurück, wenn der Parameter 'text' nicht definiert ist.
  * @method text
  * @param {String} text
  * @chainable
  *
  * @example
  * Liest den Text des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeText'.
  * @example
  *     var shapeText = $shapes('selector').text()
  *
  * @example
  * Setzt den Text des PowerPoint-Objekts auf 'Fu Bar'.
  * @example
  *     $shapes('selector').text('Fu Bar')
  */
  text: function (text) {
    return this.attr('Text', text)
  },

  /**
  * Setzt den Wert des Abstands nach oben eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert des Abstands nach oben des PowerPoint-Objekts zurück, wenn der Parameter 'top' nicht definiert ist.
  * @method top
  * @param {Number} top
  * @chainable
  *
  * @example
  * Liest den Wert des Abstands nach oben des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeTop'.
  * @example
  *     var shapeTop = $shapes('selector').top()
  *
  * @example
  * Setzt den Wert des Abstands nach oben des PowerPoint-Objekts auf 1337.
  * @example
  *     $shapes('selector').top(1337)
  */
  top: function (top) {
    return this.attr('Top', top)
  },

  /**
  * Setzt den Wert des Abstands nach links eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert des Abstands nach links des PowerPoint-Objekts zurück, wenn der Parameter 'left' nicht definiert ist.
  * @method left
  * @param {Number} left
  * @chainable
  *
  * @example
  * Liest den Wert des Abstands nach links des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeLeft'.
  * @example
  *     var shapeLeft = $shapes('selector').left()
  *
  * @example
  * Setzt den Wert des Abstands nach links des PowerPoint-Objekts auf 1337.
  * @example
  *     $shapes('selector').left(1337)
  */
  left: function (left) {
    return this.attr('Left', left)
  },

  /**
  * Setzt die Höhe eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert der Höhe des PowerPoint-Objekts zurück, wenn der Parameter 'height' nicht definiert ist.
  * @method height
  * @param {Number} height
  * @chainable
  *
  * @example
  * Liest den Wert der Höhe des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeHeight'.
  * @example
  *     var shapeHeight = $shapes('selector').height()
  *
  * @example
  * Setzt den Wert der Höhe des PowerPoint-Objekts auf 1337.
  * @example
  *     $shapes('selector').height(1337)
  */
  height: function (height) {
    return this.attr('Height', height)
  },

  /**
  * Setzt die Breite eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert der Breite des PowerPoint-Objekts zurück, wenn der Parameter 'width' nicht definiert ist.
  * @method width
  * @param {Number} width
  * @chainable
  *
  * @example
  * Liest den Wert der Breite des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeWidth'.
  * @example
  *     var shapeWidth = $shapes('selector').width()
  *
  * @example
  * Setzt den Wert der Breite des PowerPoint-Objekts auf 1337.
  * @example
  *     $shapes('selector').width(1337)
  */
  width: function (width) {
    return this.attr('Width', width)
  },

  /**
  * Rotiert ein PowerPoint-Objekt um den übergebenen Wert in Grad nach rechts oder gibt den Wert der Rotation eines PowerPoint-Objekts zurück, wenn der Parameter 'rotation' nicht definbiert ist.
  * @method rotation
  * @param {Number} rotation
  * @chainable
  *
  * @example
  * Liest den Wert der Rotation des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeRotation'.
  * @example
  *     var shapeRotation = $shapes('selector').rotation()
  *
  * @example
  * Dreht das PowerPoint-Objekt um 90 Grad nach rechts.
  * @example
  *     $shapes('selector').rotation(90)
  */
  rotation: function (rotation) {
    return this.attr('Rotation', rotation)
  },

  /**
  * Befüllt ein PowerPoint-Objekt mit der Farbe mit dem übergebenen Wert oder gibt den Wert der Farbe mit welcher das PowerPoint-Objekt gefüllt ist aus, wenn der Parameter 'fill' nicht definiert ist.
  * @method fill
  * @param {Number} fill
  * @chainable
  *
  * @example
  * Liest den Wert der Farbe des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeColor'.
  * @example
  *     var shapeColor = $shapes('selector').fill()
  *
  * @example
  *  Befüllt das PowerPoint-Objekt mit der Farbe mit dem Wert '#FF9900'.
  * @example
  *     $shapes('selector').fill('FF9900')
  */
  fill: function (fill) {
    return this.attr('Fill', fill)
  },

  /**
  * Gibt den nächsten sog. 'Parent' (Elternteil, übergeordnetes Objekt) eines PowerPoint-Objekts zurück.
  * @method parent
  * @chainable
  *
  * @example
  * Liest den Parent des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeParent'.
  * @example
  *     var shapeParent = $shapes('selector').parent()
  */
  parent: function () {
    throw new Error('Not implemented.')
  // return this.attr('Parent')
  },

  altText: function (altText) {
    return this.attr('AltText', altText)
  },

  title: function (title) {
    return this.attr('Title', title)
  },

  type: function () {
    return this.attr('Type')
  }

}

module.exports = attributes
