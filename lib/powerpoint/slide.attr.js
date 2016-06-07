var attributes = {
  /**
    * Setzt die PowerPoint-Folie an die übergebene Name oder gibt den Namen einer PowerPoint-Folie aus, wenn der der Parameter 'name' nicht definiert wird.
    * @method pos
    * @param {Number} pos
    * @chainable
    *
    * @example
    * Gibt dName der Folie zurück und schreibt diese in die Variable 'slideName'.
    * @example
    *     var slideName = slide.name()
    *
    * @example
    * Setzt den Folie Name.
    * @example
    *     slide.name('Slide Fu Bar')
    */
  name: function (name) {
    return this.attr('Name', name)
  },

  /**
  * Setzt die PowerPoint-Folie an die übergebene Position oder gibt die Position einer PowerPoint-Folie aus, wenn der der Parameter 'pos' nicht definiert wird.
  * @method pos
  * @param {Number} pos
  * @chainable
  *
  * @example
  * Gibt die Position der Folie aus und schreibt diese in die Variable 'sildePos'.
  * @example
  *     var slidePos = slide.pos()
  *
  * @example
  * Schiebt die Folie an die dritte Stelle.
  * @example
  *     slide.pos(3)
  */
  pos: function (pos) {
    return this.attr('Pos', pos)
  },

  /**
  * Gibt die Nummer einer PowerPoint-Folie aus.
  * @method number
  * @chainable
  * @readonly
  *
  * @example
  * Gibt die Nummer der Folie aus und schreibt diese in die Variable 'sildeNum'.
  * @example
  *     var slideNum = slide.number()
  */
  number: function () {
    return this.attr('Number')
  }

}

module.exports = attributes
