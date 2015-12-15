
var attributes = {
    se: this,
     text:  function(text) {
        return this.attr('Text', text);
    },
    
    count: function() {
        return this.attr('Count');
    },
    
    align: function (align) {
            return this.attr('Alignment', align, 'format');
    },
    
    fontName: function (fontName) {
            return this.attr('Name', fontName, 'font');
    },
    
    fontSize: function (fontName) {
            return this.attr('Size', fontName, 'font');
    },
    
    fontColor: function (fontName) {
            return this.attr('Color', fontName, 'font');
    },
    
    italic: function (fontName) {
            return this.attr('Italic', fontName, 'font');
    },
    
    bold: function (fontName) {
            return this.attr('Bold', fontName, 'font');
    }
}

module.exports = attributes;
