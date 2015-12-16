
var attributes = {
     text:  function(text) {
        return this.attr('Text', text);
    },
    
    count: function() {
        return this.attr('Count');
    },
    
    //Font properties
    fontName: function (fontName) {
            return this.attr('Name', fontName, 'font');
    },
    
    fontSize: function (fontName) {
            return this.attr('Size', fontName, 'font');
    },
    
    fontColor: function (fontName) {
            return this.attr('Color', fontName, 'font');
    },
    
    fontItalic: function (fontName) {
            return this.attr('Italic', fontName, 'font');
    },
    
    fontBold: function (fontName) {
            return this.attr('Bold', fontName, 'font');
    },
    
    fontCopy: function () {
            throw new Error('Not implemented.');
    },
    
    //Format properties
    align: function (align) {
            return this.attr('Alignment', align, 'format');
    },
    
    bullet: function (bullet) {
            return this.attr('Bullet', bullet, 'format');
    },
    
    indent: function (indent) {
            return this.attr('IndentLevel', indent, 'format');
    },
    
    formatCopy: function () {
            throw new Error('Not implemented.');
    }
    
}

module.exports = attributes;
