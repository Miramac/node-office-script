var Shape = require('./shape');

function Paragraph(nativeParagraph) {
    var paragraph = this;
    
    paragraph.attr = function(name, value) {
        if(value) {
            return new Paragraph(nativeParagraph.attr({name:name, value:value}, true));
        } 
        return nativeParagraph.attr(name, true);
    }
    
    paragraph.text =  function(text) {
        return this.attr('Text', text);
    }
    
    paragraph.count = function() {
        return this.attr('Count');
    }
    
    paragraph.remove = function() {
        return nativeParagraph.remove(null, true);
    }
    
    
}

module.exports = Paragraph;