var _ = require('lodash');

function Paragraph(nativeParagraph) {
    var paragraph = this;
    var native = {
        paragraph: nativeParagraph,
        format: nativeParagraph.format(null, true),
        font: nativeParagraph.font(null, true)
    }
    paragraph.attr = function(name, value, target) {
        target = target || 'paragraph';
        if(typeof value !== 'undefined') {
            console.log("SLDMSKLD");
            return new Paragraph(native[target].attr({name:name, value:value}, true));
        } 
        return native[target].attr(name, true);
    }
    
    //inject attr
    _.assign(paragraph, require('./paragraph.attr'));
    
   
    
    paragraph.remove = function() {
        return nativeParagraph.remove(null, true);
    }
    
    paragraph._format =function() {
        return native.format;
    }
    
    paragraph._font = function() {
        return native.font;
    }
    
}

module.exports = Paragraph;