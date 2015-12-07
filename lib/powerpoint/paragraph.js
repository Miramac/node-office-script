var Shape = require('./shape.js');

function Paragraph(nativeParagraph) {
    var paragraph = this;
    
    paragraph.text = function(string) {
        if(string) {
            return nativeParagraph.text = string;
        } else {
            return nativeParagraph.text;   
        }
    }
    
    paragraph.remove = function() {
        return nativeParagraph.remove(null, true);
    }
    
    paragraph.count = function() {
        return nativeParagraph.count;
    }
}

module.exports = Paragraph;