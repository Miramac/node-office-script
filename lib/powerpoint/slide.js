
function Slide(nativeSlide) {
    var slide = {};
    
    slide.attr = function(name, value) {
        if(value) {
            return new Slide(nativeSlide.attr({name:name, value:value}, true));
        } 
        return nativeSlide.attr(name, true);
    }
    
    slide.pos = function(value) {
        if(value) {
            return slide.attr('Pos', value);
        }
        return slide.attr('Pos');
    }
    
    slide.remove = function() {
         return nativeSlide.remove(null, true);
    }
    
    slide.dublicate = function() {
        return nativeSlide.dublicate(null, true);
    }
    
    slide.addTextBox = function({options}) {
        if(!options) {
            options = {'left': 0, 'top': 0, 'height': 100, 'width': 100};
        } else {
            options.left ? options.left : 0;
            options.top ? options.top : 0;
            options.height ? options.height : 100;
            options.width ? options.width : 100;
        }

        return nativeSlide.addTextBox(options, true);
        
    }
    
    slide.shapes = function() {
        
    }

    return slide;
}

module.exports = Slide;
