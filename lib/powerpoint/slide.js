var _ = require('lodash');
var Shape = require('./shape.js');

function Slide(nativeSlide) {
    var slide = this;
    
    slide.attr = function(name, value) {
        if(value) {
            return new Slide(nativeSlide.attr({name:name, value:value}, true));
        } 
        return nativeSlide.attr(name, true);
    }
    
    //inject attr
    _.assign(slide, require('./slide.attr'));
    
    slide.remove = function() {
         return nativeSlide.remove(null, true);
    }
    
    slide.dublicate = function() {
        return new Slide(nativeSlide.dublicate(null, true));
    }
    
    slide.addTextBox = function(options) {
        if(!options) {
            options = {'left': 0, 'top': 0, 'height': 100, 'width': 100};
        } else {
            options.left ? options.left : 0;
            options.top ? options.top : 0;
            options.height ? options.height : 100;
            options.width ? options.width : 100;
        }
        return new Shape(nativeSlide.addTextBox(options, true));
    }
    
    slide.shapes = function(selector) {
        var shapes = [];
        var i;
        var nativeShapes = slide.shapes(selector, true)
        for(i=0; i<nativeShapes.length; i++) {
            shapes.push(new Shape(nativeShapes[i]))
        }
        return shapes;
    }
}

module.exports = Slide;
