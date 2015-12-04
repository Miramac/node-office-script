var _ = require('lodash');
var Slide = require('./slide.js');

function Shape(nativeShape) {
    var shape = this;
    
    shape.attr = function(name, value) {
        if(value) {
            return new Shape(nativeShape.attr({name:name, value:value}, true));
        } 
        return nativeShape.attr(name, true);
    }
    
    //inject attr
    _.assign(shape, require('./shape.attr'));
    
    shape.remove = function() {
        return nativeShape.remove(null, true);
    }
    
    shape.dublicate = function() {
        return new Shape(nativeShape.dublicate(null, true));
    }
    
    shape.paragraphs = function(start, length) {
        
    }
    
    shape.textReplace = function(findString, replaceString) {
        return new Shape(nativeShape.textReplace({'find': findString, 'replace': replaceString}));
    }
    
    shape.exportAs = function(options) {
        if(typeof options === 'string') {
            var path = options;
            options = {'path': path}
            return nativeShape.exportAs(options);
        } else if(typeof options === 'object') {
            return nativeShape.exportAs(options);
        }
    }
}

module.exports = Shape;