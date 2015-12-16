var _ = require('lodash');
var Slide = require('./slide');
var Paragraph = require('./paragraph')

function Shape(nativeShape) {
    var shape = this;
    
    shape.attr = function(name, value) {
         if(typeof value !== 'undefined') {
            return new Shape(nativeShape.attr({name:name, value:value}, true));
        } 
        return nativeShape.attr(name, true);
    }
    
    //inject attr
    _.assign(shape, require('./shape.attr'));
    
    shape.remove = function() {
        return nativeShape.remove(null, true);
    }
    
    shape.duplicate = function() {
        return new Shape(nativeShape.duplicate(null, true));
    }
    
    shape.paragraph = function(start, length) {
        start = start || -1;
        length = length || -1;
        return new Paragraph(nativeShape.paragraph({'start': start, 'length': length}, true));
    }
    
    shape.p = shape.paragraph;
    
    shape.addLine = function(text, pos) {
        if(typeof pos !== 'number') {
            shape.paragraph(shape.paragraph().count() + 1, -1).text(text);
            return shape;
        } else {
            shape.paragraph(pos, -1).text(text);
            return shape;
        }
    }
    
    shape.removeLine = function(pos) {
        if(typeof pos !== 'number') {
            shape.paragraph(shape.paragraph().count(), -1).remove();
            return shape;
        } else {
            shape.paragraph(pos, -1).remove();
            return shape;
        }
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
    
    shape.zIndex = function(cmd) {
        if(typeof cmd !== 'string') {
            return nativeShape.getZindex(null, true);   
        } else {
            return nativeShape.setZindex(cmd, true);   
        }
    }
    
    shape.z = shape.zIndex;
    
    shape.tag = {
        get: function(name) {
            return nativeShape.tags(null, true).get(name, true);
        },
        set: function(name, value) {
            nativeShape.tags(null, true).set({name: name, value:value}, true);
            return shape;
        },
        remove: function(name) {
           nativeShape.tags(null, true).remove(name, true);
           return shape; 
        }
    }
    
    shape.tags = nativeShape.tags(null, true).all(null, true);
    
    shape.getType = function() {
        return nativeShape.getType(null, true);
    }
}

module.exports = Shape;