var application = require('./application')
var Slide = require('./slide')

function Presentation(path) {
    var ppt = application.open(path, true);
    var presentation = this;
    
    presentation.slides = function(selector) {
        var slides = [];
        var i;
        var nativeSlides = ppt.slides(selector, true)
        for(i=0; i<nativeSlides.length; i++) {
            slides.push(new Slide(nativeSlides[i]))
        }
        return slides;
    }
    
   presentation.nativeSlides = function(selector) {
        return ppt.slides(selector, true);
    }
    
    presentation.shapes = function(selector, slides) {
         slides = (typeof slides === 'string') ? presentation.nativeSlides(slides) : slides;
         slides = (typeof slides === 'undefined') ? presentation.nativeSlides() : slides;
         slides = (slides.length)  ? slides : [slides]
         
         var shapes = [], i;
         for(i=0; i<slides.length; i++) {
             shapes = shapes.concat(slides[i].shapes(selector, true))
         }
         
        return shapes;
    }
    
   
    presentation.attr = function(input, cb) {
       return ppt.attr(input, cb)
    };
   
    presentation.quit = function(cb) {
        application.quit(null, cb)
    };
    presentation.close = presentation.quit;
    
    presentation.save = function(cb) {
        ppt.save(null, cb)
    };
    
    presentation.saveAs = function(input, cb) {
        ppt.saveAs(input, cb)
    };
    
    presentation.saveAsCopy = function(input, cb) {
        ppt.saveAs(input, cb)
    };
    
    //* Attr shortcuts
     presentation.path = function() {
        return presentation.attr('Path', true); 
    }
    
    presentation.name = function() {
        return presentation.attr( 'Name', true); 
    }
    
    presentation.fullName = function() {
        return presentation.attr( 'FullName', true); 
    }
    
}

module.exports = Presentation;
