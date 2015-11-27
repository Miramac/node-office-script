var application = require('./application')

function Presentation(path) {
    var ppt = application.open(path, true);
    var presentation = {};
    
    presentation.slides = function(selector) {
        return ppt.slides(selector, true);
    }
    
    presentation.shapes = function(selector, slides) {
         slides = (typeof slides === 'string') ? presentation.slides(slides) : slides;
         slides = (typeof slides === 'undefined') ? presentation.slides() : slides;
         
         var shapes = [], i;
         for(i=0; i<slides.length; i++) {
             shapes = shapes.concat(slides[i].shapes(selector, true))
         }
         
        return shapes;
    }
    
    presentation.quit = function() {
        application.quit(null, true)
    };
    
    return presentation;
}

module.exports = Presentation;
