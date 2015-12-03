
function Slide(nativeSlide) {
    var slide = {};
    
    slide.attr = function(name, value , cb) {
        cb = cb || true;
        if(value) {
        return new Slide(nativeSlide.attr({name:name, value:value}, cb));
        } 
        return  nativeSlide.attr(name, cb);
    }
    
    slide.pos = function (value, cb) {
        cb = cb || true;
        if(value) {
            return slide.attr('Pos', value, cb);
        } 
        return  slide.attr('Pos', false, cb);
    };
    
    return slide;
}

module.exports = Slide;
