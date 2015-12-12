# OfficeScript

*Early alpha stage...*

Scripting MS Office application with node.js using [NetOffice](http://netoffice.codeplex.com/) and edge.js.

## Install
```sh
npm install office-script --save
```

# PowerPoint

PowerPoint application automation. 
```javascript
    var path = require('path');
    var powerpoint = require('office-script').powerpoint;

    //Create a new instance of PowerPoint an try to open Presentation
    powerpoint.open(path.join(__dirname, 'Presentation01.pptx'), function(err, presentation) {
        if(err) throw err;
        //use presentation object
        console.log('Presentation path:', presentation.attr({name:'Path'}, true));
        //Get slides
        presentation.slides(null, function(err, slides) {
            if(err) throw err;
            console.log('Slides count:', slides.length);
            //get shapes on slide 1
            slides[0].shapes(null, function(err, shapes) {
                console.log('Shape count on slide 1:', shapes.length);
                shapes[0].attr({'name':'Text', 'value': 'Fu Bar'}, true); //Set text value
                console.log('Get text first shape:', shapes[0].attr({'name':'Text'}, true));
                //close presentation
                presentation.close(null, function(err) {
                    if(err) throw err;
                    //exit powerpoint
                    powerpoint.quit()
                });
            });
        });
    });
```
# sync vs. async
Office-Script is written in an async pattern, but application automation has serious problems with async... 

Because of this, I recommend to use the sync presentation wrapper. Also it has more readable API. 

```javascript
    var path = require('path');
    var Presentation = require('office-script').Presentation;
    
    //open a new PPT Presentation  
    var presentation = new Presentation(path.join(__dirname,'Presentation01.pptx'));
    
    //get presentation slides  
    var slides = presentation.slides();
    console.log('Slide count:', slides.length);
    
    //Get all shapes of the first slide 1 
    var shapes = slides[0].shapes();
    console.log('Title shape count:', shapes.length);
    
    //get name and text of the first shape
    console.log('shape name:', shapes[0].name());
    console.log('Title shape count:', shapes[0].text());
    
    //change name of the first
    shapes[0].name('First Shape');
    shapes[0].text('FuBar');
    
    //Setter retun the destination object so you can chain them
    shapes[0].top(10).left(10).height(100).width(100);
      
    //Save presentation as PDF (sync)
    presentation.saveAs({name:path.join(__dirname,'Presentation01.pdf'), type:'pdf'});
    //SaveAs new presentation and quit application 
    presentation.saveAs(path.join(__dirname,'Presentation01_New.pptx'));
    presentation.quit(); //Close presentation & quit application  


```
# Synchronous API
## Presentation([path]);
If path exists the presentation will be open. 
If `path` not exists, an allready open presentation with the name of `path` will be used. 
If `path` is missing or `null`, the active presentation is used.
### Property methods 
* .name() `String readonly` Presentation name
* .path() `String readonly` Presentation path
* .fullName() `String readonly` Presentation path with presentation name

### presentation methods
* .addSlide([pos]) *returns Slide*

## presentation.slides([selector])
Get presentation slides. Optional filterd by the selector.
### Property methods
If `value` is set, is set the property and return the slide
* .name([value]) `String` 
* .pos([value] `Int`
* .number([value])  `Int readonly`

### Slide methods
* .remove()
* .addTextbox(options) *returns Shape*
* .addPicture(options) *returns Shape*


## presentation.shapes([selector] [, context])
Get presentation shapes. Optional filterd by the selector. Context is an optional slides array.
## slide.shapes([selector])
Get slide shapes. Optional filterd by the selector. Context is an optional slides array.
### Property methods
If `value` is set, is set the property and return the shape
* .name([value]) `String`
* .text([value]) `String`
* .top([value])  `Float`
* .left([value])  `Float`
* .height([value])  `Float`
* .width([value])  `Float`
* .rotation([value])  `Float`
* .fill([value])  `String`
* .altText([value])  `String`
