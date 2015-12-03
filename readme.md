# OfficeScript

*Early alpha stage...*

Scripting MS Office application with node.js using [NetOffice](http://netoffice.codeplex.com/) and edge.js.

## Install
```sh
npm install office-script --save
```

## PowerPoint

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

Using sync presentation object
```javascript
    var path = require('path');
    var Presentation = require('office-script').Presentation;
    
    //open a new PPT Presentation  
    var presentation = new Presentation(path.join(__dirname,'Presentation01.pptx'));
    
    //get presentation slides  
    var slides = presentation.slides();
    console.log('Slide count:', slides.length);
    
    //Get only Shapes with name 'Title 1' on slide 1 
    var shapes = presentation.shapes({'attr:Name':'Title 1'}, slides[0]);
    console.log('Title shape text:', shapes[0].attr('Text', true));
    
    //Save presentation as PDF (sync)
    presentation.saveAs({name:path.join(__dirname,'Presentation01.pdf'), type:'pdf'},true);
    //SaveAs new presentation and quit application 
    presentation.saveAs(path.join(__dirname,'Presentation01_New.pptx'), function(err) {
        if (err) throw err;
        presentation.quit()
    });


```

### .slides([selector])
Get presentation slides. Optional filterd by the selector.

### .shapes([selector] [,context])
Get presentation shapes. Optional filterd by the selector. Context is an optional slides array.

### .attr(params, callback)
If callback is `true`, it's a sync function and return the current object for chaining
* Getter

```javascript
presentation.attr('Path', function(pptPat) {
    console.log(pptPat)
})
//or sync
var pptPath = presentation.attr('Path', true)
console.log(pptPat)
```
* Setter

```javascript
shape.attr({name:'Text', value:'FuBar'}, function() {
    console.log("Shape text set to 'FuBar")
})
//or sync
shape.attr({name:'Text', value:'FuBar'}, true)
console.log("Shape text set to 'FuBar")
```
#### Presentation properties
* Name `String readonly`
* Path `String readonly`
* FullName `String readonly`

#### Slide properties
* Name `String`
* Pos `Int`
* Number  `Int readonly`

#### Shape properties
* Name `String`
* Text `String`
* Top  `Float`
* Left  `Float`
* Height  `Float`
* Width  `Float`
* Rotation  `Float`
* Fill  `String`
* AltText  `String`
