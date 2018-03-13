# OfficeScript

Office-Script is a Microsoft Office application automation with node.js.
It does not work with the Open Office XML document, instead it accesses the COM interop interface of the Offices application.
Therefore, you must have Office installed! Also be carefull, **Microsoft strongly recommends against Office Automation from software solutions**  https://support.microsoft.com/en-us/kb/257757

*Only on tested with Office 2007 and Office 2016.*

> *Work in progress.. Just ask if you have any question or feature requests!*

## Install
```sh
npm install office-script --save
```

# PowerPoint

PowerPoint application automation. 
```javascript
  var path = require('path');
  var powerpoint = require('office-script').powerpoint;

  //Create a new instance of PowerPoint and try to open the presentation
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
Office-Script is written in an async pattern, but application automation can has serious problems with async... 

Because of this, I recommend to use the sync presentation wrapper. It also has the more readable API. 

```javascript
var path = require('path')
var Presentation = require('office-script').Presentation
var presentation
try {
  // open a new PPT Presentation
  presentation = new Presentation(path.join(__dirname, 'Presentation01.pptx'))

  // get presentation slides
  var slides = presentation.slides()
  console.log('Slide count: ', slides.length)

  // Get all shapes of the first slide
  var shapes = slides[0].shapes()
  console.log('Title shape count:', shapes.length)

  // get name and text of the first shape
  console.log('shape name:', shapes[0].name())
  console.log('Title shape text:', shapes[0].text())

  // change name of the first
  shapes[0].name('First Shape')
  // change text of the first
  shapes[0].text('FuBar')

  // Setter retun the destination object so you can chain them
  shapes[0].top(10).left(10).height(100).width(100)

  // Save presentation as PDF (sync)
  presentation.saveAs({name: path.join(__dirname, 'Presentation01.pdf'), type: 'pdf'})
  // SaveAs new presentation and quit application 
  presentation.saveAs(path.join(__dirname, 'Presentation01_New.pptx'))
} catch (e) {
  console.error(e)
}
if (presentation) {
  presentation.quit() // Close presentation & quit application
}
  
```
# Synchronous API
## Presentation([path]);
If path exists the presentation will be open. 
If `path` does not exist, an allready open presentation with the name of `path` will be used. 
If `path` is missing or `null`, the active presentation is used.
___
  
### Presentation methods

#### .addSlide([pos]) *returns slide object*
Adds a new empty slide on the provided postiton an returns it. If no postiton was provided, the new slide will be added at the end.
___
#### .close([callback]) 
Closes the presentation without exiting powerpoint itself.
___
#### .quit([callback])
Closes the presentation and powerpoint itself.
___
#### .save([callback])
Saves the presentation.
___
#### .saveAs(fullName [, callback])
Saves the presentation to the provided path and name.
___
#### .saveAsCopy(fullName [, callback])
Saves the presentation as copy to the provided path and name.
___
#### .textReplace(find, replace)
Find and replace text in the entire presentation.
___


### Property methods

#### .builtinProp([property, value]) `multifunctional` 
Without parameters, returns all builtin properties with their vlaues.
With `property`, returns value of the specific builtin property.
With `property` and `value`, sets value of specific builtin property.
___
#### .customProp([property, value]) `multifunctional` 
Without parameters, returns all custom properties with their vlaues.
With `property`, returns value of the specific custom property.
With `property` and `value`, sets/ccustomreates value of specific custom property
___
#### .fullName() `String readonly` Presentation path with presentation name
#### .name() `String readonly` Presentation name
#### .path() `String readonly` Presentation path
#### .slideHeight() `Number readonly` Slide/presentation height
#### .slideWidth() `Number readonly` Slide/presentation width
#### .type() `String readonly` Presentation type

### Tag methods
#### .tags returns all tags
#### .tag *object with tag functions*
#### .tag.get(name) returns value of specific tag
#### .tag.set(name, value) set value of specific tag
#### .tag.remove(name) removes tag

## presentation.slides([selector])
Get presentation slides. Optional filterd by the selector.
## presentation.activeSlide()
Get active slide.
#### presentation.pasteSlide([index]) 
Pastes the slides on the Clipboard into the Slides collection for the presentation. Returns the pasted slide. Index `-1` moves the slide to the end of the presentation.

### Slide methods
#### .addTextbox(options) *returns shape object*
#### .addPicture(options) *returns shape object*
#### .duplicate() *returns slide object*
#### .remove() *delete the slide*
#### .copy() *copy the slide to the Clipboard* (to paste the slide in an other presentation)
```javascript
presentation01.slides()[0].copy() // Copy first slide from presentation presentation01
presentation02.pasteSlide() // Paste it on the end in presentation presentation02
```

### Property methods
If `value` is provided, it will set the property and return the slide
#### .name([value]) `String` 
#### .number([value]) `Int readonly`
#### .pos([value]) `Int`

### Slide tag methods
#### .tags returns all tags
#### .tag *object with tag functions*
#### .tag.get(name) returns value of specific tag
#### .tag.set(name, value) set value of specific tag
#### .tag.remove(name) removes tag

## presentation.shapes([selector] [, context])
Get presentation shapes. Optional filterd by the selector. Context is an optional slides array.
## presentation.selectedShape()
Get selected shape.
## slide.shapes([selector])
Get slide shapes. Optional filterd by the selector.

### Shape methods
#### .addline(text[, pos]) *returns paragraph object*
#### .duplicate() *returns shape object*
#### .exportAs(options) *returns shape object*
#### .remove()
#### .shape.removeLine(pos) *returns paragraph object*
#### .textReplace(findString, replaceString) *returns shape object*
#### .zIndex([command]) *returns shape object*
#### .has(name) *Check if the current shape has a table, chart or textframe*
```javascript
// export chart-shapes as EMF
if (shapes[i].has('chart')) {
  shapes[i].exportAs({path: path.join(__dirname, shapes[i].name() + '.emf'), type: 'emf'})
}
```
### Property methods

If `value` is provided, it will set the property and return the shape. If not, it will return the value.

#### .altText([value]) `String`
#### .title([value]) `String`
#### .fill([value]) `String`
#### .height([value]) `Float`
#### .left([value]) `Float`
#### .name([value]) `String`
#### .parent() *Not implemented yet*
#### .rotation([value]) `Float`
#### .table() `array` *Read-Only*
#### .text([value]) `String`
#### .top([value]) `Float`
#### .width([value]) `Float`

### Shape tag methods
#### .tags returns all tags
#### .tag *object with tag functions*
#### .tag.get(name) returns value of specific tag
#### .tag.set(name, value) set value of specific tag
#### .tag.remove(name) removes tag

## shape.paragraph(start, length)
Get paragraph object. Optional filterd by start and length.

### Property methods
If `value` is provided, it will set the property and return the shape. If not, it will return the value.
#### .text([value])
#### .count()
#### .fontName([value])
#### .fontSize([value])
#### .fontColor([value])
#### .fontItalic([value])
#### .fontBold([value])
#### .align([value])
#### .indent([value])
#### .bulletCharacter([value])
#### .bulletFontName([value])
#### .bulletFontBold([value])
#### .bulletFontSize([value])
#### .bulletFontColor([value])
#### .bulletVisible([value])
#### .bulletRelativeSize([value])
#### .firstLineIndent([value])
#### .leftIndent([value])
#### .lineRuleBefore([value])
#### .lineRuleAfter([value])
#### .hangingPunctuation([value])
#### .spaceBefore([value])
#### .spaceAfter([value])
#### .spaceWithin([value])

### Paragraph methods
#### .copyFont(srcParagraph)
#### .copyFormat(srcParagraph)
#### .copyStyle(srcParagraph)
#### .remove()
Delete paragraph


 
