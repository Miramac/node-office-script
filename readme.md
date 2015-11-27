# OfficeScript

Office Application scripting in node.js using [NetOffice](http://netoffice.codeplex.com/) and edge.js.


## PowerPoint

PowerPoint Application automation. MS PowerPoint installation on the machine is required!

```javascript
    var path = require('path');
    var pptApplication = require('officescript').report.application;

    //Create a new instance of PowerPoint an try to open Presentation
    pptApplication.open(path.join(__dirname, 'Presentation01.pptx'), function(err, presentation) {
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
                    app.quit()
                });
            });
        });
    });

```


### .attr(params, callback)





