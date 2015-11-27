var path = require('path');
var Presentation = require('../').Presentation

var presentation = new Presentation(path.join(__dirname,'/data/Testpptx_00.pptx'));
var slides = presentation.slides({"attr:Pos":'1'});
console.log('slides', slides.length);

var shapes = presentation.shapes({'tag:ctobjectdata.id':'FABI'}, slides);
console.log('shapes', shapes.length);

console.log('shapes', shapes[0].attr('Left',true));
console.log('shapes', shapes[0].attr({name:'Left', value:10},true));
console.log('shapes', shapes[0].attr('Left',true));


setTimeout(presentation.quit, 500);


/*
var Shape = require('../lib/report/wrapper/shape');
var Shapes = require('../lib/report/wrapper/shapes');
//, reportApp = report.application


var presentation = application.open(__dirname+'/data/Testpptx_01.pptx', true);
var slides = presentation.slides({"attr:Name":'Slide1,Slide2'}, true)

var i, j;

// for(i=0; i<slides.length; i++) {
    // console.log("Slide:", slides[i].attr('Name', true))
    // var shapes = slides[i].shapes({'attr:Name':'TextBox 3','tag:ctobjectdata.id':'shape1'},true);
    // for(j=0; j<shapes.length; j++) {
        // console.log("Shape:", shapes[j].attr('Name', true))
    // }
// }


var slides = presentation.slides({"attr:Name":'Slide2'}, function(err, slides){
    var i, j;

    for(i=0; i<slides.length; i++) {
        console.log("Slide:", slides[i].attr('Name', true))
        var shapes = slides[i].shapes({'attr:Name':'TextBox 3','tag:ctobjectdata.id':'shape1'},true);
        for(j=0; j<shapes.length; j++) {
            console.log("Shape:", shapes[j].attr('Name', true))
        }
    }
})
var slides = presentation.slides({"attr:Name":'Slide1'}, function(err, slides){
    var i, j;

    for(i=0; i<slides.length; i++) {
        console.log("Slide:", slides[i].attr('Name', true))
        var shapes = slides[i].shapes({'attr:Name':'TextBox 3','tag:ctobjectdata.id':'shape1'},true);
        for(j=0; j<shapes.length; j++) {
            console.log("Shape:", shapes[j].attr('Name', true))
        }
    }
})

*/

// var chart = slides[1].shapes('chart1',true)[0]

// Q.nfcall(chart.exportAs ,__dirname+'\\data\\Testpptx_01.png').done(function() {presentation.close(null,application.quit)})

//console.log(presentation.getType(null, true));

//var shapes = Shapes(shapes);

//console.log(shapes.count());

//console.log($shape.tag('Fu', 'Bar').tag('Fu'));

//console.log(shapes[0].tags(null, true).set({name:'Fu', value:'Bar'}, true).set({name:'Hans', value:'Dampf'}, true).get('FU',true));

// console.log($shape.attr('Name' , 'Foo'));

// console.log($shape.name('bar'));

// console.log($shape.attr('Name'));

// slides[0].addTextbox({top:100, left:100, height:200, width:200}, function (err, shape) {
    // console.log(shape);
    // var s = Shape(shape);
    // s.text('Foo Bar');
    // console.log(shape.attr({ name: "Height" }, true))
    // console.log(s.left());
// })
//presentation.close(null,application.quit)





/*
 report.open(__dirname+'\\data\\Testpptx_02.pptx', function(err, presentation) {
    //use presentation object
    console.log('Presentation Name:', presentation.attr({name:'Name'}, true)); 
    console.log('Presentation Path:', presentation.attr({name:'Path'}, true)); 
    console.log('Presentation FullName:',presentation.attr({'name':'FullName'}, true));
    presentation.slides(null, function(err, slides) {
        if(err) throw err;
        console.log('Slides count:', slides.length);
        console.log('Slides props:', slides);
        slides[1].shapes(null, function(err, shapes) {
            var shape0 = shapes[0];
            var shape1 = shapes[1];
            console.log('Shape count on slide 1:', shapes.length);
            shape0.attr({'name':'Text', 'value': 'Fu Bar'}, true); //Set text value
            console.log('get Text first shape:', shape0.attr({'name':'Text'}, true));
            
            console.log('get Text first shape:', shape0.attr({'name':'Text'}, true));
            
            console.log(slides[1].addTextbox({}, true).attr({name:'Name'},true));
            console.log(shape1.paragraph({'start':5}, true).attr({name:'Text', value:"test"}, true));
           
            // close presentation
             presentation.close(null, function(err){
                if(err) throw err;
                report.quit()
            });
        });
    });
});

    */
    
    