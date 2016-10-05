var path = require('path')
var Presentation = require('..').Presentation

var presentation01 = new Presentation(path.join(__dirname, '/data/Testpptx_01.pptx'))
// var presentation02 = new Presentation(path.join(__dirname, '/data/Testpptx_02.pptx'))
//var presentation01 = new Presentation()
var shape = presentation01.slides()[0].shapes()[0]
var para = shape.paragraph(1)
console.log(para.text())
var char = shape.char(4, 2)
console.log(char.text())
console.log(char.fontBold(true))

// var i, j
// var slides = presentation01.slides()
// var shapes
// for (i = 0; i < slides.length; i++) {
//   shapes = slides[i].shapes()
//   for (j = 0; j < shapes.length; j++) {
//     shapes[j].textReplace('<CurrWave>', '12345')
//   }
// }


var replaces = {
  '<Txt_IV_005>': 1234,
  '<Txt_IV_532>': '<xyz> Fu',
  'Testpptx': '<Txt_IV_532> Bar',
  '<xyz>': 'ABC'
}
try {
  // presentation01.textReplace('<CurrWave>', '12345')
  presentation01.textReplace(replaces)
} catch (e) {
  console.error(e)
}

setTimeout(() => {
  presentation01.quit()
}, 5000)


/*


var slide1 = presentation01.slides()[0].copy()
console.log(slide1.name())
var slide2 = presentation02.pasteSlide(-1)
console.log(slide2.pos())
setTimeout(() => {
  presentation01.close()
  presentation02.quit()
}, 5000)
*/

/* var shapes = presentation.shapes()

for (var i = 0; i < shapes.length; i++) {
  if (shapes[i].altText().indexOf('coc_object_id=0698;') > 0) {
  console.log(shapes[i].name())
  } else if (shapes[i].altText() !== '') {
  console.log(shapes[i].altText())
  }
}
*/

// presentation.quit()

/*
var presentation = new Presentation(path.join(__dirname, '/data/Testpptx_00.pptx'))
//var presentation = new Presentation()
// get presentation slides
var slides = presentation.slides()
console.log('Slide count:', slides.length)
var powerpoint = require('../').powerpoint
var shapes = presentation.shapes()

for (var i = 0; i < shapes.length; i++) {
  if(shapes[i].has('chart')) {
    shapes[i].exportAs({path: path.join(__dirname, 'data/chart_4.emf'), type: 'emf'})
  } 
  if(shapes[i].has('table')) {
    console.log(shapes[i].name())
  }
  if(shapes[i].has('text')) {
    console.log(shapes[i].text())
  }
}
presentation.quit()
powerpoint.quit(true, true)
*/
// var i, j, shapes

// console.log('Slide count:', slides.length)
// for (i = 0; i < slides.length; i++) {
//   shapes = slides[i].shapes()
//   console.log('Slide Num:', slides[i].pos())
//   console.log('Shape count:', shapes.length)
//   for (j = 0; j < shapes.length; j++) {
//   console.log(shapes[j].name(), shapes[j].text(), shapes[j].table())
//   }
// }

// var shape = presentation.getSelectedShape()
// console.log( shape.textReplace('a', 'X'))

// shape.textReplace('a', 'X').textReplace('X', 'a')

// console.log(slides[2].shapes()[0].table().length)


// presentation.quit()

// Get all shapes of the first slide 1 
// var shapes = slides[0].shapes();

// var p = shapes[0].p();
// var form = p.format()
// console.log(form.attr('Bullet',true)); 

//   console.log('Title shape count:', shapes.length);
//
//   //get name and text of the first shape
//   var shape =  shapes[0];
//    console.log('Title shape count:', shape.text());
//    shape.tag.set('FU', 'bar')
//   console.log('tag:', shape.tag.get('fu'));
//   shape.tag.remove('FU')
//  console.log('allTags:', shape.tags);
//   
  
  
  //change name of the first
  // shape.name('First Shape');
  // shape.text('FuBar');
  //  console.log(shape.name(), shape.text());
  //Setter retun the destination object so you can chain them
  //  shape.top(10).left(10).height(100).width(200);
    
  
  // //Save presentation as PDF (sync)
  // presentation.saveAs({name:path.join(__dirname,'Presentation01.pdf'), type:'pdf'});
  // //SaveAs new presentation and quit application 
   // presentation.saveAs(path.join(__dirname,'Presentation01_New.pptx'));
  // presentation.quit(); //Close presentation & quit application  

  
// console.log('shapes', shapes.length);
// console.log('shapes', shapes[0].attr('Name',true));
// console.log('shapes', shapes[0].attr('Left',true));
// console.log('shapes', shapes[0].attr({name:'Left', value:10},true));
// console.log('shapes', shapes[0].attr('Left',true));

/*
//get the Slide 
var slides = presentation.slides();
console.log('slides', slides.length);

console.log( presentation.attr('Name', true))

var shapes = presentation.shapes({'attr:Name':'Title 1'});

console.log( shapes[0].attr('Text', true))
// console.log('shapes', shapes[0].attr('Rotation', true));

// shapes[0].attr({name:'Rotation',value:200}, true)
attr = presentation.tags(null, true)
console.log(attr);
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
  
  