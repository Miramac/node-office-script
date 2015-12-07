/* global describe,it,after,__dirname */
var assert = require('assert');
var path = require('path');
var Presentation = require('../').Presentation;
var powerpoint = require('../').powerpoint;
var testPPT01 = 'Testpptx_01.pptx';
var testDataPath = path.join(__dirname, 'data');


describe('report', function() {
    this.timeout(15000);
    after( function(done) {powerpoint.quit(null, done);} );
    describe('presentation', function() {
        describe('#open&close', function() {
            it('should open and close the file', function(done) {
                var presentation = new Presentation( path.join(testDataPath,testPPT01));
                //Close file
                presentation.close(done);
            });
        });
        describe('#attr', function() {
            it('should get a name, path, and fullPath attribute from presentation', function(done) {
                 var presentation = new Presentation(path.join(testDataPath,testPPT01));
                //get Path Sync
                assert.equal(presentation.path(), testDataPath);
                assert.equal(presentation.name(), testPPT01);
                assert.equal(presentation.fullName(), path.join(testDataPath,testPPT01));
                //Close file
                presentation.close(done);
            });
        });
        describe('#slides', function() {
            it('should have 2 slides', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slides = presentation.slides(); 
                assert.equal(slides.length, 2);
                presentation.close(done);
            });
            
            it('should have the Attr Name', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slides = presentation.slides(); 
                assert.equal(slides[1].name(), 'Slide2');
                presentation.close(done);
            });
            
            it('should have the Attr Pos', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slides = presentation.slides(); 
                slides.forEach(function(slide, index) {
                    assert.equal(slide.pos(), index+1);
                });
                presentation.close(done);
            });
            it('should be changeable the pos of Slide2', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slides = presentation.slides(); 
                assert.equal(slides[1].pos(), 2);
                assert.equal(slides[1].pos(1).pos(), 1);
                presentation.close(done);
            });
            it('should be able to delete Slide2', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slides = presentation.slides(); 
                slides[1].remove();
                slides = presentation.slides(); 
                assert.equal(slides.length, 1);
                presentation.close(done);
            });
            it('should be able to duplicate Slide1', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                assert.equal(presentation.slides()[0].duplicate().pos(), 2);
                var slides = presentation.slides();
                assert.equal(slides.length, 3);
                presentation.close(done);
            });
            it('should be able to create a shape on slide1', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slide = presentation.slides()[1];
                var shapeCount = slide.shapes().length;
                slide.addTextbox();
                assert.equal(slide.shapes().length, shapeCount + 1);
                presentation.close(done);
            });
            it('should be able to create a shape with top=100,left=100,height=200,width=200', function (done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var slide = presentation.slides()[1];
                var shapeCount = slide.shapes().length;
                var shape = slide.addTextbox({top:100, left:100, height:200, width:200});
                assert.equal(slide.shapes().length, shapeCount + 1);
                assert.equal(shape.top(), 100);
                assert.equal(shape.left(), 100);
                assert.equal(shape.height(), 200);
                assert.equal(shape.width(), 200);
                presentation.close(done);
            });
        });
        describe('#shapes', function() {
            it('should have 2 shapes on slide one', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                assert.equal(presentation.slides()[0].shapes().length, 2);
                presentation.close(done);   
            });
            it('should have the Attr Name', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                assert.equal(presentation.slides()[0].shapes()[0].name(), 'Title 1');
                presentation.close(done);   
            });
            it('should be changeable the Attribute Name', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                assert.equal(presentation.slides()[0].shapes()[0].name('Test').name(), 'Test');
                presentation.close(done);   
            });
            it('should be able to duplicate shape1', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                var shapes = presentation.slides()[0].shapes();
                var shapeCount = shapes.length; 
                var shape = shapes[0].duplicate();
                assert.equal(presentation.slides()[0].shapes().length, shapeCount + 1);
                assert.equal(shape.text(), shapes[0].text());
                presentation.close(done); 
            });
            it('should be able to remove shape1', function(done) {
                var presentation = new Presentation(path.join(testDataPath,testPPT01));
                presentation.slides()[0].shapes()[0].remove();
                presentation.close(done); 
            });
        });
    });
});