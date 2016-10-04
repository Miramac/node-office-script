/* global describe,it,after,__dirname */
var assert = require('assert')
var path = require('path')
var powerpoint = require('../').powerpoint
var testPPT01 = 'Testpptx_01.pptx'
var testDataPath = path.join(__dirname, 'data')

describe('report', function () {
  this.timeout(15000)
  after(function (done) { powerpoint.quit(null, done) })
  describe('presentation', function () {
    describe('#open&close', function () {
      it('should open and close the file', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.close(null, done)
        })
      })
    })
    describe('#fetch', function () {
      it('should fetch an open presentation', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err) {
          if (err) throw err
          powerpoint.fetch(null, function (err, presentation) {
            if (err) throw err
            presentation.close(null, done)
          })
        })
      })
    })
    describe('#attr', function () {
      it('should get a name and path attribute from presentation', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          // get Path Sync
          assert.equal(presentation.attr({name: 'Path'}, true), testDataPath)
          assert.equal(presentation.attr('Path', true), testDataPath)
          // get name async
          presentation.attr({name: 'Name'}, function (err, data) {
            if (err) throw err
            assert.equal(data, testPPT01)
            presentation.close(null, done)
          })
        })
      })
    })
    describe('#slides', function () {
      it('should have 2 slides', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            assert.equal(slides.length, 2)
            presentation.close(null, done)
          })
        })
      })

      it('should have the Attr Name', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            assert.equal(slides[1].attr({name: 'Name'}, true), 'Slide2')
            presentation.close(null, done)
          })
        })
      })

      it('should have the Attr Pos', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            slides.forEach(function (slide, index) {
              assert.equal(slide.attr({name: 'Pos'}, true), index + 1)
            })
            presentation.close(null, done)
          })
        })
      })
      it('should be changeable the pos of Slide2', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            assert.equal(slides[1].attr({name: 'Pos'}, true), 2)
            assert.equal(slides[1].attr({name: 'Pos', value: 1}, true).attr({name: 'Pos'}, true), 1)
            presentation.close(null, done)
          })
        })
      })
      it('should be able to delete Slide2', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, true)[1].remove(null, function (err) {
            presentation.slides(null, function (err, slides) {
              if (err) throw err
              assert.equal(slides.length, 1)
              presentation.close(null, done)
            })
          })
        })
      })
      it('should be able to duplicate Slide1', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, true)[0].duplicate(null, function (err, slide) {
            if (err) throw err
            assert.equal(slide.attr({name: 'Pos'}, true), 2)
            presentation.slides(null, function (err, slides) {
              if (err) throw err
              assert.equal(slides.length, 3)
              presentation.close(null, done)
            })
          })
        })
      })
      it('should be able to create a shape on slide1', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          var slide = presentation.slides(null, true)[1]
          var shapeCount = slide.shapes(null, true).length
          slide.addTextbox(null, function (err, shape) {
            if (err) throw err
            assert.equal(slide.shapes(null, true).length, shapeCount + 1)
            presentation.close(null, done)
          })
        })
      })
      it('should be able to create a shape with top=100,left=100,height=200,width=200', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          var slide = presentation.slides(null, true)[1]
          var shapeCount = slide.shapes(null, true).length
          slide.addTextbox({top: 100, left: 100, height: 200, width: 200}, function (err, shape) {
            if (err) throw err
            assert.equal(slide.shapes(null, true).length, shapeCount + 1)
            assert.equal(shape.attr({ name: 'Top' }, true), 100)
            assert.equal(shape.attr({ name: 'Left' }, true), 100)
            assert.equal(shape.attr({ name: 'Height' }, true), 200)
            assert.equal(shape.attr({ name: 'Width' }, true), 200)
            presentation.close(null, done)
          })
        })
      })
    })
    describe('#shapes', function () {
      it('should have 2 shapes on slide one', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            slides[0].shapes(null, function (err, shapes) {
              assert.equal(shapes.length, 2)
              presentation.close(null, done)
            })
          })
        })
      })
      it('should have the Attr Name', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            slides[0].shapes(null, function (err, shapes) {
              assert.equal(shapes[0].attr({name: 'Name'} , true), 'Title 1')
              presentation.close(null, done)
            })
          })
        })
      })
      it('should be changeable the Attribute Name', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            slides[0].shapes(null, function (err, shapes) {
              assert.equal(shapes[0].attr({name: 'Name', value: 'Test'} , true).attr({name: 'Name'}, true), 'Test')
              presentation.close(null, done)
            })
          })
        })
      })
      it('should be able to duplicate shape1', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            slides[0].shapes(null, function (err, shapes) {
              var shapeCount = shapes.length
              shapes[0].duplicate(null, function (err, shape) {
                if (err) throw err
                assert.equal(slides[0].shapes(null, true).length, shapeCount + 1)
                assert.equal(shape.attr({name: 'Text'}, true), shapes[0].attr({name: 'Text'}, true))
                presentation.close(null, done)
              })
            })
          })
        })
      })
      it('should be able to remove shape1', function (done) {
        powerpoint.open(path.join(testDataPath, testPPT01), function (err, presentation) {
          if (err) throw err
          presentation.slides(null, function (err, slides) {
            if (err) throw err
            slides[0].shapes(null, function (err, shapes) {
              shapes[0].remove(null, function (err) {
                if (err) throw err
                presentation.close(null, done)
              })
            })
          })
        })
      })
    })
  })
})
