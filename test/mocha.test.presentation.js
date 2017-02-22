/* global describe,it,after,__dirname */
var assert = require('assert')
var path = require('path')
var Presentation = require('../').Presentation
var powerpoint = require('../').powerpoint
var testPPT01 = 'Testpptx_mocha.pptx'
var testDataPath = path.join(__dirname, 'data')

describe('report', function () {
  this.timeout(15000)
  after(function (done) { powerpoint.quit(null, done) })
  describe('presentation', function () {
    describe('#open&close', function () {
      it('should open and close the file', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        // Close file
        presentation.close(done)
      })
    })
    describe('#attr', function () {
      it('should get a name, path, and fullPath attribute from presentation', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        // get Path Sync
        assert.equal(presentation.path(), testDataPath)
        assert.equal(presentation.name(), testPPT01)
        assert.equal(presentation.fullName(), path.join(testDataPath, testPPT01))
        // Close file
        presentation.close(done)
      })
      it('should have slideWidth and slide Height', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slideHeight(), 540)
        assert.equal(presentation.slideWidth(), 720)
        presentation.close(done)
      })
    })
    describe('#properties', function () {
      it('should have 13 builtin properties', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var props = presentation.builtinProp()
        var counter = 0
        var key
        for (key in props) {
          if (props.hasOwnProperty(key)) {
            counter++
          }
        }
        assert.equal(counter, 13)
        presentation.close(done)
      })
      it('should have and set builtin property "Last author"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        presentation.builtinProp('Last author', 'Oliver Queen')
        assert.equal(presentation.builtinProp('Last author'), 'Oliver Queen')
        presentation.close(done)
      })
      it('should have and set custom property', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        presentation.customProp('Hero', 'Green Arrow')
        assert.equal(presentation.customProp('Hero'), 'Green Arrow')
        presentation.close(done)
      })
      it('should set a tag', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        presentation.tag.set('Green', 'Arrow')
        assert.equal(1, 1)
        presentation.close(done)
      })
      it('should find and replace text', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        presentation.textReplace('Testpptx', 'Fu Bar')
        assert.equal(presentation.slides()[0].shapes()[0].text(), 'Fu Bar_01')
        presentation.close(done)
      })
      it('should batch find and replace text', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var replaces = {
          '<Txt_IV_005>': 1234,
          '<Txt_IV_532>': 'Fu',
          'Testpptx': '<Txt_IV_532> Bar'
        }
        presentation.textReplace(replaces)
        assert.equal(presentation.slides()[0].shapes()[0].text(), 'Fu Bar_01')
        presentation.close(done)
      })
    })
    describe('#slides', function () {
      it('should have 4 slides', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slides = presentation.slides()
        assert.equal(slides.length, 4)
        presentation.close(done)
      })
      it('should have the attribute "name"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slides = presentation.slides()
        assert.equal(slides[1].name(), 'Slide2')
        presentation.close(done)
      })
      it('should have the attribute "pos"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slides = presentation.slides()
        slides.forEach(function (slide, index) {
          assert.equal(slide.pos(), index + 1)
        })
        presentation.close(done)
      })
      it('should be able to change the position of slide2', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slides = presentation.slides()
        assert.equal(slides[1].pos(), 2)
        assert.equal(slides[1].pos(1).pos(), 1)
        presentation.close(done)
      })
      it('should be able to delete slide2', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slides = presentation.slides()
        slides[1].remove()
        slides = presentation.slides()
        assert.equal(slides.length, 3)
        presentation.close(done)
      })
      it('should be able to duplicate slide1', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].duplicate().pos(), 2)
        var slides = presentation.slides()
        assert.equal(slides.length, 5)
        presentation.close(done)
      })
      it('should be able to create a shape on slide1', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slide = presentation.slides()[1]
        var shapeCount = slide.shapes().length
        slide.addTextbox()
        assert.equal(slide.shapes().length, shapeCount + 1)
        presentation.close(done)
      })
      it('should be able to create a shape with top=100,left=100,height=200,width=200', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slide = presentation.slides()[1]
        var shapeCount = slide.shapes().length
        var shape = slide.addTextbox({top: 100, left: 100, height: 200, width: 200})
        assert.equal(slide.shapes().length, shapeCount + 1)
        assert.equal(shape.top(), 100)
        assert.equal(shape.left(), 100)
        assert.equal(shape.height(), 200)
        assert.equal(shape.width(), 200)
        presentation.close(done)
      })
      it('should be able to insert a picture', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slide = presentation.slides()[2]
        var shapeCount = slide.shapes().length
        slide.addPicture(path.join(testDataPath, 'ga.png'))
        assert.equal(slide.shapes().length, shapeCount + 1)
        presentation.close(done)
      })
      it('should be able to insert a picture with top=100,left=100', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slide = presentation.slides()[2]
        var shapeCount = slide.shapes().length
        var shape = slide.addPicture(path.join(testDataPath, 'ga.png'), {top: 100, left: 100})
        assert.equal(slide.shapes().length, shapeCount + 1)
        assert.equal(shape.top(), 100)
        assert.equal(shape.left(), 100)
        presentation.close(done)
      })
      it('should be able to set and have a tag on slide1', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slide = presentation.slides()[1]
        slide.tag.set('Hero', 'Green Arrow')
        slide = presentation.slides()[1]
        assert.equal(slide.tag.get('Hero'), 'Green Arrow')
        presentation.close(done)
      })
      it('should be able to set, have and remove a tag on slide1', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slide = presentation.slides()[1]
        slide.tag.set('Oliver Queen', 'Green Arrow')
        slide.tag.set('Batman', 'Bruce Wayne')
        slide.tag.set('Flash', 'Barry Allen')
        slide = presentation.slides()[1]
        var tags = slide.tags
        var counter = 0
        var key
        for (key in tags) {
          if (tags.hasOwnProperty(key)) {
            counter++
          }
        }
        assert.equal(counter, 3)
        slide.tag.remove('Flash')
        slide = presentation.slides()[1]
        tags = slide.tags
        counter = 0
        for (key in tags) {
          if (tags.hasOwnProperty(key)) {
            counter++
          }
        }
        assert.equal(counter, 2)
        presentation.close(done)
      })
      it('should be able to copy/paste a slide', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var slideCount = presentation.slides().length
        presentation.slides()[1].copy()
        presentation.pasteSlide(-1)
        assert.equal(presentation.slides().length, slideCount + 1)
        presentation.close(done)
      })
    })
    describe('#shapes', function () {
      it('should have 2 shapes on slide one', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes().length, 2)
        presentation.close(done)
      })
      it('should have the attribute "name"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].name(), 'Title 1')
        presentation.close(done)
      })
      it('should be able to change the attribute "name"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].name('Test').name(), 'Test')
        presentation.close(done)
      })
      it('should have the attribute "top"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].top(), 195.33204650878906)
        presentation.close(done)
      })
      it('should be able to change the attribute "top"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].top(10.5).top(), 10.5)
        presentation.close(done)
      })
      it('should have the attribute "left"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].left(), 54)
        presentation.close(done)
      })
      it('should be able to change the attribute "left"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].left(10.5).left(), 10.5)
        presentation.close(done)
      })
      it('should have the attribute "height"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].height(), 60.585906982421875)
        presentation.close(done)
      })
      it('should be able to change the attribute "height"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].height(10.5).height(), 10.5)
        presentation.close(done)
      })
      it('should have the attribute "width"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].width(), 612)
        presentation.close(done)
      })
      it('should be able to change the attribute "width"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].width(10.5).width(), 10.5)
        presentation.close(done)
      })
      it('should have the attribute "rotation"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].rotation(), 0)
        presentation.close(done)
      })
      it('should be able to change the attribute "rotation"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].rotation(10.5).rotation(), 10.5)
        presentation.close(done)
      })
      it('should be able to change the attribute "altText"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].altText('Fu Bar').altText(), 'Fu Bar')
        presentation.close(done)
      })
      it('should be able to change the attribute "title"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[0].shapes()[0].title('Fu Bar').title(), 'Fu Bar')
        presentation.close(done)
      })

      it('should be able to duplicate shape1', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shapes = presentation.slides()[0].shapes()
        var shapeCount = shapes.length
        var shape = shapes[0].duplicate()
        assert.equal(presentation.slides()[0].shapes().length, shapeCount + 1)
        assert.equal(shape.text(), shapes[0].text())
        presentation.close(done)
      })
      it('should be able to remove shape1', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        presentation.slides()[0].shapes()[0].remove()
        presentation.close(done)
      })
      it('should be able to add a new line with text', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var counter = presentation.slides()[2].shapes()[1].paragraph().count()
        presentation.slides()[2].shapes()[1].addLine('Text2')
        assert.equal(presentation.slides()[2].shapes()[1].paragraph(counter + 1).text(), 'Text2')
        presentation.close(done)
      })
      it('should be able to add a new line on 5 and remove it', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        presentation.slides()[2].shapes()[1].addLine('TextTest', 5)
        assert.equal(presentation.slides()[2].shapes()[1].removeLine().paragraph().count(), 4)
        presentation.close(done)
      })
      it('should be able to add a new picture', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var counter = presentation.slides()[2].shapes().length
        presentation.slides()[2].addPicture(path.join(testDataPath, 'ga.png'))
        assert.equal(presentation.slides()[2].shapes().length, counter + 1)
        presentation.close(done)
      })
      it('should be able to set and have a tag', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[2].shapes()[1]
        shape.tag.set('Hero', 'Green Arrow')
        shape = presentation.slides()[2].shapes()[1]
        assert.equal(shape.tag.get('Hero'), 'Green Arrow')
        presentation.close(done)
      })
      it('should be able to set, have and remove a tag', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[2].shapes()[1]
        shape.tag.set('Oliver Queen', 'Green Arrow')
        shape.tag.set('Batman', 'Bruce Wayne')
        shape.tag.set('Flash', 'Barry Allen')
        shape = presentation.slides()[2].shapes()[1]
        var tags = shape.tags
        var counter = 0
        var key
        for (key in tags) {
          if (tags.hasOwnProperty(key)) {
            counter++
          }
        }
        assert.equal(counter, 14)
        shape.tag.remove('Flash')
        shape = presentation.slides()[2].shapes()[1]
        tags = shape.tags
        counter = 0
        for (key in tags) {
          if (tags.hasOwnProperty(key)) {
            counter++
          }
        }
        assert.equal(counter, 13)
        presentation.close(done)
      })
      it('should have and set zIndex', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[2].shapes()[1]
        assert.equal(shape.zIndex(), 2)
        shape.zIndex('back')
        assert.equal(shape.zIndex(), 1)
        shape.zIndex('forward')
        assert.equal(shape.zIndex(), 2)
        presentation.close(done)
      })
      it('should be able to replace a text', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        assert.equal(shape.textReplace('01', 'XX').text(), 'Testpptx_XX')
        presentation.close(done)
      })
    })
    describe('#paragraphs', function () {
      it('should have the attribute "text"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.text(), 'Testpptx_01')
        presentation.close(done)
      })
      it('should be able to change the attribute "text"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.text('Testi_1').text(), 'Testi_1')
        presentation.close(done)
      })
      it('should have the attribute "count"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        assert.equal(presentation.slides()[2].shapes()[1].addLine('yay').addLine('yay').paragraph().count(), 3)
        presentation.close(done)
      })
      it('should have the attribute "fontName"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontName(), 'Calibri')
        presentation.close(done)
      })
      it('should be able to change the attribute "fontName"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontName('Arial').fontName(), 'Arial')
        presentation.close(done)
      })
      it('should have the attribute "fontColor"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontColor(), '#000000')
        presentation.close(done)
      })
      it('should be able to change the attribute "fontColor"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontColor('#FF0000').fontColor(), '#ff0000')
        presentation.close(done)
      })
      it('should have the attribute "fontItalic"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontItalic(), false)
        presentation.close(done)
      })
      it('should be able to change the attribute "fontItalic"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontItalic(true).fontItalic(), true)
        presentation.close(done)
      })
      it('should have the attribute "fontBold"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontBold(), false)
        presentation.close(done)
      })
      it('should be able to change the attribute "fontBold"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.fontBold(true).fontBold(), true)
        presentation.close(done)
      })
      it('should have the attribute "align"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.align(), 'center')
        presentation.close(done)
      })
      it('should be able to change the attribute "align"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[0].shapes()[0].paragraph(1)
        assert.equal(para.align('left').align(), 'left')
        presentation.close(done)
      })
      it('should have the attribute "indent"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[2].shapes()[1].paragraph()
        assert.equal(para.indent(), 1)
        presentation.close(done)
      })
      it('should be able to change the attribute "indent"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var para = presentation.slides()[2].shapes()[1].paragraph()
        assert.equal(para.indent(2).indent(), 2)
        presentation.close(done)
      })
    })
    describe('#characters', function () {
      it('should have the attribute "text"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.text(), 'est')
        presentation.close(done)
      })
      it('should be able to change the attribute "text"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        char.text('XXX')
        assert.equal(shape.text(), 'TXXXpptx_01')
        presentation.close(done)
      })
      it('should have the attribute "count"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.count(), 3)
        presentation.close(done)
      })
      it('should have the attribute "fontName"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontName(), 'Calibri')
        presentation.close(done)
      })
      it('should be able to change the attribute "fontName"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontName('Arial').fontName(), 'Arial')
        presentation.close(done)
      })
      it('should have the attribute "fontColor"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontColor(), '#000000')
        presentation.close(done)
      })
      it('should be able to change the attribute "fontColor"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontColor('#FF0000').fontColor(), '#ff0000')
        presentation.close(done)
      })
      it('should have the attribute "fontItalic"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontItalic(), false)
        presentation.close(done)
      })
      it('should be able to change the attribute "fontItalic"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontItalic(true).fontItalic(), true)
        presentation.close(done)
      })
      it('should have the attribute "fontBold"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontBold(), false)
        presentation.close(done)
      })
      it('should be able to change the attribute "fontBold"', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[0].shapes()[0]
        var char = shape.character(2, 3)
        assert.equal(char.fontBold(true).fontBold(), true)
        presentation.close(done)
      })
    })
    describe('#tables', function () {
      it('should have a table with 2 rows and 5 columns', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[3].shapes()[1]
        assert.equal(shape.has('table'), true)
        var table = shape.table()
        assert.equal(table.length, 2)
        assert.equal(table[0].length, 5)
        presentation.close(done)
      })
      it('should have cell text and can change it', function (done) {
        var presentation = new Presentation(path.join(testDataPath, testPPT01))
        var shape = presentation.slides()[3].shapes()[1]
        var table = shape.table()
        assert.equal(table[0][0].text(), 'A')
        assert.equal(table[0][0].text('XX').text(), 'XX')
        presentation.close(done)
      })
    })
  })
})
