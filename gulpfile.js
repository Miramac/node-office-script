var gulp = require('gulp')
var exec = require('child_process').exec
var path = require('path')
var del = require('del')

var dest = './dist'
var src = './src/OfficeScript/OfficeScript/bin/Debug/*.dll'

gulp.task('build', ['deploy'], function () {})

gulp.task('compile', function (cb) {
  exec(`MSBuild ${path.normalize('src/OfficeScript/OfficeScript.sln')} /clp:ErrorsOnly`, function (err, stdout, stderr) {
    console.log(stdout)
    console.log(stderr)
    cb(err)
  })
})

gulp.task('deploy', ['clean', 'compile'], function () {
  // Copy .NET functionsto /dist
  return gulp.src(src)
    .pipe(gulp.dest(dest))
})

gulp.task('clean', function () {
  // clean /dist
  return del(dest)
})
