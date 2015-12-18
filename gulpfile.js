var gulp = require('gulp');
var exec = require('child_process').exec;
var path = require('path');

gulp.task('build', ['compile', 'deploy'], function() {
    
});

gulp.task('compile', function(cb) {
    exec(`MSBuild ${path.normalize('src/OfficeScript/OfficeScript.sln')} /clp:ErrorsOnly /clp:ErrorsOnly`, function (err, stdout, stderr) {
        console.log(stdout);
        console.log(stderr);
        cb(err);
      });
});


gulp.task('deploy', function() {
    
    //Copy .NET functionsto /dist
    return gulp.src('./src/OfficeScript/OfficeScript/bin/Debug/*.dll')
    .pipe(gulp.dest('./dist'));
});
