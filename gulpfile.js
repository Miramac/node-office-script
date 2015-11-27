var gulp = require('gulp');

gulp.task('build', function() {

  // Any deployment logic should go here
   return gulp.src('./src/OfficeScript/OfficeScript/bin/Debug/*.dll')
    .pipe(gulp.dest('./dist'));
 
});