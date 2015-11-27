var gulp = require('gulp');

gulp.task('build', function() {
    //Copy .NET functionsto /dist
    return gulp.src('./src/OfficeScript/OfficeScript/bin/Debug/*.dll')
    .pipe(gulp.dest('./dist'));´
});