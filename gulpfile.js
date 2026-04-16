'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');

const postcss = require('gulp-postcss');
const atimport = require('postcss-import');
const tailwindcss = require('tailwindcss');
const autoprefixer = require('autoprefixer');

function buildTailwind() {
  return gulp
    .src('assets/main.css')
    .pipe(
      postcss([
        atimport(),
        tailwindcss(),
        autoprefixer()
      ])
    )
    .pipe(gulp.dest('assets/dist'));
}

// 🔹 Run once before build/bundle
const tailwindTask = build.subTask('tailwindcss', function (gulpInst, buildOptions, done) {
  buildTailwind();
  done();
});

// 🔹 WATCH for gulp serve
gulp.task('tailwind:watch', function () {
  gulp.watch(
    ['assets/main.css', 'src/**/*.{ts,tsx}'],
    buildTailwind
  );
});

build.rig.addPreBuildTask(tailwindTask);
build.initialize(gulp);
