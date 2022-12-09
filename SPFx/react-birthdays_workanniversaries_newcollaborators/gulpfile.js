'use strict';

const build = require('@microsoft/sp-build-web');
build.tslintCmd.enabled = false;

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

const eslint = require('gulp-eslint7');

const eslintSubTask = build.subTask('eslint', function (gulp, buildOptions, done) {
  return gulp.src(['src/**/*.{ts,tsx}'])
    // eslint() attaches the lint output to the "eslint" property
    // of the file object so it can be used by other modules.
    .pipe(eslint())
    // eslint.format() outputs the lint results to the console.
    // Alternatively use eslint.formatEach() (see Docs).
    .pipe(eslint.format())
    // To have the process exit with an error code (1) on
    // lint error, return the stream and pipe to failAfterError last.
    .pipe(eslint.failAfterError());
});

build.rig.addPreBuildTask(build.task('eslint-task', eslintSubTask));

// // Uncomment to analyze bundle
// const bundleAnalyzer = require('webpack-bundle-analyzer');
// const path = require('path');

// build.configureWebpack.mergeConfig({
//   additionalConfiguration: (generatedConfiguration) => {
//     const lastDirName = path.basename(__dirname);
//     const dropPath = path.join(__dirname, 'temp', 'stats');
//     generatedConfiguration.plugins.push(new bundleAnalyzer.BundleAnalyzerPlugin({
//       openAnalyzer: false,
//       analyzerMode: 'static',
//       reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
//       generateStatsFile: true,
//       statsFilename: path.join(dropPath, `${lastDirName}.stats.json`),
//       logLevel: 'error'
//     }));

//     return generatedConfiguration;
//   }
// });

const path = require('path');

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if (!generatedConfiguration.resolve.alias) {
      generatedConfiguration.resolve.alias = {};
    }

    //root src folder
    generatedConfiguration.resolve.alias['@app'] = path.resolve(__dirname, 'lib')

    return generatedConfiguration;
  }
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));

