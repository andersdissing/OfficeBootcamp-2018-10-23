module.exports = wallaby => {

  return {
      files: ['src/**/*', 'jest.config.js', 'package.json', 'tsconfig.json', 'tests/**/*' , '!src/**/*.test.ts', '!tests/**/*.test.ts',
          {pattern: 'node_modules/enzyme-adapter-react-15/**/*', instrument: false},
      ],

      tests: ['tests/**/*.test.ts','src/**/*.test.ts'],

      env: {
          type: 'node',
          runner: 'node',
      },
      
      preprocessors: {
      },

      compilers: {
          '**/*.ts?(x)': wallaby.compilers.typeScript({
              module: 'commonjs',
              jsx: 'React'
          })
      },

      setup(wallaby) {
          const jestConfig = require('./package').jest || require('./jest.config')
          delete jestConfig.transform['^.+\\.tsx?$']
          Object.keys(jestConfig.moduleNameMapper).forEach(k => (jestConfig.moduleNameMapper[k] = jestConfig.moduleNameMapper[k].replace('<rootDir>', wallaby.localProjectDir)))
          wallaby.testFramework.configure(jestConfig)
      },

      testFramework: 'jest',

      debug: true
  }
}