module.exports = {
  moduleFileExtensions: [
      'ts',
      'tsx',
      'js',
      'jsx',
      'json'
  ],
  transform: {
      '^.+\\.tsx?$': 'ts-jest'
  },
  moduleNameMapper: {
      '^@/(.*)$': '<rootDir>/src/$1',
  },
  testMatch: [
      '<rootDir>/(tests/unit/**/*.test.(js|jsx|ts|tsx)|src/**/*.test.(js|jsx|ts|tsx))'
  ]
}