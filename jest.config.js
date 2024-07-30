// jest.config.js
export default {
    testEnvironment: 'node',
    moduleFileExtensions: ['js', 'json', 'jsx', 'mjs'],
    transform: {
      '^.+\\.jsx?$': 'babel-jest', // Use babel-jest to transpile JavaScript/JSX files
    },
    testPathIgnorePatterns: ['/node_modules/', '/dist/'],
  };
  