module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  roots: ["<rootDir>/src"],
  testMatch: ["**/*.test.(ts|tsx)"],
  moduleNameMapper: {
    "\\.(css|scss)$": "identity-obj-proxy",
    "\\.(png|jpg|jpeg|gif|svg)$": "<rootDir>/__mocks__/fileMock.js",
    "EformWebPartStrings": "<rootDir>/src/__mocks__/strings.ts",
    "^@microsoft/sp-core-library$": "<rootDir>/__mocks__/@microsoft/sp-core-library.js",
    "^@microsoft/sp-http$": "<rootDir>/__mocks__/@microsoft/sp-http.js",
    "^@microsoft/sp-webpart-base$": "<rootDir>/__mocks__/@microsoft/sp-webpart-base.js"
  },
  setupFilesAfterEnv: ["<rootDir>/jest.setup.ts"],
  collectCoverageFrom: [
    "src/webparts/eform/components/**/*.{ts,tsx}",
    "!src/**/*.d.ts",
    "!src/**/index.ts"
  ],
  coverageThreshold: {
    global: {
      branches: 60,
      functions: 70,
      lines: 80,
      statements: 80
    }
  }
};
