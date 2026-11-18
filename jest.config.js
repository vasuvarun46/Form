module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  roots: ["<rootDir>/src"],
  testMatch: ["**/*.test.(ts|tsx)"],
  moduleNameMapper: {
    "\\.(css|scss)$": "identity-obj-proxy",
    "\\.(png|jpg|jpeg|gif|svg)$": "<rootDir>/__mocks__/fileMock.js",
    "EformWebPartStrings": "<rootDir>/src/lib/strings.ts",
    "^@microsoft/sp-core-library$": "<rootDir>/src/lib/spShim.ts",
    "^@microsoft/sp-http$": "<rootDir>/src/lib/spShim.ts",
    "^@microsoft/sp-webpart-base$": "<rootDir>/src/lib/spShim.ts",
    "^@microsoft/sp-component-base$": "<rootDir>/src/lib/spShim.ts",
    "^@microsoft/sp-property-pane$": "<rootDir>/src/lib/spShim.ts"
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
