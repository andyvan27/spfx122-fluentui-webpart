module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  roots: ["<rootDir>/tests"],
  moduleNameMapper: {
    "^@pnp/sp$": "<rootDir>/tests/mocks/pnp-sp.mock.ts",
    "^@pnp/sp/(.*)$": "<rootDir>/tests/mocks/pnp-sp.mock.ts",
    "^@pnp/graph$": "<rootDir>/tests/mocks/pnp-graph.mock.ts",
    "^@pnp/graph/(.*)$": "<rootDir>/tests/mocks/pnp-graph.mock.ts"
  }
};
