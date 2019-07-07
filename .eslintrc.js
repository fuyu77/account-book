module.exports = {
    "env": {
        "browser": true,
        "es6": true
    },
    "plugins": ["@typescript-eslint"],
    "parser": "@typescript-eslint/parser",
    "extends": "eslint:recommended",
    "globals": {
        "Atomics": "readonly",
        "SharedArrayBuffer": "readonly"
    },
    "parserOptions": {
        "ecmaVersion": 2018,
        "sourceType": "module"
    },
    "extends": "plugin:@typescript-eslint/recommended",
    "rules": {
      "@typescript-eslint/indent": ["error", 2],
      "@typescript-eslint/no-unused-vars": ["error", { "varsIgnorePattern": "onEdit", "argsIgnorePattern": "^_" }],
      "@typescript-eslint/no-explicit-any": "off"
    }
};