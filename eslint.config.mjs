import js from '@eslint/js';
import googleappsscript from 'eslint-plugin-googleappsscript';

export default [
  { ignores: ['node_modules/'] },
  js.configs.recommended,
  {
    files: ['src/**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      // All Apps Script files share one global scope; there are no imports or exports.
      sourceType: 'script',
      globals: googleappsscript.environments.googleappsscript.globals,
    },
    rules: {
      // Cross-file globals are resolved by the `checkJs` type check (npm run typecheck).
      'no-undef': 'off',
      // Top-level functions are the public API, called from outside; keep the rule for locals and params.
      'no-unused-vars': ['error', { vars: 'local', args: 'after-used' }],
      'no-var': 'error',
      'prefer-const': 'error',
    },
  },
];
