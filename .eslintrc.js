module.exports = {
  env: {
    browser: true,
    es2021: true,
    node: true,
    jquery: true,
  },
  extends: ['eslint:recommended', 'prettier'],
  parserOptions: {
    ecmaVersion: 'latest',
    sourceType: 'module',
  },
  globals: {
    '$': 'readonly',
    'jQuery': 'readonly',
    'luckysheet': 'writable'
  },
  rules: {
    // 核心错误（保持 error）
    'no-debugger': 'error',
    'no-unused-vars': 'off',
    'no-console': 'warn',

    // 代码质量（降级为 warn，避免阻断开发）
    'eqeqeq': 'warn',  // 从 error 改为 warn
    'curly': 'warn',   // 从 error 改为 warn
    'no-prototype-builtins': 'warn',

    // 代码风格（降级为 warn 或 off）
    'prefer-const': 'warn',
    'no-var': 'warn',  // 从 error 改为 warn，老项目 var 太多
    'one-var': 'off',  // 关闭，这个规则太严格
    'prefer-template': 'warn',
    'no-multiple-empty-lines': ['warn', { max: 2 }],  // 放宽到 2 行
    'eol-last': 'warn',
    'no-trailing-spaces': 'warn',
    quotes: ['warn', 'single'],
    semi: ['warn', 'always'],
    indent: ['warn', 2],
    'no-redeclare': 'warn',  // 添加这个规则
  },
  ignorePatterns: [
    'node_modules/',
    'dist/',
    'docs/',
    '*.min.js',
  ],
};
