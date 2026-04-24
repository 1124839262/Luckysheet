# 第一步：只检查并修复 utils 目录（最简单）
npx eslint "src/utils/**/*.js" --fix

# 第二步：修复 methods 目录
npx eslint "src/methods/**/*.js" --fix

# 第三步：修复 locale 目录
npx eslint "src/locale/**/*.js" --fix

# 第四步：修复 data 和 demoData
npx eslint "src/data/**/*.js" "src/demoData/**/*.js" --fix

# 最后再处理复杂的 controllers 和 global
