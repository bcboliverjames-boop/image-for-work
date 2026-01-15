const fs = require('fs');
const path = require('path');

const filePath = path.join(__dirname, 'src/taskpane-new.ts');
let content = fs.readFileSync(filePath, 'utf8');

// 更新版本号
content = content.replace(/v10/g, 'v11');

fs.writeFileSync(filePath, content, 'utf8');
console.log('版本号已更新到 v11');
