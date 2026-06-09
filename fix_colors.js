const fs = require('fs');
const files = ['index.html', 'dashboard.html', 'list.html'];

files.forEach(file => {
  if (fs.existsSync(file)) {
    let content = fs.readFileSync(file, 'utf8');

    // Revert the accidentally generated invalid classes back to original
    content = content.replace(/bg-orange-50 dark:bg-orange-900\/300/g, 'bg-orange-500');
    content = content.replace(/bg-blue-50 dark:bg-blue-900\/300/g, 'bg-blue-500');

    // Also fix the comma separated ones from earlier classList fix
    content = content.replace(/'bg-orange-50', 'dark:bg-orange-900\/300'/g, "'bg-orange-500'");
    
    fs.writeFileSync(file, content, 'utf8');
  }
});
console.log('Fixed colors');
