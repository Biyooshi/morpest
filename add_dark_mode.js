const fs = require('fs');
const path = require('path');

const files = ['index.html', 'dashboard.html', 'list.html'];

files.forEach(file => {
  let content = fs.readFileSync(file, 'utf8');

  // Replace bg-white with bg-white dark:bg-gray-800
  // Note: Only targeting main containers and inputs. 
  // Let's be more specific.

  // Inputs, selects, textareas that have class="..."
  // Adding dark:bg-gray-700 dark:border-gray-600 dark:text-white
  content = content.replace(/(<(?:input|select|textarea)[^>]*class="[^"]*)(?<!dark:bg-gray-700)([^"]*")/gi, (match, p1, p2) => {
    // Avoid double adding
    if (match.includes('dark:bg-gray-700')) return match;
    return p1 + ' dark:bg-gray-700 dark:border-gray-600 dark:text-white dark:placeholder-gray-400' + p2;
  });

  // text-gray-500 -> text-gray-500 dark:text-gray-400
  content = content.replace(/text-gray-500(?!\s*dark:text-gray-400)/g, 'text-gray-500 dark:text-gray-400');
  
  // text-gray-400 -> text-gray-400 dark:text-gray-300
  // content = content.replace(/text-gray-400(?!\s*dark:text-gray-300)/g, 'text-gray-400 dark:text-gray-300');

  // text-blue-800 -> text-blue-800 dark:text-blue-400
  content = content.replace(/text-blue-800(?!\s*dark:text-blue-400)/g, 'text-blue-800 dark:text-blue-400');

  // bg-blue-50 -> bg-blue-50 dark:bg-blue-900/30
  content = content.replace(/bg-blue-50(?!\s*dark:bg-blue-900\/30)/g, 'bg-blue-50 dark:bg-blue-900/30');

  // bg-orange-50 -> bg-orange-50 dark:bg-orange-900/30
  content = content.replace(/bg-orange-50(?!\s*dark:bg-orange-900\/30)/g, 'bg-orange-50 dark:bg-orange-900/30');

  // bg-gray-50 -> bg-gray-50 dark:bg-gray-700
  content = content.replace(/bg-gray-50(?!\s*dark:bg-gray-700)/g, 'bg-gray-50 dark:bg-gray-700');

  // text-orange-800 -> text-orange-800 dark:text-orange-400
  content = content.replace(/text-orange-800(?!\s*dark:text-orange-400)/g, 'text-orange-800 dark:text-orange-400');

  // text-gray-800 -> text-gray-800 dark:text-gray-200
  content = content.replace(/text-gray-800(?!\s*dark:text-gray-200)/g, 'text-gray-800 dark:text-gray-200');

  // text-gray-700 -> text-gray-700 dark:text-gray-300
  content = content.replace(/text-gray-700(?!\s*dark:text-gray-300)/g, 'text-gray-700 dark:text-gray-300');

  fs.writeFileSync(file, content);
  console.log('Processed', file);
});
