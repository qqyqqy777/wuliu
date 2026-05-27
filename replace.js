import fs from 'fs';

let text = fs.readFileSync('src/App.tsx', 'utf8');

// The main goal is to optimize the look and typography.
// 1. Replace 'slate' with 'zinc' for a crisper layout.
text = text.replace(/slate-/g, 'zinc-');
// 2. Refine text colors
text = text.replace(/text-zinc-800/g, 'text-zinc-900');
// 3. Make main wrappers look more elegant
text = text.replace(/bg-zinc-50/g, 'bg-zinc-50/50');
text = text.replace(/rounded-2xl/g, 'rounded-xl');
// 4. Update the input fields and text areas to have a cleaner focus state
text = text.replace(/focus:ring-2 focus:ring-indigo-500/g, 'focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500');

fs.writeFileSync('src/App.tsx', text);
console.log('Done refactoring classes in App.tsx');
