import fs from 'fs';

let text = fs.readFileSync('src/App.tsx', 'utf8');

// Modernize primary buttons (indigo-900 / indigo-600)
// The original uses 'bg-slate-900 text-white', which was changed to 'zinc-900'. Let's refine it.
text = text.replace(/bg-zinc-900 hover:bg-zinc-800 disabled:opacity-50 disabled:cursor-not-allowed text-white/g, 
  'bg-zinc-900 hover:bg-zinc-800 disabled:opacity-50 disabled:cursor-not-allowed text-white shadow-sm ring-1 ring-zinc-900/10');

text = text.replace(/bg-indigo-600 hover:bg-indigo-700/g, 
  'bg-indigo-600 hover:bg-indigo-500 shadow-sm ring-1 ring-indigo-600/20');
  
// Refine secondary emerald buttons
text = text.replace(/bg-emerald-600 hover:bg-emerald-700/g, 
  'bg-emerald-600 hover:bg-emerald-500 shadow-sm ring-1 ring-emerald-600/20');

// Fix border radius on input elements
text = text.replace(/rounded-lg/g, 'rounded-xl'); // uniform rounding
text = text.replace(/rounded-xl/g, 'rounded-xl'); // just in case

// Use softer typography for the main labels
text = text.replace(/text-xs font-bold text-zinc-500 uppercase tracking-wider/g, 'text-[11px] font-semibold text-zinc-500 uppercase tracking-widest');
text = text.replace(/text-xs font-bold text-zinc-500 uppercase/g, 'text-[11px] font-semibold text-zinc-500 uppercase tracking-widest');

fs.writeFileSync('src/App.tsx', text);
console.log('Button and form styling refined.');
