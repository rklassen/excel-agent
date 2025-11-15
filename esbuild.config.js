require('esbuild').build({
  entryPoints: ['src/agent.js'],
  bundle: true,
  minify: true,
  format: 'iife',
  outfile: 'public/bundle.js',
  target: 'es2020'
}).catch(() => process.exit(1));