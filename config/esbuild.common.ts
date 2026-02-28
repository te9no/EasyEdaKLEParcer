import type esbuild from 'esbuild';

// Mirrors EasyEDA Pro's official SDK build style:
// iife bundle + global export name `edaEsbuildExportName`.
export default {
	entryPoints: {
		index: './src/index',
	},
	entryNames: '[name]',
	assetNames: '[name]',
	bundle: true,
	minify: false,
	loader: {},
	outdir: './dist/',
	sourcemap: undefined,
	platform: 'browser',
	format: 'iife',
	globalName: 'edaEsbuildExportName',
	// Downlevel to ES2015 to avoid parse errors (optional chaining, ??, async/await) in some EasyEDA builds.
	target: 'es2015',
	treeShaking: true,
	ignoreAnnotations: true,
	define: {},
	external: [],
} satisfies Parameters<(typeof esbuild)['build']>[0];
