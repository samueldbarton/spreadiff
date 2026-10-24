import ts from 'typescript';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../packages/spreadiff/src/', import.meta.url));
const files = ['index.ts', 'html/index.ts', 'react/index.tsx', 'worker/index.ts'].map((f) => root + f);
const program = ts.createProgram(files, { jsx: ts.JsxEmit.ReactJSX, moduleResolution: ts.ModuleResolutionKind.Bundler, module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022, skipLibCheck: true });
const c = program.getTypeChecker();
const seen = new Set();
for (const f of files) {
  const sym = c.getSymbolAtLocation(program.getSourceFile(f));
  for (const e of c.getExportsOfModule(sym)) {
    const t = e.flags & ts.SymbolFlags.Alias ? c.getAliasedSymbol(e) : e;
    if (seen.has(t)) continue;
    seen.add(t);
    const d = t.declarations?.[0];
    if (d && ts.isFunctionDeclaration(d)) {
      const sig = c.getSignatureFromDeclaration(d);
      console.log(`${e.name}${c.signatureToString(sig, undefined, ts.TypeFormatFlags.NoTruncation)}`);
    }
  }
}
