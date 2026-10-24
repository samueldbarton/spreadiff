// Lists every public export per entry point, with whether it has JSDoc.
// Usage: node scripts/exports.mjs [--missing]
import ts from 'typescript';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../packages/spreadiff/', import.meta.url));
const pkg = JSON.parse(readFileSync(root + 'package.json', 'utf8'));
const entries = Object.entries(pkg.exports)
  .map(([k, v]) => [k === '.' ? 'spreadiff' : `spreadiff/${k.slice(2)}`, root + v.import.replace('./dist/', 'src/').replace(/\.js$/, '.ts')])
  .map(([k, f]) => [k, f.endsWith('react/index.ts') ? f.replace(/\.ts$/, '.tsx') : f]);
// spreadiff-lite is a package of its own; its API is documented in docs/api.md too.
entries.push(['spreadiff-lite', fileURLToPath(new URL('../packages/spreadiff-lite/src/index.ts', import.meta.url))]);
const program = ts.createProgram(entries.map(([, f]) => f), { paths: { 'spreadiff/core': [root + 'src/core/index.ts'] }, allowJs: false, jsx: ts.JsxEmit.ReactJSX, moduleResolution: ts.ModuleResolutionKind.Bundler, module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2022, skipLibCheck: true });
const checker = program.getTypeChecker();
export function listExports() {
  const out = [];
  for (const [entry, file] of entries) {
    const sf = program.getSourceFile(file);
    if (!sf) continue;
    const sym = checker.getSymbolAtLocation(sf);
    if (!sym) continue;
    for (const e of checker.getExportsOfModule(sym)) {
      const target = e.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(e) : e;
      const decl = target.declarations?.[0];
      const doc = ts.displayPartsToString(target.getDocumentationComment(checker)).trim();
      const kind = decl ? ts.SyntaxKind[decl.kind].replace(/Declaration$/, '') : '?';
      // Interface / class members declared in our sources must be documented too.
      const members = [];
      if (decl && (ts.isInterfaceDeclaration(decl) || ts.isClassDeclaration(decl))) {
        for (const m of decl.members) {
          if (!m.name || (ts.isClassDeclaration(decl) && (ts.getCombinedModifierFlags(m) & (ts.ModifierFlags.Private | ts.ModifierFlags.Protected)))) continue;
          if (ts.isConstructorDeclaration(m)) continue;
          const msym = checker.getSymbolAtLocation(m.name);
          const mdoc = msym ? ts.displayPartsToString(msym.getDocumentationComment(checker)).trim() : '';
          members.push({ name: m.name.getText(), documented: mdoc.length > 0 });
        }
      }
      out.push({ entry, name: e.name, kind, documented: doc.length > 0, members, file: decl?.getSourceFile().fileName.replace(root, '') });
    }
  }
  return out;
}
if (import.meta.url === `file://${process.argv[1]}`) {
  const all = listExports();
  const missing = process.argv.includes('--missing');
  const seen = new Set();
  let members = 0;
  let membersDoc = 0;
  for (const e of all) {
    const key = `${e.file}:${e.name}`;
    const undocMembers = e.members.filter((m) => !m.documented).map((m) => m.name);
    if (!seen.has(key)) {
      members += e.members.length;
      membersDoc += e.members.length - undocMembers.length;
    }
    if (missing && seen.has(key)) continue;
    seen.add(key);
    if (!missing || !e.documented || undocMembers.length)
      console.log(`${e.documented ? '✓' : '✗'} ${e.entry.padEnd(24)} ${e.name.padEnd(28)} ${e.kind.padEnd(20)} ${e.file}${undocMembers.length ? `  members: ${undocMembers.join(', ')}` : ''}`);
  }
  console.log(`\nexports: ${all.filter((e) => e.documented).length}/${all.length} documented; members: ${membersDoc}/${members}`);
}
