# spreadiff-lite

**What changed between two spreadsheets, and what it moved downstream, as
Markdown or JSON.** No UI.

For an agent that edits a workbook and has to say what its edits did, a
CI job that comments on a pull request, or a script that checks a model
before it goes out. It reads `.xlsx`, `.csv` and `.tsv` and reports four
things:
- **each edit:** what the sheet calls the cell (`India › Revenue · 2025`),
  and its value before and after;
- **each downstream result:** with the edits that moved it and how much
  each contributed, found by recalculating the workbook;
- **data-feed values:** Bloomberg, Capital IQ, FactSet, LSEG, `RTD`,
  links to other workbooks. They are reported as inputs, never
  recalculated;
- **shape changes:** rows, columns and sheets added, removed or moved.

```bash
npx spreadiff-lite model-v1.xlsx model-v2.xlsx          # Markdown; exit 0 same, 1 changed, 2 error
npx spreadiff-lite model-v1.xlsx model-v2.xlsx --json   # the report as data
```

```ts
import { readFileSync } from 'node:fs';
import { compare, toMarkdown } from 'spreadiff-lite';

const report = compare(
  { name: 'model-v1.xlsx', data: readFileSync('model-v1.xlsx') },
  { name: 'model-v2.xlsx', data: readFileSync('model-v2.xlsx') },
);
console.log(toMarkdown(report));
```

<!-- excerpt-of: lite-regions -->
```markdown
| Cell | What | Before | After | Change | Driven by |
| --- | --- | --- | --- | --- | --- |
| `Model!C5` | India › Profit (USD) · 2025 | 55 | 55.3333333333333 | +0.3333333333333 (+0.61%) | Growth +0.9881; USD/INR −0.6667 |
```

It runs on the same engine as [spreadiff](https://github.com/samueldbarton/spreadiff),
whose causes and contributions are checked against LibreOffice.

**Documentation:**
- [Guide](https://github.com/samueldbarton/spreadiff/blob/main/docs/spreadiff-lite.md)
- [API reference](https://github.com/samueldbarton/spreadiff/blob/main/docs/api.md#spreadiff-lite)
- [How it can be wrong](https://github.com/samueldbarton/spreadiff/blob/main/docs/spreadiff-lite.md#how-it-can-be-wrong)

MIT licensed.
