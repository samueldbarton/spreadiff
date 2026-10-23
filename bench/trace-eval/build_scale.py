import sys, os
import lo, models, build
build.OUT = 'scale-cases'
os.makedirs('scale-cases', exist_ok=True); os.makedirs('bases', exist_ok=True)
for name, fn, n in [('chain20k', models.chain, 20000), ('grid30k', models.lookup_grid, 30000), ('grid100k', models.lookup_grid, 100000)]:
    path = f'bases/{name}.xlsx'
    edits = fn(path, n)
    build.build(name, path, edits, singles=False)
lo.stop()
