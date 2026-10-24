# Apache POI test files

From the [Apache POI](https://github.com/apache/poi) test corpus
(`test-data/spreadsheet/`, commit c751824), under the Apache License 2.0. They are real
Excel files with charts and pivot tables, for `test/objects.test.ts`:

- `WithTwoCharts.xlsx`: two charts on two sheets;
- `ExcelPivotTableSample.xlsx`: two pivot tables, over a range and over a table;
- `123233_charts.xlsx`: four charts, in parts whose names differ in case from their links.

And two small .xls (BIFF8) files, for `test/xls.test.ts`:

- `3dFormulas.xls`: formulas with references to other sheets, one named like a cell (`S2`);
- `1904DateWindowing.xls`: dates in the 1904 date system.
