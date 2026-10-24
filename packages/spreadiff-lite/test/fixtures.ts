import ExcelJS from 'exceljs';

export interface ModelInputs {
  growth: number;
  /** USD/INR, from Bloomberg. */
  inr: number;
  /** USD/CNY, from Bloomberg. */
  cny: number;
  indiaRevenue2024: number;
  chinaRevenue2024: number;
  indiaCosts2024: number;
  chinaCosts2024: number;
}

/**
 * A regional model as an analyst would build it, saved with cached values
 * as Excel would:
 *
 *   Inputs   Growth | USD/INR = BDP(...) | USD/CNY = BDP(...)
 *   Model    FY2024 / FY2025 columns; bold region headings, indented lines:
 *            India › Revenue, Costs, Profit (local) · Profit (USD); China › same
 *   Summary  Total profit (USD) per year
 */
export async function regionalModel(v: ModelInputs): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const inputs = wb.addWorksheet('Inputs');
  inputs.getCell('A1').value = 'Growth';
  inputs.getCell('B1').value = v.growth;
  inputs.getCell('A2').value = 'USD/INR';
  inputs.getCell('B2').value = { formula: '_xll.BDP("USDINR Curncy","PX_LAST")', result: v.inr };
  inputs.getCell('A3').value = 'USD/CNY';
  inputs.getCell('B3').value = { formula: '_xll.BDP("USDCNY Curncy","PX_LAST")', result: v.cny };

  const m = wb.addWorksheet('Model');
  m.getCell('B1').value = 2024;
  m.getCell('C1').value = { formula: 'B1+1', result: 2025 };
  const region = (top: number, name: string, rev: number, costs: number, fx: string, fxValue: number) => {
    const head = m.getCell(`A${top}`);
    head.value = name;
    head.font = { bold: true };
    const line = (row: number, label: string) => {
      const c = m.getCell(`A${row}`);
      c.value = label;
      c.alignment = { indent: 1 };
    };
    const g = 1 + v.growth;
    line(top + 1, 'Revenue');
    m.getCell(`B${top + 1}`).value = rev;
    m.getCell(`C${top + 1}`).value = { formula: `B${top + 1}*(1+Inputs!$B$1)`, result: rev * g };
    line(top + 2, 'Costs');
    m.getCell(`B${top + 2}`).value = costs;
    m.getCell(`C${top + 2}`).value = { formula: `B${top + 2}*(1+Inputs!$B$1)`, result: costs * g };
    line(top + 3, 'Profit');
    for (const [col, k] of [['B', 1], ['C', g]] as const) {
      m.getCell(`${col}${top + 3}`).value = { formula: `${col}${top + 1}-${col}${top + 2}`, result: (rev - costs) * k };
    }
    line(top + 4, 'Profit (USD)');
    for (const [col, k] of [['B', 1], ['C', g]] as const) {
      m.getCell(`${col}${top + 4}`).value = { formula: `${col}${top + 3}/Inputs!$B$${fx}`, result: ((rev - costs) * k) / fxValue };
    }
  };
  region(2, 'India', v.indiaRevenue2024, v.indiaCosts2024, '2', v.inr);
  region(7, 'China', v.chinaRevenue2024, v.chinaCosts2024, '3', v.cny);

  const s = wb.addWorksheet('Summary');
  s.getCell('B1').value = { formula: 'Model!B1', result: 2024 };
  s.getCell('C1').value = { formula: 'Model!C1', result: 2025 };
  s.getCell('A2').value = 'Total profit (USD)';
  const g = 1 + v.growth;
  for (const [col, k] of [['B', 1], ['C', g]] as const) {
    const usd = ((v.indiaRevenue2024 - v.indiaCosts2024) * k) / v.inr + ((v.chinaRevenue2024 - v.chinaCosts2024) * k) / v.cny;
    s.getCell(`${col}2`).value = { formula: `Model!${col}6+Model!${col}11`, result: usd };
  }
  return new Uint8Array(await wb.xlsx.writeBuffer());
}
