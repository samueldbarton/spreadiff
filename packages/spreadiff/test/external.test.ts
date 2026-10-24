import { describe, expect, it } from 'vitest';
import { externalSource } from '../src/core';

describe('externalSource', () => {
  it('recognises market-data add-ins, with or without the _xll. prefix', () => {
    expect(externalSource('=_xll.BDP("IBM US Equity","PX_LAST")')).toEqual({ provider: 'Bloomberg', function: 'BDP' });
    expect(externalSource('=bdh(A1,"PX_LAST",B1,C1)')).toEqual({ provider: 'Bloomberg', function: 'BDH' });
    expect(externalSource('=CIQAVG("IBM","IQ_EPS_EST")')?.provider).toBe('S&P Capital IQ');
    expect(externalSource('=FDS("AAPL-US","P_PRICE(0)")')?.provider).toBe('FactSet');
    expect(externalSource('=TR("IBM.N","TR.PriceClose")')?.provider).toBe('LSEG');
    expect(externalSource('=_xlfn.STOCKHISTORY("MSFT",A1)')?.provider).toBe('Excel');
    expect(externalSource('=RTD("x",,"y")')?.function).toBe('RTD');
  });
  it('treats any other XLL add-in function as external', () => {
    expect(externalSource('=_xll.VADATA("x")')).toEqual({ provider: 'add-in', function: 'VADATA' });
  });
  it('finds links to other workbooks, but not structured references', () => {
    expect(externalSource('=[1]Sheet1!A1*2')).toEqual({ provider: 'link', function: '[1]' });
    expect(externalSource("='[2]My sheet'!B3")).toEqual({ provider: 'link', function: '[2]' });
    expect(externalSource('=SUM(Sales[2024])')).toBeNull();
  });
  it('ignores names inside strings, longer names and ordinary formulas', () => {
    expect(externalSource('=CONCAT("BDP(",A1,")")')).toBeNull();
    expect(externalSource('=SUM(A1:A3)*TRUNC(B1)')).toBeNull(); // TRUNC isn't TR
    expect(externalSource('=MY.BDP(A1)')).toBeNull();
    expect(externalSource('')).toBeNull();
  });
  it('accepts extra functions', () => {
    expect(externalSource('=PRICE_OF(A1)', { 'In-house': ['PRICE_OF'] })).toEqual({ provider: 'In-house', function: 'PRICE_OF' });
  });
});
