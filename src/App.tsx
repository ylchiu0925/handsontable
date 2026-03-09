import { useRef, useEffect } from 'react';
import { HotTable, HotTableClass } from '@handsontable/react';
import { registerAllModules } from 'handsontable/registry';
import { HyperFormula } from 'hyperformula';
import { CellValue } from 'handsontable/common';
import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.min.css';

registerAllModules();

const initialData: CellValue[][] = [
  ['', 'Tesla', 'Volvo', 'Toyota', 'Ford'],
  ['2019', 10, 11, 12, 13],
  ['2020', 20, 11, 14, 13],
  ['2021', 30, 15, 12, 13],
  ['SUM', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', '=SUM(E2:E4)'],
  ['AVERAGE', '=AVERAGE(B2:B4)', '=AVERAGE(C2:C4)', '=AVERAGE(D2:D4)', '=AVERAGE(E2:E4)'],
];

const isNum = (v: CellValue): boolean =>
  v !== '' && v !== null && v !== undefined && !isNaN(Number(v));

const getLinearStep = (nums: number[]): number | null => {
  if (nums.length < 2) return null;
  const step = nums[1] - nums[0];
  const isLinear = nums.every((v, i) => i === 0 || Math.abs(v - nums[i - 1] - step) < 0.0001);
  return isLinear ? step : null;
};

function App() {
  const hotRef = useRef<HotTableClass>(null);

  useEffect(() => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;

    const handleAutofill = (
      selectionData: CellValue[][],
      _sourceRange: Handsontable.CellRange,
      targetRange: Handsontable.CellRange,
      direction: string
    ): CellValue[][] => {
      const tFrom = targetRange.getTopStartCorner();
      const tTo = targetRange.getBottomEndCorner();
      const fillRows = tTo.row - tFrom.row + 1;
      const fillCols = tTo.col - tFrom.col + 1;

      if (direction === 'down' || direction === 'up') {
        const numCols = selectionData[0]?.length ?? 0;
        const result: CellValue[][] = Array.from({ length: fillRows }, () => Array(numCols).fill(''));
        for (let c = 0; c < numCols; c++) {
          const colValues = selectionData.map(row => row[c]);
          if (colValues.every(isNum)) {
            const nums = colValues.map(Number);
            const step = getLinearStep(nums);
            if (step !== null) {
              if (direction === 'down') {
                let last = nums[nums.length - 1];
                for (let r = 0; r < fillRows; r++) result[r][c] = (last += step);
              } else {
                let first = nums[0];
                const seq: number[] = [];
                for (let r = 0; r < fillRows; r++) seq.unshift((first -= step));
                seq.forEach((v, r) => { result[r][c] = v; });
              }
              continue;
            }
          }
          for (let r = 0; r < fillRows; r++) result[r][c] = colValues[r % colValues.length];
        }
        return result;
      }

      if (direction === 'right' || direction === 'left') {
        const numRows = selectionData.length;
        const result: CellValue[][] = Array.from({ length: numRows }, () => Array(fillCols).fill(''));
        for (let r = 0; r < numRows; r++) {
          const rowValues = selectionData[r];
          if (rowValues.every(isNum)) {
            const nums = rowValues.map(Number);
            const step = getLinearStep(nums);
            if (step !== null) {
              if (direction === 'right') {
                let last = nums[nums.length - 1];
                for (let c = 0; c < fillCols; c++) result[r][c] = (last += step);
              } else {
                let first = nums[0];
                const seq: number[] = [];
                for (let c = 0; c < fillCols; c++) seq.unshift((first -= step));
                seq.forEach((v, c) => { result[r][c] = v; });
              }
              continue;
            }
          }
          for (let c = 0; c < fillCols; c++) result[r][c] = rowValues[c % rowValues.length];
        }
        return result;
      }

      return selectionData;
    };

    hot.addHook('beforeAutofill', handleAutofill as any);
    return () => { hot.removeHook('beforeAutofill', handleAutofill as any); };
  }, []);

  const addRow = (): void => {
    const hot = hotRef.current?.hotInstance;
    hot?.alter('insert_row_below', hot.countRows());
  };

  const addCol = (): void => {
    const hot = hotRef.current?.hotInstance;
    hot?.alter('insert_col_end', hot.countCols());
  };

  return (
    <div style={{ padding: '20px' }}>
      <h2>Handsontable with React</h2>
      <div style={{ marginBottom: '10px', display: 'flex', gap: '8px' }}>
        <button onClick={addRow}>+ Add Row</button>
        <button onClick={addCol}>+ Add Column</button>
      </div>
      <HotTable
        ref={hotRef}
        data={initialData}
        rowHeaders={true}
        colHeaders={true}
        height="auto"
        contextMenu={true}
        fillHandle={{ autoInsertRow: true }}
        filters={true}
        dropdownMenu={true}
        columnSorting={true}
        formulas={{ engine: HyperFormula }}
        licenseKey="non-commercial-and-evaluation"
      />
    </div>
  );
}

export default App;
