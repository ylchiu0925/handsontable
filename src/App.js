import { useRef } from 'react';
import { HotTable } from '@handsontable/react';
import { registerAllModules } from 'handsontable/registry';
import { HyperFormula } from 'hyperformula';
import 'handsontable/dist/handsontable.full.min.css';

registerAllModules();

const initialData = [
  ['', 'Tesla', 'Volvo', 'Toyota', 'Ford'],
  ['2019', 10, 11, 12, 13],
  ['2020', 20, 11, 14, 13],
  ['2021', 30, 15, 12, 13],
  ['SUM', '=SUM(B2:B4)', '=SUM(C2:C4)', '=SUM(D2:D4)', '=SUM(E2:E4)'],
  ['AVERAGE', '=AVERAGE(B2:B4)', '=AVERAGE(C2:C4)', '=AVERAGE(D2:D4)', '=AVERAGE(E2:E4)'],
];


const isNum = v => v !== '' && v !== null && !isNaN(Number(v));

const getLinearStep = (nums) => {
  if (nums.length < 2) return null;
  const step = nums[1] - nums[0];
  const isLinear = nums.every((v, i) => i === 0 || Math.abs(v - nums[i - 1] - step) < 0.0001);
  return isLinear ? step : null;
};

const beforeAutofill = (selectionData, sourceRange, targetRange, direction) => {
  if (direction === 'down') {
    const numCols = selectionData[0].length;
    const fillRows = targetRange.to.row - targetRange.from.row + 1;
    const result = Array.from({ length: fillRows }, () => Array(numCols).fill(''));

    for (let c = 0; c < numCols; c++) {
      const colValues = selectionData.map(row => row[c]);
      if (colValues.every(isNum)) {
        const nums = colValues.map(Number);
        const step = getLinearStep(nums);
        if (step !== null) {
          let last = nums[nums.length - 1];
          for (let r = 0; r < fillRows; r++) result[r][c] = (last += step);
          continue;
        }
      }
      for (let r = 0; r < fillRows; r++) result[r][c] = colValues[r % colValues.length];
    }
    return result;
  }

  if (direction === 'up') {
    const numCols = selectionData[0].length;
    const fillRows = targetRange.to.row - targetRange.from.row + 1;
    const result = Array.from({ length: fillRows }, () => Array(numCols).fill(''));

    for (let c = 0; c < numCols; c++) {
      const colValues = selectionData.map(row => row[c]);
      if (colValues.every(isNum)) {
        const nums = colValues.map(Number);
        const step = getLinearStep(nums);
        if (step !== null) {
          let first = nums[0];
          for (let r = fillRows - 1; r >= 0; r--) result[r][c] = (first -= step);
          continue;
        }
      }
      for (let r = 0; r < fillRows; r++) result[r][c] = colValues[r % colValues.length];
    }
    return result;
  }

  if (direction === 'right') {
    const numRows = selectionData.length;
    const fillCols = targetRange.to.col - targetRange.from.col + 1;
    const result = Array.from({ length: numRows }, () => Array(fillCols).fill(''));

    for (let r = 0; r < numRows; r++) {
      const rowValues = selectionData[r];
      if (rowValues.every(isNum)) {
        const nums = rowValues.map(Number);
        const step = getLinearStep(nums);
        if (step !== null) {
          let last = nums[nums.length - 1];
          for (let c = 0; c < fillCols; c++) result[r][c] = (last += step);
          continue;
        }
      }
      for (let c = 0; c < fillCols; c++) result[r][c] = rowValues[c % rowValues.length];
    }
    return result;
  }

  if (direction === 'left') {
    const numRows = selectionData.length;
    const fillCols = targetRange.to.col - targetRange.from.col + 1;
    const result = Array.from({ length: numRows }, () => Array(fillCols).fill(''));

    for (let r = 0; r < numRows; r++) {
      const rowValues = selectionData[r];
      if (rowValues.every(isNum)) {
        const nums = rowValues.map(Number);
        const step = getLinearStep(nums);
        if (step !== null) {
          let first = nums[0];
          for (let c = fillCols - 1; c >= 0; c--) result[r][c] = (first -= step);
          continue;
        }
      }
      for (let c = 0; c < fillCols; c++) result[r][c] = rowValues[c % rowValues.length];
    }
    return result;
  }
};

function App() {
  const hotRef = useRef(null);

  const addRow = () => {
    const hot = hotRef.current.hotInstance;
    hot.alter('insert_row_below', hot.countRows());
  };

  const addCol = () => {
    const hot = hotRef.current.hotInstance;
    hot.alter('insert_col_end', hot.countCols());
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
        beforeAutofill={beforeAutofill}
        formulas={{ engine: HyperFormula, licenseKey: 'internal-use-in-handsontable' }}
        licenseKey="non-commercial-and-evaluation"
      />
    </div>
  );
}

export default App;
