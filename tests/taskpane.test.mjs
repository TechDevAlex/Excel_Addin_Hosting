// excelAddin.test.js

import { ExcelMock } from "office-addin-mock";
import { run, insertcolumn } from "../src/taskpane/taskpane.js";  

// Mock the Excel object
const mockContext = new ExcelMock.MockRequestContext();
ExcelMock.Excel.run = jest.fn((callback) => callback(mockContext));

test("run function should set the fill color of the selected range to yellow", async () => {
  const range = mockContext.workbook.getSelectedRange();
  await run();
  expect(range.format.fill.color).toBe("yellow");
});

test("insertcolumn function should set the topmost cell to 'non-faktura Zeiten'", async () => {
  const range = mockContext.workbook.getSelectedRange();
  range.getCell = jest.fn().mockReturnValue({
    values: [[""]],
  });

  await insertcolumn();
  const topmostCell = range.getCell(0, 0);
  expect(topmostCell.values[0][0]).toBe("non-faktura Zeiten");
});
