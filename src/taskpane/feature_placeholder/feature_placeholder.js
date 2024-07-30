/* global Excel, Office */

let entries = [];
const months = [
  "Januar",
  "Februar",
  "März",
  "April",
  "Mai",
  "Juni",
  "Juli",
  "August",
  "September",
  "Oktober",
  "November",
  "Dezember",
];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    populateDropdowns();
    loadEntriesFromSettings();
    populateMonthDropdown();
    document.getElementById("addEntryButton").onclick = addEntry;
    document.getElementById("addFormulaButton").onclick = addFormula;
    document.getElementById("formulaInput").oninput = (event) => resizeTextarea(event.target);
  }
});

async function populateDropdowns() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
      const rangeC = sheet.getRange("C:C").getUsedRange();
      const rangeD = sheet.getRange("D:D").getUsedRange();
      const rangeE = sheet.getRange("E:E").getUsedRange();
      rangeC.load("values");
      rangeD.load("values");
      rangeE.load("values");

      await context.sync();

      const employeesSheet = context.workbook.worksheets.getItem("ExportDaten");

      const rangeEmployees = employeesSheet.getRange("E:E").getUsedRange();
      const rangeColumns = employeesSheet.getRange("1:1").getUsedRange();
      rangeEmployees.load("values");
      rangeColumns.load("values");

      await context.sync();

      const employees = ["Alle", ...(await getEmployeeNames(context))];

      const columnEEntries = [...new Set(rangeE.values.flat())];

      populateDropdown("employeeSelect", employees);
      populateDropdown("columnESelect", columnEEntries);
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 *
 * @param {*} context
 * @returns {Promise<string[]>}
 */
async function getEmployeeNames(context) {
  const sheet = context.workbook.worksheets.getItem("ExportDaten");
  const range = sheet.getRange("E:E").getUsedRange();

  range.load("values");

  await context.sync();

  const employees = [...new Set(range.values.flat().slice(1))];

  return employees;
}

function populateDropdown(elementId, values) {
  const select = document.getElementById(elementId);
  select.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.text = value;
    select.appendChild(option);
  });
}

function populateMonthDropdown() {
  const monthSelect = document.getElementById("monthSelect");
  months.forEach((month) => {
    const option = document.createElement("option");
    option.value = month;
    option.text = month;
    monthSelect.appendChild(option);
  });
}

function updateTable() {
  const tbody = document.getElementById("entriesTableBody");
  tbody.innerHTML = "";

  entries.forEach((entry, index) => {
    const row = document.createElement("tr");

    const nameCell = document.createElement("td");
    nameCell.textContent = entry.employee;
    row.appendChild(nameCell);


    const columnECell = document.createElement("td");
    columnECell.textContent = entry.columnE;
    row.appendChild(columnECell);

    const formulaCell = document.createElement("td");
    const formulaInput = document.createElement("textarea");
    formulaInput.type = "text";
    formulaInput.value = entry.formula;
    formulaInput.onchange = () => updateFormula(index, formulaInput.value);
    formulaInput.style.height = "150px";
    formulaInput.style.width = "300px";
    formulaInput.style.display = "none";
    formulaInput.oninput = () => resizeTextarea(formulaInput);

    const showformulaButton = document.createElement("button");
    showformulaButton.textContent = "Show Formula";
    showformulaButton.onclick = () => showFormulaInTable(index);


    formulaCell.appendChild(showformulaButton);
    formulaCell.appendChild(formulaInput);
    row.appendChild(formulaCell);

    // Actions cell
    const actionsCell = document.createElement("td");

    const editButton = document.createElement("button");
    editButton.textContent = "Edit";
    editButton.onclick = () => toggleEditMode(index, editButton);
    actionsCell.appendChild(editButton);

    const removeButton = document.createElement("button");
    removeButton.textContent = "Remove";
    removeButton.onclick = () => removeEntry(index);
    actionsCell.appendChild(removeButton);

    const testButton = document.createElement("button");
    testButton.textContent = "Test";
    testButton.onclick = () => {
      onTestButtoncClick(index);
    };
    actionsCell.appendChild(testButton);

    row.appendChild(actionsCell);

    tbody.appendChild(row);
  });
}

function showFormulaInTable(index) {
  const formulaInputs = document.querySelectorAll('table textarea');
  const formulaInput = formulaInputs[index];
  const showFormulaButton = document.querySelectorAll('table button')[index];
  if (formulaInput.style.display === "block") {
    formulaInput.style.display = "none";
    showFormulaButton.textContent = "Show Formula";
  }
  else  {
    formulaInput.style.display = "block";
    showFormulaButton.textContent = "Hide Formula";

}
}

function resizeTextarea(textarea) {
  textarea.style.height = "auto";
  textarea.style.height = textarea.scrollHeight + "px";
}


function addEntry() {
  const employee = document.getElementById("employeeSelect").value;
  const columnE = document.getElementById("columnESelect").value;

  const entry = {
    employee: employee,
    columnE: columnE,
    formula: "",
  };

  entries.push(entry);
  updateTable();
  saveEntriesToSettings(); // Save entries to settings
}

function addFormula() {
  const employee = document.getElementById("employeeSelect").value;
  const columnE = document.getElementById("columnESelect").value;
  const formula = document.getElementById("formulaInput").value;

  const entry = entries.find(
    (e) => e.employee === employee && e.columnE === columnE
  );

  if (entry) {
    entry.formula = formula;
    updateTable();
    saveEntriesToSettings(); // Save entries to settings
  } else {
    console.error("Entry not found");
  }
}

function loadEntriesFromSettings() {
  const savedEntries = Office.context.document.settings.get("savedEntries");
  if (savedEntries) {
    entries = savedEntries;
    updateTable();
  }
}

function saveEntriesToSettings() {
  Office.context.document.settings.set("savedEntries", entries);
  Office.context.document.settings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to save settings: " + result.error.message);
    } else {
      console.log("Settings saved.");
    }
  });
}

function updateFormula(index, newFormula) {
  entries[index].formula = newFormula;
  saveEntriesToSettings();
}

function removeEntry(index) {
  entries.splice(index, 1);
  updateTable();
  saveEntriesToSettings();
}

function toggleEditMode(index, button) {
  // const row = document.querySelector(`[data-index='${index}']`).parentNode;

  if (button.textContent === "Edit") {
    button.textContent = "Done";

    // Set the dropdowns and input fields to the corresponding values
    const entry = entries[index];
    document.getElementById("employeeSelect").value = entry.employee;
    document.getElementById("columnESelect").value = entry.columnE;
    document.getElementById("formulaInput").value = entry.formula;

    const input = document.createElement("input");
    input.type = "text";
  } else {
    button.textContent = "Edit";
  }
}

async function findCellByEntry(index, context) {
  console.log(`Finding cell for entry ${index}`);
  const entry = entries[index];

  const [firstName, lastName] = entry.employee.split(" ").map((name) => name.trim());

  const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
  sheet.load("name");
  const rangeC = sheet.getRange("C:C").getUsedRange();
  const rangeD = sheet.getRange("D:D").getUsedRange();

  rangeC.load("values");
  rangeD.load("values");

  await context.sync();

  console.log(`Looking for: First Name = '${firstName}', Last Name = '${lastName}'`);

  let startRow = -1;
  for (let i = 0; i < rangeC.values.length; i++) {
    const currentFirstName = rangeC.values[i][0].trim();
    const currentLastName = rangeD.values[i][0].trim();
    if (currentFirstName === firstName && currentLastName === lastName) {
      startRow = i;
      entries[index].startRow = startRow;
      console.log(`Found name at row: ${startRow + 1}`);
      break;
    }
  }

  return startRow;
}

async function loadExportDaten(context) {
  const employeesSheet = context.workbook.worksheets.getItem("ExportDaten");
  const range = employeesSheet.getUsedRange();
  range.load("values");
  await context.sync();
  return range.values;
}

async function findCellByName(name, context) {
  const firstName = name.split(" ")[0];
  const lastName = name.split(" ")[1];

  const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
  sheet.load("name");

  const rangeC = sheet.getRange("C:C").getUsedRange();
  const rangeD = sheet.getRange("D:D").getUsedRange();

  // TODO this loading (and syncing) is necessary once, but not n times if we want to use the "All" feature.
  // We should load the values once and then use them in the findCellByName function.
  // Presumably, the function does not need to be async anymore.
  rangeC.load("values");
  rangeD.load("values");

  await context.sync();

  // await context.sync();

  let startRow = -1;
  for (let i = 0; i < rangeC.values.length; i++) {
    const currentFirstName = rangeC.values[i][0].trim();
    const currentLastName = rangeD.values[i][0].trim();
    if (currentFirstName === firstName && currentLastName === lastName) {
      startRow = i;
      console.log(`Found name at row: ${startRow + 1}`);
      break;
    }
  }

  if (startRow === -1) {
    console.log(`Name not found`);
  }
  // console.log(startRow);
  return startRow;
}

/**
 *
 * @param {number} index
 */
async function onTestButtoncClick(index) {
  // call findCellByName, load data from the ExportDaten sheet, execute formula, findMatchingRowAndWriteSum
  await Excel.run(async (context) => {
    const entry = entries[index];
    const data = await loadExportDaten(context);
    const employeeNames = await getEmployeeNames(context);

    if (entry.employee === "Alle") {
      const employeeCellPositions = await findAllRelevantCells(context);
      console.log("employeepositions", employeeCellPositions);
      for (let employeeName of employeeNames) {
        let startRow = await findCellByName(employeeName, context);
        if (startRow != -1) {
          let result = await executeFormula(entry, data, employeeName);
          console.log("result", result);
          await findMatchingRowAndWriteResultInColumnE(context, entry, startRow, result);
        }
      }
    } else {
      const startRow = await findCellByEntry(index, context);

      if (startRow !== -1 && data) {
        // Within the onTestButtoncClick function
        const result = await executeFormula(entries[index], data, entry.employee);
        await findMatchingRowAndWriteResultInColumnE(context, entries[index], startRow, result);
      }
      console.log("onTestButtoncClick");
    }
  });
}

/**
 *
 * @param {*} context
 * @returns {Promise<number[]>}
 */
async function findAllRelevantCells(context) {
  const employeeNames = await getEmployeeNames(context);
  const matchingRows = [];

  console.log(employeeNames);

  for (const name of employeeNames) {
    matchingRows.push(await findCellByName(name, context));
  }
  return matchingRows;
}

async function executeFormula(entry, data, employeeName) {
  try {
    // Use the formula from the entry
    const formula = `
      ${entry.formula}
    `;

    // Convert formula string to a function
    const formulaFunction = new Function("data", "employeeName", formula);

    // Execute the formula with the data and employeeName
    return formulaFunction(data, employeeName);
  } catch (error) {
    console.error("Error executing formula: ", error);
  }
}

async function findMatchingRowAndWriteResultInColumnE(context, columnEEntry, startRow, sum) {
  const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
  const rangeE = sheet.getRange("E:E").getUsedRange();
  rangeE.load("values");
  await context.sync();

  // Declare months locally
  let months = [
    "Januar",
    "Februar",
    "März",
    "April",
    "Mai",
    "Juni",
    "Juli",
    "August",
    "September",
    "Oktober",
    "November",
    "Dezember",
  ];

  const monthSelect = document.getElementById("monthSelect").value;
  const columnIndex = months.indexOf(monthSelect) + 6; // Assuming month columns start from column G (index 6)

  if (columnIndex === -1) {
    console.error("Selected month not found in months array.");
    return;
  }

  for (let j = startRow; j < rangeE.values.length; j++) {
    const currentE = rangeE.values[j][0].trim();

    if (currentE === columnEEntry.columnE) {
      const cell = sheet.getRangeByIndexes(j + 4, columnIndex, 1, 1);
      cell.load("address");
      cell.values = [[sum]];
      await context.sync();
      console.log(`Wrote sum ${sum} to cell ${cell.address}`);
      break;
    }
  }
}
