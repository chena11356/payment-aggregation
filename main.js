/**
 * Reads spreadsheet of payments into a matrix.
 * @returns {number[][]} Raw payments matrix.
 */
const readMatrix = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const range = sheet.getDataRange();
  const values = range.getValues();

  let numPeople = 0;
  let whoPaidCol = -1;
  let firstOwesCol = -1;

  if (values.length == 1) {
    Logger.log("ERROR: no payments found");
    return;
  }

  // Get the number of people, and the column of who paid
  for (let j = 0; j < values[0].length; j++) {
    if (values[0][j].toLowerCase().includes("owes")) {
      numPeople++;
      if (firstOwesCol == -1) {
        firstOwesCol = j;
      }
    }
    if (values[0][j].toLowerCase().includes("who paid")) {
      whoPaidCol = j;
    }
  }

  if (whoPaidCol == -1) {
    Logger.log("ERROR: could not find WHO PAID column");
    return;
  }

  // Create empty raw payment matrix [i, j] is how much person i owes person j
  let rawPayments = Array(numPeople).fill().map(() => Array(numPeople).fill(0));

  // Fill raw payment matrix
  for (let i = 1; i < values.length; i++) {
    let debtee = values[i][whoPaidCol];
    for (let j = firstOwesCol; j < firstOwesCol + numPeople; j++) {
      let debtor = j - firstOwesCol;
      rawPayments[debtor][debtee] += values[i][j];
    }
  }

  return rawPayments;
}

/**
 * Subtract redundancies from matrix. E.g., if person 0 owes $5 to person 1 and person 1 owes $2 to person 0, person 0 now owes $3 to person 1 and person 1 owes $0 to person 0. 
 * @param {number[][]} matrix - Raw payments matrix.
 * @returns {number[][]} Subtracted payments matrix.
 */
const matrixSubtractRedundancies = (matrix) => {
  for (let i = 0; i < matrix.length; i++) {
    for (let j = i + 1; j < matrix[i].length; j++) {
      if (matrix[i][j] == matrix[j][i]) {
        matrix[i][j] = 0;
        matrix[j][i] = 0;
      } else if (matrix[i][j] > matrix[j][i]) {
        matrix[i][j] -= matrix[j][i];
        matrix[j][i] = 0;
      } else {
        matrix[j][i] -= matrix[i][j];
        matrix[i][j] = 0; 
      }
    }
  }
  return matrix;
}

/**
 * Determine if there are remaining minimizations to be made.
 * @param {number[][]} matrix - Partially minimized payments matrix.
 * @returns {boolean} Whether matrix has remaining minimizations to be made.
 */
const matrixRemainingPaymentsToMinimize = (matrix) => {
  for (let i = 0; i < matrix.length; i++) {
    for (let j = 0; j < matrix.length; j++) {
      for (let k = 0; k < matrix.length; k++) {
        if (i != j && j != k && i != k && matrix[i][j] > 0 && matrix[j][k] > 0 && matrix[i][k] > 0) {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * Minimize number of payments in matrix. E.g., if person 0 owes $7 to person 1, person 0 owes $10 to person 2 and person 1 owes $15 to person 2, person 0 now owes $0 to person 1, person 0 now owes $17 to person 2, and person 1 now owes $8 to person 2. 
 * Repeat until there is no longer any case where person A owes person B, and person B owes person C, and person A owes person C. 
 * @param {number[][]} matrix - Subtracted payments matrix.
 * @returns {number[][]} Minimized payments matrix.
 */
const matrixMinimizePayments = (matrix) => {
  while (matrixRemainingPaymentsToMinimize(matrix)) {
    for (let i = 0; i < matrix.length; i++) {
      for (let j = 0; j < matrix.length; j++) {
        for (let k = 0; k < matrix.length; k++) {
          if (i != j && j != k && i != k) {
            if (matrix[i][j] > 0 && matrix[j][k] > 0 && matrix[i][k] > 0) {
              // Reduce the flow from i to j to k to 0, and increase the flow from i to k to compensate
              if (matrix[i][j] == matrix[j][k]) {
                matrix[i][k] += matrix[i][j];
                matrix[i][j] = 0;
                matrix[j][k] = 0;
              } else if (matrix[i][j] > matrix[j][k]) {
                matrix[i][k] += matrix[j][k];
                matrix[i][j] -= matrix[j][k];
                matrix[j][k] = 0;
              } else {
                matrix[i][k] += matrix[i][j];
                matrix[j][k] -= matrix[i][j];
                matrix[i][j] = 0;
              }
            }
          }
        }
      }
    }
  }
  return matrix;
}

/**
 * Write new minimized matrix underneath original payments matrix.
 * @param {number[][]} matrix - Minimized payments matrix.
 */
const writeMatrix = (matrix) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const range = sheet.getDataRange();
  const values = range.getValues();

  const startRow = values.length + 2;

  // Create x-axis
  const xRange = sheet.getRange(startRow, 2, 1, matrix.length);
  xRange.setFontWeight("bold");
  let xRangeValues = []; 
  let temp = [];
  for (let i = 0; i < matrix.length; i++) {
    temp.push([i]);
  }
  xRangeValues.push(temp);
  xRange.setValues(xRangeValues);

  // Create y-axis
  const yRange = sheet.getRange(startRow + 1, 1, matrix.length, 1);
  yRange.setFontWeight("bold");
  let yRangeValues = []; 
  for (let i = 0; i < matrix.length; i++) {
    yRangeValues.push([i]);
  }
  yRange.setValues(yRangeValues);

  // Fill in the data
  const dataRange = sheet.getRange(startRow + 1, 2, matrix.length, matrix.length);
  dataRange.setValues(matrix);
}

const main = () => {
  const matrix = matrixMinimizePayments(matrixSubtractRedundancies(readMatrix()));
  writeMatrix(matrix);
}

const onOpen = () => {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Payment Aggregation')
      .addItem('Generate aggregated payments', 'main')
      .addToUi();
}
