const tbody = document.querySelector("#sheet tbody");
const thead = document.querySelector("#sheet thead");
let isMouseDown = false;
let selectedCells = [];
let borderColor = '#000000'; // Default border color

// Function to create a new cell
function createCell(row, isHeader = false) {
  const td = document.createElement("td");
  if (isHeader) {
    td.classList.add("header");
    td.contentEditable = false;
  } else {
    td.contentEditable = "true";
    td.classList.add("cell-container");
    td.addEventListener("mousedown", () => {
      clearSelection();
      isMouseDown = true;
      td.classList.add("selected");
      selectedCells.push(td);
    });
    td.addEventListener("mouseover", () => {
      if (isMouseDown && !td.classList.contains("selected")) {
        td.classList.add("selected");
        selectedCells.push(td);
      }
    });

    // Add resizers
    const rightResizer = document.createElement("div");
    rightResizer.className = "resizer right";
    td.appendChild(rightResizer);

    const bottomResizer = document.createElement("div");
    bottomResizer.className = "resizer bottom";
    td.appendChild(bottomResizer);

    // Resizing logic
    let startX, startY, startWidth, startHeight;
    rightResizer.addEventListener("mousedown", function (e) {
      e.preventDefault();
      startX = e.clientX;
      startWidth = parseInt(document.defaultView.getComputedStyle(td).width, 10);
      document.documentElement.addEventListener("mousemove", resizeCol);
      document.documentElement.addEventListener("mouseup", stopResize);
    });

    bottomResizer.addEventListener("mousedown", function (e) {
      e.preventDefault();
      startY = e.clientY;
      startHeight = parseInt(document.defaultView.getComputedStyle(td).height, 10);
      document.documentElement.addEventListener("mousemove", resizeRow);
      document.documentElement.addEventListener("mouseup", stopResize);
    });

    function resizeCol(e) {
      td.style.width = (startWidth + e.clientX - startX) + "px";
    }

    function resizeRow(e) {
      td.style.height = (startHeight + e.clientY - startY) + "px";
    }

    function stopResize() {
      document.documentElement.removeEventListener("mousemove", resizeCol);
      document.documentElement.removeEventListener("mousemove", resizeRow);
      document.documentElement.removeEventListener("mouseup", stopResize);
    }
  }

  return td;
}

// Function to update headers
function updateHeaders() {
  const rows = tbody.rows.length;
  const cols = tbody.rows[0]?.cells.length || 3;

  thead.innerHTML = "";
  const headerRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.classList.add("header");
  corner.innerHTML = "#";
  headerRow.appendChild(corner);

  for (let i = 0; i < cols - 1; i++) {
    const th = document.createElement("th");
    th.classList.add("header");
    th.innerText = String.fromCharCode(65 + i);
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  [...tbody.rows].forEach((row, i) => {
    row.cells[0].innerHTML = i + 1;
    row.cells[0].contentEditable = false;
  });
}

// Add a row to the table
function addRow() {
  const row = document.createElement("tr");
  const cols = tbody.rows[0]?.cells.length || 4;
  for (let i = 0; i < cols; i++) {
    const cell = createCell(row, i === 0);
    row.appendChild(cell);
  }
  tbody.appendChild(row);
  updateHeaders();
}

// Add a column to the table
function addColumn() {
  for (let row of tbody.rows) {
    const cell = createCell(row, row.cells.length === 0);
    row.appendChild(cell);
  }
  updateHeaders();
}

// Delete a row
function deleteRow() {
  if (tbody.rows.length > 0) {
    tbody.deleteRow(tbody.rows.length - 1);
    updateHeaders();
  }
}

// Delete a column
function deleteColumn() {
  if (tbody.rows[0]?.cells.length > 1) {
    for (let row of tbody.rows) {
      row.deleteCell(row.cells.length - 1);
    }
    updateHeaders();
  }
}

// Add a date column to the table
function addDateColumn() {
  const today = new Date().toLocaleDateString();
  if (tbody.rows.length === 0) addRow();
  for (let row of tbody.rows) {
    const cell = document.createElement("td");
    cell.contentEditable = false;
    cell.innerText = today;
    row.appendChild(cell);
  }
  updateHeaders();
}

// Calculate the sum of a selected column
function calculateSum() {
  const columnIndex = prompt("Enter column letter to sum (A, B, C...)");
  if (!columnIndex) return;

  const col = columnIndex.toUpperCase().charCodeAt(0) - 65 + 1;
  let sum = 0;

  for (let row of tbody.rows) {
    const cell = row.cells[col];
    if (cell && !isNaN(cell.innerText)) {
      sum += parseFloat(cell.innerText);
    }
  }

  alert("Sum of column " + columnIndex + " = " + sum);
}

// Apply text formatting (bold, italic, etc.)
function formatText(command) {
  document.execCommand(command, false, null);
}

// Save the sheet to localStorage
function saveSheet() {
  const data = [];
  for (let row of tbody.rows) {
    const rowData = [];
    for (let cell of row.cells) {
      rowData.push({
        html: cell.innerHTML,
        colspan: cell.colSpan || 1,
        rowspan: cell.rowSpan || 1
      });
    }
    data.push(rowData);
  }
  localStorage.setItem("excelSheet", JSON.stringify(data));
  alert("✅ Sheet saved!");
}

// Load the sheet from localStorage
function loadSheet() {
  const data = JSON.parse(localStorage.getItem("excelSheet") || "[]");
  tbody.innerHTML = "";
  for (let rowData of data) {
    const row = document.createElement("tr");
    for (let i = 0; i < rowData.length; i++) {
      const cellData = rowData[i];
      const cell = createCell(row, i === 0);
      cell.innerHTML = cellData.html;
      cell.colSpan = cellData.colspan;
      cell.rowSpan = cellData.rowspan;
      row.appendChild(cell);
    }
    tbody.appendChild(row);
  }
  updateHeaders();
}

// Export the table data as CSV
function exportCSV() {
  let csv = "";
  for (let row of tbody.rows) {
    const rowData = Array.from(row.cells).map(cell => `"${cell.innerText}"`);
    csv += rowData.join(",") + "\n";
  }
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "sheet.csv";
  a.click();
}

// Clear the contents of the sheet
function clearSheet() {
  for (let row of tbody.rows) {
    for (let i = 1; i < row.cells.length; i++) {
      row.cells[i].innerHTML = "";
      row.cells[i].colSpan = 1;
      row.cells[i].rowSpan = 1;
    }
  }
}

// Search the table based on input
function searchTable() {
  const value = document.getElementById("searchInput").value.toLowerCase();
  for (let row of tbody.rows) {
    row.style.display = [...row.cells].some(cell =>
      cell.innerText.toLowerCase().includes(value)
    ) ? "" : "none";
  }
}

// Clear selection of cells
function clearSelection() {
  selectedCells.forEach(cell => cell.classList.remove("selected"));
  selectedCells = [];
}

// Function to open the border color picker
function openBorderColorPicker() {
  const colorPicker = document.createElement("input");
  colorPicker.type = "color";
  colorPicker.value = borderColor;
  colorPicker.addEventListener("input", (event) => {
    borderColor = event.target.value;
    applyBorderColor();
  });

  document.body.appendChild(colorPicker);
  colorPicker.click();
  document.body.removeChild(colorPicker);
}

// Apply the selected border color to the selected cells
function applyBorderColor() {
  if (selectedCells.length > 0) {
    selectedCells.forEach(cell => {
      cell.style.border = `2px solid ${borderColor}`;
    });
  }
}

// Merge the selected cells
function mergeCells() {
  if (selectedCells.length <= 1) return alert("Select more than 1 cell to merge.");

  const firstCell = selectedCells[0];
  let mergedContent = selectedCells.map(c => c.innerHTML).join(" ");
  let rowspan = 1, colspan = 1;

  const rows = [...new Set(selectedCells.map(c => c.parentElement))];
  rowspan = rows.length;
  const cols = [...new Set(selectedCells.map(c => c.cellIndex))];
  colspan = cols.length;

  selectedCells.forEach((cell, i) => {
    if (i === 0) {
      cell.innerHTML = mergedContent;
      cell.rowSpan = rowspan;
      cell.colSpan = colspan;
    } else {
      cell.remove();
    }
  });

  clearSelection();
}

// Unmerge the selected cells
function unmergeCells() {
  for (let row of tbody.rows) {
    for (let cell of [...row.cells]) {
      if (cell.rowSpan > 1 || cell.colSpan > 1) {
        const rowIndex = row.rowIndex - 1;
        const colIndex = cell.cellIndex;

        const content = cell.innerHTML;
        const rs = cell.rowSpan;
        const cs = cell.colSpan;

        cell.rowSpan = 1;
        cell.colSpan = 1;
        cell.innerHTML = "";

        for (let i = 0; i < rs; i++) {
          for (let j = 0; j < cs; j++) {
            if (i === 0 && j === 0) {
              cell.innerHTML = content;
            } else {
              const newCell = document.createElement("td");
              newCell.contentEditable = true;
              tbody.rows[rowIndex + i].insertBefore(newCell, tbody.rows[rowIndex + i].cells[colIndex + j] || null);
            }
          }
        }
      }
    }
  }
}

// Export the table data to PDF using jsPDF
function exportToPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  // Add title to PDF
  doc.text("Excel-like Sheet", 10, 10);

  // Prepare table data
  const tableRows = [];
  for (let row of tbody.rows) {
    const rowData = [];
    for (let cell of row.cells) {
      rowData.push(cell.innerText);
    }
    tableRows.push(rowData);
  }

  // Add table to PDF
  doc.autoTable({
    head: [['#', 'Column A', 'Column B', 'Column C']], // Update headers if necessary
    body: tableRows,
    startY: 20,
    theme: 'grid',
  });

  // Save the PDF
  doc.save('sheet.pdf');
}

// Initialize the sheet
window.onload = () => {
  for (let i = 0; i < 10; i++) {
    addRow();
  }
};
function exportToPDF() {
  // Initialize jsPDF
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  // Capture HTML table data
  const table = document.querySelector("#sheet"); // Your sheet ID

  // Add table content to PDF
  doc.autoTable({ html: table });

  // Add footer text
  const footerText = "Made with SheetCraft";
  doc.setFontSize(10);
  doc.text(footerText, 10, doc.internal.pageSize.height - 10); // Position at the bottom

  // Save the PDF with the name "SheetCraftChiranjeev.pdf"
  doc.save("SheetCraftChiranjeev.pdf");
}
doc.autoTable({
  html: table,
  startY: startY,
  styles: {
    fontSize: 8,
    cellPadding: 2,
  },
  columnStyles: {
    0: { cellWidth: 'auto' }, // example
  }
});
