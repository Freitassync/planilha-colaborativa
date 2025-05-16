const fileInput = document.getElementById("fileInput");
const spreadsheetContainer = document.getElementById("spreadsheetContainer");
const socket = io();

fileInput.addEventListener("change", handleFileUpload);

function handleFileUpload(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    renderTable(jsonData);
  };

  reader.readAsArrayBuffer(file);
}

function renderTable(data) {
  const table = document.createElement("table");

  data.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");

    row.forEach((cell, colIndex) => {
      const td = document.createElement("td");
      td.textContent = cell;
      td.contentEditable = "true";

      td.addEventListener("input", () => {
        socket.emit("cell-update", {
          row: rowIndex,
          col: colIndex,
          value: td.textContent
        });
      });

      td.dataset.row = rowIndex;
      td.dataset.col = colIndex;

      tr.appendChild(td);
    });

    table.appendChild(tr);
  });

  spreadsheetContainer.innerHTML = "";
  spreadsheetContainer.appendChild(table);
}

// Atualiza células modificadas por outros usuários
socket.on("cell-update", ({ row, col, value }) => {
  const cell = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (cell) {
    cell.textContent = value;
  }
});
document.getElementById("saveBtn").addEventListener("click", () => {
  const table = document.querySelector("table");
  if (!table) {
    alert("Nenhuma planilha carregada!");
    return;
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, "Planilha");

  XLSX.writeFile(wb, "planilha_editada.xlsx");
});
