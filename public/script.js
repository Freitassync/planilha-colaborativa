const socket = io();
const fileInput = document.getElementById("fileInput");
const saveBtn = document.getElementById("saveBtn");
const tableContainer = document.getElementById("tableContainer");

let tableRef = null;

// Carrega o Excel
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const html = XLSX.utils.sheet_to_html(sheet);

    tableContainer.innerHTML = html;
    tableRef = document.querySelector("table");
    makeEditable(tableRef);
  };

  reader.readAsArrayBuffer(file);
});

function makeEditable(table) {
  const cells = table.querySelectorAll("td");
  cells.forEach((cell, index) => {
    cell.contentEditable = "true";

    const row = cell.parentNode.rowIndex;
    const col = cell.cellIndex;

    cell.addEventListener("input", () => {
      const value = cell.innerText;
      socket.emit("cell-update", { row, col, value });
    });
  });
}

// Recebe alteração dos outros usuários
socket.on("cell-update", ({ row, col, value }) => {
  const rows = document.querySelectorAll("table tr");
  if (rows[row]) {
    const cell = rows[row].children[col];
    if (cell) {
      cell.innerText = value;
    }
  }
});

// Exportar como Excel
saveBtn.addEventListener("click", () => {
  if (!tableRef) {
    alert("Carregue uma planilha primeiro.");
    return;
  }

  const ws = XLSX.utils.table_to_sheet(tableRef);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Planilha");

  XLSX.writeFile(wb, "planilha_editada.xlsx");
});
