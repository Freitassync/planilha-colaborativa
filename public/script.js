const socket = io(); // Conecta ao servidor WebSocket

const table = document.getElementById("table");
const fileInput = document.getElementById("fileInput");

// Carrega o Excel
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const html = XLSX.utils.sheet_to_html(sheet);
    table.innerHTML = html;

    makeCellsEditable();
  };

  reader.readAsArrayBuffer(file);
});

// Torna as células editáveis e emite eventos via socket
function makeCellsEditable() {
  const cells = document.querySelectorAll("table td");
  cells.forEach((cell) => {
    cell.setAttribute("contenteditable", "true");

    cell.addEventListener("input", () => {
      const row = cell.parentNode.rowIndex;
      const col = cell.cellIndex;
      const value = cell.innerText;

      // Envia para o servidor
      socket.emit("cell-update", { row, col, value });
    });
  });
}

// Quando receber atualização de outro usuário:
socket.on("cell-update", ({ row, col, value }) => {
  const rowElement = document.querySelectorAll("table tr")[row];
  if (rowElement) {
    const cell = rowElement.children[col];
    if (cell) {
      cell.innerText = value;
    }
  }
});

document.getElementById("saveBtn").addEventListener("click", () => {
  const tableEl = document.querySelector("table");
  if (!tableEl) {
    alert("Nenhuma planilha carregada!");
    return;
  }

  // Converte tabela HTML para worksheet
  const worksheet = XLSX.utils.table_to_sheet(tableEl);

  // Cria novo workbook e adiciona a worksheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Planilha");

  // Exporta como arquivo .xlsx
  XLSX.writeFile(workbook, "planilha_editada.xlsx");
});

