Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("btn-escrever").onclick = escreverNaCelula;
  }
});

async function escreverNaCelula() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [["Funcionou!"]];
      range.format.autofitColumns();
      await context.sync();
      document.getElementById("status").innerText = "Texto inserido com sucesso!";
    });
  } catch (error) {
    document.getElementById("status").innerText = "Erro: " + error.message;
  }
}
