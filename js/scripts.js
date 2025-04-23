const dados = [];

document.getElementById("formulario").addEventListener("submit", function (e) {
  e.preventDefault();
  const form = e.target;
  const tipoSelecionado =
    [...form.tipo].find((input) => input.checked)?.value || "orçamento";
  const data = {
    Setor: form.setor.value,
    tipoEquipamento: form.tipoEquipamento.value,
    Cliente: form.nomeCliente.value,
    Marca: form.marca.value,
    btu: form.btu.value,
    dataorçamento: new Date(form.dataOrcamento.value).toLocaleDateString(),
    descricaoServico: form.descricaoServico.value,
    tipo: tipoSelecionado,
  };
  dados.push(data);
  form.reset();
  document.getElementById(
    "contador"
  ).textContent = `Formulários adicionados: ${dados.length}`;
  alert("Formulário adicionado!");
});

function exportarExcel() {
  if (dados.length === 0) {
    alert("Nenhum formulário para exportar!");
    return;
  }
  const worksheet = XLSX.utils.json_to_sheet(dados);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Orçamentos");
  const excelBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });
  const blob = new Blob([excelBuffer], {
    type: "application/octet-stream",
  });
  saveAs(blob, `orcamentos_${Date.now()}.xlsx`);
  dados.length = 0;
  document.getElementById("contador").textContent = "";
  alert("Arquivo Excel exportado com sucesso!");
}

function toggleCheckbox(selected) {
  const checkboxes = document.querySelectorAll('input[name="tipo"]');
  checkboxes.forEach((cb) => {
    cb.checked = false;
  });
  selected.checked = true;
}
