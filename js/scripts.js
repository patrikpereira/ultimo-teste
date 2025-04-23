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
    observacao: form.observacao.value,
    tipo: tipoSelecionado,
  };
  dados.push(data);

  const cliente = form.nomeCliente.value; // salva o nome antes
  form.reset(); // reseta tudo
  form.nomeCliente.value = cliente; // restaura o nome

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

  const dadosFormatados = dados.map((item) => ({
    Setor: item.Setor,
    "Tipo do Equipamento": item.tipoEquipamento,
    Cliente: item.Cliente,
    Marca: item.Marca,
    BTU: item.btu,
    "Data do Orçamento": item.dataorçamento,
    "Descrição do Serviço": item.descricaoServico,
    Tipo: item.tipo,
  }));

  const worksheet = XLSX.utils.json_to_sheet(dadosFormatados);

  // Definir largura automática das colunas
  const colWidths = Object.keys(dadosFormatados[0]).map((key) => {
    const maxLen = Math.max(
      key.length,
      ...dadosFormatados.map((row) =>
        row[key] ? row[key].toString().length : 0
      )
    );
    return { wch: maxLen + 2 };
  });
  worksheet["!cols"] = colWidths;

  // Estilizar células
  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      if (!worksheet[cellAddress]) continue;

      worksheet[cellAddress].s = {
        font: { bold: R === 0, color: { rgb: R === 0 ? "FFFFFF" : "000000" } },
        fill: {
          fgColor: {
            rgb: R === 0 ? "00C853" : R % 2 === 0 ? "F1F1F1" : "FFFFFF", // header / zebra striping
          },
        },
        alignment: {
          horizontal: "center",
          vertical: "center",
        },
        border: {
          top: { style: "thin", color: { rgb: "AAAAAA" } },
          bottom: { style: "thin", color: { rgb: "AAAAAA" } },
          left: { style: "thin", color: { rgb: "AAAAAA" } },
          right: { style: "thin", color: { rgb: "AAAAAA" } },
        },
      };
    }
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Orçamentos");

  const excelBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
    cellStyles: true,
  });

  const blob = new Blob([excelBuffer], { type: "application/octet-stream" });
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
