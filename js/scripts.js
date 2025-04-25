const dados = [];
document.getElementById("formulario").addEventListener("submit", function (e) {
  e.preventDefault();

  // Captura segura de todos os campos
  const formData = new FormData(this);

  // Verificação explícita do campo técnicoEquipe
  const tecnicoEquipe = formData.get("tecnicoEquipe");
  if (!tecnicoEquipe) {
    alert("Por favor, preencha o campo 'Técnico / Equipe'");
    return;
  }

  const data = {
    tipo: formData.get("tipo"),
    tecnicoEquipe: formData.get("tecnicoEquipe"),
    Cliente: formData.get("nomeCliente"),
    Setor: formData.get("setor"),
    tipoEquipamento: formData.get("tipoEquipamento"),
    Marca: formData.get("marca"),
    btu: formData.get("btu"),
    dataorçamento:
      formData.get("dataOrcamento")?.split("-").reverse().join("/") || "",
    descricaoServico: Array.from(
      document.querySelectorAll('input[name="descricaoServico"]:checked')
    )
      .map((cb) => cb.value)
      .join(", "),
    observacao: formData.get("observacao"),
    responsavel: formData.get("responsavel"),
    numeroMaquina: formData.get("numeroMaquina"),
    assinatura:
      document.getElementById("assinatura")?.toDataURL("image/png") || "",
  };

  // Adiciona aos dados
  dados.push(data);

  // Limpa o formulário e os checkboxes selecionados
  this.reset();

  // Reset do select de servicos
  const selectServico = document.getElementById("selectServico");
  if (selectServico) {
    selectServico.selectedIndex = 0;
  }

  // Remove todos os checkboxes de descrição do serviço
  const checkboxesContainer = document.getElementById("checkboxesContainer");
  if (checkboxesContainer) {
    checkboxesContainer.innerHTML = ""; // Limpa todo o conteúdo
  }

  // Limpa a assinatura
  const canvas = document.getElementById("assinatura");
  if (canvas) {
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);
  }

  // Mantém o nome do cliente se necessário
  this.nomeCliente.value = data.Cliente;
  this.tecnicoEquipe.value = data.tecnicoEquipe;
  this.dataOrcamento.value = formData.get("dataOrcamento");
  this.tipo.value = data.tipo;

  // Atualiza o contador
  document.getElementById(
    "contador"
  ).textContent = `Itens adicionados á O.S: ${dados.length}`;
  alert("Item salvo com sucesso! Pronto para adicionar novo item.");
});

function exportarExcel() {
  return new Promise((resolve, reject) => {
    try {
      if (dados.length === 0) {
        alert("Nenhum formulário para exportar!");
        reject(new Error("Nenhum dado para exportar"));
        return;
      }

      const workbook = XLSX.utils.book_new();

      // Cabeçalho EXATO com todas as colunas
      const dadosFormatados = [
        [
          "NOME CLIENTE", // A
          "EXECUÇÃO", // B
          "ORÇAMENTO", // C
          "RESPONSAVEL (ACOMPANHAMENTO)", // D
          "", // E (vazia)
          "SEÇÃO/SETOR", // F
          "", // G (vazia)
          "TIPO DE EQUIPAMENTO", // H
          "MARCA", // I
          "BTU", // J
          "DATA ORÇ.", // K
          "Nº MAQ", // L
          "DESCRIÇÃO DO SERVIÇO", // M
          "", // N (vazia)
          "", // O (vazia)
          "DATA DO SV", // P
          "OBSERVAÇÃO", // Q
          "", // R (vazia)
          "", // S (vazia)
          "TECNICO EQUIPE", // T
        ],
      ];

      // Adicionar os dados do formulário
      dados.forEach((item) => {
        // Processar a descrição do serviço para que cada item fique em uma linha separada
        let descricaoServico = item.descricaoServico || "MPT";
        // Se for um array, juntar com quebras de linha
        if (Array.isArray(descricaoServico)) {
          descricaoServico = descricaoServico.join("\n");
        }
        // Substituir vírgulas ou outros separadores por quebras de linha se necessário
        else if (
          typeof descricaoServico === "string" &&
          descricaoServico.includes(",")
        ) {
          descricaoServico = descricaoServico
            .split(",")
            .map((s) => s.trim())
            .join("\n");
        }

        dadosFormatados.push([
          item.Cliente, // A: NOME CLIENTE
          item.tipo === "execução" ? "X" : "", // B: EXECUÇÃO
          item.tipo === "orçamento" ? "X" : "", // C: ORÇAMENTO
          item.responsavel || "FULANO", // D: RESPONSAVEL
          "", // E: (vazia)
          item.Setor || "ALMOX", // F: SEÇÃO/SETOR
          "", // G: (vazia)
          item.tipoEquipamento || "SPLIT", // H: TIPO DE EQUIPAMENTO
          item.Marca || "PHILCO", // I: MARCA
          item.btu || "18000", // J: BTU
          item.tipo === "orçamento" ? item.dataorçamento : "", // K: DATA ORÇ.
          item.numeroMaquina || "1 OU 00", // L: Nº MAQ
          descricaoServico, // M: DESCRIÇÃO (com quebras de linha)
          "", // N: (vazia)
          "", // O: (vazia)
          item.tipo === "execução" ? item.dataorçamento : "", // P: DATA DO SV
          item.observacao || "OBS", // Q: OBSERVAÇÃO
          "", // R: (vazia)
          "", // S: (vazia)
          item.tecnicoEquipe, // T: TECNICO EQUIPE
        ]);
      });

      const worksheet = XLSX.utils.aoa_to_sheet(dadosFormatados);

      // Ajustar larguras das colunas (aumentei um pouco a coluna M para melhor visualização)
      worksheet["!cols"] = [
        { wch: 15 }, // A
        { wch: 10 }, // B
        { wch: 10 }, // C
        { wch: 25 }, // D
        { wch: 5 }, // E (vazia)
        { wch: 15 }, // F
        { wch: 5 }, // G (vazia)
        { wch: 20 }, // H
        { wch: 15 }, // I
        { wch: 10 }, // J
        { wch: 12 }, // K
        { wch: 10 }, // L
        { wch: 40 }, // M (aumentada para acomodar texto com quebras)
        { wch: 5 }, // N (vazia)
        { wch: 5 }, // O (vazia)
        { wch: 12 }, // P
        { wch: 20 }, // Q
        { wch: 5 }, // R (vazia)
        { wch: 5 }, // S (vazia)
        { wch: 15 }, // T
      ];

      // Adicionar bordas e estilo para quebrar texto
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = XLSX.utils.encode_cell({ r: R, c: C });
          if (!worksheet[cell]) worksheet[cell] = { t: "s", v: "" };

          // Configurar estilo para célula
          worksheet[cell].s = {
            border: {
              top: { style: "thin", color: { rgb: "000000" } },
              bottom: { style: "thin", color: { rgb: "000000" } },
              left: { style: "thin", color: { rgb: "000000" } },
              right: { style: "thin", color: { rgb: "000000" } },
            },
            // Aplicar quebra de texto especialmente para a coluna M (índice 12)
            alignment: {
              wrapText: C === 12, // Apenas para coluna M (DESCRIÇÃO DO SERVIÇO)
              vertical: "top", // Alinhar no topo para múltiplas linhas
            },
          };
        }
      }

      XLSX.utils.book_append_sheet(workbook, worksheet, "Ordens de Serviço");

      // Adicionar aba de assinaturas se houver assinaturas
      const hasSignatures = dados.some((item) => item.assinatura);
      if (hasSignatures) {
        const assinaturasData = [["CLIENTE", "DATA", "ASSINATURA (Base64)"]];

        dados.forEach((item) => {
          if (item.assinatura) {
            assinaturasData.push([
              item.Cliente,
              item.dataorçamento,
              item.assinatura,
            ]);
          }
        });

        const assinaturasSheet = XLSX.utils.aoa_to_sheet(assinaturasData);

        // Ajustar larguras das colunas para a aba de assinaturas
        assinaturasSheet["!cols"] = [
          { wch: 25 }, // Cliente
          { wch: 15 }, // Data
          { wch: 100 }, // Assinatura (Base64 - precisa de mais espaço)
        ];

        XLSX.utils.book_append_sheet(workbook, assinaturasSheet, "Assinaturas");
      }

      XLSX.writeFile(
        workbook,
        `ordens_servico_${new Date().toISOString().slice(0, 10)}.xlsx`
      );

      // Limpar dados após exportação
      dados.length = 0;
      document.getElementById("contador").textContent = "";

      if (hasSignatures) {
        alert(
          "Excel exportado com sucesso! As assinaturas estão na aba 'Assinaturas'."
        );
      } else {
        alert("Excel exportado com sucesso!");
      }

      resolve();
    } catch (error) {
      reject(error);
    }
  });
}

// Função para exportar PDF corrigida
async function exportarPDF(dadosParaExportar = dados) {
  return new Promise((resolve, reject) => {
    try {
      if (dadosParaExportar.length === 0) {
        alert("Nenhum formulário para exportar no PDF!");
        reject(new Error("Nenhum dado para exportar no PDF"));
        return;
      }

      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();

      // Configurações iniciais
      doc.setFont("helvetica");
      doc.setTextColor(0, 0, 0);

      // Função auxiliar para adicionar dados
      const adicionarDados = (item, yStart = 20) => {
        let y = yStart;

        // Cabeçalho
        doc.setFontSize(16);
        doc.setFont("helvetica", "bold");
        doc.text("ORDEM DE SERVIÇO - MSJR REP.", 105, y, { align: "center" });
        y += 10;

        // Linha divisória
        doc.setDrawColor(0);
        doc.setLineWidth(0.5);
        doc.line(10, y, 200, y);
        y += 10;

        // Dados principais - usando os valores diretamente do objeto item
        doc.setFontSize(12);
        const campos = [
          {
            label: "Tipo",
            value: item.tipo ? item.tipo.toUpperCase() : "NÃO INFORMADO",
          },
          {
            label: "Técnico/Equipe",
            value: item.tecnicoEquipe || "NÃO INFORMADO",
          },
          { label: "Cliente", value: item.Cliente || "NÃO INFORMADO" },
          { label: "Setor", value: item.Setor || "ALMOX" },
          {
            label: "Equipamento",
            value: `${item.tipoEquipamento || "SPLIT"} ${
              item.Marca || "PHILCO"
            } ${item.btu || "18000"}`,
          },
          { label: "Nº Máquina", value: item.numeroMaquina || "1 OU 00" },
          { label: "Data", value: item.dataorçamento || "NÃO INFORMADA" },
          { label: "Responsável", value: item.responsavel || "FULANO" },
        ];

        campos.forEach((campo) => {
          doc.setFont("helvetica", "bold");
          doc.text(`${campo.label}:`, 15, y);
          doc.setFont("helvetica", "normal");
          doc.text(campo.value, 60, y);
          y += 8;
        });

        // Descrição do serviço
        y += 5;
        doc.setFont("helvetica", "bold");
        doc.text("Descrição do Serviço:", 15, y);
        y += 7;

        // Processar descrição
        let descricao = item.descricaoServico || "MPT";
        if (Array.isArray(descricao)) {
          descricao = descricao.join("\n");
        } else if (typeof descricao === "string" && descricao.includes(",")) {
          descricao = descricao
            .split(",")
            .map((s) => s.trim())
            .join("\n");
        }

        // Adicionar descrição com quebras de linha
        const descricaoLines = doc.splitTextToSize(descricao, 180);
        doc.setFont("helvetica", "normal");
        descricaoLines.forEach((line) => {
          doc.text(line, 20, y);
          y += 7;
        });

        // Observações
        y += 7;
        doc.setFont("helvetica", "bold");
        doc.text("Observações:", 15, y);
        doc.setFont("helvetica", "normal");
        doc.text(item.observacao || "Nenhuma", 20, y + 7);
        y += 14;

        // Assinatura
        if (item.assinatura) {
          doc.setFont("helvetica", "bold");
          doc.text("Assinatura do Cliente:", 15, y);
          try {
            doc.addImage(item.assinatura, "PNG", 15, y + 5, 80, 30);
            y += 40;
          } catch (e) {
            console.error("Erro ao adicionar assinatura:", e);
            doc.text("(Assinatura não pôde ser carregada)", 15, y + 5);
            y += 15;
          }
        }

        // Rodapé
        y += 10;
        doc.setFontSize(10);
        doc.text("MSJR Representações - Tel: (21) 97956-0103", 105, y, {
          align: "center",
        });

        return y;
      };

      // Adiciona cada formulário como uma página separada
      dadosParaExportar.forEach((item, index) => {
        if (index > 0) doc.addPage();
        adicionarDados(item);
      });

      // Salva o PDF
      doc.save(`ordens_servico_${new Date().toISOString().slice(0, 10)}.pdf`);
      resolve();
    } catch (error) {
      console.error("Erro ao gerar PDF:", error);
      reject(error);
    }
  });
}
// Função para exportar ambos os formatos
async function exportarAmbos() {
  const btnExportar = document.getElementById("btnExportar");
  if (!btnExportar) return;

  // Guarda o texto original do botão
  const btnOriginalText = btnExportar.textContent;

  try {
    if (dados.length === 0) {
      alert("Nenhum formulário para exportar!");
      return;
    }

    // Atualiza o botão para estado de loading
    btnExportar.textContent = "Exportando...";
    btnExportar.disabled = true;

    // Cria uma cópia dos dados para usar no PDF
    const dadosParaPDF = [...dados];

    // Exporta Excel
    await exportarExcel();

    // Pequeno delay entre as exportações
    await new Promise((resolve) => setTimeout(resolve, 500));

    // Exporta PDF usando a cópia dos dados
    await exportarPDF(dadosParaPDF);

    // Limpa os dados apenas após ambas exportações
    dados.length = 0;
    document.getElementById("contador").textContent = "";

    alert("Arquivos exportados com sucesso!");
  } catch (error) {
    console.error("Erro na exportação:", error);
    alert(`Erro ao exportar: ${error.message}`);
  } finally {
    // SEMPRE restaura o botão, mesmo se houver erro
    btnExportar.textContent = btnOriginalText;
    btnExportar.disabled = false;
  }
}

// Configuração do botão de exportação
document.addEventListener("DOMContentLoaded", function () {
  const btnExportar = document.getElementById("btnExportar");
  if (btnExportar) {
    btnExportar.addEventListener("click", exportarAmbos);
  }
});

// Código da assinatura (mantido igual)
const canvas = document.getElementById("assinatura");
if (canvas) {
  const ctx = canvas.getContext("2d");
  let desenhando = false;

  // Eventos de desenho (mantidos iguais)
  canvas.addEventListener("mousedown", () => (desenhando = true));
  canvas.addEventListener("mouseup", () => (desenhando = false));
  canvas.addEventListener("mouseout", () => (desenhando = false));
  canvas.addEventListener("mousemove", desenhar);

  // Versão mobile
  canvas.addEventListener("touchstart", (e) => {
    e.preventDefault();
    desenhando = true;
  });
  canvas.addEventListener("touchend", () => (desenhando = false));
  canvas.addEventListener("touchmove", desenharMobile);

  function desenhar(e) {
    if (!desenhando) return;
    ctx.lineWidth = 2;
    ctx.lineCap = "round";
    ctx.strokeStyle = "#000";

    const rect = canvas.getBoundingClientRect();
    ctx.lineTo(e.clientX - rect.left, e.clientY - rect.top);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(e.clientX - rect.left, e.clientY - rect.top);
  }

  function desenharMobile(e) {
    if (!desenhando) return;
    e.preventDefault();
    const touch = e.touches[0];
    const rect = canvas.getBoundingClientRect();
    const x = touch.clientX - rect.left;
    const y = touch.clientY - rect.top;

    ctx.lineWidth = 2;
    ctx.lineCap = "round";
    ctx.strokeStyle = "#000";

    ctx.lineTo(x, y);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(x, y);
  }

  window.limparAssinatura = function () {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.beginPath();
  };
}

// Código dos serviços (mantido igual)
const select = document.getElementById("selectServico");
const checkboxesContainer = document.getElementById("checkboxesContainer");

if (select && checkboxesContainer) {
  select.addEventListener("change", (event) => {
    const valor = event.target.value;
    if (!valor || document.querySelector(`input[value="${valor}"]`)) return;

    const container = document.createElement("div");
    container.style.display = "flex";
    container.style.alignItems = "center";
    container.style.margin = "5px 0";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "descricaoServico";
    checkbox.value = valor;
    checkbox.checked = true;
    checkbox.style.marginRight = "10px";
    checkbox.style.width = "auto";

    checkbox.addEventListener("change", function () {
      if (!this.checked) container.remove();
    });

    const texto = document.createElement("span");
    texto.textContent = valor;
    texto.style.flexGrow = "1";

    container.appendChild(checkbox);
    container.appendChild(texto);
    checkboxesContainer.appendChild(container);
  });
}

// Opções para cada tipo de serviço
const opcoesServico = {
  preventiva: [
    "MANUTENÇÃO PREVENTIVA DA CONDENSADORA",
    "MANUTENÇÃO PREVENTIVA DA EVAPORADORA",
    "MANUTENÇÃO PREVENTIVA - LIMPEZA DE FILTROS",
  ],
  corretiva: [
    "REVISÃO ELÉTRICA",
    "RECARGA DE GÁS 1KG",
    "RECARGA DE GÁS 2KGS",
    "RECARGA DE GÁS 3KGS",
    "RECARGA DE GÁS 4KGS",
    "RECARGA DE GÁS 5KGS",
    "RECARGA DE GÁS 6KGS",
    "RECARGA DE GÁS 7KGS",
    "RECARGA DE GÁS 8KGS",
    "RECARGA DE GÁS 9KGS",
    "RECARGA DE GÁS 10KG",
    "TROCA DE CAPACITOR",
    "TROCA DE TERMOSTATO",
    "TROCA DE SENSOR DE TEMPERATURA",
    "TROCA DE CONTROLE REMOTO",
    "TROCA DA PLACA ELETRONICA RECEPTORA",
    "TROCA DA PLACA ELETRONICA DA CONDENSADORA",
    "TROCA DA PLACA ELETRONICA DA EVAPORADORA",
    "TROCA DA PLACA ELETRONICA UNIVERSAL",
    "TROCA DO COMPRESSOR",
    "TROCA DE MOTOR DE VENTILADOR DA EVAPORADORA",
    "TROCA DE MOTOR DE VENTILADOR DA CONDENSADORA",
    "TROCA DE SERPENTINA DA CONDENSADORA",
    "TROCA DE SERPENTINA DA EVAPORADORA",
    "TROCA DO CONJUNTO DE FILTROS",
    "TROCA DAS ALETAS",
    "TROCA DA VÁLVULA DE EXPANSÃO",
    "TROCA DA HÉLICE",
    "BOMBA DE DRENO",
    "INSTALAÇÃO DE AR CONDICIONADO (DISTÂNCIA VIDE OBS)",
    "DESISTALAÇÃO DO APARELHO",
    "TROCA DA CHAVE INVERSORA",
    "TROCA DO FILTRO SECADOR",
    "TROCA DO GABINETE",
    "TROCA DO MOTOR DA ALETA",
    "TROCA DE RELÊ",
    "TROCA DA TURBINA",
    "TROCA DO MOTOR DA TURBINA",
    "TROCA DA GRADE DA CONDENSADORA",
    "TROCA DA CALHA",
    "TROCA DA CHAVE TERMOSTÁTICA",
    "TROCA DA CÂMARA DO VENTILADOR",
    "TROCA DA MANGUEIRA DO DRENO",
    "TROCA DO PAINEL",
    "TROCA DO SENSOR DE DEGELO",
    "TROCA DO DUTO DE AR",
    "TROCA DO BOTÃO",
    "TROCA DOS CALCOS DE BORRACHA",
    "TROCA DO SUPORTE DA CONDENSADORA",
    "TROCA DO SUPORTE DA EVAPORADORA",
    "TROCA DO CAPILAR",
    "TROCA DAS CONEXÕES DE COBRE",
    "TROCA DO PISTÃO",
    "TROCA DA CHAVE DE FLUXO",
    "TROCA DE CONTROLE DE TEMPERATURA",
    "TROCA DA BOBINA SOLENÓIDE",
    "TROCA DA CHAVE SELETORA",
    "TROCA DO PRESSOSTATO",
    "TROCA DA VÁLVULA DE SERVIÇO",
    "TROCA DA CONTATORA",
    "TROCA DA BOBINA DA VÁLVULA",
    "TROCA DO PAINEL DE CONTROLE",
    "APLICAÇÃO DE NITROGÊNIO 1KG",
    "APLICAÇÃO DE NITROGÊNIO 2KG",
    "APLICAÇÃO DE NITROGÊNIO 3KG",
    "SOLDA PARA SANAR VAZAMENTO",
    "APERTO DAS PORCAS",
    "DESMONTAGEM PARCIAL DA EVAPORADORA",
    "DESMONTAGEM PARCIAL DA CONDENSADORA",
    "OUTROS (VIDE OBS)",
  ],
};

// Função para carregar as opções no select
function carregarOpcoesServico(tipo) {
  const select = document.getElementById("selectServico");
  select.innerHTML = ""; // Limpa as opções atuais

  // Adiciona um placeholder
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Selecione um serviço";
  placeholder.disabled = true;
  placeholder.selected = true;
  select.appendChild(placeholder);

  // Adiciona as opções do tipo selecionado
  opcoesServico[tipo].forEach((opcao) => {
    const option = document.createElement("option");
    option.value = opcao;
    option.textContent = opcao;
    select.appendChild(option);
  });
}

// Event listener para os radios
document.querySelectorAll('input[name="tipoServico"]').forEach((radio) => {
  radio.addEventListener("change", function () {
    carregarOpcoesServico(this.value);
    // Limpa os checkboxes quando mudar o tipo
    document.getElementById("checkboxesContainer").innerHTML = "";
  });
});

// Inicializa com as opções preventivas (padrão)
carregarOpcoesServico("preventiva");

// Mantenha o resto do seu código para adicionar os checkboxes quando selecionar
const select1 = document.getElementById("selectServico");
const checkboxesContainer1 = document.getElementById("checkboxesContainer");

if (select1 && checkboxesContainer1) {
  select1.addEventListener("change", (event) => {
    const valor = event.target.value;
    if (!valor || document.querySelector(`input[value="${valor}"]`)) return;

    const container = document.createElement("div");
    container.style.display = "flex";
    container.style.alignItems = "center";
    container.style.margin = "5px 0";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.name = "descricaoServico";
    checkbox.value = valor;
    checkbox.checked = true;
    checkbox.style.marginRight = "10px";
    checkbox.style.width = "auto";

    checkbox.addEventListener("change", function () {
      if (!this.checked) container.remove();
    });

    const texto = document.createElement("span");
    texto.textContent = valor;
    texto.style.flexGrow = "1";

    container.appendChild(checkbox);
    container.appendChild(texto);
    checkboxesContainer.appendChild(container);
  });
}
