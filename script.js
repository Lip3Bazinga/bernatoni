document.addEventListener("DOMContentLoaded", () => {
  const uploader = document.getElementById("uploader");
  const csvFileInput = document.getElementById("csv-file");
  const statusText = document.getElementById("status-text");
  const spinner = document.getElementById("spinner");
  const gerarBtn = document.getElementById("gerar-etiquetas");
  const exportarBtn = document.getElementById("exportar-pdf");
  const outputArea = document.getElementById("output-area");
  const templateEtiqueta = document.getElementById("template-etiqueta");
  const brandToggle = document.getElementById("brand-toggle");

  let produtosData = [];
  let selectedFile = null;

  uploader.addEventListener("click", () => csvFileInput.click());
  csvFileInput.setAttribute("accept", ".xlsx");

  csvFileInput.addEventListener("change", () => {
    if (csvFileInput.files.length > 0) {
      handleXlsxFile(csvFileInput.files[0]);
    }
  });

  ["dragenter", "dragover", "dragleave", "drop"].forEach(eventName =>
    uploader.addEventListener(eventName, preventDefaults, false)
  );
  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }
  uploader.addEventListener("dragenter", () => uploader.classList.add("dragover"));
  uploader.addEventListener("dragleave", () => uploader.classList.remove("dragover"));
  uploader.addEventListener("drop", e => {
    uploader.classList.remove("dragover");
    handleXlsxFile(e.dataTransfer.files[0]);
  });

  function handleXlsxFile(file) {
    if (file && file.name.toLowerCase().endsWith(".xlsx")) {
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        produtosData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        statusText.textContent = `Arquivo selecionado: ${file.name}`;
        gerarBtn.disabled = false;
        exportarBtn.hidden = true;
        outputArea.innerHTML = "";
        selectedFile = file;
      };
      reader.readAsArrayBuffer(file);
    } else {
      statusText.textContent = "Formato inválido. Por favor, selecione um arquivo .XLSX";
      selectedFile = null;
      gerarBtn.disabled = true;
    }
  }

  gerarBtn.addEventListener("click", () => {
    if (!produtosData.length) return;
    gerarBtn.disabled = true;
    spinner.hidden = false;
    statusText.textContent = "Processando arquivo...";
    renderizarEtiquetas();
    spinner.hidden = true;
    statusText.textContent = `${produtosData.length} etiquetas geradas com sucesso!`;
    exportarBtn.hidden = false;
    gerarBtn.disabled = false;
  });

  exportarBtn.addEventListener("click", () => {
    const etiquetasVisiveis = document.querySelectorAll("#output-area .etiqueta-wrapper");
    if (etiquetasVisiveis.length === 0) {
      alert("Nenhuma etiqueta para exportar.");
      return;
    }

    exportarBtn.disabled = true;
    spinner.hidden = false;
    statusText.textContent = "Gerando PDF de alta qualidade...";

    const { jsPDF } = window.jspdf;
    const labelWidthMM = 84.6; // Equivalente a 320px a 96dpi
    const labelHeightMM = 47.6; // Equivalente a 180px a 96dpi
    const doc = new jsPDF({ orientation: "l", unit: "mm", format: [labelWidthMM, labelHeightMM] });

    const isMasculino = brandToggle.checked;
    const brandName = isMasculino ? "BERNATONI" : "VIOLANTA";
    const FONT_FAMILY = "Helvetica";

    produtosData.forEach((produto, i) => {
      // Validação de dados essenciais
      if (!produto["Código do produto"] || !produto["Opção de estoque"])
        return;

      if (i > 0) {
        doc.addPage([labelWidthMM, labelHeightMM], "l");
      }

      const margin = 4;
      const codigoProduto = produto["Código do produto"];
      doc.setFont(FONT_FAMILY, "bold");
      doc.setFontSize(19.5);
      doc.text(String(codigoProduto), margin, margin + 7);

      // --- Caixa de Tamanho (canto inferior direito) ---
      const boxSize = 21.1;
      const boxBorder = 0.8;
      const boxX = labelWidthMM - boxSize - margin;
      const boxY = labelHeightMM - boxSize - margin;

      doc.setLineWidth(boxBorder);
      doc.rect(boxX, boxY, boxSize, boxSize, "S"); // "S" para apenas contorno

      // --- Texto dentro da Caixa de Tamanho ---
      const tamanho = produto["Opção de estoque"];
      doc.setFont(FONT_FAMILY, "bold");
      doc.setFontSize(36);
      const tamanhoWidth = doc.getTextWidth(String(tamanho));
      doc.text(String(tamanho), boxX + (boxSize - tamanhoWidth) / 2, boxY + 14.5);

      doc.setFontSize(7.5);
      const brandNameWidth = doc.getTextWidth(brandName);
      doc.text(brandName, boxX + (boxSize - brandNameWidth) / 2, boxY + 18);

      // --- Bloco de Informações (canto inferior esquerdo) ---
      let currentY = 22;
      doc.setFont(FONT_FAMILY, "normal");
      doc.setFontSize(10.5);

      doc.text(`MODELO: ${tamanho}`, margin, currentY); currentY += 4.5;
      doc.text("FICHA: --", margin, currentY); currentY += 4.5;
      doc.text("PART.: --", margin, currentY);

      // --- Código de Barras com removeAcentos ---
      const campo1 = produto["Código do produto"] || "";
      const campo2 = produto["Descrição"] ? removeAcentos(produto["Descrição"]) : "";
      const campo3 = produto["Opção de estoque"] || "";
      const strCodigoBarra = `${campo1} ${campo2} ${campo3}`.trim();

      const tempCanvas = document.createElement("canvas");
      try {
        JsBarcode(tempCanvas, strCodigoBarra, {
          format: "CODE128",
          displayValue: false,
          width: 4,
          height: 80,
          margin: 0,
        });

        const barcodeImgData = tempCanvas.toDataURL("image/png");
        const barcodeHeightMM = 10;
        const barcodeWidthMM = 45;

        doc.addImage(
          barcodeImgData,
          "PNG",
          margin,
          currentY + 2,
          barcodeWidthMM,
          barcodeHeightMM
        );
      } catch (e) {
        console.error("Erro ao gerar o código de barras:", e);
        doc.text("Erro no barcode", margin, currentY + 6);
      }
    });

    doc.save("etiquetas_produtos.pdf");
    spinner.hidden = true;
    statusText.textContent = "PDF de alta qualidade exportado com sucesso!";
    exportarBtn.disabled = false;
  });

  brandToggle.addEventListener("change", () => {
    if (produtosData.length > 0) {
      renderizarEtiquetas();
    }
  });

  function removeAcentos(str) {
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  function renderizarEtiquetas() {
    outputArea.innerHTML = "";
    const isMasculino = brandToggle.checked;
    const brandName = isMasculino ? "BERNATONI" : "VIOLANTA";
    for (const produto of produtosData) {
      if (!produto["Código do produto"] || !produto["Opção de estoque"])
        continue;
      const novaEtiqueta = templateEtiqueta.cloneNode(true);
      novaEtiqueta.id = "";
      novaEtiqueta.style.position = "static";
      novaEtiqueta.style.visibility = "visible";
      novaEtiqueta.querySelector(".numeracao-nome").textContent = produto["Código do produto"];
      novaEtiqueta.querySelector(".modelo-value").textContent = produto["Opção de estoque"];
      novaEtiqueta.querySelector(".tamanho-value").textContent = produto["Opção de estoque"];
      novaEtiqueta.querySelector(".brand-name").textContent = brandName;

      outputArea.appendChild(novaEtiqueta);

      // Monta a string para o barcode igual exportação
      const campo1 = produto["Código do produto"] || "";
      const campo2 = produto["Descrição"] ? removeAcentos(produto["Descrição"]) : "";
      const campo3 = produto["Opção de estoque"] || "";
      const strCodigoBarra = `${campo1} ${campo2} ${campo3}`.trim();

      try {
        JsBarcode(
          novaEtiqueta.querySelector(".barcode-svg"),
          strCodigoBarra,
          {
            format: "CODE128",
            displayValue: false,
            width: 2,
            height: 40,
            margin: 0,
          }
        );
      } catch (e) {
        console.error("Erro no JsBarcode:", e);
      }
    }
  }
});
