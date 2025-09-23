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
  uploader.addEventListener("dragenter", () =>
    uploader.classList.add("dragover")
  );
  uploader.addEventListener("dragleave", () =>
    uploader.classList.remove("dragover")
  );
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
        produtosData = XLSX.utils.sheet_to_json(worksheet, {
          defval: "",
        });
        statusText.textContent = `Arquivo selecionado: ${file.name}`;
        gerarBtn.disabled = false;
        exportarBtn.hidden = true;
        outputArea.innerHTML = "";
        selectedFile = file;
      };
      reader.readAsArrayBuffer(file);
    } else {
      statusText.textContent =
        "Formato inválido. Por favor, selecione um arquivo .XLSX";
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

  exportarBtn.addEventListener("click", async () => {

    const etiquetasVisiveis = document.querySelectorAll("#output-area .etiqueta-wrapper");

    if (etiquetasVisiveis.length === 0) {
      alert("Nenhuma etiqueta para exportar.");
      return;
    }

    exportarBtn.disabled = true;
    spinner.hidden = false;
    statusText.textContent = "Gerando PDF (1 por página tamanho da etiqueta)...";

    const { jsPDF } = window.jspdf;

    let doc = null;

    for (let i = 0; i < etiquetasVisiveis.length; i++) {
      const etiquetaElement = etiquetasVisiveis[i];

      // Renderiza a etiqueta como imagem
      const canvas = await html2canvas(etiquetaElement, {
        scale: 3,
        allowTaint: true,
        useCORS: true,
      });
      const imgData = canvas.toDataURL("image/png");

      // Calcula largura/altura do canvas em mm
      const dpi = 96; // padrão web
      const canvasWidthMM = (canvas.width / dpi) * 25.4;
      const canvasHeightMM = (canvas.height / dpi) * 25.4;

      // Cria um PDF com o tamanho igual ao da etiqueta
      if (i === 0) {
        doc = new jsPDF({
          orientation: "p",
          unit: "mm",
          format: [canvasWidthMM, canvasHeightMM],
        });
      } else {
        doc.addPage([canvasWidthMM, canvasHeightMM], 'p');
      }

      doc.addImage(imgData, "PNG", 0, 0, canvasWidthMM, canvasHeightMM);
    }

    doc.save("etiquetas_produtos_individual.pdf");
    spinner.hidden = true;
    statusText.textContent = "PDF exportado com sucesso!";
    exportarBtn.disabled = false;
  });

  brandToggle.addEventListener("change", () => {
    if (produtosData.length > 0) {
      renderizarEtiquetas();
    }
  });

  function renderizarEtiquetas() {
    outputArea.innerHTML = "";
    const isMasculino = brandToggle.checked;
    const brandName = isMasculino ? "BERNATONI" : "VIOLANTA";
    for (const produto of produtosData) {
      if (
        !produto["Código do produto"] ||
        !produto["Opção de estoque"] ||
        !produto["Número do pedido"]
      )
        continue;
      const novaEtiqueta = templateEtiqueta.cloneNode(true);
      novaEtiqueta.id = "";
      novaEtiqueta.style.position = "static";
      novaEtiqueta.style.visibility = "visible";
      novaEtiqueta.querySelector(".numeracao-nome").textContent =
        produto["Código do produto"];
      novaEtiqueta.querySelector(".modelo-value").textContent =
        produto["Opção de estoque"];
      novaEtiqueta.querySelector(".tamanho-value").textContent =
        produto["Opção de estoque"];
      novaEtiqueta.querySelector(".brand-name").textContent = brandName;
      outputArea.appendChild(novaEtiqueta);
      try {
        JsBarcode(
          novaEtiqueta.querySelector(".barcode-svg"),
          String(produto["Número do pedido"]),
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