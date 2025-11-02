// ===========================
// Painel de Demandas - Affari
// ===========================

const PLANILHA = "painel-affari.xlsx";
const INTERVALO = 10 * 1000; // 10 segundos

// -------- Função para converter número Excel em data/hora legível --------
function excelSerialToDate(serial) {
  if (serial == null || serial === "") return "";

  // Se já for texto (ex: "02/11/2025"), apenas retorna
  if (typeof serial === "string") return serial;

  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  let total_seconds = Math.floor(86400 * fractional_day);

  const seconds = total_seconds % 60;
  total_seconds -= seconds;
  const hours = Math.floor(total_seconds / (60 * 60));
  const minutes = Math.floor(total_seconds / 60) % 60;

  // Formatar data
  const dia = date_info.getUTCDate().toString().padStart(2, "0");
  const mes = (date_info.getUTCMonth() + 1).toString().padStart(2, "0");
  const ano = date_info.getUTCFullYear();

  // Traduzir dia da semana
  const diasSemana = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  const diaSemana = diasSemana[date_info.getUTCDay()];

  // Se for apenas data
  if (serial > 30000 && serial % 1 === 0)
    return `${diaSemana}, ${dia}/${mes}/${ano}`;

  // Se for data + hora
  if (serial > 30000 && serial % 1 !== 0)
    return `${diaSemana}, ${dia}/${mes}/${ano} – ${hours
      .toString()
      .padStart(2, "0")}:${minutes.toString().padStart(2, "0")}`;

  // Se for apenas hora
  if (serial < 1)
    return `${hours.toString().padStart(2, "0")}:${minutes
      .toString()
      .padStart(2, "0")}`;

  return `${dia}/${mes}/${ano}`;
}

// -------- Função principal para ler a planilha --------
async function carregarPlanilha() {
  try {
    const response = await fetch(PLANILHA);
    if (!response.ok) throw new Error("Erro ao carregar a planilha");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const primeiraAba = workbook.SheetNames[0];
    const planilha = workbook.Sheets[primeiraAba];
    const dados = XLSX.utils.sheet_to_json(planilha);
    renderizarCards(dados);
  } catch (erro) {
    console.error("Erro ao ler planilha:", erro);
    document.getElementById("cardsContainer").innerHTML =
      "<p style='color:white'>Erro ao carregar planilha. Verifique se o servidor está rodando e o arquivo existe.</p>";
  }
}

// -------- Função que cria e exibe os cards --------
function renderizarCards(dados) {
  const container = document.getElementById("cardsContainer");
  container.innerHTML = "";

  // Recupera status salvos no navegador
  const statusSalvos = JSON.parse(localStorage.getItem("statusCards") || "{}");

  dados.forEach((linha) => {
    const card = document.createElement("div");
    card.className = "card";

    const secretaria = linha.Secretaria || "";
    const idUnico =
      secretaria.trim() || Math.random().toString(36).substring(2);

    const data = excelSerialToDate(linha.Data);
    const horario = excelSerialToDate(linha.Horário || linha.Hora);
    const local = linha.Local || "";
    const diretor = linha.Diretor || "";
    const itens = linha.Itens || "";
    const observacoes = linha.Observações || "";
    const pax = linha.Pax || 0;

    // Recupera status salvo ou padrão
    let estado = statusSalvos[idUnico] || linha.Status || "Pendente";

    card.innerHTML = `
      <img src="download.png" alt="Affari logo">
      <h2>${secretaria}</h2>
      <div class="info">
        <strong>Diretor:</strong> ${diretor}<br>
        <strong>Local:</strong> ${local}<br>
        <strong>Data:</strong> ${data}<br>
        <strong>Hora:</strong> ${horario}<br>
        <strong>Pax:</strong> ${pax}
      </div>
      <div class="itens"><strong>Itens:</strong><br>${itens}</div>
      <div class="observacoes"><strong>Observações:</strong><br>${observacoes}</div>
      <div class="status"></div>
    `;

    const statusEl = card.querySelector(".status");

    // Atualiza aparência conforme status
    function atualizarVisual() {
      statusEl.textContent = estado;
      card.classList.remove("pendente", "Entregue", "retirado");

      switch (estado.toLowerCase()) {
        case "pendente":
          card.classList.add("pendente");
          break;
        case "enviado":
          card.classList.add("Entregue");
          break;
        case "retirado":
          card.classList.add("retirado");
          break;
      }
    }

    // Clique alterna status e salva no localStorage
    card.addEventListener("click", () => {
      if (estado.toLowerCase() === "pendente") estado = "Entregue";
      else if (estado.toLowerCase() === "enviado") estado = "Retirado";
      else estado = "Pendente";

      statusSalvos[idUnico] = estado;
      localStorage.setItem("statusCards", JSON.stringify(statusSalvos));
      atualizarVisual();
    });

    atualizarVisual();
    container.appendChild(card);
  });

  const agora = new Date();
  document.getElementById(
    "updateInfo"
  ).textContent = `Última atualização: ${agora.toLocaleTimeString()} — Atualizando automaticamente a cada 10 minutos`;
}

// -------- Inicialização --------
carregarPlanilha();
setInterval(carregarPlanilha, INTERVALO);

