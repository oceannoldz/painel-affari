// === CONFIGURAÇÕES ===
const API_URL_JSON =
  "https://script.google.com/macros/s/AKfycbxdoKGM2guipchqe8R2F9YAQWVEc-GVv-1fyr6EDwXCk3U8T60DK50XU9Wsvtb1d17iIA/exec"; // endpoint do Apps Script
const UPDATE_INTERVAL = 30 * 1000; // atualiza a cada 30 segundos

// === FORMATADORES ===
function excelSerialToDate(serial) {
  if (!serial) return "";
  if (typeof serial === "string") return serial.trim();

  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial) + 0.0000001;
  const total_seconds = Math.floor(86400 * fractional_day);
  const hours = Math.floor(total_seconds / 3600);
  const minutes = Math.floor((total_seconds % 3600) / 60);

  const d = String(date_info.getUTCDate()).padStart(2, "0");
  const m = String(date_info.getUTCMonth() + 1).padStart(2, "0");
  const y = date_info.getUTCFullYear();
  const diasSemana = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  const diaSemana = diasSemana[date_info.getUTCDay()];

  if (serial > 30000 && serial % 1 === 0) return `${diaSemana}, ${d}/${m}/${y}`;
  if (serial > 30000 && serial % 1 !== 0)
    return `${diaSemana}, ${d}/${m}/${y} – ${String(hours).padStart(
      2,
      "0"
    )}:${String(minutes).padStart(2, "0")}`;
  if (serial < 1)
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(
      2,
      "0"
    )}`;
  return `${d}/${m}/${y}`;
}

function normalizarHora(h) {
  if (!h) return "";
  const s = String(h).trim();
  const m = s.match(/(\d{1,2})[:\.](\d{2})/);
  if (m) return `${m[1].padStart(2, "0")}:${m[2]}`;
  return s;
}

function apenasData(str) {
  if (!str) return "";
  const m = String(str).match(/(\d{2}\/\d{2}\/\d{4})/);
  return m ? m[1] : String(str).trim();
}

// === CARREGAMENTO DE DADOS ===
async function carregarDados() {
  try {
    const response = await fetch(`${API_URL_JSON}?t=${Date.now()}`, {
      cache: "no-store",
    });
    if (!response.ok) throw new Error("Erro ao buscar dados da planilha.");
    const dados = await response.json();

    console.log(`✅ Dados recebidos: ${dados.length} registros`);
    renderizarCards(dados);
  } catch (erro) {
    console.error("Erro:", erro);
    document.getElementById("cardsContainer").innerHTML = `
      <p style="color:white; text-align:center;">⚠️ ${erro.message}</p>`;
  }
}

// === RENDERIZAÇÃO DOS CARDS ===
function renderizarCards(dados) {
  const container = document.getElementById("cardsContainer");
  const contador = document.getElementById("contadorStatus");
  if (!container) return;

  container.innerHTML = "";
  let pendentes = 0;
  let entregues = 0;

  // Ordena por hora
  dados.sort((a, b) => {
    const horaA = normalizarHora(a["Horário"] || a.Hora || "");
    const horaB = normalizarHora(b["Horário"] || b.Hora || "");
    return horaA.localeCompare(horaB);
  });

  dados.forEach((linha) => {
    const secretaria = linha.Secretaria?.trim() || "(Sem Secretaria)";
    const dataFmt = excelSerialToDate(linha.Data);
    const soData = apenasData(dataFmt) || String(linha.Data || "").trim();
    const horaFmt = excelSerialToDate(linha["Horário"] || linha.Hora);
    const horaHHMM = normalizarHora(horaFmt || linha["Horário"] || linha.Hora);

    const local = (linha.Local || "").trim();
    const diretor = (linha.Diretor || "").trim();
    const itens = linha.Itens || "";
    const observacoes = linha.Observações || linha.Observacoes || "";
    const statusEntregaPlanilha = linha.Status_Entrega || "";
    const pax = linha.Pax || 0;

    let estado = statusEntregaPlanilha || "Pendente";

    if (estado.toLowerCase() === "entregue") {
      entregues++;
      return;
    } else {
      pendentes++;
    }

    // Cria card visual
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <img src="download.png" alt="logo">
      <h2>${secretaria}</h2>
      <div class="info">
        <strong>Diretor:</strong> ${diretor || "-"}<br>
        <strong>Local:</strong> ${local || "-"}<br>
        <strong>Data:</strong> ${dataFmt}<br>
        <strong>Hora:</strong> ${horaHHMM}<br>
        <strong>Pax:</strong> ${pax}
      </div>
      <div class="itens"><strong>Itens:</strong><br>${itens || "-"}</div>
      <div class="observacoes"><strong>Observações:</strong><br>${
        observacoes || "-"
      }</div>
      <div class="status"></div>
    `;

    const statusEl = card.querySelector(".status");
    function atualizarVisual() {
      statusEl.textContent = estado;
      card.classList.remove("pendente", "entregue");
      card.classList.add(estado.toLowerCase());
    }

    // Clique → marcar como entregue
    card.addEventListener("click", async () => {
      if (estado.toLowerCase() !== "pendente") return;

      const confirmar = confirm("Marcar este item como ENTREGUE?");
      if (!confirmar) return;

      estado = "Entregue";
      atualizarVisual();
      card.style.transition = "opacity 0.4s ease";
      card.style.opacity = "0";
      setTimeout(() => card.remove(), 400);

      try {
        const body = {
  Data: soData, // ex: "03/11/2025"
  Hora: horaHHMM, // ex: "08:00"
  Local: local.trim(),
  Diretor: diretor.trim(),
  Status: "Entregue",
};


        await fetch(API_URL_JSON, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });

        mostrarNotificacao("✅ Marcado como entregue!");
        setTimeout(carregarDados, 1000);
      } catch (erro) {
        console.error("Erro ao enviar:", erro);
        mostrarNotificacao("⚠️ Erro ao atualizar!");
      }
    });

    atualizarVisual();
    container.appendChild(card);
  });

  // Contadores e info
  contador.innerHTML = `
    <span style="color:#ff4d4d;">🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp;
    <span style="color:#00b050;">🟢 Entregues: ${entregues}</span>
  `;
  document.getElementById("updateInfo").textContent =
    `Última atualização: ${new Date().toLocaleTimeString()} — Atualizando a cada 30 segundos`;
}

// === NOTIFICAÇÃO VISUAL ===
function mostrarNotificacao(texto) {
  const alerta = document.createElement("div");
  alerta.textContent = texto;
  Object.assign(alerta.style, {
    position: "fixed",
    bottom: "20px",
    right: "20px",
    background: "#333",
    color: "#fff",
    padding: "10px 16px",
    borderRadius: "8px",
    boxShadow: "0 2px 6px rgba(0,0,0,.3)",
    zIndex: 9999,
    fontSize: "13px",
    opacity: 0,
    transition: "opacity .3s",
  });
  document.body.appendChild(alerta);
  requestAnimationFrame(() => (alerta.style.opacity = "1"));
  setTimeout(() => {
    alerta.style.opacity = "0";
    setTimeout(() => alerta.remove(), 300);
  }, 2500);
}

// === INICIALIZAÇÃO ===
document.addEventListener("DOMContentLoaded", () => {
  carregarDados();
  setInterval(carregarDados, UPDATE_INTERVAL);
});

