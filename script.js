const PLANILHA =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vSR3FyZKCXFS5Mi4UaRc6GLCSfH0erH_rraD87M0ZFo6jeDT0hEnpvUfEH2-cxXI0-ionFDxLFFUuvg/pub?output=csv";

const API_URL =
  "https://script.google.com/macros/s/AKfycbzBhaHbGAnNC5u-D3sSKqDWVsGVUzyDWfLsRqzpiHEuyOFVxsRtIytL_rv4dkxynBdorw/exec";

const INTERVALO = 60 * 1000; // 1 minuto

// 🕒 Converte formato do Excel para data legível
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

// 🕓 Normaliza hora (ex: "8.3" -> "08:30")
function normalizarHora(h) {
  if (!h) return "";
  const s = String(h).trim();
  const m = s.match(/(\d{1,2})[:\.](\d{2})/);
  if (m) return `${m[1].padStart(2, "0")}:${m[2]}`;
  return s;
}

// 📅 Extrai apenas data (dd/mm/yyyy)
function apenasData(str) {
  if (!str) return "";
  const m = String(str).match(/(\d{2}\/\d{2}\/\d{4})/);
  return m ? m[1] : String(str).trim();
}

// 🔄 Carrega planilha pública e renderiza cards
async function carregarPlanilha() {
  try {
    const url = PLANILHA + "&t=" + new Date().getTime();
    const response = await fetch(url);
    if (!response.ok) throw new Error("Erro ao carregar planilha Google.");

    const texto = (await response.text()).replace(/^\uFEFF/, "").trim();

    if (texto.startsWith("<"))
      throw new Error("⚠️ A planilha não está publicada como CSV.");

    const primeiraLinha = texto.split("\n")[0];
    const delimitador = primeiraLinha.includes(";") ? ";" : ",";
    const parsed = Papa.parse(texto, {
      header: true,
      skipEmptyLines: true,
      delimiter: delimitador,
    });

    const dados = parsed.data;
    if (!dados.length) throw new Error("Planilha vazia.");

    renderizarCards(dados);
  } catch (erro) {
    console.error("Erro ao ler planilha:", erro);
    document.getElementById("cardsContainer").innerHTML = `
      <p style="color:white; text-align:center;">⚠️ ${erro.message}</p>`;
  }
}

// 🧩 Monta os cards no front-end
function renderizarCards(dados) {
  const container = document.getElementById("cardsContainer");
  const contador = document.getElementById("contadorStatus");
  if (!container) return;

  container.innerHTML = "";
  let pendentes = 0;
  let entregues = 0;

  dados.sort((a, b) => {
    const horaA = normalizarHora(a["Horário"] || "");
    const horaB = normalizarHora(b["Horário"] || "");
    return horaA.localeCompare(horaB);
  });

  dados.forEach((linha) => {
    const secretaria = linha.Secretaria?.trim() || "(Sem Secretaria)";
    const dataFmt = excelSerialToDate(linha.Data);
    const soData = apenasData(dataFmt);
    const horaFmt = excelSerialToDate(linha["Horário"]);
    const horaHHMM = normalizarHora(horaFmt || linha["Horário"]);

    const local = (linha.Local || "").trim();
    const diretor = (linha.Diretor || "").trim();
    const itens = linha.Itens || "";
    const observacoes = linha.Observações || linha.Observacoes || "";
    const statusEntregaPlanilha = (linha.Status_Entrega || "").trim();
    const pax = linha.Pax || 0;

    let estado = statusEntregaPlanilha || "Pendente";

    if (estado.toLowerCase() === "entregue") {
      entregues++;
      return; // Não mostra entregues
    } else {
      pendentes++;
    }

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <img src="download.png" alt="logo">
      <h2>${secretaria}</h2>
      <div class="info">
        <strong>Diretor:</strong> ${diretor}<br>
        <strong>Local:</strong> ${local}<br>
        <strong>Data:</strong> ${dataFmt}<br>
        <strong>Hora:</strong> ${horaHHMM}<br>
        <strong>Pax:</strong> ${pax}
      </div>
      <div class="itens"><strong>Itens:</strong><br>${itens}</div>
      <div class="observacoes"><strong>Observações:</strong><br>${observacoes}</div>
      <div class="status">${estado}</div>
    `;

    // 🖱️ Clique para marcar como entregue
    card.addEventListener("click", async () => {
      if (estado.toLowerCase() !== "pendente") return;

      const confirmar = confirm(
        "Tem certeza que deseja marcar este item como ENTREGUE?"
      );
      if (!confirmar) return;

      estado = "Entregue";
      card.querySelector(".status").textContent = estado;
      card.style.transition = "opacity 0.4s ease";
      card.style.opacity = "0";

      setTimeout(() => card.remove(), 400);

      try {
        const body = {
          Data: soData,
          Hora: horaHHMM,
          Local: local,
          Diretor: diretor,
          Status: "Entregue",
        };

        const resposta = await fetch(API_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });

        const resultado = await resposta.json();
        console.log(resultado.mensagem);

        mostrarNotificacao(resultado.mensagem);
        setTimeout(carregarPlanilha, 1500);
      } catch (erro) {
        console.error("Erro ao enviar atualização:", erro);
        mostrarNotificacao("⚠️ Falha ao atualizar na planilha!");
      }
    });

    container.appendChild(card);
  });

  contador.innerHTML = `
    <span style="color:#ff4d4d;">🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp;
    <span style="color:#00b050;">🟢 Entregues: ${entregues}</span>
  `;

  document.getElementById(
    "updateInfo"
  ).textContent = `Última atualização: ${new Date().toLocaleTimeString()}`;
}

// 🔔 Notificação visual
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

// 🚀 Inicialização
document.addEventListener("DOMContentLoaded", () => {
  carregarPlanilha();
  setInterval(carregarPlanilha, INTERVALO);
});
