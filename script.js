const PLANILHA =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vSR3FyZKCXFS5Mi4UaRc6GLCSfH0erH_rraD87M0ZFo6jeDT0hEnpvUfEH2-cxXI0-ionFDxLFFUuvg/pub?output=csv";

const API_URL =
  "https://script.google.com/macros/s/AKfycbymCML92y2aeDoem6dqC4tXdVMV0kMTvf2GgFf4SahjOjNex1weiVJOAq6gueO4QZ2Y/exec";

const INTERVALO = 30 * 1000; // 30 segundos

function excelSerialToDate(serial) {
  if (!serial) return "";
  if (typeof serial === "string") return serial.trim();

  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400;
  const date_info = new Date(utc_value * 1000);
  const fractional_day = serial - Math.floor(serial);
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

const normalizarHora = (h) => {
  if (!h) return "";
  const s = String(h).trim();
  const m = s.match(/(\d{1,2})[:\.](\d{2})/);
  return m ? `${m[1].padStart(2, "0")}:${m[2]}` : s;
};

const apenasData = (str) =>
  !str ? "" : (String(str).match(/(\d{2}\/\d{2}\/\d{4})/) || [])[1] || str;

async function carregarPlanilha() {
  try {
    const url = `${PLANILHA}&t=${Date.now()}`;
    let response = await fetch(url, { cache: "no-store" });
    if (!response.ok) throw new Error("Erro ao carregar planilha.");

    let texto = (await response.text()).replace(/^\uFEFF/, "").trim();
    if (texto.startsWith("<")) throw new Error("⚠️ Planilha não está publicada em CSV.");

    const delimitador = texto.split("\n")[0].includes(";") ? ";" : ",";
    const parsed = Papa.parse(texto, {
      header: true,
      skipEmptyLines: true,
      delimiter: delimitador,
    });

    if (!parsed.data.length) throw new Error("⚠️ Planilha vazia.");
    renderizarCards(parsed.data);
  } catch (erro) {
    console.error("Erro ao ler planilha:", erro);
    document.getElementById("cardsContainer").innerHTML =
      `<p style="color:white;text-align:center;">⚠️ ${erro.message}</p>`;
  }
}

function renderizarCards(dados) {
  const container = document.getElementById("cardsContainer");
  const contador = document.getElementById("contadorStatus");
  if (!container) return;

  container.innerHTML = "";
  let pendentes = 0, entregues = 0;

  const entreguesLocal = JSON.parse(localStorage.getItem("entreguesLocal") || "[]");

  dados
    .sort((a, b) =>
      normalizarHora(a["Horário"] || a.Hora || "").localeCompare(
        normalizarHora(b["Horário"] || b.Hora || "")
      )
    )
    .forEach((linha) => {
      const secretaria = linha.Secretaria?.trim() || "(Sem Secretaria)";
      const dataFmt = excelSerialToDate(linha.Data);
      const soData = apenasData(dataFmt) || String(linha.Data || "").trim();
      const horaHHMM = normalizarHora(linha["Horário"] || linha.Hora);
      const local = (linha.Local || "").trim();
      const diretor = (linha.Diretor || "").trim();
      const itens = linha.Itens || "";
      const observacoes = linha.Observações || linha.Observacoes || "";
      const statusEntregaPlanilha = linha.Status_Entrega || "";
      const pax = linha.Pax || 0;

      const idUnico = `${soData}-${horaHHMM}-${local}-${diretor}`.replace(/\W+/g, "_");
      let estado = statusEntregaPlanilha || "Pendente";

      // 🚫 Ignorar cards já entregues
      if (estado.toLowerCase() === "entregue" || entreguesLocal.includes(idUnico)) {
        entregues++;
        return;
      }

      pendentes++;

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
        <div class="observacoes"><strong>Observações:</strong><br>${observacoes || "-"}</div>
        <div class="status"></div>
      `;

      const statusEl = card.querySelector(".status");
      statusEl.textContent = estado;
      card.classList.add(estado.toLowerCase());

      card.addEventListener("click", async () => {
        if (estado.toLowerCase() !== "pendente") return;
        if (!confirm("Tem certeza que deseja marcar como ENTREGUE?")) return;

        estado = "Entregue";
        statusEl.textContent = estado;
        card.classList.add("fade-out");
        setTimeout(() => card.remove(), 400);

        entreguesLocal.push(idUnico);
        localStorage.setItem("entreguesLocal", JSON.stringify(entreguesLocal));

        try {
          await fetch(API_URL, {
            method: "POST",
            mode: "no-cors",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              Data: soData,
              Hora: horaHHMM,
              Local: local,
              Diretor: diretor,
              Status: "Entregue",
            }),
          });
          mostrarNotificacao("✅ Marcado como entregue!");
        } catch (e) {
          console.error("Erro ao atualizar planilha:", e);
          mostrarNotificacao("⚠️ Erro ao atualizar planilha!");
        }
      });

      container.appendChild(card);
    });

  contador.innerHTML = `
    <span style="color:#ff4d4d;">🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp;
    <span style="color:#00b050;">🟢 Entregues: ${entregues}</span>
  `;
  document.getElementById("updateInfo").textContent =
    `Última atualização: ${new Date().toLocaleTimeString()} — Atualiza a cada 30s`;
}

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

document.addEventListener("DOMContentLoaded", () => {
  carregarPlanilha();
  setInterval(carregarPlanilha, INTERVALO);
});
