// Config
const API_URL = "https://script.google.com/macros/s/AKfycbxurrqpPqtzeBlCXJgx6P_o-Ms_CRNGMy0jHdDBHJvc3_NgbIYGYLAI4PR2zvqYH-l6QQ/exec";
const INTERVALO = 60 * 1000; // 1 minuto

// --- Utilitários de formatação ---
// Tenta normalizar vários formatos que podem vir do Apps Script:
// - número serial Excel (pouco provável vindo via doGet, mas mantido)
// - string "dd/mm/yyyy" ou "dd/mm/yyyy – HH:MM"
// - ISO string "2025-11-03T12:00:00.000Z"
// - Date-like "Mon Nov 03 2025 00:00:00 GMT-0300 (BRT)"
function formatDateSmart(value) {
  if (!value && value !== 0) return "";

  // número (serial Excel)
  if (typeof value === "number") {
    // converte serial Excel -> Date UTC
    const utc_days = Math.floor(value - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional_day = value - Math.floor(value) + 0.0000001;
    const total_seconds = Math.floor(86400 * fractional_day);
    const hours = Math.floor(total_seconds / 3600);
    const minutes = Math.floor((total_seconds % 3600) / 60);
    const d = String(date_info.getUTCDate()).padStart(2, "0");
    const m = String(date_info.getUTCMonth() + 1).padStart(2, "0");
    const y = date_info.getUTCFullYear();
    const diasSemana = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
    const diaSemana = diasSemana[date_info.getUTCDay()];
    if (value > 30000 && value % 1 === 0) return `${diaSemana}, ${d}/${m}/${y}`;
    if (value > 30000 && value % 1 !== 0) return `${diaSemana}, ${d}/${m}/${y} – ${String(hours).padStart(2,"0")}:${String(minutes).padStart(2,"0")}`;
    if (value < 1) return `${String(hours).padStart(2,"0")}:${String(minutes).padStart(2,"0")}`;
    return `${d}/${m}/${y}`;
  }

  // string: tenta extrair dd/mm/yyyy e hora
  const s = String(value).trim();
  // já no formato dd/mm/yyyy ou dd/mm/yyyy – HH:MM
  const m1 = s.match(/(\d{2}\/\d{2}\/\d{4})(?:\s*[–-]\s*(\d{1,2}:\d{2}))?/);
  if (m1) {
    const dia = m1[1];
    if (m1[2]) return `${dia} – ${m1[2].padStart(5,'0')}`;
    return dia;
  }

  // ISO ou Date string -> formata para dd/mm/yyyy – HH:MM
  const parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const dt = new Date(parsed);
    const d = String(dt.getDate()).padStart(2, "0");
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    const y = dt.getFullYear();
    const hh = String(dt.getHours()).padStart(2, "0");
    const mm = String(dt.getMinutes()).padStart(2, "0");
    const diasSemana = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
    const diaSemana = diasSemana[dt.getDay()];
    return `${diaSemana}, ${d}/${m}/${y} – ${hh}:${mm}`;
  }

  // fallback: retorna a string original
  return s;
}

function normalizarHora(h) {
  if (!h) return "";
  const s = String(h).trim();
  const m = s.match(/(\d{1,2})[:\.](\d{2})/);
  if (m) return `${m[1].padStart(2, "0")}:${m[2]}`;
  // se vier "dd/mm/yyyy – HH:MM"
  const mm = s.match(/–\s*(\d{1,2}:\d{2})$/);
  if (mm) return mm[1];
  // ISO date
  const parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const dt = new Date(parsed);
    return `${String(dt.getHours()).padStart(2,"0")}:${String(dt.getMinutes()).padStart(2,"0")}`;
  }
  return s;
}

// --- Carregar dados via doGet (API) ---
async function carregarPlanilha() {
  try {
    const resposta = await fetch(API_URL + "?_=" + Date.now(), { cache: "no-store" });
    const dados = await resposta.json();
    if (!dados || !dados.sucesso || !Array.isArray(dados.rows)) {
      throw new Error(dados && dados.erro ? dados.erro : "Resposta inválida do servidor");
    }
    renderizarCards(dados.rows);
  } catch (err) {
    console.error("Erro ao carregar planilha:", err);
    const container = document.getElementById("cardsContainer");
    if (container) container.innerHTML = `<p style="color:white; text-align:center;">⚠️ ${err.message}</p>`;
  }
}

// --- Renderizar cards (sem localStorage) ---
function renderizarCards(rows) {
  const container = document.getElementById("cardsContainer");
  const contador = document.getElementById("contadorStatus");
  if (!container) return;
  container.innerHTML = "";

  let pendentes = 0;
  let entregues = 0;

  // ordenar por horário se existir
  rows.sort((a, b) => {
    const ha = normalizarHora(a["Horário"] || a.Hora || "");
    const hb = normalizarHora(b["Horário"] || b.Hora || "");
    return ha.localeCompare(hb);
  });

  rows.forEach((row) => {
    const status = (row.Status_Entrega || "").toString().trim();
    const secretaria = (row.Secretaria || "").toString().trim() || "(Sem Secretaria)";
    const dataRaw = row.Data;
    const horaRaw = row["Horário"] || row.Hora || "";
    const dataFmt = formatDateSmart(dataRaw);
    const horaFmt = normalizarHora(horaRaw || dataFmt);
    const local = (row.Local || "").toString().trim();
    const diretor = (row.Diretor || "").toString().trim();
    const itens = row.Itens || "";
    const obs = row.Observações || row.Observacoes || "";
    const pax = row.Pax || "";

    if (status.toLowerCase() === "entregue") {
      entregues++;
      return; // não exibe entregues
    } else pendentes++;

    const card = document.createElement("div");
    card.className = "card pendente";

    card.innerHTML = `
      <div class="top">
        <img class="logo"download.png" alt="Affari logo" onerror="this.style.display='none'">
        <h2>${escapeHtml(secretaria)}</h2>
      </div>
      <div class="info">
        <strong>Diretor:</strong> ${escapeHtml(diretor || "-")}<br>
        <strong>Local:</strong> ${escapeHtml(local || "-")}<br>
        <strong>Data:</strong> ${escapeHtml(dataFmt || "-")}<br>
        <strong>Hora:</strong> ${escapeHtml(horaFmt || "-")}<br>
        <strong>Pax:</strong> ${escapeHtml(pax || "-")}
      </div>
      <div class="itens"><strong>Itens:</strong><br>${escapeHtml(String(itens || "-"))}</div>
      <div class="observacoes"><strong>Observações:</strong><br>${escapeHtml(String(obs || "-"))}</div>
      <div class="status">Pendente</div>
    `;

    // clique para marcar entregue
    card.addEventListener("click", async () => {
      if (!confirm("Deseja marcar este item como ENTREGUE?")) return;

      // animação de saída
      card.style.opacity = "0";
      setTimeout(() => card.remove(), 350);

      // envia atualização ao servidor
      try {
        const body = {
          Data: formatDateForPost(dataRaw, dataFmt),
          Hora: horaFmt,
          Local: local,
          Diretor: diretor,
          Status: "Entregue",
        };
        const res = await fetch(API_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        const json = await res.json();
        if (json && json.sucesso) {
          mostrarNotificacao("✅ Marcado como entregue!");
        } else {
          mostrarNotificacao("⚠️ Falha ao atualizar planilha!");
          console.warn("Resposta do servidor:", json);
        }
      } catch (err) {
        console.error("Erro ao enviar POST:", err);
        mostrarNotificacao("⚠️ Erro na comunicação com a planilha!");
      }
    });

    container.appendChild(card);
  });

  // contador
  if (contador) {
    contador.innerHTML = ` <span style="color:#ff4d4d;">🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp; <span style="color:#00b050;">🟢 Entregues: ${entregues}</span>`;
  }

  document.getElementById("updateInfo").textContent =
    `Última atualização: ${new Date().toLocaleTimeString()} — Atualizando automaticamente a cada 1 minuto`;
}

// util: proteger texto para innerHTML
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// tenta preparar Data para postar: se veio serial numérico, converte para dd/mm/yyyy; se já string dd/mm/yyyy usa ela; se ISO -> converte para dd/mm/yyyy
function formatDateForPost(rawValue, formattedFallback) {
  // se rawValue é number (serial excel)
  if (typeof rawValue === "number") {
    const utc_days = Math.floor(rawValue - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const d = String(date_info.getUTCDate()).padStart(2, "0");
    const m = String(date_info.getUTCMonth() + 1).padStart(2, "0");
    const y = date_info.getUTCFullYear();
    return `${d}/${m}/${y}`;
  }

  const s = String(rawValue || "").trim();
  const m1 = s.match(/(\d{2}\/\d{2}\/\d{4})/);
  if (m1) return m1[1];

  // tenta parsear ISO
  const parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    const dt = new Date(parsed);
    return `${String(dt.getDate()).padStart(2,"0")}/${String(dt.getMonth()+1).padStart(2,"0")}/${dt.getFullYear()}`;
  }

  // fallback: usar o formatted (que já foi gerado) e pegar só a data parte
  const f = String(formattedFallback || "");
  const mf = f.match(/(\d{2}\/\d{2}\/\d{4})/);
  if (mf) return mf[1];

  return s; // último recurso
}

// notificação pequena
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
    transition: "opacity .25s",
  });
  document.body.appendChild(alerta);
  requestAnimationFrame(() => (alerta.style.opacity = "1"));
  setTimeout(() => {
    alerta.style.opacity = "0";
    setTimeout(() => alerta.remove(), 300);
  }, 2500);
}

// inicialização
document.addEventListener("DOMContentLoaded", () => {
  carregarPlanilha();
  setInterval(carregarPlanilha, INTERVALO);
});
