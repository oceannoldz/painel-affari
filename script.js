const PLANILHA =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vSR3FyZKCXFS5Mi4UaRc6GLCSfH0erH_rraD87M0ZFo6jeDT0hEnpvUfEH2-cxXI0-ionFDxLFFUuvg/pub?output=csv";

const API_URL =
  "https://script.google.com/macros/s/AKfycbyoNSATAd22R7ArTq09e3rk5reg3u2kyWQVGV80TLU4TK_NPVq__3wvtxTXoIV5TUnJ9A/exec";

const INTERVALO = 60 * 1000; // 1 minuto

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

async function carregarPlanilha() {
  try {
    const url = PLANILHA + "&t=" + new Date().getTime();
    let response = await fetch(url);
    if (!response.ok) {
      response = await fetch("https://cors.isomorphic-git.org/" + url);
    }
    if (!response.ok) throw new Error("Erro ao carregar planilha Google.");

    let texto = await response.text();
    texto = texto.replace(/^\uFEFF/, "").trim();

    if (texto.startsWith("<")) {
      throw new Error("⚠️ A planilha não está publicada em formato CSV.");
    }

    const primeiraLinha = texto.split("\n")[0];
    const delimitador = primeiraLinha.includes(";") ? ";" : ",";

    const parsed = Papa.parse(texto, {
      header: true,
      skipEmptyLines: true,
      delimiter: delimitador,
    });

    const dados = parsed.data;
    console.log(
      `✅ CSV lido: ${dados.length} registros (delimitador "${delimitador}")`
    );

    if (!dados.length) {
      throw new Error("⚠️ A planilha foi carregada, mas não contém registros.");
    }

    // 🔸 NÃO renderiza mais os cards (apenas exibe no console)
    console.log("Planilha carregada com sucesso:", dados.length, "registros.");
  } catch (erro) {
    console.error("Erro ao ler planilha:", erro);
    const container = document.getElementById("cardsContainer");
    if (container) {
      container.innerHTML = `<p style="color:white; text-align:center;">⚠️ ${erro.message}</p>`;
    }
  }
}

// 🔸 Ajuste no fetch para envio ao Apps Script (sem 'no-cors')
async function marcarComoEntregue(dadosEnvio) {
  try {
    const resposta = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dadosEnvio),
    });

    if (!resposta.ok) throw new Error("Erro HTTP ao enviar dados");

    const resultado = await resposta.json();
    console.log("✅ Dados enviados com sucesso:", resultado);
    mostrarNotificacao("✅ Marcado como entregue!");
  } catch (erro) {
    console.error("Erro ao enviar para planilha:", erro);
    mostrarNotificacao("⚠️ Erro ao atualizar a planilha!");
  }
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

