// ===========================
// Painel de Demandas - Affari (versão otimizada)
// ===========================

// URL do seu Apps Script publicado como Web App (versão /exec)
const API_URL = 'https://script.google.com/macros/s/AKfycbxurrqpPqtzeBlCXJgx6P_o-Ms_CRNGMy0jHdDBHJvc3_NgbIYGYLAI4PR2zvqYH-l6QQ/exec';
const INTERVALO = 60 * 1000; // 1 minuto

// --------- Função principal para carregar dados ---------
async function carregarPlanilha() {
  try {
    const resposta = await fetch(API_URL + '?_=' + Date.now(), { cache: 'no-store' });
    const dados = await resposta.json();

    if (!dados.sucesso || !dados.rows) throw new Error("Erro ao obter dados");
    renderizarCards(dados.rows);
  } catch (erro) {
    console.error("Erro ao carregar dados:", erro);
    document.getElementById("cardsContainer").innerHTML =
      "<p style='color:white'>Erro ao carregar planilha. Verifique conexão e permissões.</p>";
  }
}

// --------- Renderiza os cards com base nos dados ---------
function renderizarCards(linhas) {
  const container = document.getElementById("cardsContainer");
  const contador = document.getElementById("contador") || document.createElement("div");
  contador.id = "contador";
  contador.style.marginBottom = "15px";
  contador.style.color = "white";
  container.innerHTML = "";
  container.parentNode.insertBefore(contador, container);

  const statusSalvos = JSON.parse(localStorage.getItem("statusCards") || "{}");
  let pendentes = 0;
  let entregues = 0;

  linhas.forEach((linha) => {
    const secretaria = (linha.Secretaria || "").trim();
    const dataFmt = (linha.Data || "").trim();
    const horaFmt = (linha.Horário || linha.Hora || "").trim();
    const local = (linha.Local || "").trim();
    const diretor = (linha.Diretor || "").trim();
    const itens = linha.Itens || "";
    const obs = linha.Observações || linha.Observacoes || "";
    const statusPlanilha = (linha.Status_Entrega || "").trim();

    const idUnico = `${dataFmt}-${horaFmt}-${local}-${diretor}`.replace(/\W+/g, "_");
    let estado = statusSalvos[idUnico] || statusPlanilha || "Pendente";

    if (estado.toLowerCase() === "entregue") {
      entregues++;
      return; // não exibe cards entregues
    } else {
      pendentes++;
    }

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <h2>${secretaria}</h2>
      <div class='info'>
        <strong>Diretor:</strong> ${diretor || "-"}<br>
        <strong>Local:</strong> ${local || "-"}<br>
        <strong>Data:</strong> ${dataFmt}<br>
        <strong>Hora:</strong> ${horaFmt}
      </div>
      <div class='itens'><strong>Itens:</strong><br>${itens || "-"}</div>
      <div class='observacoes'><strong>Observações:</strong><br>${obs || "-"}</div>
      <div class='status'></div>
    `;

    const statusEl = card.querySelector(".status");

    function atualizarVisual() {
      statusEl.textContent = estado;
      card.classList.remove("pendente", "entregue");
      if (estado.toLowerCase() === "pendente") {
        card.classList.add("pendente");
      } else if (estado.toLowerCase() === "entregue") {
        card.classList.add("entregue");
        card.style.backgroundColor = "#ccffcc"; // verde claro
      }
    }

    card.addEventListener("click", async () => {
      if (estado.toLowerCase() !== "pendente") return;
      if (!confirm("Deseja marcar este item como ENTREGUE?")) return;

      estado = "Entregue";
      statusSalvos[idUnico] = estado;
      localStorage.setItem("statusCards", JSON.stringify(statusSalvos));

      card.style.transition = "opacity 0.4s";
      card.style.opacity = "0";
      setTimeout(() => card.remove(), 400);

      try {
        const body = {
          Data: dataFmt,
          Hora: horaFmt,
          Local: local,
          Diretor: diretor,
          Status: "Entregue"
        };
        const res = await fetch(API_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });
        const json = await res.json();
        if (json.sucesso) {
          mostrarNotificacao("✅ Marcado como entregue.");
        } else {
          mostrarNotificacao("⚠️ Falha ao atualizar planilha.");
        }
      } catch (err) {
        mostrarNotificacao("⚠️ Erro de comunicação com a planilha.");
        console.error(err);
      }
    });

    atualizarVisual();
    container.appendChild(card);
  });

  contador.innerHTML = `
    <span style='color:#ff4d4d;'>🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp;
    <span style='color:#00b050;'>🟢 Entregues: ${entregues}</span>
  `;

  document.getElementById("updateInfo").textContent =
    `Última atualização: ${new Date().toLocaleTimeString()} — Atualizando automaticamente a cada 1 minuto`;
}

// --------- Notificação visual ---------
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

// --------- Inicialização ---------
document.addEventListener("DOMContentLoaded", () => {
  carregarPlanilha();
  setInterval(carregarPlanilha, INTERVALO);
});
