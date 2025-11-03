const API_URL = "https://script.google.com/macros/s/AKfycbxurrqpPqtzeBlCXJgx6P_o-Ms_CRNGMy0jHdDBHJvc3_NgbIYGYLAI4PR2zvqYH-l6QQ/exec";
const INTERVALO = 60 * 1000; // 1 minuto

async function carregarPlanilha() {
  try {
    const resposta = await fetch(API_URL + "?_=" + Date.now());
    const dados = await resposta.json();
    if (!dados.sucesso || !dados.rows) throw new Error("Erro ao obter dados da planilha");
    renderizarCards(dados.rows);
  } catch (erro) {
    console.error("Erro ao ler planilha:", erro);
    document.getElementById("cardsContainer").innerHTML = `<p style='color:white; text-align:center;'>⚠️ ${erro.message}</p>`;
  }
}

function renderizarCards(dados) {
  const container = document.getElementById("cardsContainer");
  const contador = document.getElementById("contadorStatus");
  container.innerHTML = "";
  let pendentes = 0;
  let entregues = 0;

  dados.forEach((linha) => {
    const status = (linha.Status_Entrega || "").trim();
    if (status.toLowerCase() === "entregue") {
      entregues++;
      return;
    } else pendentes++;

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `
      <h2>${linha.Secretaria || "(Sem Secretaria)"}</h2>
      <div class="info">
        <strong>Diretor:</strong> ${linha.Diretor || "-"}<br>
        <strong>Local:</strong> ${linha.Local || "-"}<br>
        <strong>Data:</strong> ${linha.Data}<br>
        <strong>Hora:</strong> ${linha["Horário"] || linha.Hora || "-"}
      </div>
      <div class="itens"><strong>Itens:</strong><br>${linha.Itens || "-"}</div>
      <div class="observacoes"><strong>Observações:</strong><br>${linha.Observações || linha.Observacoes || "-"}</div>
      <div class="status">${status || "Pendente"}</div>
    `;

    card.addEventListener("click", async () => {
      if (!confirm("Marcar este item como ENTREGUE?")) return;
      card.style.opacity = "0";
      setTimeout(() => card.remove(), 400);

      try {
        const body = {
          Data: linha.Data,
          Hora: linha["Horário"] || linha.Hora,
          Local: linha.Local,
          Diretor: linha.Diretor,
          Status: "Entregue",
        };

        const resposta = await fetch(API_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(body),
        });

        const resultado = await resposta.json();
        if (resultado.sucesso)
          mostrarNotificacao("✅ Marcado como entregue!");
        else
          mostrarNotificacao("⚠️ Erro ao atualizar planilha!");
      } catch (erro) {
        console.error("Erro ao atualizar:", erro);
        mostrarNotificacao("⚠️ Falha na comunicação!");
      }
    });

    container.appendChild(card);
  });

  contador.innerHTML = `
    <span style='color:#ff4d4d;'>🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp;
    <span style='color:#00b050;'>🟢 Entregues: ${entregues}</span>`;

  document.getElementById("updateInfo").textContent =
    `Última atualização: ${new Date().toLocaleTimeString()} — Atualizando automaticamente a cada 1 minuto`;
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
