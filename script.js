const API_URL = 'https://script.google.com/macros/s/AKfycbzxDUsDpFbAlPf3BI5woUvVP0bziZtC2kxko73W9Fof/dev';
const local=(linha.Local||'').trim();
const diretor=(linha.Diretor||'').trim();
const itens=linha.Itens||'';
const obs=linha.Observações||linha.Observacoes||'';
const statusPlanilha=linha.Status_Entrega||'';
const idUnico=`${dataFmt}-${horaFmt}-${local}-${diretor}`.replace(/\W+/g,'_');
let estado=statusSalvos[idUnico]||statusPlanilha||'Pendente';


if(estado.toLowerCase()==='entregue'){entregues++;return;}else{pendentes++;}


const card=document.createElement('div');
card.className='card';
card.innerHTML=`
<h2>${secretaria}</h2>
<div class='info'>
<strong>Diretor:</strong> ${diretor||'-'}<br>
<strong>Local:</strong> ${local||'-'}<br>
<strong>Data:</strong> ${dataFmt}<br>
<strong>Hora:</strong> ${horaFmt}
</div>
<div class='itens'><strong>Itens:</strong><br>${itens||'-'}</div>
<div class='observacoes'><strong>Observações:</strong><br>${obs||'-'}</div>
<div class='status'></div>
`;


const statusEl=card.querySelector('.status');
function atualizarVisual(){
statusEl.textContent=estado;
card.classList.remove('pendente','entregue');
card.classList.add(estado.toLowerCase());
}


card.addEventListener('click',async()=>{
if(estado.toLowerCase()!=='pendente')return;
if(!confirm('Deseja marcar este item como ENTREGUE?'))return;
estado='Entregue';
statusSalvos[idUnico]=estado;
localStorage.setItem('statusCards',JSON.stringify(statusSalvos));
card.style.opacity='0';
setTimeout(()=>card.remove(),400);
try{
const body={Data:dataFmt,Hora:horaFmt,Local:local,Diretor:diretor,Status:'Entregue'};
const res=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
const json=await res.json();
if(json.sucesso) mostrarNotificacao('✅ Marcado como entregue.');
else mostrarNotificacao('⚠️ Falha ao atualizar planilha.');
}catch(err){
mostrarNotificacao('⚠️ Erro de comunicação com a planilha.');
}
});


atualizarVisual();
container.appendChild(card);
});


contador.innerHTML=`<span style='color:#ff4d4d;'>🔴 Pendentes: ${pendentes}</span> &nbsp;&nbsp; <span style='color:#00b050;'>🟢 Entregues: ${entregues}</span>`;
document.getElementById('updateInfo').textContent=`Última atualização: ${new Date().toLocaleTimeString()} — Atualizando automaticamente a cada 1 minuto`;
}


function mostrarNotificacao(texto){
const alerta=document.createElement('div');
alerta.textContent=texto;
Object.assign(alerta.style,{position:'fixed',bottom:'20px',right:'20px',background:'#333',color:'#fff',padding:'10px 16px',borderRadius:'8px',boxShadow:'0 2px 6px rgba(0,0,0,.3)',zIndex:9999,fontSize:'13px',opacity:0,transition:'opacity .3s'});
document.body.appendChild(alerta);
requestAnimationFrame(()=>alerta.style.opacity='1');
setTimeout(()=>{alerta.style.opacity='0';setTimeout(()=>alerta.remove(),300);},2500);
}


document.addEventListener('DOMContentLoaded',()=>{carregarPlanilha();setInterval(carregarPlanilha,INTERVALO);});
