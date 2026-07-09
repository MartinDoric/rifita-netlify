const API_URL = '/.netlify/functions/api';
let codigoRifa = '';
let adminPin = '';
let isAdmin = false;
let appData = null;
let dashboard = null;
let vendedores = [];
let config = null;
let bgDataUrl = '';
let pickedBgFile = null;
let lastSaleReceipt = null;

const $ = (id) => document.getElementById(id);
function showMsg(id, text, ok = true){ const el=$(id); if(!el) return; el.textContent=text||''; el.className='msg '+(ok?'ok':'err'); }
function escapeHtml(s){ return String(s||'').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c])); }
function normalizeCode(s){ return String(s||'').trim().toLowerCase().replace(/\s+/g,''); }
async function api(action, data={}){
  const res = await fetch(API_URL, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ action, codigoRifa, ...data }) });
  const text = await res.text();
  let json; try { json = JSON.parse(text); } catch(e){ throw new Error('Respuesta inválida del servidor: '+text.slice(0,120)); }
  if(!json.ok) throw new Error(json.error || 'Error desconocido');
  return json.data;
}
function downloadText(name, text, type='text/csv;charset=utf-8'){
  const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([text],{type})); a.download=name; a.click(); URL.revokeObjectURL(a.href);
}
function csv(rows){ return rows.map(r => r.map(v => `"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n'); }
function fillSelect(id, values, placeholder){ const sel=$(id); if(!sel) return; sel.innerHTML=''; if(placeholder){ const o=document.createElement('option'); o.value=''; o.textContent=placeholder; sel.appendChild(o); } values.forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent=v; sel.appendChild(o); }); }

async function enterRifa(){
  const code = normalizeCode($('rifaCodeInput').value);
  const remember = $('rememberDeviceCheck').checked;
  if(!code){ showMsg('rifaLoginMsg','Ingresá el código de rifa.',false); return; }
  showMsg('rifaLoginMsg','Cargando…',true);
  try{
    await api('checkRifa', { codigoRifa: code });
    codigoRifa = code;
    if(remember) localStorage.setItem('rifita_codigo', code); else localStorage.removeItem('rifita_codigo');
    $('rifaLogin').classList.add('hidden'); $('appShell').classList.remove('hidden');
    await loadAll(true);
  }catch(e){ showMsg('rifaLoginMsg','❌ '+e.message,false); }
}
function changeRifa(){ localStorage.removeItem('rifita_codigo'); location.reload(); }

async function loadAll(show=false){
  try{
    if(show) showMsg('saleMsg','Cargando datos…',true);
    appData = await api('init');
    dashboard = appData.dashboard; vendedores = appData.vendedores || []; config = appData.config || {};
    $('titleMain').textContent = '🎟️ Rifita - ' + (appData.rifa?.nombre || codigoRifa);
    $('subline').textContent = 'Código: '+codigoRifa+' · Actualizado: '+new Date().toLocaleString();
    renderSummary(); renderSaleForm(); renderBoardDefaults(); renderPubDefaults(); renderConfig(); renderWinners();
    await loadBackgroundImage();
    if(show) showMsg('saleMsg','',true);
  }catch(e){ showMsg('saleMsg','❌ '+e.message,false); }
}
function renderSummary(){ $('soldCount').textContent=dashboard?.soldCount??'–'; $('availCount').textContent=dashboard?.availableCount??'–'; $('rangeTxt').textContent=`${dashboard?.startNum||'–'}–${dashboard?.endNum||'–'}`; }
function renderSaleForm(){
  fillSelect('saleNumero', dashboard.availableNumbers || [], 'No hay números disponibles');
  fillSelect('saleNumero2', dashboard.availableNumbers || [], 'Segundo número');
  fillSelect('saleVend', vendedores, 'Elegí vendedor'); fillSelect('pubVend', vendedores, 'Elegí vendedor'); fillSelect('deleteSellerSelect', vendedores, 'Elegí vendedor');
  $('saleTypeBox').classList.toggle('hidden', !(config?.promoActiva));
  $('saleNumero2Box').classList.toggle('hidden', $('saleTipoPrecio').value !== '2');
  const sold = dashboard.soldNumbers || []; fillSelect('delNumero', sold, sold.length?'Elegí número vendido':'No hay vendidos');
}
function renderBoardDefaults(){ if(!$('boardFrom').value) $('boardFrom').value=dashboard.startNum; if(!$('boardTo').value) $('boardTo').value=Math.min(dashboard.startNum+49,dashboard.endNum); renderBoardRange(); }
function renderPubDefaults(){ if(!$('pubFrom').value) $('pubFrom').value=dashboard.startNum; if(!$('pubTo').value) $('pubTo').value=Math.min(dashboard.startNum+49,dashboard.endNum); const key='rifita_footer_'+codigoRifa; const saved=localStorage.getItem(key); $('pubFooterMsg').value = saved != null ? saved : (config?.publicidad?.mensajeInferior || ''); $('pubImageSavedHint').textContent = config?.publicidad?.imagenNombre ? 'Imagen guardada: '+config.publicidad.imagenNombre : 'Sin imagen guardada.'; }
function renderConfig(){
  $('cfgStart').value=dashboard.startNum; $('cfgEnd').value=dashboard.endNum; $('cfgPrecio1').value=config?.precio1||0; $('cfgPrecio2').value=config?.precio2||0; $('cfgPromo').checked=!!config?.promoActiva;
  const p=config?.publicidad||{}; $('cfgPubTitulo').value=p.titulo||''; $('cfgPubPremios').value=p.premios||''; $('cfgPubContacto').value=p.contacto||''; $('cfgPubExtra').value=p.textoExtra||''; $('cfgPubFooter').value=p.mensajeInferior||'';
}
function renderWinners(){ const w=dashboard?.winners||[]; $('winnerBox').innerHTML = w.length ? w.map(x=>`🏆 Nº ${escapeHtml(x.numero)} · ${escapeHtml(x.nombreApellido)} · ${escapeHtml(x.vendedor)}`).join('\n') : 'Todavía no hay ganadores.'; }

function showTab(tab){
  document.querySelectorAll('.tab[data-tab]').forEach(t=>t.classList.toggle('active',t.dataset.tab===tab));
  ['sale','board','pub','del','draw','config','report'].forEach(t=> $('tab'+t[0].toUpperCase()+t.slice(1))?.classList.add('hidden'));
  const map={sale:'tabSale',board:'tabBoard',pub:'tabPub',del:'tabDel',draw:'tabDraw',config:'tabConfig',report:'tabReport'}; $(map[tab])?.classList.remove('hidden');
}
async function requestAdmin(){ $('pinOverlay').classList.remove('hidden'); $('pinInput').focus(); }
async function confirmPin(){ const pin=$('pinInput').value.trim(); if(!pin){ showMsg('pinMsg','Ingresá el PIN.',false); return; } showMsg('pinMsg','Verificando…',true); try{ await api('adminLogin',{pin}); adminPin=pin; isAdmin=true; $('appShell').classList.add('adminOn'); $('adminState').textContent='Modo: administrador'; $('pinOverlay').classList.add('hidden'); $('pinInput').value=''; showTab('sale'); await loadAll(); }catch(e){ showMsg('pinMsg','❌ '+e.message,false); } }
function closePin(){ $('modeSelect').value='public'; $('pinOverlay').classList.add('hidden'); }
function setPublicMode(){ isAdmin=false; adminPin=''; $('appShell').classList.remove('adminOn'); $('adminState').textContent='Modo: vendedor'; showTab('sale'); }

async function doSale(){
  const payload={ tipoPrecio:$('saleTipoPrecio').value, numero:$('saleNumero').value, numero2:$('saleNumero2').value, nombreApellido:$('saleNombre').value, dni:$('saleDni').value, telefono:$('saleTel').value, vendedor:$('saleVend').value };
  showMsg('saleMsg','Guardando venta…',true); $('saleWhatsappBox').classList.add('hidden');
  try{ const res=await api('confirmSale',{payload}); lastSaleReceipt=res.venta; showMsg('saleMsg','✅ Venta confirmada: '+res.vendidos.join(', '),true); $('saleNombre').value=''; $('saleDni').value=''; $('saleTel').value=''; $('saleWhatsappBox').classList.remove('hidden'); await loadAll(); }catch(e){ showMsg('saleMsg','❌ '+e.message,false); }
}
function sendWhatsappReceipt(){ if(!lastSaleReceipt) return; const v=lastSaleReceipt; const msg=`🎟️ Comprobante de compra\n\nRifa: ${v.rifa}\nNúmero/s: ${v.numeros.join(', ')}\nNombre: ${v.nombreApellido}\nDNI: ${v.dni}\nVendedor: ${v.vendedor}\nFecha: ${new Date(v.fecha).toLocaleString()}\n\nGracias por participar.`; window.open('https://wa.me/?text='+encodeURIComponent(msg),'_blank'); }

function setQuick(prefix, str){ const [a,b]=str.split(',').map(Number); $(prefix+'From').value=a; $(prefix+'To').value=b; if(prefix==='board') renderBoardRange(); }
function renderBoardRange(){ const from=Number($('boardFrom').value)||dashboard.startNum; const to=Number($('boardTo').value)||Math.min(from+49,dashboard.endNum); const sold=new Set(dashboard.soldNumbers||[]); let html=''; for(let n=from;n<=to;n++){ const isSold=sold.has(n); html += `<button type="button" class="numCell ${isSold?'numSold':'numAvailable'}" data-num="${n}" ${isSold?'disabled':''}>${n}</button>`; } $('boardGrid').innerHTML=html; }
function pickNumberFromBoard(n){ if($('saleTipoPrecio').value==='2' && $('saleNumero').value && !$('saleNumero2').value){ $('saleNumero2').value=n; } else { $('saleNumero').value=n; } showTab('sale'); }

async function loadBackgroundImage(){ try{ const res=await api('getPublicityBackgroundData'); bgDataUrl=res.dataUrl||''; }catch(e){ bgDataUrl=''; } }
function fileToDataUrl(file){ return new Promise((resolve,reject)=>{ const r=new FileReader(); r.onload=()=>resolve(r.result); r.onerror=reject; r.readAsDataURL(file); }); }
async function uploadBackground(){ if(!pickedBgFile){ showMsg('pubMsg','Elegí una imagen primero.',false); return; } showMsg('pubMsg','Subiendo imagen…',true); try{ const dataUrl=await fileToDataUrl(pickedBgFile); await api('uploadPublicityBackground',{pin:adminPin,dataUrl,fileName:pickedBgFile.name}); showMsg('pubMsg','✅ Fondo guardado.',true); pickedBgFile=null; $('pubImageFileName').textContent='Ningún archivo elegido.'; await loadAll(); }catch(e){ showMsg('pubMsg','❌ '+e.message,false); } }
function saveFooterMsg(){ localStorage.setItem('rifita_footer_'+codigoRifa, $('pubFooterMsg').value); }
function useGeneralFooter(){ $('pubFooterMsg').value=config?.publicidad?.mensajeInferior||''; saveFooterMsg(); }
function generateFlyer(){ saveFooterMsg(); const c=$('flyerCanvas'); const ctx=c.getContext('2d'); const from=Number($('pubFrom').value)||dashboard.startNum; const to=Number($('pubTo').value)||from+49; const sold=new Set(dashboard.soldNumbers||[]); const footer=$('pubFooterMsg').value.trim(); const draw=()=>{ ctx.fillStyle='#f8fafc'; ctx.fillRect(0,0,c.width,c.height); if(bgDataUrl&&window._bgImg){ ctx.drawImage(window._bgImg,0,0,c.width,c.height); } else { ctx.fillStyle='#10223c'; ctx.fillRect(0,0,c.width,c.height); ctx.fillStyle='#fff'; ctx.font='900 72px system-ui'; ctx.textAlign='center'; ctx.fillText(config?.publicidad?.titulo || 'RIFITA', c.width/2, 170); ctx.font='500 36px system-ui'; wrapText(ctx,config?.publicidad?.premios||'',c.width/2,240,900,44,'center'); }
    const startY = bgDataUrl ? 760 : 520; const cols=10; const cellW=88; const cellH=62; const gap=10; const gridW=cols*cellW+(cols-1)*gap; let x0=(c.width-gridW)/2; let y=startY; ctx.textAlign='center'; ctx.font='900 28px system-ui';
    for(let n=from;n<=to;n++){ const i=n-from; const col=i%cols; const row=Math.floor(i/cols); const x=x0+col*(cellW+gap); const yy=y+row*(cellH+gap); ctx.fillStyle=sold.has(n)?'#FFC7CE':'#C6EFCE'; roundRect(ctx,x,yy,cellW,cellH,10,true,false); ctx.fillStyle='#000'; ctx.fillText(String(n),x+cellW/2,yy+41); }
    const rows=Math.ceil((to-from+1)/cols); const legendY=y+rows*(cellH+gap)+30; ctx.fillStyle='rgba(0,0,0,.38)'; roundRect(ctx,c.width/2-210,legendY,420,42,12,true,false); ctx.font='700 20px system-ui'; ctx.fillStyle='#35d07f'; ctx.fillText('● Disponible',c.width/2-80,legendY+27); ctx.fillStyle='#ff4d4d'; ctx.fillText('● Vendido',c.width/2+90,legendY+27);
    if(footer){ ctx.fillStyle='rgba(255,255,255,.88)'; roundRect(ctx,70,1540,940,190,24,true,false); ctx.fillStyle='#1f3d2b'; ctx.font='900 38px system-ui'; wrapText(ctx,footer,540,1600,860,46,'center'); }
    showMsg('pubMsg','✅ Publicidad generada.',true); };
  if(bgDataUrl){ const img=new Image(); img.onload=()=>{window._bgImg=img; draw();}; img.onerror=draw; img.src=bgDataUrl; } else draw(); }
function roundRect(ctx,x,y,w,h,r,fill,stroke){ ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); ctx.closePath(); if(fill)ctx.fill(); if(stroke)ctx.stroke(); }
function wrapText(ctx,text,x,y,maxWidth,lineHeight,align='left'){ ctx.textAlign=align; String(text||'').split('\n').forEach(line=>{ const words=line.split(' '); let cur=''; for(const word of words){ const test=cur?cur+' '+word:word; if(ctx.measureText(test).width>maxWidth && cur){ ctx.fillText(cur,x,y); y+=lineHeight; cur=word; } else cur=test; } if(cur){ ctx.fillText(cur,x,y); y+=lineHeight; } }); return y; }
function downloadFlyer(){ generateFlyer(); setTimeout(()=>{ const a=document.createElement('a'); a.href=$('flyerCanvas').toDataURL('image/png'); a.download='rifita_publicidad.png'; a.click(); },250); }
function copyPubMessage(){ const msg=`🎟️ ${appData.rifa.nombre}\nNúmeros ${$('pubFrom').value} al ${$('pubTo').value}\n${$('pubFooterMsg').value||''}`; navigator.clipboard.writeText(msg).then(()=>showMsg('pubMsg','✅ Mensaje copiado.',true)).catch(()=>showMsg('pubMsg','No se pudo copiar.',false)); }

async function loadSaleForDelete(){ const n=$('delNumero').value; if(!n) return; try{ const d=await api('getSale',{numero:n}); $('delInfo').textContent=`Nº ${d.numero}\n${d.nombreApellido}\nDNI: ${d.dni}\nTel: ${d.telefono}\nVendedor: ${d.vendedor}`; }catch(e){ $('delInfo').textContent='❌ '+e.message; } }
async function deleteSale(){ if(!confirm('¿Seguro que querés eliminar esta venta?')) return; try{ await api('deleteSale',{pin:adminPin,numero:$('delNumero').value}); showMsg('delMsg','✅ Venta eliminada.',true); await loadAll(); }catch(e){ showMsg('delMsg','❌ '+e.message,false); } }
async function drawWinners(){ try{ const res=await api('drawWinners',{pin:adminPin,count:$('winnerCount').value}); showMsg('drawMsg','✅ Sorteo realizado.',true); await loadAll(); renderWinners(); }catch(e){ showMsg('drawMsg','❌ '+e.message,false); } }
async function resetWinners(){ if(!confirm('¿Borrar ganadores guardados?')) return; try{ await api('resetWinners',{pin:adminPin}); showMsg('drawMsg','✅ Ganadores reseteados.',true); await loadAll(); }catch(e){ showMsg('drawMsg','❌ '+e.message,false); } }
async function saveRange(){ try{ await api('saveRange',{pin:adminPin,startNum:$('cfgStart').value,endNum:$('cfgEnd').value}); showMsg('configMsg','✅ Rango guardado.',true); await loadAll(); }catch(e){ showMsg('configMsg','❌ '+e.message,false); } }
async function saveConfig(){ const publicidad={titulo:$('cfgPubTitulo').value,premios:$('cfgPubPremios').value,contacto:$('cfgPubContacto').value,textoExtra:$('cfgPubExtra').value,mensajeInferior:$('cfgPubFooter').value,imagenId:config?.publicidad?.imagenId,imagenNombre:config?.publicidad?.imagenNombre}; try{ await api('saveConfig',{pin:adminPin,config:{precio1:$('cfgPrecio1').value,precio2:$('cfgPrecio2').value,promoActiva:$('cfgPromo').checked,publicidad}}); showMsg('configMsg','✅ Configuración guardada.',true); await loadAll(); }catch(e){ showMsg('configMsg','❌ '+e.message,false); } }
async function addSeller(){ try{ await api('addSeller',{pin:adminPin,nombre:$('newSeller').value}); $('newSeller').value=''; showMsg('configMsg','✅ Vendedor agregado.',true); await loadAll(); }catch(e){ showMsg('configMsg','❌ '+e.message,false); } }
async function deleteSeller(){ if(!confirm('¿Eliminar vendedor?'))return; try{ await api('deleteSeller',{pin:adminPin,nombre:$('deleteSellerSelect').value}); showMsg('configMsg','✅ Vendedor eliminado.',true); await loadAll(); }catch(e){ showMsg('configMsg','❌ '+e.message,false); } }
async function loadReports(){ try{ const summary=await api('sellerSummary'); $('reportBox').textContent = summary.length ? summary.map(r=>`${r.vendedor}: ${r.cantidad} ventas · $${r.totalEstimado}`).join('\n') : 'Sin ventas.'; }catch(e){ showMsg('reportMsg','❌ '+e.message,false); } }
function downloadSoldCsv(){ const rows=[['Numero','Estado']]; (dashboard.soldNumbers||[]).forEach(n=>rows.push([n,'Vendido'])); downloadText('vendidos.csv',csv(rows)); }
async function downloadSellerCsv(){ const summary=await api('sellerSummary'); const rows=[['Vendedor','Cantidad','Total']].concat(summary.map(r=>[r.vendedor,r.cantidad,r.totalEstimado])); downloadText('resumen_vendedores.csv',csv(rows)); }

function setup(){
  $('enterRifaBtn').onclick=enterRifa; $('rifaCodeInput').addEventListener('keydown',e=>{ if(e.key==='Enter') enterRifa(); });
  $('tChangeRifa').onclick=changeRifa; document.querySelectorAll('.tab[data-tab]').forEach(t=>t.onclick=()=>showTab(t.dataset.tab)); $('refreshBtn').onclick=()=>loadAll(true);
  $('modeSelect').onchange=e=> e.target.value==='admin'?requestAdmin():setPublicMode(); $('confirmPinBtn').onclick=confirmPin; $('cancelPinBtn').onclick=closePin; $('pinInput').addEventListener('keydown',e=>{ if(e.key==='Enter') confirmPin(); });
  $('saleTipoPrecio').onchange=()=>renderSaleForm(); $('confirmSaleBtn').onclick=doSale; $('saleWhatsappBtn').onclick=sendWhatsappReceipt;
  document.querySelectorAll('[data-board]').forEach(b=>b.onclick=()=>setQuick('board',b.dataset.board)); $('renderBoardBtn').onclick=renderBoardRange; $('boardGrid').onclick=e=>{ const n=e.target.dataset.num; if(n) pickNumberFromBoard(n); };
  document.querySelectorAll('[data-pub]').forEach(b=>b.onclick=()=>setQuick('pub',b.dataset.pub)); $('pubImageFileAdmin').onchange=e=>{ pickedBgFile=e.target.files[0]; $('pubImageFileName').textContent=pickedBgFile?pickedBgFile.name:'Ningún archivo elegido.'; }; $('uploadBgBtn').onclick=uploadBackground; $('pubFooterMsg').oninput=saveFooterMsg; $('useGeneralFooterBtn').onclick=useGeneralFooter; $('generateFlyerBtn').onclick=generateFlyer; $('downloadFlyerBtn').onclick=downloadFlyer; $('copyPubMsgBtn').onclick=copyPubMessage;
  $('loadSaleBtn').onclick=loadSaleForDelete; $('delNumero').onchange=loadSaleForDelete; $('deleteSaleBtn').onclick=deleteSale; $('drawBtn').onclick=drawWinners; $('resetWinnersBtn').onclick=resetWinners; $('saveRangeBtn').onclick=saveRange; $('saveConfigBtn').onclick=saveConfig; $('addSellerBtn').onclick=addSeller; $('deleteSellerBtn').onclick=deleteSeller; $('loadReportsBtn').onclick=loadReports; $('downloadSoldCsvBtn').onclick=downloadSoldCsv; $('downloadSellerCsvBtn').onclick=downloadSellerCsv;
  const saved=localStorage.getItem('rifita_codigo'); if(saved){ $('rifaCodeInput').value=saved; $('rememberDeviceCheck').checked=true; }
}
setup();
