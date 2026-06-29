// =====================================================================
//  D&A | MOTOR DE CUSTO INDUSTRIAL  (porte 1:1 da planilha Estrutura)
//  Cada função abaixo replica a coluna correspondente do Excel.
//  Nada foi simplificado: branches Metros/BATIDAS, setup por modelo
//  técnico, perdas acumuladas e impostos por fora/por dentro estão fiéis.
// =====================================================================

const SIM = v => String(v||'').trim().toUpperCase() === 'SIM';
const NAO = v => !SIM(v);
const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const up  = v => String(v||'').trim().toUpperCase();
const round = (v,d) => { const f = Math.pow(10,d); return Math.round(v*f)/f; };

// ---- Normaliza um SKU "cru" (cabeçalhos do Excel) p/ objeto limpo ----
function normalizarSKU(s){
  const M = (n,p) => ({nome: s[n]||'NÃO HÁ', proc: s[p]||'NÃO HÁ'});
  return {
    codigo:   s['Produto'],
    descricao:s['Descrição do Produto'],
    grupo:    s['Grupo do Produto'] || 'ETIQUETAS',
    qtd:      num(s['Quantidade']),
    modelo:   up(s['Modelo Técnico']),
    larg:     num(s['Largura']),
    entreCarr:num(s['Entre Carreiras']),
    esqueleto:num(s['Esqueleto']),
    alt:      num(s['Altura']),
    gap:      num(s['Gap']),
    rep:      num(s['Repetições']),
    carrProd: num(s['Carreiras de Produção']) || 1,
    carrEntrega:num(s['Carreiras de Entrega']) || 1,
    bocasTuro:num(s['Bocas Turo']) || 1,
    qtdModelos:num(s['Qntd. Modelos']) || 1,
    coresCMYK:num(s['Cores (CMYK)']),
    coresBranco:s['Cores (Branco)'] || 'NÃO',
    ocupTinta:num(s['Ocupação Tinta %']),
    cliches:  s['(Clichês) Sim ou Não?'] || 'NÃO',
    vpSN:     s['(Verniz Parcial) Sim ou não?'] || 'NÃO',
    vpTipo:   s['(Verniz Parcial) Tipo'] || 'NÃO',
    vtSN:     s['(Verniz Total) Sim ou não?'] || 'NÃO',
    vtTipo:   s['(Verniz Total) Tipo'] || 'NÃO',
    ocupVerniz:num(s['Ocupação Verniz %']),
    cold:     s['Cold Stamping'] || 'NÃO',
    hot:      s['Hot Stamping'] || 'NÃO',
    sub1:s['Substrato 1']||'NÃO', sub1Larg:num(s['(Substrato 1) Largura']),
    sub2:s['Substrato 2']||'NÃO', sub2Larg:num(s['(Substrato 2) Largura']),
    sub3:s['Substrato 3']||'NÃO', sub3Larg:num(s['(Substrato 3) Largura']),
    ribbon:s['Ribbon']||'NÃO', carrRibbon:num(s['Carreiras Ribbon']) || 1,
    maquinas:[ M('(Máquina 1)','(Máquina 1) Processo'), M('(Máquina 2)','(Máquina 2) Processo'),
               M('(Máquina 3)','(Máquina 3) Processo'), M('(Máquina 4)','(Máquina 4) Processo'),
               M('(Máquina 5)','(Máquina 5) Processo'), M('(Máquina 6)','(Máquina 6) Processo') ],
    faca:     num(s['$ Faca']),
    investEtiq:num(s['Invest. P/ Etiq.']),
    comissao: num(s['% Comissão']),       // já é fração (ex.: 0.03)
    precoUnit:num(s['$ Preço Unit.']),
    nf:       s['NF'] || 'SIM',
    icms:     0.18                         // alíquota cliente (editável p/ SKU genérico)
  };
}

// ---- lookups seguros nas tabelas de referência ----
const maq    = n => REF.maquinas[up(n)] ? null : REF.maquinas[n] || null; // (mantém chaves originais)
function rowMaq(n){ return REF.maquinas[n] || null; }
function rowProc(n){ return REF.processos[n] || null; }
function outros(n){ return REF.outros[up(n)] || {g:0,valor:0}; }
function subPrint(n){ return num(REF.substratos[n]); }
function rowRibbon(n){ return REF.ribbons[n] || null; }

// =====================================================================
//  calcularCustoMaquina  ->  resolve dados de UMA máquina (BB..BK p/ m1,
//  BL..BK equivalentes p/ m2..m6). Retorna constantes de tabela.
//  Faz o override de velocidade/acerto p/ processo HOT STAMPING.
// =====================================================================
function calcularCustoMaquina(maquina){
  const r = rowMaq(maquina.nome);
  if(!r) return null;
  const isHot = up(maquina.proc) === 'HOT STAMPING';
  const proc = rowProc('HOT STAMPING') || {acerto:0,vel:0};
  return {
    nome: maquina.nome,
    proc: maquina.proc,
    fator: r.fator,                                   // BB / BL
    vel:   isHot ? proc.vel : r.vel,                  // BC / BM (override HOT)
    un:    up(r.un),                                  // BD / BN  (METROS|BATIDAS)
    hora:  r.hora,                                    // BE / BO
    acerto:r.acerto,                                  // BF / BP  (acerto por cor)
    perda: r.perda,                                   // BG / BQ  (perda de processo)
    tempoAcerto: isHot ? proc.acerto : r.tempoAcerto, // BH / BR
    tempoMin: r.tempoMin,                             // col J  (h)
    metrosMin:r.metrosMin                             // col K  (m)
  };
}

// =====================================================================
//  calcularConsumoMaterial -> DJ, DK, DL, DM, DN, DO
//  DJ perda acumulada | DK acerto(setup) em metros | DL m/l c/ setup+perda
//  DN m/l líquido do produto | DM/DO m²
// =====================================================================
function calcularConsumoMaterial(p, mc){
  // DJ = soma das perdas de processo das 6 máquinas (DE+CU+CK+CA+BQ+BG)
  const DJ = mc.reduce((a,m)=> a + (m? m.perda:0), 0);

  // acertos por cor de cada máquina (BF,BP,BZ,CJ,CT,DD)
  const ac = i => (mc[i] ? mc[i].acerto : 0);
  const acExtras = ac(1)+ac(2)+ac(3)+ac(4)+ac(5);
  let DK;
  if(p.modelo === 'FLEXOGRÁFICO'){
    const cores = (p.coresCMYK + (SIM(p.vpSN)?1:0) + (SIM(p.vtSN)?1:0) + (SIM(p.cold)?1:0)) * p.qtdModelos;
    DK = (cores * ac(0)) + acExtras;
  } else if(p.modelo === 'DIGITAL'){
    DK = (p.qtdModelos > 1) ? (ac(0) + (ac(0)/4)*(p.qtdModelos-1)) + acExtras : ac(0) + acExtras;
  } else { // SEM IMPRESSÃO (e fallback)
    DK = ac(0) + acExtras;
  }

  const passo = ((p.alt + p.gap)/1000) * p.qtd / p.carrProd; // metros lineares brutos
  const DL = (passo / (1 - DJ)) + DK;          // DL  (m/l setup+perda)
  const DN = DL - DK;                          // DN  (m/l do produto, base p/ velocidade)
  const DM = DL * (p.sub1Larg/1000);           // DM  (m² setup+perda)
  const DO = passo * (p.sub1Larg/1000);        // DO  (m² produto)
  return {DJ, DK, DL, DM, DN, DO};
}

// ---- tempo total + custo de cada máquina (BJ/BK e equivalentes) ----
function tempoEcusto(m, idx, p, cons){
  if(!m) return {setup:0, tempo:0, custo:0};
  const {DN, DJ} = cons;
  // setup (BI p/ m1 / BS p/ demais)
  let setup;
  if(idx === 0 && p.modelo === 'FLEXOGRÁFICO'){
    const cores = (p.coresCMYK + (SIM(p.vpSN)?1:0) + (SIM(p.cold)?1:0) + (SIM(p.vtSN)?1:0)) * p.qtdModelos;
    setup = cores * m.tempoAcerto * m.fator;
  } else {
    setup = m.tempoAcerto * m.fator;
  }
  // m1 usa a própria perda; m2..m6 usam DJ (perda acumulada) — fiel ao Excel
  const perdaBatida = idx === 0 ? m.perda : DJ;
  let tempo = 0;
  if(m.un === 'METROS'){
    tempo = (DN < m.metrosMin) ? (m.tempoMin*m.fator)+setup
                               : round(DN/m.vel,3)*m.fator + setup;
  } else if(m.un === 'BATIDAS'){
    const batidas = p.qtd / p.bocasTuro;
    tempo = (batidas < m.metrosMin) ? (m.tempoMin*m.fator)+setup
          : round(((p.qtd/(1-perdaBatida))/p.bocasTuro/m.vel)*m.fator,3) + setup;
  }
  if(!isFinite(tempo)) tempo = 0;
  return {setup, tempo, custo: tempo*m.hora};
}

// =====================================================================
//  calcularCustoMateriaPrima -> EN  (+ detalhamento por insumo)
// =====================================================================
function calcularCustoMateriaPrima(p, cons){
  const {DM} = cons;
  const Z = ((p.alt + p.gap) * p.rep)/3.175;                 // AK (dentes -> circunferência)

  const DQ = subPrint(p.sub1) * DM;                          // substrato 1
  const DS = subPrint(p.sub2) * DM;                          // substrato 2
  const DU = subPrint(p.sub3) * DM;                          // substrato 3

  // tinta (DV/DW) — ramos DIGITAL e FLEXOGRÁFICO
  let DV = 0, DW = 0;
  if(p.modelo === 'DIGITAL'){
    DV = outros('INSUMO BOBST (CMYK)').valor + (SIM(p.coresBranco) ? outros('INSUMO BOBST (BRANCO)').valor : 0);
    DW = DV * DM;
  } else if(p.modelo === 'FLEXOGRÁFICO'){
    const tf = outros('TINTA FLEXOGRÁFICA');
    DW = ((tf.valor * tf.g) * p.ocupTinta) * DM;
  }

  // clichês (DX/DY/DZ) — área da arte em cm²
  let DX = 0;
  if(SIM(p.cliches)) DX = (SIM(p.vpSN)?1:0) + (SIM(p.cold)?1:0) + (p.qtdModelos*p.coresCMYK);
  else if(p.modelo === 'DIGITAL') DX = (SIM(p.vpSN)?1:0) + (SIM(p.cold)?1:0);
  const DY = DX === 0 ? 0 : outros('CLICHÊ').valor;
  const larguraArte = Math.floor((((p.larg*p.carrProd)+(p.entreCarr*(p.carrProd-1))+p.esqueleto)/10)+4);
  const alturaArte  = Math.round(((Z*3.175)/10)+4);
  const DZ = DX === 0 ? 0 : (larguraArte * alturaArte) * DX * DY;

  // verniz parcial / total
  const EA = NAO(p.vpSN) ? 0 : outros(p.vpTipo).g;
  const EB = EA === 0 ? 0 : outros(p.vpTipo).valor * EA * DM * p.ocupVerniz;
  const EC = NAO(p.vtSN) ? 0 : outros(p.vtTipo).g;
  const ED = EC === 0 ? 0 : outros(p.vtTipo).valor * EC * DM * p.ocupVerniz;

  // cold / hot stamping
  const EE = NAO(p.cold) ? 0 : outros('COLD STAMPING').valor;
  const EF = NAO(p.cold) ? 0 : (EE*DM) + (DM * outros('VERNIZ COLD STAMPING').g * outros('VERNIZ COLD STAMPING').valor);
  const EG = NAO(p.hot) ? 0 : outros('HOT STAMPING').valor;
  const EH = EG * DM;

  // ribbon
  const rb = rowRibbon(p.ribbon);
  const EI = (NAO(p.ribbon) || !rb) ? 0 : rb.altura;                 // "metros" (altura mm)
  let EJ = 0;
  if(EI > 0) EJ = Math.ceil((((p.alt+p.gap)/1000)/p.carrRibbon * p.qtd)/EI);
  const EK = (NAO(p.ribbon) || !rb) ? 0 : rb.unit;
  const EL = EK * EJ;

  const EN = DQ+DS+DU+DW+DZ+EB+ED+EF+EH+EL;
  return {EN, det:{DQ,DS,DU,DW,DZ,EB,ED,EF,EH,EL}};
}

// =====================================================================
//  calcularImpostos -> ER (PIS/COFINS por fora), ET outros, EU ICMS, EV frete
// =====================================================================
function calcularImpostos(p, EQ){
  const ER = SIM(p.nf) ? 0.0365 : 0;       // PIS/COFINS "por fora"
  const ES = EQ/(1-ER);                     // custo c/ PIS/COFINS embutido
  const ET = SIM(p.nf) ? 0.0228 : 0;       // outros impostos "por dentro"
  const EU = SIM(p.nf) ? p.icms : 0;       // ICMS por cliente
  const EV = 0.025;                         // frete + tubete (fixo)
  return {ER, ES, ET, EU, EV};
}

// calcularPrecoVenda -> EY = quantidade * preço unitário
function calcularPrecoVenda(p){ return p.qtd * p.precoUnit; }

// calcularLucro -> EW (% lucro) e EX (% lucro s/ investimento)
function calcularLucro(p, EY, base, imp){
  const {EN,EO,EP,EM,ES,ER} = base;
  const taxa = imp.ET + imp.EU + imp.EV + p.comissao;
  const EW = EY === 0 ? 0 : (EY - (ES + taxa*EY)) / EY;
  const EX = EY === 0 ? 0 : (EY - ((EN+EO+EP)/(1-ER) + taxa*EY)) / EY;  // sem investimento na base
  return {EW, EX, lucroRS: EY*EW};
}

// calcularMargem -> EZ ($) e FA (%)
function calcularMargem(p, EY, base, imp){
  const {EN,EO,EM,ES,ER} = base;
  const EZ = (EY-EO-EN-EM) - (ES*ER) - ((imp.ET+imp.EU+imp.EV)*EY) - (p.comissao*EY);
  return {EZ, FA: EY===0?0:EZ/EY};
}

// =====================================================================
//  ORQUESTRADOR — calcularCustoTotalProduto
//  override: {qtd, precoUnit, icms}  (usado p/ simular venda em outra qtd)
// =====================================================================
function calcularCustoTotalProduto(sku, override){
  const p = {...sku};
  if(override){
    if(override.qtd != null)       p.qtd = num(override.qtd);
    if(override.precoUnit != null) p.precoUnit = num(override.precoUnit);
    if(override.icms != null)      p.icms = num(override.icms);
  }
  // 1) constantes de cada máquina
  const mc = p.maquinas.map(calcularCustoMaquina);
  // 2) consumo de material
  const cons = calcularConsumoMaterial(p, mc);
  // 3) tempo/custo por máquina (depende de DN)
  const mt = mc.map((m,i)=> tempoEcusto(m,i,p,cons));
  const EP = mt.slice(0,5).reduce((a,x)=>a+x.custo,0);   // MO = máquinas 1..5 (fiel)
  // 4) matéria-prima
  const {EN, det} = calcularCustoMateriaPrima(p, cons);
  // 5) totais
  const EO = p.faca;                 // terceiros
  const EM = p.investEtiq * p.qtd;   // investimento
  const EQ = EM + EN + EO + EP;      // total custos
  // 6) impostos + resultados
  const imp = calcularImpostos(p, EQ);
  const EY  = calcularPrecoVenda(p);
  const base = {EN,EO,EP,EM,EQ,ES:imp.ES,ER:imp.ER};
  const luc = calcularLucro(p, EY, base, imp);
  const mar = calcularMargem(p, EY, base, imp);

  return {
    consumo: cons, maquinas: mc.map((m,i)=> m && {...m, ...mt[i]}),
    mp: {total: EN, det},
    EM, EO, EP, EN, EQ, EY,
    imp,
    precoTotal: EY, custoTotal: EQ,
    lucroRS: luc.lucroRS, lucroPct: luc.EW, lucroSInvest: luc.EX,
    mcRS: mar.EZ, mcPct: mar.FA,
    p
  };
}

// ---------- formatação ----------
const fmtBRL = v => (v||0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const fmtPct = v => ((v||0)*100).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'%';
const fmtNum = v => (v||0).toLocaleString('pt-BR',{maximumFractionDigits:0});
// =====================================================================
//  D&A | APLICAÇÃO (carregamento externo + importação em massa)
//  Fonte da verdade: dados.json (referência+SKUs) e vendas.json (diário).
//  Membros da empresa abrem em modo visualização (só fetch, sem escrita).
// =====================================================================

// ---- CONFIGURAÇÃO ----------------------------------------------------
// Para o modo COMPARTILHADO, aponte para as URLs raw do seu repo público
// (mesmo padrão do PCP). Ex.:
//   base: 'https://raw.githubusercontent.com/lhtorres1995/pcp-sync/main/'
const CONFIG = {
  base: '',                       // prefixo das URLs (vazio = arquivos locais)
  dadosUrl: 'dados.json',         // referência + SKUs
  vendasUrl: 'vendas.json',       // base de vendas (diária)
  editavel: true                  // false = modo visualização p/ membros (oculta cadastro/importação)
};

let REF = {maquinas:{},processos:{},outros:{},substratos:{},ribbons:{}};
let produtos = [];
let vendas   = [];

const findSku = cod => produtos.find(p=>p.codigo===cod);

// financeiro de uma venda: usa o valor PRÉ-CALCULADO; se faltar, cai no motor
function fin(v){
  if(v.receita!=null && v.lucro!=null && v.custo!=null)
    return {receita:v.receita,custo:v.custo,lucro:v.lucro,lucroPct:v.lucroPct||0,mc:v.mc||0,mcPct:v.mcPct||0,fonte:'base'};
  const s = findSku(v.produto);
  if(s){ const e=calcularCustoTotalProduto(s,{qtd:v.qtd,precoUnit:v.preco});
    return {receita:e.precoTotal,custo:e.custoTotal,lucro:e.lucroRS,lucroPct:e.lucroPct,mc:e.mcRS,mcPct:e.mcPct,fonte:'motor'}; }
  return {receita:(v.qtd*v.preco)||0,custo:null,lucro:0,lucroPct:0,mc:0,mcPct:0,fonte:'sem custo'};
}

// ---- carregamento ----------------------------------------------------
async function fetchJson(url){
  const r = await fetch(CONFIG.base+url+'?t='+Date.now(), {cache:'no-store'});
  if(!r.ok) throw new Error(url+' → HTTP '+r.status);
  return r.json();
}
async function boot(){
  const root=document.getElementById('content');
  root.innerHTML='<div class="loading">Carregando dados…</div>';
  try{
    const [d,v] = await Promise.all([fetchJson(CONFIG.dadosUrl), fetchJson(CONFIG.vendasUrl)]);
    REF = {maquinas:d.maquinas||{},processos:d.processos||{},outros:d.outros||{},substratos:d.substratos||{},ribbons:d.ribbons||{}};
    produtos = (d.skus||[]).map(normalizarSKU);
    vendas   = v||[];
  }catch(e){
    root.innerHTML=`<div class="erro"><h2>Não foi possível carregar os dados</h2>
      <p>${e.message}</p><p class="sub">Verifique se <code>${CONFIG.dadosUrl}</code> e <code>${CONFIG.vendasUrl}</code>
      estão acessíveis na URL configurada. Em modo compartilhado, ajuste <code>CONFIG.base</code> para o raw do repositório.</p></div>`;
    return;
  }
  configurarNav();
  render();
}

function configurarNav(){
  // oculta itens de edição no modo visualização
  document.querySelectorAll('[data-edit]').forEach(n=>n.style.display = CONFIG.editavel?'':'none');
  document.querySelectorAll('.nav-item').forEach(n=>n.onclick=()=>go(n.dataset.view));
  const banner=document.getElementById('viewer-banner');
  if(banner) banner.style.display = CONFIG.editavel?'none':'flex';
}

// ---- download (publicar) ---------------------------------------------
function baixar(nome,obj){
  const blob=new Blob([JSON.stringify(obj)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=nome;
  document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(a.href);
}
function publicarVendas(){
  // re-calcula financeiro p/ congelar na base publicada
  const out=vendas.map(v=>{ const f=fin(v); return {...v,
    receita:+ (f.receita||0).toFixed(2), custo:f.custo==null?null:+f.custo.toFixed(2),
    lucro:+(f.lucro||0).toFixed(2), lucroPct:+(f.lucroPct||0).toFixed(4),
    mc:+(f.mc||0).toFixed(2), mcPct:+(f.mcPct||0).toFixed(4)}; });
  baixar('vendas.json', out);
}
function publicarDados(){
  baixar('dados.json', {maquinas:REF.maquinas,processos:REF.processos,outros:REF.outros,
    substratos:REF.substratos,ribbons:REF.ribbons, skus:produtos.map(desnormalizarSKU)});
}

// ---- navegação -------------------------------------------------------
let view='dashboard';
function go(v){ view=v; render(); }
function render(){
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.toggle('active',n.dataset.view===view));
  const root=document.getElementById('content');
  if(view==='dashboard') renderDashboard(root);
  else if(view==='produtos') renderProdutos(root);
  else if(view==='vendas') renderVendas(root);
  else if(view==='importar') renderImportar(root);
}

// =====================================================================
//  DASHBOARD  (lê financeiro pronto da base)
// =====================================================================
let charts={}; let filtro={ini:null,fim:null,cliente:'',vendedor:''};
function vendasFiltradas(){
  return vendas.filter(v=>{
    if(filtro.ini && v.data<filtro.ini) return false;
    if(filtro.fim && v.data>filtro.fim) return false;
    if(filtro.cliente && v.familia!==filtro.cliente) return false;
    if(filtro.vendedor && v.vendedor!==filtro.vendedor) return false;
    return true;
  });
}
function renderDashboard(root){
  const vs=vendasFiltradas();
  let receita=0,lucro=0,mc=0,custo=0,semCusto=0;
  const porMes={},porVend={},porGrupo={},porCliente={},porSeg={};
  vs.forEach(v=>{ const f=fin(v);
    receita+=f.receita; lucro+=f.lucro; mc+=f.mc; custo+=(f.custo||0);
    if(f.fonte==='sem custo') semCusto++;
    const add=(o,k)=>{o[k]=o[k]||{r:0,l:0};o[k].r+=f.receita;o[k].l+=f.lucro;};
    add(porMes,v.data.slice(0,7)); add(porVend,v.vendedor||'—'); add(porGrupo,v.grupo||'—');
    add(porCliente,v.familia||v.cliente||'—'); add(porSeg,v.segmento||'—');
  });
  const pctLucro=receita?lucro/receita:0, pctMC=receita?mc/receita:0;
  const clientes=[...new Set(vendas.map(v=>v.familia).filter(Boolean))].sort();
  const vendedores=[...new Set(vendas.map(v=>v.vendedor).filter(Boolean))].sort();

  root.innerHTML=`
    <div class="page-head"><div><h1>Dashboard comercial</h1>
      <p class="sub">${vs.length} pedidos${semCusto?` · ${semCusto} sem custo na base`:''}</p></div></div>
    <div class="filtros">
      <label>De <input type="date" id="f-ini" value="${filtro.ini||''}"></label>
      <label>Até <input type="date" id="f-fim" value="${filtro.fim||''}"></label>
      <label>Cliente <select id="f-cli"><option value="">Todos</option>${clientes.map(c=>`<option ${filtro.cliente===c?'selected':''}>${c}</option>`).join('')}</select></label>
      <label>Vendedor <select id="f-vend"><option value="">Todos</option>${vendedores.map(c=>`<option ${filtro.vendedor===c?'selected':''}>${c}</option>`).join('')}</select></label>
      <button class="btn-ghost" id="f-limpar">Limpar</button>
    </div>
    <div class="kpis">
      ${kpi('Receita total',fmtBRL(receita),'accent')}
      ${kpi('Lucro total',fmtBRL(lucro))}
      ${kpi('% Lucro',fmtPct(pctLucro))}
      ${kpi('Margem de contribuição',fmtBRL(mc))}
      ${kpi('% MC',fmtPct(pctMC))}
      ${kpi('Custo total',fmtBRL(custo))}
    </div>
    <div class="grid-2">
      <div class="card"><h3>Vendas por mês</h3><canvas id="c-mes"></canvas></div>
      <div class="card"><h3>Resultado por grupo de produto</h3><canvas id="c-grupo"></canvas></div>
      <div class="card"><h3>Vendas por vendedor</h3><canvas id="c-vend"></canvas></div>
      <div class="card"><h3>% Lucro por mês</h3><canvas id="c-pctmes"></canvas></div>
      <div class="card"><h3>Top 5 clientes</h3>${tabelaTop(porCliente,'Família')}</div>
      <div class="card"><h3>Top 5 segmentos</h3>${tabelaTop(porSeg,'Segmento')}</div>
    </div>`;
  const upd=()=>{filtro.ini=val('f-ini');filtro.fim=val('f-fim');filtro.cliente=val('f-cli');filtro.vendedor=val('f-vend');render();};
  ['f-ini','f-fim','f-cli','f-vend'].forEach(id=>document.getElementById(id).addEventListener('change',upd));
  document.getElementById('f-limpar').onclick=()=>{filtro={ini:null,fim:null,cliente:'',vendedor:''};render();};
  const RED='#C8102E',GREY='#4b5563';
  const mk=Object.keys(porMes).sort();
  mkChart('c-mes','bar',mk.map(rotuloMes),[{data:mk.map(m=>porMes[m].r),backgroundColor:RED}]);
  const gk=Object.keys(porGrupo).sort((a,b)=>porGrupo[b].r-porGrupo[a].r);
  mkChart('c-grupo','bar',gk,[{data:gk.map(k=>porGrupo[k].r),backgroundColor:RED}],true);
  const vk=Object.keys(porVend).sort((a,b)=>porVend[b].r-porVend[a].r).slice(0,8);
  mkChart('c-vend','bar',vk,[{data:vk.map(k=>porVend[k].r),backgroundColor:GREY}],true);
  mkChart('c-pctmes','line',mk.map(rotuloMes),[{data:mk.map(m=>porMes[m].r?100*porMes[m].l/porMes[m].r:0),borderColor:RED,backgroundColor:'rgba(200,16,46,.1)',tension:.3,fill:true}],false,true);
}
const val=id=>document.getElementById(id).value;
function kpi(l,v,c=''){return `<div class="kpi ${c}"><span class="kpi-label">${l}</span><span class="kpi-value">${v}</span></div>`;}
function rotuloMes(m){const[y,mm]=m.split('-');return['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'][+mm-1]+'/'+y.slice(2);}
function tabelaTop(o,label){const k=Object.keys(o).sort((a,b)=>o[b].r-o[a].r).slice(0,5);
  return `<table class="mini"><thead><tr><th>${label}</th><th class="r">R$ Vendido</th><th class="r">% Lucro</th></tr></thead><tbody>${
    k.map(x=>`<tr><td>${x}</td><td class="r">${fmtBRL(o[x].r)}</td><td class="r">${fmtPct(o[x].r?o[x].l/o[x].r:0)}</td></tr>`).join('')}</tbody></table>`;}
function mkChart(id,type,labels,datasets,horizontal,pct){
  if(charts[id])charts[id].destroy();
  charts[id]=new Chart(document.getElementById(id),{type,data:{labels,datasets},options:{
    indexAxis:horizontal?'y':'x',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
    scales:{x:{ticks:{callback:v=>(type==='line'||pct)?v+'%':abbr(v),font:{size:10}},grid:{display:!horizontal}},
            y:{ticks:{callback:v=>horizontal?undefined:(pct?v+'%':abbr(v)),font:{size:10}},grid:{display:horizontal}}}}});
}
function abbr(v){if(typeof v!=='number')return v;if(Math.abs(v)>=1e6)return 'R$'+(v/1e6).toFixed(1)+'M';if(Math.abs(v)>=1e3)return 'R$'+(v/1e3).toFixed(0)+'k';return v;}

// =====================================================================
//  PRODUTOS  (simulador / motor de custo)
// =====================================================================
function renderProdutos(root){
  root.innerHTML=`
    <div class="page-head"><div><h1>Produtos (SKUs)</h1><p class="sub">${produtos.length} cadastrados · custo pela estrutura técnica</p></div>
      <div class="head-actions">
        <button class="btn-ghost" data-edit id="pub-dados">Baixar dados.json</button>
        <button class="btn" data-edit id="novo-sku">+ Novo SKU</button></div></div>
    <div class="card no-pad"><table class="full"><thead><tr>
      <th>Código</th><th>Descrição</th><th>Grupo</th><th>Modelo</th><th class="r">Qtd</th>
      <th class="r">Custo total</th><th class="r">Preço total</th><th class="r">% Lucro</th><th class="r">% MC</th></tr></thead>
      <tbody>${produtos.map(p=>{const e=calcularCustoTotalProduto(p);return `
        <tr class="clickable" data-cod="${p.codigo}"><td class="mono">${p.codigo}</td><td>${p.descricao}</td>
        <td><span class="tag">${p.grupo}</span></td><td>${p.modelo}</td><td class="r">${fmtNum(p.qtd)}</td>
        <td class="r">${fmtBRL(e.custoTotal)}</td><td class="r">${fmtBRL(e.precoTotal)}</td>
        <td class="r ${e.lucroPct<0.18?'neg':'pos'}">${fmtPct(e.lucroPct)}</td><td class="r">${fmtPct(e.mcPct)}</td></tr>`;}).join('')}
      </tbody></table></div>`;
  document.querySelectorAll('.clickable').forEach(tr=>tr.onclick=()=>detalheProduto(tr.dataset.cod));
  if(CONFIG.editavel){
    document.getElementById('novo-sku').onclick=()=>formProduto();
    document.getElementById('pub-dados').onclick=publicarDados;
  }
  configurarNav();
}
function detalheProduto(cod){
  const p=findSku(cod), e=calcularCustoTotalProduto(p), root=document.getElementById('content');
  const linha=(l,v,s)=>`<tr class="${s?'strong':''}"><td>${l}</td><td class="r">${v}</td></tr>`;
  const maqRows=e.maquinas.map(m=>m?`<tr><td>${m.nome}</td><td>${m.un}</td><td class="r">${m.vel}</td><td class="r">${m.tempo.toFixed(3)}h</td><td class="r">${fmtBRL(m.custo)}</td></tr>`:'').join('');
  const d=e.mp.det;
  const mpRows=[['Substrato 1',d.DQ],['Substrato 2',d.DS],['Substrato 3',d.DU],['Tinta',d.DW],['Clichês',d.DZ],['Verniz parcial',d.EB],['Verniz total',d.ED],['Cold stamping',d.EF],['Hot stamping',d.EH],['Ribbon',d.EL]]
    .filter(x=>x[1]>0).map(x=>`<tr><td>${x[0]}</td><td class="r">${fmtBRL(x[1])}</td></tr>`).join('');
  root.innerHTML=`
    <div class="page-head"><div><button class="btn-ghost" id="voltar">← Produtos</button>
      <h1>${p.descricao}</h1><p class="sub mono">${p.codigo} · ${p.grupo} · ${p.modelo}</p></div>
      <button class="btn" data-edit id="editar">Editar SKU</button></div>
    <div class="kpis">${kpi('Custo total',fmtBRL(e.custoTotal),'accent')}${kpi('Preço total',fmtBRL(e.precoTotal))}
      ${kpi('Lucro',fmtBRL(e.lucroRS))}${kpi('% Lucro',fmtPct(e.lucroPct))}${kpi('% Lucro s/ invest.',fmtPct(e.lucroSInvest))}${kpi('% MC',fmtPct(e.mcPct))}</div>
    <div class="grid-led">
      <div class="card"><h3>Consumo de material</h3><table class="led">
        ${linha('M/L produto (DN)',fmtNum(e.consumo.DN)+' m')}${linha('M/L setup + perda (DL)',fmtNum(e.consumo.DL)+' m')}
        ${linha('M² setup + perda (DM)',e.consumo.DM.toFixed(1)+' m²')}${linha('Perda acumulada (DJ)',fmtPct(e.consumo.DJ))}
        ${linha('Acerto / setup (DK)',fmtNum(e.consumo.DK)+' m')}</table></div>
      <div class="card"><h3>Matéria-prima · ${fmtBRL(e.mp.total)}</h3><table class="led">${mpRows||'<tr><td>—</td></tr>'}</table></div>
      <div class="card"><h3>Mão de obra (máquinas) · ${fmtBRL(e.EP)}</h3>
        <table class="led"><thead><tr><th>Máquina</th><th>Un.</th><th class="r">Vel.</th><th class="r">Tempo</th><th class="r">Custo</th></tr></thead>${maqRows}</table></div>
      <div class="card"><h3>Composição do preço</h3><table class="led">
        ${linha('Matéria-prima',fmtBRL(e.EN))}${linha('Mão de obra',fmtBRL(e.EP))}${linha('Terceiros (faca)',fmtBRL(e.EO))}
        ${linha('Investimento',fmtBRL(e.EM))}${linha('Total de custos',fmtBRL(e.EQ),true)}
        ${linha('PIS/COFINS (por fora)',fmtPct(e.imp.ER))}${linha('Outros impostos',fmtPct(e.imp.ET))}
        ${linha('ICMS',fmtPct(e.imp.EU))}${linha('Frete + tubete',fmtPct(e.imp.EV))}${linha('Comissão',fmtPct(p.comissao))}
        ${linha('Preço total',fmtBRL(e.precoTotal),true)}${linha('Margem de contribuição',fmtBRL(e.mcRS))}</table></div>
    </div>`;
  document.getElementById('voltar').onclick=()=>go('produtos');
  if(CONFIG.editavel) document.getElementById('editar').onclick=()=>formProduto(p);
  configurarNav();
}

// reconstrói o objeto "cru" (cabeçalhos Excel) a partir do SKU normalizado
function desnormalizarSKU(p){
  return {'Produto':p.codigo,'Descrição do Produto':p.descricao,'Grupo do Produto':p.grupo,'Quantidade':p.qtd,
    'Modelo Técnico':p.modelo,'Largura':p.larg,'Entre Carreiras':p.entreCarr,'Esqueleto':p.esqueleto,'Altura':p.alt,
    'Gap':p.gap,'Repetições':p.rep,'Carreiras de Produção':p.carrProd,'Carreiras de Entrega':p.carrEntrega,
    'Bocas Turo':p.bocasTuro,'Qntd. Modelos':p.qtdModelos,'Cores (CMYK)':p.coresCMYK,'Cores (Branco)':p.coresBranco,
    'Ocupação Tinta %':p.ocupTinta,'(Clichês) Sim ou Não?':p.cliches,'(Verniz Parcial) Sim ou não?':p.vpSN,
    '(Verniz Parcial) Tipo':p.vpTipo,'(Verniz Total) Sim ou não?':p.vtSN,'(Verniz Total) Tipo':p.vtTipo,
    'Ocupação Verniz %':p.ocupVerniz,'Cold Stamping':p.cold,'Hot Stamping':p.hot,
    'Substrato 1':p.sub1,'(Substrato 1) Largura':p.sub1Larg,'Substrato 2':p.sub2,'(Substrato 2) Largura':p.sub2Larg,
    'Substrato 3':p.sub3,'(Substrato 3) Largura':p.sub3Larg,'Ribbon':p.ribbon,'Carreiras Ribbon':p.carrRibbon,
    '(Máquina 1)':p.maquinas[0].nome,'(Máquina 2)':p.maquinas[1].nome,'(Máquina 3)':p.maquinas[2].nome,
    '(Máquina 4)':p.maquinas[3].nome,'(Máquina 5)':p.maquinas[4].nome,'(Máquina 6)':p.maquinas[5].nome,
    '$ Faca':p.faca,'Invest. P/ Etiq.':p.investEtiq,'% Comissão':p.comissao,'$ Preço Unit.':p.precoUnit,'NF':p.nf};
}

function formProduto(edit){
  const p=edit||{maquinas:[]}, root=document.getElementById('content');
  const f=(l,k,t='number',v='')=>`<label class="fld"><span>${l}</span><input name="${k}" type="${t}" value="${edit?(p[k]??''):v}"></label>`;
  const sel=(l,k,opts,v='')=>`<label class="fld"><span>${l}</span><select name="${k}">${opts.map(o=>`<option ${(edit?p[k]:v)===o?'selected':''}>${o}</option>`).join('')}</select></label>`;
  const sn=(l,k)=>sel(l,k,['NÃO','SIM']);
  const maqOpts=['NÃO HÁ',...Object.keys(REF.maquinas).filter(x=>x!=='NÃO HÁ')];
  const subOpts=['NÃO',...Object.keys(REF.substratos)], ribOpts=['NÃO',...Object.keys(REF.ribbons)];
  root.innerHTML=`
    <div class="page-head"><div><button class="btn-ghost" id="voltar">← Produtos</button><h1>${edit?'Editar SKU':'Novo SKU'}</h1></div></div>
    <form id="form-sku" class="form">
      <fieldset><legend>Identificação</legend><div class="fgrid">
        <label class="fld"><span>Código</span><input name="codigo" value="${edit?p.codigo:''}"></label>
        <label class="fld wide"><span>Descrição</span><input name="descricao" value="${edit?p.descricao:''}"></label>
        ${sel('Grupo','grupo',['ETIQUETAS','RÓTULOS','COMODATO','RIBBON','PEÇA','LOCAÇÃO'])}
        ${sel('Modelo técnico','modelo',['SEM IMPRESSÃO','FLEXOGRÁFICO','DIGITAL'])}${f('Quantidade','qtd')}</div></fieldset>
      <fieldset><legend>Dados técnicos</legend><div class="fgrid">
        ${f('Largura (mm)','larg')}${f('Altura (mm)','alt')}${f('Gap (mm)','gap')}${f('Entre carreiras','entreCarr')}
        ${f('Esqueleto','esqueleto')}${f('Repetições','rep')}${f('Carreiras produção','carrProd')}${f('Carreiras entrega','carrEntrega')}
        ${f('Bocas Turo','bocasTuro')}${f('Qntd. modelos','qtdModelos')}${f('Cores CMYK','coresCMYK')}${sn('Cores branco','coresBranco')}${f('Ocupação tinta %','ocupTinta')}</div></fieldset>
      <fieldset><legend>Substratos & ribbon</legend><div class="fgrid">
        ${sel('Substrato 1','sub1',subOpts)}${f('Largura sub 1','sub1Larg')}${sel('Substrato 2','sub2',subOpts)}${f('Largura sub 2','sub2Larg')}
        ${sel('Substrato 3','sub3',subOpts)}${f('Largura sub 3','sub3Larg')}${sel('Ribbon','ribbon',ribOpts)}${f('Carreiras ribbon','carrRibbon')}</div></fieldset>
      <fieldset><legend>Acabamentos</legend><div class="fgrid">
        ${sn('Clichês','cliches')}${sn('Cold stamping','cold')}${sn('Hot stamping','hot')}${sn('Verniz parcial','vpSN')}${f('Tipo verniz parcial','vpTipo','text')}
        ${sn('Verniz total','vtSN')}${f('Tipo verniz total','vtTipo','text')}${f('Ocupação verniz %','ocupVerniz')}</div></fieldset>
      <fieldset><legend>Processo (máquinas)</legend><div class="fgrid">
        ${[0,1,2,3,4,5].map(i=>sel('Máquina '+(i+1),'maq'+i,maqOpts)).join('')}</div></fieldset>
      <fieldset><legend>Comercial</legend><div class="fgrid">
        ${f('$ Faca','faca')}${f('Invest. p/ etiq.','investEtiq')}${f('Comissão (fração, ex 0.03)','comissao')}
        ${f('Preço unitário','precoUnit')}${f('ICMS (fração, ex 0.18)','icms')}${sn('Emite NF','nf')}</div></fieldset>
      <div class="form-actions"><button type="button" class="btn-ghost" id="cancelar">Cancelar</button><button type="button" class="btn" id="salvar-sku">Salvar SKU</button></div>
    </form>`;
  document.getElementById('voltar').onclick=()=>go('produtos');
  document.getElementById('cancelar').onclick=()=>go('produtos');
  document.getElementById('salvar-sku').onclick=()=>{
    const fd=new FormData(document.getElementById('form-sku')), g=k=>fd.get(k);
    const novo=normalizarSKU({'Produto':g('codigo'),'Descrição do Produto':g('descricao'),'Grupo do Produto':g('grupo'),
      'Quantidade':g('qtd'),'Modelo Técnico':g('modelo'),'Largura':g('larg'),'Entre Carreiras':g('entreCarr'),
      'Esqueleto':g('esqueleto'),'Altura':g('alt'),'Gap':g('gap'),'Repetições':g('rep'),'Carreiras de Produção':g('carrProd'),
      'Carreiras de Entrega':g('carrEntrega'),'Bocas Turo':g('bocasTuro'),'Qntd. Modelos':g('qtdModelos'),
      'Cores (CMYK)':g('coresCMYK'),'Cores (Branco)':g('coresBranco'),'Ocupação Tinta %':g('ocupTinta'),
      '(Clichês) Sim ou Não?':g('cliches'),'(Verniz Parcial) Sim ou não?':g('vpSN'),'(Verniz Parcial) Tipo':g('vpTipo'),
      '(Verniz Total) Sim ou não?':g('vtSN'),'(Verniz Total) Tipo':g('vtTipo'),'Ocupação Verniz %':g('ocupVerniz'),
      'Cold Stamping':g('cold'),'Hot Stamping':g('hot'),'Substrato 1':g('sub1'),'(Substrato 1) Largura':g('sub1Larg'),
      'Substrato 2':g('sub2'),'(Substrato 2) Largura':g('sub2Larg'),'Substrato 3':g('sub3'),'(Substrato 3) Largura':g('sub3Larg'),
      'Ribbon':g('ribbon'),'Carreiras Ribbon':g('carrRibbon'),'(Máquina 1)':g('maq0'),'(Máquina 2)':g('maq1'),
      '(Máquina 3)':g('maq2'),'(Máquina 4)':g('maq3'),'(Máquina 5)':g('maq4'),'(Máquina 6)':g('maq5'),
      '$ Faca':g('faca'),'Invest. P/ Etiq.':g('investEtiq'),'% Comissão':g('comissao'),'$ Preço Unit.':g('precoUnit'),'NF':g('nf')});
    novo.icms=num(g('icms'))||0.18;
    if(edit){const i=produtos.findIndex(x=>x.codigo===edit.codigo);produtos[i]=novo;} else produtos.push(novo);
    alert('SKU salvo em memória. Use "Baixar dados.json" para publicar a estrutura atualizada.');
    go('produtos');
  };
}

// =====================================================================
//  VENDAS
// =====================================================================
function renderVendas(root){
  root.innerHTML=`
    <div class="page-head"><div><h1>Vendas</h1><p class="sub">${vendas.length} pedidos na base</p></div>
      <div class="head-actions">
        <button class="btn-ghost" data-edit id="pub-vendas">Baixar vendas.json</button>
        <button class="btn-ghost" data-edit id="ir-importar">Importar do ERP</button>
        <button class="btn" data-edit id="nova-venda">+ Nova venda</button></div></div>
    <div class="card no-pad"><table class="full"><thead><tr>
      <th>Pedido</th><th>Data</th><th>Cliente</th><th>Vendedor</th><th>SKU</th><th class="r">Qtd</th>
      <th class="r">Receita</th><th class="r">Custo</th><th class="r">Lucro</th><th class="r">% Lucro</th><th class="r">% MC</th></tr></thead>
      <tbody>${vendas.slice(0,400).map(v=>{const f=fin(v);return `<tr>
        <td class="mono">${v.pedido}</td><td>${v.data}</td><td>${v.cliente||v.familia||'—'}</td><td>${v.vendedor||'—'}</td>
        <td class="mono">${v.produto}</td><td class="r">${fmtNum(v.qtd)}</td>
        <td class="r">${fmtBRL(f.receita)}</td><td class="r">${f.custo==null?'<span class="muted">—</span>':fmtBRL(f.custo)}</td>
        <td class="r">${fmtBRL(f.lucro)}</td><td class="r ${f.lucroPct<0.18?'neg':'pos'}">${fmtPct(f.lucroPct)}</td>
        <td class="r">${fmtPct(f.mcPct)}</td></tr>`;}).join('')}</tbody></table></div>
    ${vendas.length>400?`<p class="sub" style="margin-top:10px">Exibindo 400 de ${vendas.length} pedidos. O dashboard considera a base completa.</p>`:''}`;
  if(CONFIG.editavel){
    document.getElementById('nova-venda').onclick=()=>formVenda();
    document.getElementById('ir-importar').onclick=()=>go('importar');
    document.getElementById('pub-vendas').onclick=publicarVendas;
  }
  configurarNav();
}
function formVenda(){
  const root=document.getElementById('content');
  root.innerHTML=`
    <div class="page-head"><div><button class="btn-ghost" id="voltar">← Vendas</button><h1>Nova venda</h1></div></div>
    <form id="form-venda" class="form"><fieldset><legend>Pedido</legend><div class="fgrid">
      <label class="fld"><span>Pedido</span><input name="pedido" type="number"></label>
      <label class="fld"><span>Data</span><input name="data" type="date" value="${new Date().toISOString().slice(0,10)}"></label>
      <label class="fld"><span>Cliente</span><input name="cliente"></label><label class="fld"><span>Família</span><input name="familia"></label>
      <label class="fld"><span>Segmento</span><input name="segmento"></label><label class="fld"><span>Vendedor</span><input name="vendedor"></label>
      <label class="fld wide"><span>SKU</span><select name="produto">${produtos.map(p=>`<option value="${p.codigo}">${p.codigo} — ${p.descricao}</option>`).join('')}</select></label>
      <label class="fld"><span>Quantidade</span><input name="qtd" type="number"></label>
      <label class="fld"><span>Preço unitário</span><input name="preco" type="number" step="0.0001"></label></div></fieldset>
      <div class="form-actions"><button type="button" class="btn-ghost" id="cancelar">Cancelar</button><button type="button" class="btn" id="salvar-venda">Registrar venda</button></div>
      <div id="previa"></div></form>`;
  document.getElementById('voltar').onclick=()=>go('vendas');
  document.getElementById('cancelar').onclick=()=>go('vendas');
  const prev=()=>{const fd=new FormData(document.getElementById('form-venda'));
    if(!fd.get('qtd')||!fd.get('preco')){document.getElementById('previa').innerHTML='';return;}
    const e=calcularCustoTotalProduto(findSku(fd.get('produto')),{qtd:fd.get('qtd'),precoUnit:fd.get('preco')});
    document.getElementById('previa').innerHTML=`<div class="kpis" style="margin-top:18px">${kpi('Receita',fmtBRL(e.precoTotal),'accent')}${kpi('Custo',fmtBRL(e.custoTotal))}${kpi('Lucro',fmtBRL(e.lucroRS))}${kpi('% Lucro',fmtPct(e.lucroPct))}${kpi('% MC',fmtPct(e.mcPct))}</div>`;};
  document.getElementById('form-venda').addEventListener('input',prev);
  document.getElementById('salvar-venda').onclick=()=>{
    const fd=new FormData(document.getElementById('form-venda')),g=k=>fd.get(k);
    const s=findSku(g('produto')), e=calcularCustoTotalProduto(s,{qtd:g('qtd'),precoUnit:g('preco')});
    vendas.unshift({pedido:+g('pedido'),data:g('data'),cliente:g('cliente'),familia:g('familia')||g('cliente'),
      segmento:g('segmento')||'OUTROS',vendedor:g('vendedor'),produto:g('produto'),descricao:s.descricao,grupo:s.grupo,
      qtd:num(g('qtd')),preco:num(g('preco')),receita:+e.precoTotal.toFixed(2),custo:+e.custoTotal.toFixed(2),
      lucro:+e.lucroRS.toFixed(2),lucroPct:+e.lucroPct.toFixed(4),mc:+e.mcRS.toFixed(2),mcPct:+e.mcPct.toFixed(4)});
    go('vendas');
  };
}

// =====================================================================
//  IMPORTAÇÃO EM MASSA (extrato Excel do ERP)
// =====================================================================
const CAMPOS = [
  {k:'pedido',l:'Pedido',req:true},{k:'data',l:'Data',req:true},{k:'cliente',l:'Cliente',req:true},
  {k:'familia',l:'Família'},{k:'segmento',l:'Segmento'},{k:'vendedor',l:'Vendedor',req:true},
  {k:'produto',l:'Produto (código)',req:true},{k:'descricao',l:'Descrição'},{k:'grupo',l:'Grupo'},
  {k:'qtd',l:'Quantidade',req:true},{k:'preco',l:'Preço unit.',req:true},
  {k:'receita',l:'Receita (R$)',fin:true},{k:'custo',l:'Custo (R$)',fin:true},
  {k:'lucro',l:'Lucro (R$)',fin:true},{k:'mc',l:'MC (R$)',fin:true}
];
let importBuffer=null;
function renderImportar(root){
  root.innerHTML=`
    <div class="page-head"><div><button class="btn-ghost" id="voltar">← Vendas</button>
      <h1>Importar vendas do ERP</h1><p class="sub">Extrato diário em Excel ou CSV. Mapeie as colunas uma vez.</p></div></div>
    <div class="card">
      <input type="file" id="arq" accept=".xlsx,.xls,.csv" class="file-input">
      <p class="sub" style="margin-top:10px">A 1ª linha do arquivo deve conter os cabeçalhos das colunas.
      Se o extrato já trouxer receita/custo/lucro, o dashboard usa esses valores. Se não, o motor de custo calcula a partir do SKU.</p>
    </div>
    <div id="map-area"></div>`;
  document.getElementById('voltar').onclick=()=>go('vendas');
  document.getElementById('arq').onchange=lerArquivo;
}
function lerArquivo(ev){
  const file=ev.target.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{
    let rows;
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      rows=XLSX.utils.sheet_to_json(ws,{raw:false,defval:''});
    }catch(err){ document.getElementById('map-area').innerHTML=`<div class="erro"><p>Falha ao ler o arquivo: ${err.message}</p></div>`; return; }
    if(!rows.length){ document.getElementById('map-area').innerHTML='<div class="card"><p class="sub">Nenhuma linha encontrada.</p></div>'; return; }
    importBuffer=rows;
    renderMapeamento(Object.keys(rows[0]), rows.length);
  };
  reader.readAsArrayBuffer(file);
}
function autoMap(headers,k){
  // dobra acentos (Emissão->emissao) e trata "$" como "s" (R$ Un.->rsun)
  const norm=s=>String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\$/g,'s').replace(/[^a-z0-9]/g,'');
  const alvo={pedido:['pedido','nrpedido','numero','numpedido'],data:['emissao','data','dtemissao','dataemissao'],
    cliente:['cliente','razaosocial','nomecliente'],familia:['familiadecliente','familia'],
    segmento:['segmentodeatividade','segmento'],vendedor:['vendedor1','vendedor','representante'],
    produto:['produto','codproduto','sku','codigo','codprod'],descricao:['descricao','descricaodoproduto','desc'],
    grupo:['grupodoproduto','grupo'],qtd:['qtdepedido','quantidade','qtd','qtde'],
    preco:['rsun','rsunitario','precounit','precounitario','valorunit','valorunitario','vlrunit'],
    receita:['rspedido','rspedidototal','receita','valortotal','valorpedido'],custo:['custo','custototal'],
    lucro:['lucro','rslucro'],mc:['mc','margemdecontribuicao','rsmc']}[k]||[];
  return headers.find(h=>alvo.includes(norm(h)))||'';
}
function renderMapeamento(headers,n){
  const opt=sel=>['<option value="">—</option>',...headers.map(h=>`<option ${h===sel?'selected':''}>${h}</option>`)].join('');
  const linhas=CAMPOS.map(c=>`<label class="fld"><span>${c.l}${c.req?' *':''}${c.fin?' <small>(opcional)</small>':''}</span>
    <select data-campo="${c.k}">${opt(autoMap(headers,c.k))}</select></label>`).join('');
  document.getElementById('map-area').innerHTML=`
    <div class="card"><h3>Mapeamento de colunas · ${n} linhas no arquivo</h3>
      <div class="fgrid">${linhas}</div>
      <div class="map-mode"><span>Como aplicar:</span>
        <label><input type="radio" name="modo" value="substituir" checked> Substituir a base inteira</label>
        <label><input type="radio" name="modo" value="adicionar"> Adicionar / atualizar (dedup por pedido+produto)</label></div>
      <div class="form-actions"><button class="btn" id="processar">Pré-visualizar</button></div>
    </div><div id="previa-imp"></div>`;
  document.getElementById('processar').onclick=processarImport;
}
function pnum(v){ if(typeof v==='number')return v; if(v==null)return 0;
  let s=String(v).trim().replace(/[R$\s]/g,'');
  if(s.includes(',')&&s.includes('.')) s=s.replace(/\./g,'').replace(',','.');
  else if(s.includes(',')) s=s.replace(',','.');
  const n=parseFloat(s); return isNaN(n)?0:n; }
function pdata(v){ if(v instanceof Date) return v.toISOString().slice(0,10);
  let s=String(v).trim();
  let m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})/); if(m) return `${m[3]}-${m[2]}-${m[1]}`;
  m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return `${m[1]}-${m[2]}-${m[3]}`;
  const n=parseFloat(s); if(!isNaN(n)&&n>30000){ const d=new Date(Math.round((n-25569)*86400*1000)); return d.toISOString().slice(0,10); }
  return s.slice(0,10); }
function processarImport(){
  const map={}; document.querySelectorAll('[data-campo]').forEach(s=>{ if(s.value) map[s.dataset.campo]=s.value; });
  const falta=CAMPOS.filter(c=>c.req&&!map[c.k]).map(c=>c.l);
  if(falta.length){ alert('Mapeie os campos obrigatórios: '+falta.join(', ')); return; }
  const get=(row,k)=> map[k]!=null ? row[map[k]] : undefined;
  const novas=importBuffer.map(row=>{
    const produto=String(get(row,'produto')||'').trim();
    const s=findSku(produto);
    const v={pedido:get(row,'pedido'),data:pdata(get(row,'data')),cliente:String(get(row,'cliente')||'').trim(),
      familia:map.familia?String(get(row,'familia')).trim():(String(get(row,'cliente')||'').trim()),
      segmento:map.segmento?String(get(row,'segmento')).trim():'OUTROS',vendedor:String(get(row,'vendedor')||'').trim(),
      produto,descricao:map.descricao?String(get(row,'descricao')).trim():(s?s.descricao:''),
      grupo:map.grupo?String(get(row,'grupo')).trim():(s?s.grupo:''),qtd:pnum(get(row,'qtd')),preco:pnum(get(row,'preco'))};
    // financeiro: do extrato se mapeado; senão pelo motor; senão receita simples
    if(map.receita){ v.receita=pnum(get(row,'receita')); v.custo=map.custo?pnum(get(row,'custo')):null;
      v.lucro=map.lucro?pnum(get(row,'lucro')):(v.custo!=null?v.receita-v.custo:null);
      v.mc=map.mc?pnum(get(row,'mc')):null;
      v.lucroPct=v.receita&&v.lucro!=null?v.lucro/v.receita:null; v.mcPct=v.receita&&v.mc!=null?v.mc/v.receita:null;
    } else if(s){ const e=calcularCustoTotalProduto(s,{qtd:v.qtd,precoUnit:v.preco});
      v.receita=+e.precoTotal.toFixed(2);v.custo=+e.custoTotal.toFixed(2);v.lucro=+e.lucroRS.toFixed(2);
      v.lucroPct=+e.lucroPct.toFixed(4);v.mc=+e.mcRS.toFixed(2);v.mcPct=+e.mcPct.toFixed(4);
    } else { v.receita=+(v.qtd*v.preco).toFixed(2);v.custo=null;v.lucro=0;v.lucroPct=0;v.mc=0;v.mcPct=0; }
    return v;
  });
  const semSku=novas.filter(v=>!findSku(v.produto)&&v.custo==null).length;
  const modo=document.querySelector('[name="modo"]:checked').value;
  let receita=0,lucro=0; novas.forEach(v=>{receita+=v.receita||0;lucro+=v.lucro||0;});
  const meses=[...new Set(novas.map(v=>v.data.slice(0,7)))].sort();
  document.getElementById('previa-imp').innerHTML=`
    <div class="kpis" style="margin-top:18px">${kpi('Linhas',fmtNum(novas.length),'accent')}${kpi('Receita',fmtBRL(receita))}
      ${kpi('% Lucro',fmtPct(receita?lucro/receita:0))}${kpi('Período',meses.length?meses[0]+' a '+meses[meses.length-1]:'—')}
      ${kpi('Sem custo',fmtNum(semSku))}</div>
    ${semSku?`<div class="aviso">${semSku} linha(s) sem SKO correspondente nem custo no extrato. Entram só com receita. Cadastre o SKU ou inclua as colunas de custo no extrato para margem completa.</div>`:''}
    <div class="card no-pad" style="margin-top:14px"><table class="full"><thead><tr><th>Pedido</th><th>Data</th><th>Cliente</th><th>Produto</th><th class="r">Qtd</th><th class="r">Receita</th><th class="r">% Lucro</th></tr></thead>
      <tbody>${novas.slice(0,8).map(v=>`<tr><td class="mono">${v.pedido}</td><td>${v.data}</td><td>${v.cliente}</td><td class="mono">${v.produto}</td><td class="r">${fmtNum(v.qtd)}</td><td class="r">${fmtBRL(v.receita)}</td><td class="r">${fmtPct(v.lucroPct||0)}</td></tr>`).join('')}</tbody></table></div>
    <div class="form-actions" style="margin-top:14px">
      <button class="btn-ghost" id="canc-imp">Cancelar</button>
      <button class="btn" id="aplicar-imp">Aplicar (${modo==='substituir'?'substituir':'adicionar'}) e baixar vendas.json</button></div>`;
  document.getElementById('canc-imp').onclick=()=>go('importar');
  document.getElementById('aplicar-imp').onclick=()=>{
    if(modo==='substituir'){ vendas=novas; }
    else { const key=v=>v.pedido+'|'+v.produto; const idx=new Map(vendas.map(v=>[key(v),v]));
      novas.forEach(v=>idx.set(key(v),v)); vendas=[...idx.values()]; }
    vendas.sort((a,b)=>String(a.data).localeCompare(String(b.data)));
    publicarVendas();
    alert(`Base atualizada: ${vendas.length} pedidos. O arquivo vendas.json foi baixado — suba-o no repositório para publicar aos membros.`);
    go('vendas');
  };
}

// ---- bootstrap -------------------------------------------------------
boot();
