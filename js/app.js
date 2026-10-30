/* ==========================================================
   CRM Fonte Clara — app.js
   As variáveis DADOS, CONTATOS, BASEDATA, CTDATA e GERADO
   vêm do arquivo js/dados-iniciais.js, carregado antes deste.
   ========================================================== */

/* ----------------------------------------------------------
   1. LEITURA DA PLANILHA
   ---------------------------------------------------------- */
const ABAS={'CLIENTES':'Pessoa física','EMPRESAS':'Empresa','ESCOLAS':'Escola','DE OUTRA CIDADE':'Outra cidade'};
const _n=s=>(s==null?'':String(s)).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/\s+/g,' ').trim();
const _t=v=>v==null?'':String(v).replace(/\s+/g,' ').trim();

function mapear(hdr){
  const H=hdr.map(_n), find=(...ps)=>{for(const p of ps){const i=H.findIndex(h=>h&&p.test(h));if(i>=0)return i}return -1};
  return {nome:0,
    end:find(/^ENDERE/),
    loc:find(/^BAIRRO/,/^CIDADE/),
    data:find(/^DATA$/,/^DIA\/MES\/ANO/,/^DATA/),
    prod:find(/^REFIL/,/^PRODUTO/),
    vld:find(/^VLD/,/^365$/),
    venc:find(/^VENCE/,/^VENCIMENTO/),
    tel:find(/^TELEFONE(?! 2)/),
    tel2:find(/^TELEFONE 2/),
    fc:find(/^FALAR COM/),
    tp:find(/^TIPO DE TELEFONE/),
    obs:find(/^OBSERVA/),
    cod:find(/^ID$/)};
}
const _iso=d=>d?d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'):'';
function daDataDoExcel(v){
  if(v instanceof Date) return v;
  if(typeof v==='number'&&v>10000) return new Date(Math.round((v-25569)*86400000));
  if(typeof v==='string'){const m=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);if(m)return new Date(+m[3],+m[2]-1,+m[1]);}
  return null;
}
const CANONICOS=["Refil Padrão", "Refil Padrão + Pré-filtro", "Refil Carbono 9\"", "Refil Plissado", "Refil Compacto"];
function _pk(s){
  let t=(s==null?'':String(s)).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  t=t.replace(/filtro\s+padrao\s*2\s*(\/|em)\s*1/g,'refil padrao');
  t=t.replace(/\boriginal\b/g,'').replace(/[\/]/g,' + ');
  t=t.replace(/\bpre\s*-?\s*filtro\b/g,'prefiltro');
  return t.replace(/[^a-z0-9+]/g,'');
}
const CANON={}; for(const x of CANONICOS) CANON[_pk(x)]=x;

const _bk=s=>(s==null?'':String(s)).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
function unificarBairros(R){
  const cont={};
  for(const x of R){const k=_bk(x.loc); if(!k) continue; const t=_t(x.loc);
    (cont[k]=cont[k]||{})[t]=(cont[k][t]||0)+1;}
  const canon={};
  for(const k in cont){
    let melhor='',bs=null;
    for(const t in cont[k]){
      const naoCaps=(t===t.toUpperCase()&&t!==t.toLowerCase())?0:1;
      const caps=t.split(/\s+/).filter(w=>w&&w[0]===w[0].toUpperCase()&&w[0]!==w[0].toLowerCase()).length;
      const ac=/[^\x00-\x7F]/.test(t)?1:0;
      const sc=[naoCaps,caps,ac,cont[k][t]];
      if(!bs||sc.some((v,i)=>v!==bs[i]&&sc.slice(0,i).every((y,j)=>y===bs[j])&&v>bs[i])){bs=sc;melhor=t}
    }
    canon[k]=melhor;
  }
  for(const x of R){const k=_bk(x.loc); if(k) x.loc=canon[k];}
  return R;
}
function unificarProdutos(R){
  const m={};
  for(const x of R){const k=_pk(x.prod); if(!k) continue; if(!m[k]) m[k]=CANON[k]||_t(x.prod);}
  for(const x of R){const k=_pk(x.prod); if(k) x.prod=m[k];}
  return R;
}
const SEMPADRAO=new Set();
function construir(wb){
  SEMPADRAO.clear();
  const hoje=new Date();hoje.setHours(0,0,0,0);
  const R=[];
  for(const nomeAba of wb.SheetNames){
    const chave=Object.keys(ABAS).find(k=>_n(nomeAba)===k);
    if(!chave) continue;
    const linhas=XLSX.utils.sheet_to_json(wb.Sheets[nomeAba],{header:1,raw:true,cellDates:true,defval:null});
    if(!linhas.length) continue;
    const C=mapear(linhas[0]);
    if(C.end<0||C.data<0||C.tel<0) throw new Error('A aba '+nomeAba+' não tem as colunas esperadas.');
    if(C.tp<0) SEMPADRAO.add(nomeAba.trim());
    for(let i=1;i<linhas.length;i++){
      const L=linhas[i]||[], nome=_t(L[0]);
      if(!nome) continue;
      const g=k=>C[k]>=0?_t(L[C[k]]):'';
      let d=daDataDoExcel(C.data>=0?L[C.data]:null);
      if(d&&d.getFullYear()<2010) d=null;
      let vld=parseInt(String(g('vld')).replace(/\D/g,''),10); if(!vld||isNaN(vld)) vld=365;
      let venc=null,dias=99999,st='sem data';
      if(d){ venc=new Date(d.getTime()+vld*86400000); venc.setHours(0,0,0,0);
        dias=Math.round((hoje-venc)/86400000);
        if(dias>0) st='vencido';
        else if(venc.getMonth()===hoje.getMonth()&&venc.getFullYear()===hoje.getFullYear()) st='vence este mes';
        else if(dias>=-60) st='vence em breve';
        else st='em dia'; }
      R.push({id:chave.slice(0,3)+(i+1),aba:nomeAba.trim(),tipo:ABAS[chave],nome,end:g('end'),loc:g('loc'),
        prod:g('prod'),vld,ult:_iso(d),venc:_iso(venc),dias,tel:g('tel'),tel2:g('tel2'),fc:g('fc'),tp:g('tp'),
        obs:g('obs'),st,lin:i+1,cod:C.cod>=0?g('cod'):''});
    }
  }
  if(!R.length) throw new Error('Não encontrei nenhum cadastro. Essa é a Planilha de Clientes?');
  unificarProdutos(R);unificarBairros(R);
  const gr={},bp={};
  for(const x of R){const k=x.aba+'|'+_n(x.nome)+'|'+_n(x.end);(gr[k]=gr[k]||[]).push(x);
    const num=(x.tel||'').replace(/\D/g,''); if(num.length>=10){const k2=num.slice(-8);(bp[k2]=bp[k2]||[]).push(x)}}
  for(const k in bp){const v=bp[k],ns=new Set(v.map(x=>_n(x.nome)));
    if(ns.size>1) v.forEach(x=>{const o=[...new Set(v.filter(y=>_n(y.nome)!==_n(x.nome)).map(y=>y.nome))].sort();if(o.length)x.tambem=o})}
  return agrupar(R);
}

/* ----------------------------------------------------------
   1b. UM CARD POR CLIENTE (20/10/2026)
   Antes cada linha da planilha virava um card — a Real Calçados, com três
   purificadores, aparecia três vezes. Agora as linhas do mesmo cliente (mesmo
   código na coluna ID da planilha; sem código, mesmo nome e endereço na mesma
   aba) viram UM card, com a lista de aparelhos dentro (c.aps).
   - O card segue o aparelho mais urgente (o que vence primeiro).
   - O id do card é o da primeira linha do cliente; os ids das outras linhas
     ficam em ALIAS, para levar junto anotações, agendamentos e trocas antigas.
   - Cada aparelho guarda: id da linha, linha, nome, refil, última troca,
     validade e a observação da planilha (ex.: "Térreo", "1º andar").
   ---------------------------------------------------------- */
let ALIAS={};
const _APC=['id','lin','nome','prod','ult','vld','obs','cod'];
function expandir(R){
  const o=[];
  for(const c of R){
    if(c&&Array.isArray(c.aps)&&c.aps.length){
      for(const a of c.aps){const x=Object.assign({},c);
        delete x.aps;delete x.ids;delete x.nomes;delete x.aparelhos;delete x._busca;delete x._buscaAnot;
        for(const k of _APC) x[k]=a[k]!==undefined?a[k]:(k==='obs'||k==='cod'?'':x[k]);
        if(!x.nome) x.nome=c.nome;
        o.push(x)}
    } else if(c) o.push(c);
  }
  return o;
}
function agrupar(R){
  const rows=expandir(R); unificarProdutos(rows);
  const g={},ord=[];
  for(const x of rows){
    const k=x.aba+'|'+(x.cod?'#'+_n(x.cod):_n(x.nome)+'|'+_n(x.end));
    if(!g[k]){g[k]=[];ord.push(k)} g[k].push(x);
  }
  ALIAS={}; const out=[];
  for(const k of ord){
    const L=g[k];
    if(L.length===1){const c=L[0]; delete c.aparelhos; out.push(c); continue}
    L.sort((a,b)=>(a.lin||0)-(b.lin||0));
    const G=Object.assign({},L[0]);
    delete G._busca; delete G._buscaAnot;
    G.aps=L.map(x=>{const a={};for(const k2 of _APC) a[k2]=x[k2]!==undefined?x[k2]:'';return a});
    G.ids=L.map(x=>x.id);
    G.nomes=[...new Set(L.map(x=>x.nome))];
    G.aparelhos=L.length; G.obs='';
    for(const f of ['tel','tel2','fc','tp','end','loc']) if(!G[f]){const y=L.find(x=>x[f]); if(y) G[f]=y[f]}
    if(G.tambem){G.tambem=G.tambem.filter(n=>!G.nomes.some(m=>_n(m)===_n(n))); if(!G.tambem.length) delete G.tambem}
    for(const x of L) if(x.id!==G.id) ALIAS[x.id]=G.id;
    out.push(G);
  }
  return out;
}
/* o nome do registro bate com o cliente (ou com algum nome das linhas dele)? */
const nomeBate=(c,n)=>!!c&&(_n(c.nome)===_n(n)||(c.nomes||[]).some(m=>_n(m)===_n(n)));
/* índice por nome que conhece todos os nomes das linhas de cada cliente */
function indicePorNome(){const P={};
  for(const c of DADOS) for(const n of new Set([c.nome].concat(c.nomes||[]))) (P[_n(n)]=P[_n(n)]||[]).push(c);
  return P}
/* acha o cliente pelo id — inclusive pelo id antigo de uma linha que foi juntada */
const achaC=id=>DADOS.find(x=>x.id===id)||(ALIAS[id]?DADOS.find(x=>x.id===ALIAS[id]):null);
/* texto das linhas da planilha: "linha 319" ou "linhas 319, 320, 321" */
const linTxt=c=>c&&c.aps&&c.aps.length>1?'linhas '+c.aps.map(a=>a.lin).join(', '):'linha '+(c?c.lin:'');
/* junta dois registros de contato do mesmo cliente sem perder nada */
function juntarLog(a,b,nome){
  if(!a) {const r=Object.assign({},b); r.n=nome||r.n; return r}
  if(!b) return a;
  const r=Object.assign({},a); r.n=nome||a.n||b.n;
  r.h=(a.h||[]).slice();
  for(const x of (b.h||[])) if(!r.h.some(y=>y.d===x.d&&y.t===x.t&&(y.o||'')===(x.o||''))) r.h.push(x);
  if(b.a&&((b.d||'')>(r.d||'')||!r.a)){r.a=b.a;r.d=b.d}
  if(b.uc&&(!r.uc||b.uc>r.uc)) r.uc=b.uc;
  if(b.v&&(!r.v||b.v>r.v)) r.v=b.v;
  if(b.obs&&b.obs!==r.obs) r.obs=r.obs?(r.obs.split(' | ').includes(b.obs)?r.obs:r.obs+' | '+b.obs):b.obs;
  return r;
}
/* leva anotações, agendamentos e trocas das linhas juntadas para o card do cliente */
function migrarIds(){
  let mL=0,mA=0,mT=0;
  for(const old in ALIAS){
    const nid=ALIAS[old], G=DADOS.find(x=>x.id===nid); if(!G) continue;
    const r=LOG[old];
    if(r&&(!r.n||nomeBate(G,r.n))){LOG[nid]=juntarLog(LOG[nid],r,G.nome);delete LOG[old];mL++}
    for(const a of AGEND) if(a.cid===old&&(!a.nome||nomeBate(G,a.nome))){a.cid=nid;mA++}
    for(const t of TROCAS) if(t.cid===old&&(!t.nome||nomeBate(G,t.nome))){t.cid=nid;mT++}
  }
  if(mL) salvar(); if(mA) salvarAg(); if(mT) salvarTr();
  return mL+mA+mT;
}

function _fones(v){const o=new Set();const s=String(v==null?'':v);
  for(let n of (s.match(/\d{8,15}/g)||[])){ if(n.length>=12&&n.slice(0,2)==='55') n=n.slice(2); if(n.length>=8) o.add(n.slice(-8)); }
  return [...o];}
function construirContatos(wb){
  const M={'JAN':1,'FEV':2,'MAR':3,'ABRIL':4,'MAIO':5,'JUNHO':6,'JUL':7,'AGO':8,'SET':9,'OUT':10,'NOV':11,'DEZ':12};
  const C=[],S=[];
  for(const nm of wb.SheetNames){
    const N=_n(nm);
    if(N==='TOTAIS'){
      const L=XLSX.utils.sheet_to_json(wb.Sheets[nm],{header:1,raw:true,cellDates:true,defval:null});
      for(let i=1;i<L.length;i++){const r=L[i]||[],nome=_t(r[1]);
        if(!nome||_n(nome)==='CLIENTE') continue;
        const d=daDataDoExcel(r[0]); if(!d) continue;
        S.push({nome,nz:_n(nome),d:_iso(d),v:r[2],o:_t(r[3]),p:_t(r[4])});}
      continue;
    }
    const mes=M[N.split(' ')[0]]; if(!mes) continue;
    const L=XLSX.utils.sheet_to_json(wb.Sheets[nm],{header:1,raw:true,cellDates:true,defval:null});
    for(let i=3;i<L.length;i++){const r=L[i]||[],nome=_t(r[0]);
      if(!nome||_n(nome)==='CLIENTE') continue;
      const d=daDataDoExcel(r[3]);
      C.push({nome,nz:_n(nome),f:_fones(r[1]).concat(_fones(r[2])),d:d?_iso(d):'',
        s:_t(r[4]),o:_t(r[5]),m:mes,aba:nm.trim()});}
  }
  if(!C.length) throw new Error('Não achei abas de contato mensais. Essa é a planilha antiga de contatos (2026.xlsx)?');
  return {c:C,s:S};
}
function casar(R,CT){
  if(!CT||!CT.c) return R;
  const pn={},pf={},pc={},cur=x=>{const p=_n(x).split(' ');return p.length>1?p[0]+' '+p[1]:_n(x)};
  for(const c of R){ c.ct=null; c.sv=null;
    (pn[_n(c.nome)]=pn[_n(c.nome)]||[]).push(c);
    (pc[cur(c.nome)]=pc[cur(c.nome)]||[]).push(c);
    for(const f of _fones(c.tel).concat(_fones(c.tel2))) (pf[f]=pf[f]||[]).push(c); }
  const achar=(nome,fs)=>{
    if(pn[_n(nome)]) return pn[_n(nome)][0];
    for(const f of (fs||[])) if(pf[f]) return pf[f][0];
    const k=cur(nome); if(pc[k]&&pc[k].length===1) return pc[k][0];
    return null;};
  let n=0;
  for(const x of CT.c){const a=achar(x.nome,x.f); if(!a) continue; n++;
    if(!a.ct||(x.d||'')>(a.ct.d||'')) a.ct={d:x.d,s:x.s,o:x.o,m:x.m,aba:x.aba};}
  for(const x of (CT.s||[])){const a=achar(x.nome,[]); if(!a) continue;
    if(!a.sv||x.d>a.sv.d) a.sv={d:x.d,o:x.o,p:x.p};}
  CT.casados=n;
  return R;
}

/* ----------------------------------------------------------
   2. ESTADO
   ---------------------------------------------------------- */
const $=s=>document.querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
const fmt=s=>s?s.split('-').reverse().join('/'):'—';
const fmtc=s=>s?s.slice(8,10)+'/'+s.slice(5,7):'—';
const dig=s=>(s||'').replace(/\D/g,'');
const norm=s=>(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();

const ST={vencido:['s-vencido','Vencido'],'vence este mes':['s-vence','Vence este mês'],
  'vence em breve':['s-breve','Vence em breve'],'em dia':['s-dia','Em dia'],'sem data':['s-sem','Sem data']};
/* cores sempre por variável do CSS — assim o modo claro e o escuro funcionam sozinhos.
   Nunca escrever cor em hexadecimal aqui no JS. */
const COR={vencido:'var(--bad)','vence este mes':'var(--warn)','vence em breve':'var(--info)','em dia':'var(--ok)','sem data':'var(--sem)'};
/* Situação = resultado do ÚLTIMO contato. Desde 29/09/2026 só três são marcadas
   à mão (Mandei mensagem · Não atendeu/deixei recado · Falei com a pessoa) além
   de "Não quer mais". "Agendado" e "Trocou" o CRM marca sozinho (agendamento e
   "Troquei o refil"). "Vai pensar" ficou só para ler registros antigos. */
const SIT={
  aguardando:{rot:'Mandei mensagem',cor:'var(--warn)',bg:'var(--warn-b)',peso:2},
  naoatendeu:{rot:'Não atendeu · deixei recado',cor:'var(--neutro)',bg:'var(--mini-b)',peso:1},
  falei:{rot:'Falei com a pessoa',cor:'var(--info)',bg:'var(--info-b)',peso:3},
  pensar:{rot:'Vai pensar',cor:'var(--info)',bg:'var(--info-b)',peso:3},
  agendado:{rot:'Agendado',cor:'var(--ok)',bg:'var(--ok-b)',peso:4},
  trocou:{rot:'Já trocou',cor:'var(--teal)',bg:'var(--teal-b)',peso:5},
  naoquer:{rot:'Não quer mais',cor:'var(--bad)',bg:'var(--bad-b)',peso:6}};

/* ---- tipos de contato do histórico (LOG[id].h) ---- */
const TCONT={
  falei:  {rot:'Falei com a pessoa',        curto:'Falei',        ic:'📞', sit:'falei'},
  recado: {rot:'Não atendeu, deixei recado',curto:'Não atendeu',  ic:'📵', sit:'naoatendeu'},
  msg:    {rot:'Mandei mensagem',           curto:'Mensagem',     ic:'💬', sit:'aguardando'},
  naoquer:{rot:'Não quer mais',             curto:'Não quer mais',ic:'✋', sit:'naoquer'}};
const SIT2T={aguardando:'msg',naoatendeu:'recado',falei:'falei',pensar:'falei',naoquer:'naoquer'};

/* frases prontas do "o que a pessoa disse" — cada uma sugere uma data de retorno */
const FRASES=[
  ['Vai esperar o cartão virar','cartao'],
  ['Avisa no mês que vem','mes'],
  ['Vai pensar','7'],
  ['Pediu para ligar outro dia','2'],
  ['Vai trocar, marcar dia',''],
  ['Está viajando','15'],
  ['Já trocou com outra empresa',''],
  ['Não usa mais o purificador','']];

/* ---- serviços que a loja agenda ---- */
const SERVICOS=['Troca de refil','Instalação','Reinstalação','Manutenção','Higienização','Troca de peça','Orçamento','Entrega','Outro'];

let LOG={};try{LOG=JSON.parse(localStorage.getItem('crm_demo_log')||'{}')}catch(e){LOG={}}

/* ---- agendamentos de serviço (crm_demo_ag) ---- */
let AGEND=[];try{AGEND=JSON.parse(localStorage.getItem('crm_demo_ag')||'[]');if(!Array.isArray(AGEND))AGEND=[]}catch(e){AGEND=[]}
const salvarAg=()=>{try{localStorage.setItem('crm_demo_ag',JSON.stringify(AGEND))}catch(e){}};

/* ---- trocas de refil já feitas, à espera de entrar na planilha (crm_demo_tr) ---- */
let TROCAS=[];try{TROCAS=JSON.parse(localStorage.getItem('crm_demo_tr')||'[]');if(!Array.isArray(TROCAS))TROCAS=[]}catch(e){TROCAS=[]}
const salvarTr=()=>{try{localStorage.setItem('crm_demo_tr',JSON.stringify(TROCAS))}catch(e){}};

const novoId=p=>(p||'a')+Date.now().toString(36)+Math.random().toString(36).slice(2,6);

/* última mensagem montada para cada cliente, para o botão "mandar para qual número" */
const MSGPRONTA={};
const _MIG={'Vai trocar':'agendado','Remarcar':'pensar','Não atendeu':'naoatendeu','Não quer mais':'naoquer'};
for(const k in LOG){const r=LOG[k]; if(r&&_MIG[r.a]) r.a=_MIG[r.a];}

/* Migração de 29/09/2026: quem já tinha situação marcada ganha o primeiro item
   do histórico de contatos, com a data em que a situação foi marcada. A anotação
   antiga (r.obs) continua como "anotação fixa" do cliente, sem mudar. */
(function(){
  let m=0;
  for(const k in LOG){const r=LOG[k]; if(!r||Array.isArray(r.h)) continue;
    r.h=[]; const t=SIT2T[r.a];
    if(t&&(r.d||r.uc)){r.h.push({d:r.d||r.uc,t,o:r.a==='pensar'?'Vai pensar':'',mig:1}); m++}
  }
  if(m) try{localStorage.setItem('crm_demo_log',JSON.stringify(LOG))}catch(e){}
})();

const sit=c=>{const r=LOG[c.id];return r&&SIT[r.a]?r.a:''};
/* último item do histórico (o contato mais recente) */
const ultReg=c=>{const r=LOG[c.id];if(!r||!r.h||!r.h.length)return null;
  let u=r.h[0];for(const x of r.h) if((x.d||'')>=(u.d||'')) u=x; return u};
/* data de HOJE no horário do computador. Antes usava toISOString(), que é o
   horário de Londres — depois das 21h o CRM já achava que era o dia seguinte. */
const HOJE=()=>_iso(new Date());
const anot=c=>{const r=LOG[c.id];return r&&r.obs?r.obs:''};
const volta=c=>{const r=LOG[c.id];return r&&r.v?r.v:''};
const vencido_retorno=c=>{const v=volta(c);return !!v&&v<=HOJE()};
const diasDe=d=>{if(!d)return null;const a=new Date(d+'T00:00:00'),h=new Date(HOJE()+'T00:00:00');return Math.round((h-a)/86400000)};
/* "Hora de ligar": mandou mensagem (ou ligou e não atendeu) há 3 dias ou mais,
   sem data de retorno marcada para o futuro — é o fluxo da loja: mensagem
   primeiro, ligação uns dias depois. */
const DIAS_LIGAR=3;
const paraLigar=c=>{const k=sit(c);if(k!=='aguardando'&&k!=='naoatendeu')return false;
  const v=volta(c); if(v&&v>HOJE()) return false;
  const r=LOG[c.id],n=diasDe(r&&r.d);return n!==null&&n>=DIAS_LIGAR};
/* ordem da fila: retornos de hoje/atrasados → hora de ligar → nunca contatados
   → os demais pela situação */
const peso=c=>{if(vencido_retorno(c))return -2;if(paraLigar(c))return -1;const k=sit(c);return k?SIT[k].peso:0};
const salvar=()=>{try{localStorage.setItem('crm_demo_log',JSON.stringify(LOG))}catch(e){}};

/* ---- registrar um contato no histórico ----
   o = {t:'falei'|'recado'|'msg'|'naoquer', o:'o que a pessoa disse', v:'AAAA-MM-DD' (voltar em), d:data}
   A data de retorno SEMPRE é regravada: se o contato novo não tem data, a antiga
   (que já passou ou foi resolvida) sai. */
function registrarContato(id,o){
  const c=DADOS.find(x=>x.id===id); if(!c||!TCONT[o.t]) return null;
  const r=LOG[id]||{}; r.n=r.n||c.nome; if(!Array.isArray(r.h)) r.h=[];
  const d=o.d||HOJE();
  const item={d,t:o.t}; if(o.o) item.o=o.o; if(o.v) item.v=o.v; if(o.hr) item.hr=o.hr; if(o.cel) item.cel=1;
  r.h.push(item);
  // a situação e as datas acompanham o contato mais recente
  const u=ultRegR(r);
  if(u===item){
    r.a=TCONT[o.t].sit; r.d=d;
    if(o.v) r.v=o.v; else delete r.v;
  }
  if(!r.uc||d>r.uc) r.uc=d;
  LOG[id]=r; salvar();
  return item;
}
function ultRegR(r){if(!r||!r.h||!r.h.length)return null;let u=r.h[0];for(const x of r.h)if((x.d||'')>=(u.d||''))u=x;return u}
/* apaga um item do histórico e refaz a situação a partir do que sobrou */
function apagarContato(id,idx){
  const r=LOG[id]; if(!r||!r.h||!r.h[idx]) return;
  r.h.splice(idx,1);
  const u=ultRegR(r);
  if(u){r.a=TCONT[u.t].sit;r.d=u.d;if(u.v)r.v=u.v;else delete r.v;r.uc=u.d}
  else{delete r.a;delete r.d;delete r.v;delete r.uc}
  if(!r.h.length&&!r.obs&&!r.a) delete LOG[id];
  salvar();
}
/* texto curto do último contato, para a linha do cliente */
function rotUltReg(u){
  if(!u) return '';
  const n=diasDe(u.d);
  const q=n===0?'hoje':n===1?'ontem':fmtc(u.d);
  return `${TCONT[u.t]?TCONT[u.t].ic:'•'} ${q} · ${TCONT[u.t]?TCONT[u.t].curto:''}${u.o?' — “'+u.o+'”':''}`;
}

function reancorar(){
  migrarIds();
  const porId={},porNome=indicePorNome();
  for(const c of DADOS) porId[c.id]=c;
  let m=0;
  for(const k of Object.keys(LOG)){
    const r=LOG[k]; if(!r||!r.n) continue;
    const a=porId[k]; if(a&&nomeBate(a,r.n)) continue;
    const alvo=(porNome[_n(r.n)]||[])[0];
    if(alvo&&alvo.id!==k){LOG[alvo.id]=LOG[alvo.id]?juntarLog(LOG[alvo.id],r,alvo.nome):r;delete LOG[k];m++}
  }
  if(m) salvar();
  return m;
}
function recalcular(R){
  const h=new Date();h.setHours(0,0,0,0);
  const calc=c=>{
    if(!c.ult){c.dias=99999;c.st='sem data';c.venc='';return}
    const p=c.ult.split('-').map(Number), v=new Date(p[0],p[1]-1,p[2]+(c.vld||365));
    c.venc=_iso(v);
    c.dias=Math.round((h-v)/86400000);
    c.st=c.dias>0?'vencido':(v.getMonth()===h.getMonth()&&v.getFullYear()===h.getFullYear()?'vence este mes':(c.dias>=-60?'vence em breve':'em dia'));
  };
  for(const c of R){
    if(!c.aps||!c.aps.length){calc(c);continue}
    c.aps.forEach(calc);
    /* o card segue o aparelho mais urgente (o que vence primeiro) */
    const com=c.aps.filter(a=>a.ult);
    const u=com.length?com.reduce((m,a)=>a.dias>m.dias?a:m,com[0]):null;
    if(u){c.ult=u.ult;c.vld=u.vld;c.venc=u.venc;c.dias=u.dias;c.st=u.st}
    else{c.ult='';c.venc='';c.dias=99999;c.st='sem data'}
    c.prod=[...new Set(c.aps.map(a=>a.prod).filter(Boolean))].join(' / ');
    c.aparelhos=c.aps.length;
  }
  return R;
}

try{const sv=JSON.parse(localStorage.getItem('crm_demo_base')||'null');
  if(sv&&Array.isArray(sv.d)&&sv.d.length){DADOS=sv.d;BASEDATA=sv.q}}catch(e){}
try{const sc=JSON.parse(localStorage.getItem('crm_demo_ct')||'null');
  if(sc&&sc.d&&sc.d.c&&sc.d.c.length){CONTATOS=sc.d;CTDATA=sc.q}}catch(e){}
/* Se o index.html pulou o dados-iniciais.js (porque já havia base guardada) e a
   base guardada não abriu, apaga a chave estragada e recarrega — aí o retrato
   de 25/08 entra no lugar, e nada fica em branco. */
if(window.__semIniciais&&(!Array.isArray(DADOS)||!DADOS.length)){
  try{localStorage.removeItem('crm_demo_base');localStorage.removeItem('crm_demo_ct')}catch(e){}
  location.reload();
}
/* guarda os contatos antigos (planilha antiga de contatos) no navegador, para que nas
   próximas aberturas o dados-iniciais.js (517 KB) não precise mais ser lido */
try{if(!localStorage.getItem('crm_demo_ct')&&CONTATOS&&CONTATOS.c)
  localStorage.setItem('crm_demo_ct',JSON.stringify({q:CTDATA,d:CONTATOS}))}catch(e){}
DADOS=agrupar(DADOS);unificarBairros(DADOS);recalcular(DADOS);casar(DADOS,CONTATOS);reancorar();

let VISTA='fila';
let filtro='fila', tipo='todos';
const ABERTOS=new Set();
const naFila=c=>(c.st==='vencido'&&c.dias<=180)||c.st==='vence este mes'||c.st==='vence em breve';

/* ----------------------------------------------------------
   3. AVISOS FLUTUANTES
   ---------------------------------------------------------- */
function toast(txt,tipo){
  const d=document.createElement('div');
  d.className='toast'+(tipo?' '+tipo:'');
  d.innerHTML='<span class="ic">'+(tipo==='err'?'⚠️':tipo==='ok'?'✅':'💧')+'</span><span>'+esc(txt)+'</span>';
  $('#toasts').appendChild(d);
  setTimeout(()=>{d.classList.add('sai');setTimeout(()=>d.remove(),320)},3200);
}

/* ----------------------------------------------------------
   3b. JANELA FLUTUANTE (modal)
   ---------------------------------------------------------- */
function fecharModal(){$('#fundo').hidden=true;$('#modal').innerHTML='';$('#modal').className='modal'}
function abrirModal(html,depois){
  $('#modal').innerHTML=html;
  $('#fundo').hidden=false;
  if(depois) depois($('#modal'));
  const p=$('#modal').querySelector('[autofocus]'); if(p) setTimeout(()=>p.focus(),60);
}
$('#fundo').addEventListener('click',e=>{
  if(e.target.id==='fundo'||e.target.closest('[data-fechar]')) fecharModal();
});
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('#fundo').hidden)fecharModal()});

/* ----------------------------------------------------------
   3c. AGENDAMENTOS DE SERVIÇO
   ---------------------------------------------------------- */
const agOrd=()=>AGEND.slice().sort((a,b)=>(a.d+(a.h||'99'))<(b.d+(b.h||'99'))?-1:1);
const agAbertos=()=>AGEND.filter(x=>!x.feito);
const agDoDia=d=>agOrd().filter(x=>x.d===d);
const agDoCliente=id=>AGEND.filter(x=>x.cid===id&&!x.feito);

function salvarAgendamento(o){
  if(o.id){const i=AGEND.findIndex(x=>x.id===o.id); if(i>=0){AGEND[i]=Object.assign(AGEND[i],o)} else AGEND.push(o)}
  else{o.id=novoId('ag');o.criado=HOJE();AGEND.push(o)}
  salvarAg();
  // o cliente da base fica marcado como “agendado”
  if(o.cid){const c=DADOS.find(x=>x.id===o.cid);
    if(c){const r=LOG[o.cid]||{};r.a='agendado';r.d=HOJE();r.n=c.nome;LOG[o.cid]=r;salvar()}}
  atualizarTudo();
}
function apagarAgendamento(id){
  const i=AGEND.findIndex(x=>x.id===id); if(i<0) return;
  AGEND.splice(i,1); salvarAg(); atualizarTudo();
}

/* janela de novo/editar agendamento */
function modalAgendar(pref){
  pref=pref||{};
  const ed=!!pref.id;
  const hoje=HOJE();
  const serv=pref.serv||'Troca de refil';
  abrirModal(`
    <h3>${ed?'Editar agendamento':'Agendar serviço'}</h3>
    <div class="msub">${pref.cid?'Cliente da base — os dados já vêm preenchidos.':'Agendamento avulso — pode ser qualquer pessoa, mesmo fora da planilha.'}</div>
    <div class="mcampo">
      <span class="rotulo">Cliente</span>
      <input type="text" id="mnome" autofocus placeholder="Nome de quem vai ser atendido" value="${esc(pref.nome||'')}">
      <div class="msug" id="msug"></div>
      <input type="hidden" id="mcid" value="${esc(pref.cid||'')}">
    </div>
    <div class="mlin">
      <div class="mcampo"><span class="rotulo">Dia</span>
        <input type="date" id="mdata" value="${pref.d||hoje}"></div>
      <div class="mcampo"><span class="rotulo">Hora (opcional)</span>
        <input type="time" id="mhora" value="${pref.h||''}"></div>
    </div>
    <div class="mcampo">
      <span class="rotulo">Serviço</span>
      <div class="mservs" id="mservs">${SERVICOS.map(s=>
        `<button type="button" class="chip${s===serv?' on':''}" data-s="${esc(s)}">${esc(s)}</button>`).join('')}</div>
    </div>
    <div class="mlin">
      <div class="mcampo"><span class="rotulo">Telefone</span>
        <input type="text" id="mtel" placeholder="—" value="${esc(pref.tel||'')}"></div>
      <div class="mcampo"><span class="rotulo">Endereço</span>
        <input type="text" id="mend" placeholder="—" value="${esc(pref.end||'')}"></div>
    </div>
    <div class="mcampo">
      <span class="rotulo">Observação</span>
      <textarea id="mobs" rows="2" placeholder="Ex: levar refil e pré-filtro; cliente só depois das 14h">${esc(pref.obs||'')}</textarea>
    </div>
    <div class="mrod">
      ${ed?'<button class="bt" id="mdel" style="margin-right:auto">Apagar</button>':''}
      <button class="bt" data-fechar>Cancelar</button>
      <button class="bt pri" id="mok">${ed?'Salvar':'Agendar'}</button>
    </div>`, m=>{
    let sv=serv;
    m.querySelector('#mservs').addEventListener('click',e=>{
      const b=e.target.closest('.chip'); if(!b) return;
      sv=b.dataset.s;
      m.querySelectorAll('#mservs .chip').forEach(x=>x.classList.toggle('on',x===b));
    });
    // sugestão de clientes da base enquanto digita
    const nm=m.querySelector('#mnome'), sug=m.querySelector('#msug'), cid=m.querySelector('#mcid');
    nm.addEventListener('input',()=>{
      cid.value='';
      const q=norm(nm.value.trim());
      if(q.length<3){sug.innerHTML='';return}
      const achou=DADOS.filter(c=>norm(c.nome).includes(q)).slice(0,6);
      sug.innerHTML=achou.map(c=>`<button type="button" data-c="${c.id}"><b>${esc(c.nome)}</b>
        <span>${esc([c.loc,c.tel,c.prod].filter(Boolean).join(' · '))}</span></button>`).join('');
    });
    sug.addEventListener('click',e=>{
      const b=e.target.closest('button[data-c]'); if(!b) return;
      const c=DADOS.find(x=>x.id===b.dataset.c); if(!c) return;
      nm.value=c.nome; cid.value=c.id;
      m.querySelector('#mtel').value=c.tel||''; m.querySelector('#mend').value=[c.end,c.loc].filter(Boolean).join(' — ');
      sug.innerHTML='';
    });
    if(m.querySelector('#mdel')) m.querySelector('#mdel').addEventListener('click',()=>{
      if(confirm('Apagar esse agendamento?')){apagarAgendamento(pref.id);fecharModal();toast('Agendamento apagado.')}
    });
    m.querySelector('#mok').addEventListener('click',()=>{
      const nome=nm.value.trim();
      if(!nome){toast('Escreva o nome de quem vai ser atendido.','err');nm.focus();return}
      const d=m.querySelector('#mdata').value;
      if(!d){toast('Escolha o dia.','err');return}
      salvarAgendamento({id:pref.id,cid:cid.value||'',nome,d,h:m.querySelector('#mhora').value,
        serv:sv,tel:m.querySelector('#mtel').value.trim(),end:m.querySelector('#mend').value.trim(),
        obs:m.querySelector('#mobs').value.trim(),feito:pref.feito||false});
      fecharModal();
      toast((ed?'Agendamento salvo':primeiroNome(nome)+' agendado')+' para '+fmt(d)+'.','ok');
    });
  });
}
function agendarCliente(id,dia){
  const c=DADOS.find(x=>x.id===id); if(!c) return;
  modalAgendar({cid:c.id,nome:c.nome,tel:c.tel,end:[c.end,c.loc].filter(Boolean).join(' — '),
    d:dia||HOJE(),serv:'Troca de refil'});
}

/* ----------------------------------------------------------
   3d. TROCAS DE REFIL FEITAS
   ---------------------------------------------------------- */
const trPend=()=>TROCAS.filter(x=>!x.ok);
/* sel (só para cliente com vários aparelhos) = [{i: posição em c.aps, prod}] */
function registrarTroca(id,dia,prodNovo,sel){
  const c=DADOS.find(x=>x.id===id); if(!c) return;
  const d=dia||HOJE();
  const j=TROCAS.find(x=>x.cid===id&&!x.ok);
  const o=j||{id:novoId('tr'),cid:id,ok:false};
  o.nome=c.nome; o.d=d; o.aba=c.aba; o.lin=c.lin; o.tel=c.tel||'';
  o.antes=c.ult||''; o.quando=HOJE();
  if(c.aps&&c.aps.length>1&&sel&&sel.length){
    o.aps=sel.map(x=>{const a=c.aps[x.i];return {lin:a.lin,nome:a.nome,prod:x.prod||a.prod||'',obs:a.obs||''}});
    o.lin=o.aps[0].lin;
    o.prod=[...new Set(o.aps.map(a=>a.prod).filter(Boolean))].join(' / ');
    // reflete na base local: só os aparelhos trocados ganham a data nova
    sel.forEach(x=>{c.aps[x.i].ult=d});
  } else {
    delete o.aps;
    o.prod=prodNovo||c.prod||'';
    if(c.aps) c.aps.forEach(a=>{a.ult=d});
    c.ult=d;
  }
  if(!j) TROCAS.push(o);
  salvarTr();
  recalcular([c]);
  const r=LOG[id]||{}; r.a='trocou'; r.d=HOJE(); r.n=c.nome; LOG[id]=r; salvar();
  try{localStorage.setItem('crm_demo_base',JSON.stringify({q:BASEDATA,d:DADOS}))}catch(e){}
  // se havia agendamento aberto de troca, marca como feito
  agDoCliente(id).filter(a=>a.serv==='Troca de refil').forEach(a=>{a.feito=true});
  salvarAg();
  atualizarTudo();
  toast(primeiroNome(c.nome)+' — troca registrada em '+fmt(d)+'. Vá em “Trocas feitas” para atualizar a planilha.','ok');
}
function modalTroca(id){
  const c=DADOS.find(x=>x.id===id); if(!c) return;
  const varios=c.aps&&c.aps.length>1;
  abrirModal(`
    <h3>Troquei o refil</h3>
    <div class="msub">${esc(c.nome)} — isso entra na lista de trocas para atualizar a Planilha de Clientes.</div>
    <div class="mlin">
      <div class="mcampo"><span class="rotulo">Dia da troca</span>
        <input type="date" id="tdata" autofocus value="${HOJE()}"></div>
      ${varios?'':`<div class="mcampo"><span class="rotulo">Refil colocado</span>
        <input type="text" id="tprod" value="${esc(c.prod||'')}" placeholder="Refil…"></div>`}
    </div>
    ${varios?`<div class="mcampo"><span class="rotulo">Quais aparelhos foram trocados <span class="opc">(desmarque o que não foi)</span></span>
      <div class="trapar">${c.aps.map((a,i)=>`<label class="trap">
        <input type="checkbox" data-api="${i}" checked>
        <span class="trq"><b>${esc(apRot(c,a,i))}</b><small>linha ${a.lin} · última troca ${fmt(a.ult)}</small></span>
        <input type="text" data-app="${i}" value="${esc(a.prod||'')}" placeholder="Refil…" title="Refil colocado"></label>`).join('')}</div></div>`:''}
    <div class="mrod">
      <button class="bt" data-fechar>Cancelar</button>
      <button class="bt pri" id="tok">Registrar troca</button>
    </div>`, m=>{
    m.querySelector('#tok').addEventListener('click',()=>{
      const d=m.querySelector('#tdata').value;
      if(!d){toast('Escolha o dia da troca.','err');return}
      if(varios){
        const sel=[...m.querySelectorAll('[data-api]')].filter(x=>x.checked)
          .map(x=>({i:+x.dataset.api,prod:m.querySelector(`[data-app="${x.dataset.api}"]`).value.trim()}));
        if(!sel.length){toast('Marque pelo menos um aparelho.','err');return}
        registrarTroca(id,d,'',sel);
      } else registrarTroca(id,d,m.querySelector('#tprod').value.trim());
      fecharModal();
    });
  });
}

/* ----------------------------------------------------------
   3e. REGISTRAR CONTATO (29/09/2026)
   Um clique no "📞 Registrar" da linha: o que aconteceu, o que a
   pessoa disse e quando voltar a falar. Cada registro entra no
   histórico do cliente (LOG[id].h) — nada é sobrescrito.
   ---------------------------------------------------------- */
function inicioMesQueVem(){
  const h=new Date(HOJE()+'T00:00:00'), d=new Date(h.getFullYear(),h.getMonth()+1,1);
  if(d.getDay()===0) d.setDate(2); else if(d.getDay()===6) d.setDate(3);
  return _iso(d);
}
const VOLTAS=[['1','Amanhã'],['3','Em 3 dias'],['7','Em 1 semana'],['15','Em 15 dias'],['mes','Início do mês que vem'],['','Sem data']];
const dataVolta=k=>!k?'':k==='mes'?inicioMesQueVem():k==='cartao'?somaDias(HOJE(),15):somaDias(HOJE(),+k);

function modalRegistrar(id,op){
  op=op||{};
  const c=DADOS.find(x=>x.id===id); if(!c) return;
  const np=nomePartes(c.nome), N=zapNums(c);
  let T=op.t||'', V=null;   // V=null: ainda não escolheu data
  abrirModal(`
    <h3>${op.volta?'Como foi o contato?':'Registrar contato'}</h3>
    <div class="msub"><b>${esc(np.n)}</b> · ${esc([c.tel,c.loc].filter(Boolean).join(' · ')||'sem telefone')}
      ${op.volta?'<br>Você acabou de abrir o WhatsApp desse cliente.':''}</div>

    <span class="rotulo">O que aconteceu</span>
    <div class="rres" id="rres">
      ${['falei','recado','msg'].map(k=>`<button type="button" data-t="${k}" class="${T===k?'on':''}">
        <span class="ic">${TCONT[k].ic}</span>${esc(TCONT[k].rot)}</button>`).join('')}
    </div>
    <button type="button" class="rnq${T==='naoquer'?' on':''}" data-t="naoquer">✋ Não quer mais</button>

    <div class="mcampo" style="margin-top:14px">
      <span class="rotulo">O que a pessoa disse <span class="opc">(opcional)</span></span>
      <textarea id="rtxt" rows="2" placeholder="Ex.: vai esperar o cartão virar; pediu para ligar depois das 17h">${esc(op.o||'')}</textarea>
      <div class="rfrases" id="rfrases">${FRASES.map(([f,v],i)=>`<button type="button" class="chip" data-f="${i}">${esc(f)}</button>`).join('')}</div>
    </div>

    <div class="mcampo">
      <span class="rotulo">Voltar a falar em</span>
      <div class="rvoltas" id="rvoltas">${VOLTAS.map(([k,r])=>`<button type="button" class="chip" data-v="${k}">${esc(r)}</button>`).join('')}
        <input type="date" id="rdata" value="" title="Escolher outra data"></div>
      <div class="rdica" id="rdica"></div>
    </div>

    ${LOG[id]&&LOG[id].h&&LOG[id].h.length?`<details class="rhist"><summary>Contatos anteriores (${LOG[id].h.length})</summary>${histHTML(c,6)}</details>`:''}

    <div class="mrod">
      <button class="bt" data-fechar>${op.volta?'Agora não':'Cancelar'}</button>
      <button class="bt pri" id="rok">Salvar</button>
    </div>`, m=>{
    m.classList.add('mreg');
    const tx=m.querySelector('#rtxt'), dt=m.querySelector('#rdata'), dica=m.querySelector('#rdica');
    let agendarDepois=false;
    function pintaT(){
      m.querySelectorAll('#rres button,.rnq').forEach(b=>b.classList.toggle('on',b.dataset.t===T));
    }
    function pintaV(){
      m.querySelectorAll('#rvoltas .chip').forEach(b=>b.classList.toggle('on',V!==null&&dataVolta(b.dataset.v)===V&&(b.dataset.v!==''||V==='')));
      if(V!==null&&dt.value!==V) dt.value=V||'';
      dica.textContent=V?('Volta para a fila em '+fmt(V)+' ('+diaSem(V)+'). Até lá, sai da Fila de hoje.'):(V===''?'Sem data de retorno.':'');
    }
    m.querySelector('#rres').addEventListener('click',e=>{const b=e.target.closest('[data-t]');if(!b)return;T=b.dataset.t;pintaT()});
    m.querySelector('.rnq').addEventListener('click',()=>{T='naoquer';pintaT()});
    m.querySelector('#rfrases').addEventListener('click',e=>{
      const b=e.target.closest('[data-f]'); if(!b) return;
      const [f,v]=FRASES[+b.dataset.f];
      const cur=tx.value.trim();
      tx.value=cur?(cur.replace(/[.;,]\s*$/,'')+'; '+f.toLowerCase()):f;
      if(!T) {T='falei';pintaT()}
      if(/^(Já trocou com outra|Não usa mais)/.test(f)){T='naoquer';pintaT();V='';pintaV()}
      else if(v&&V===null){V=dataVolta(v);pintaV()}
      if(/^Vai trocar/.test(f)) agendarDepois=true;
      tx.focus();
    });
    m.querySelector('#rvoltas').addEventListener('click',e=>{const b=e.target.closest('.chip[data-v]');if(!b)return;V=dataVolta(b.dataset.v);pintaV()});
    dt.addEventListener('change',()=>{V=dt.value||'';pintaV()});
    m.addEventListener('click',e=>{
      const hd=e.target.closest('[data-hdel]'); if(!hd) return;
      e.stopPropagation();
      if(!confirm('Apagar este registro do histórico?')) return;
      const [cid,ix]=hd.dataset.hdel.split('|'); apagarContato(cid,+ix);
      fecharModal(); modalRegistrar(id,op); redesenharCartao(id); kpis();
    });
    function salvarReg(){
      if(!T){toast('Escolha o que aconteceu: falou, não atendeu ou mandou mensagem.','err');return}
      registrarContato(id,{t:T,o:tx.value.trim().replace(/\s+/g,' '),v:V||''});
      fecharModal();
      const nome=primeiroNome(np.n);
      toast(nome+' — '+TCONT[T].curto.toLowerCase()+(V?' · volta em '+fmt(V):'')+'.','ok');
      if(VISTA==='fila'||VISTA==='base'){kpis();redesenharCartao(id);
        const nova=$(`.linha[data-id="${id}"]`);if(nova){nova.classList.add('marcou');setTimeout(()=>nova.classList.remove('marcou'),820)}}
      else atualizarTudo();
      if(agendarDepois) setTimeout(()=>agendarCliente(id),250);
    }
    m.querySelector('#rok').addEventListener('click',salvarReg);
    m.addEventListener('keydown',e=>{if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();salvarReg()}});
    pintaT();
  });
}

/* Depois de abrir o WhatsApp de um cliente, quando você volta para o CRM ele já
   pergunta "Como foi o contato?" — assim não tem como esquecer de registrar. */
let PENDZAP=null;
function esperarVolta(id){PENDZAP={id,t:Date.now()}}
function voltouParaCRM(){
  if(!PENDZAP||document.hidden) return;
  const p=PENDZAP; if(Date.now()-p.t<2500) return;
  PENDZAP=null;
  if(!$('#fundo').hidden) return;
  modalRegistrar(p.id,{volta:true});
}
window.addEventListener('focus',()=>setTimeout(voltouParaCRM,150));
document.addEventListener('visibilitychange',()=>setTimeout(voltouParaCRM,150));

/* texto pronto para colar no Claude */
function textoPrompt(lista){
  const porData={};
  lista.forEach(t=>(porData[t.d]=porData[t.d]||[]).push(t));
  const datas=Object.keys(porData).sort();
  let s='Atualize a Planilha de Clientes com as trocas de refil abaixo:\n\n';
  for(const d of datas){
    s+='Trocaram o refil no dia '+fmt(d)+':\n';
    porData[d].forEach(t=>{
      if(t.aps&&t.aps.length) t.aps.forEach(a=>{
        s+='- '+(a.nome||t.nome)+' (aba '+t.aba+', linha '+a.lin+(a.obs?' · '+a.obs:'')+')'
          +(a.prod?' — '+a.prod:'')+'\n';
      });
      else s+='- '+t.nome+' (aba '+t.aba+', linha '+t.lin+')'
        +(t.prod?' — '+t.prod:'')+'\n';
    });
    s+='\n';
  }
  s+='Para cada um: colocar a data da troca na coluna DATA e recalcular o vencimento.';
  return s;
}
function modalPrompt(){
  const L=trPend();
  if(!L.length){toast('Não há trocas esperando para entrar na planilha.','err');return}
  const txt=textoPrompt(L);
  abrirModal(`
    <h3>Pedido pronto para o Claude</h3>
    <div class="msub">Copie o texto e cole no Claude. Depois clique em “Marcar como atualizadas”.</div>
    <div class="prompt" id="ptxt">${esc(txt)}</div>
    <div class="mrod">
      <button class="bt" data-fechar>Fechar</button>
      <button class="bt" id="pmark">✓ Marcar como atualizadas</button>
      <button class="bt pri" id="pcopy">📋 Copiar texto</button>
    </div>`, m=>{
    m.querySelector('#pcopy').addEventListener('click',()=>{
      navigator.clipboard.writeText(txt).then(()=>toast('Texto copiado — cole no Claude.','ok'))
        .catch(()=>toast('Não consegui copiar. Selecione o texto acima e copie com Ctrl+C.','err'));
    });
    m.querySelector('#pmark').addEventListener('click',()=>{
      L.forEach(t=>{t.ok=true;t.okd=HOJE()});
      salvarTr(); atualizarTudo(); fecharModal();
      toast(L.length+(L.length===1?' troca marcada':' trocas marcadas')+' como já atualizadas na planilha.','ok');
    });
  });
}

/* ----------------------------------------------------------
   4. MENSAGENS DE WHATSAPP
   ---------------------------------------------------------- */
const PREP=['de','da','do','das','dos','e','di','du','del'];
function primeiroNome(t){
  const w=String(t||'').split(/[-–(\[\/]/)[0].trim().split(/\s+/).filter(Boolean);
  let r=w[0]||''; if(w[1]&&!PREP.includes(w[1].toLowerCase())) r+=' '+w[1];
  return r;
}
const trata=c=>(c.fc||primeiroNome(c.nome));

/* Nunca citamos a marca do refil — quem escolhe o modelo é a loja. */
const MODELOS=[
 {k:'aviso',   r:'Aviso de vencimento', f:c=>c.st==='vencido'
    ? `Olá, ${trata(c)}! Aqui é da Fonte Clara. Passando para avisar que o refil do seu purificador venceu em ${fmt(c.venc)}. Quer que eu já separe o refil para a troca?`
    : `Olá, ${trata(c)}! Aqui é da Fonte Clara. O refil do seu purificador vence em ${fmt(c.venc)}. Quer agendar a troca?`},
 {k:'lembrete',r:'Lembrete (2º contato)', f:c=>
    `Olá, ${trata(c)}! Aqui é da Fonte Clara de novo. Passei uma mensagem esses dias sobre a troca do refil do seu purificador e não sei se chegou a ver. Quer que eu agende?`},
 {k:'pensar',  r:'Retomar quem ia pensar', f:c=>
    `Olá, ${trata(c)}! Aqui é da Fonte Clara. Você tinha ficado de ver a troca do refil do seu purificador. Já consegue definir uma data? Consigo encaixar essa semana.`},
 {k:'sumido',  r:'Cliente sumido', f:c=>
    `Olá, ${trata(c)}! Aqui é da Fonte Clara. Faz um tempo que não trocamos o refil do seu purificador${c.ult?' — a última troca foi em '+fmt(c.ult):''}. Está tudo certo com o aparelho? Se quiser, dou uma passada para verificar.`}
];
const modeloPadrao=c=>MODELOS[0].f(c);

/* --- números que dá para chamar no WhatsApp ---------------------------
   O telefone 1 vale pela coluna TIPO DE TELEFONE da planilha ("Celular").
   O telefone 2 não tem coluna de tipo, então a gente deduz: celular na
   Bahia tem 11 dígitos com DDD e começa com 9 depois do DDD.
   -------------------------------------------------------------------- */
const ehCel=n=>{const d=dig(n);return d.length===11&&d[2]==='9'};
function zapNums(c){
  const o=[], visto=new Set();
  const por=(n,rot)=>{const d=dig(n); if(!d||visto.has(d))return; visto.add(d); o.push({n:n,d:d,rot:rot})};
  if(c.tp==='Celular'&&dig(c.tel).length>=10) por(c.tel,'1');
  else if(ehCel(c.tel)) por(c.tel,'1');
  if(ehCel(c.tel2)) por(c.tel2,'2');
  return o;
}
const temZap=c=>zapNums(c).length>0;
function linkZap(c,texto,qual){
  const N=zapNums(c); if(!N.length) return null;
  const alvo=qual?N.find(x=>x.rot===qual)||N[0]:N[0];
  return 'https://wa.me/55'+alvo.d+'?text='+encodeURIComponent(texto);
}

/* --- última vez que falei com o cliente ------------------------------
   Vem do próprio CRM (LOG[id].uc), gravado quando você abre o WhatsApp,
   copia a mensagem ou marca uma situação. Se o cliente nunca foi contatado
   aqui, cai no registro antigo importado da planilha antiga de contatos.
   -------------------------------------------------------------------- */
function ultContato(c){
  const r=LOG[c.id];
  if(r&&r.uc) return {d:r.uc,fonte:'crm'};
  if(c.ct&&c.ct.d) return {d:c.ct.d,fonte:'planilha',aba:c.ct.aba,s:c.ct.s,o:c.ct.o};
  return null;
}
function rotUltContato(uc){
  const n=diasDe(uc.d);
  return '💬 falei '+fmt(uc.d)+(n>0?' · há '+n+(n===1?' dia':' dias'):n===0?' · hoje':'');
}
function marcarContato(id){
  const r=LOG[id]||{}, c=DADOS.find(x=>x.id===id);
  if(c&&!r.n) r.n=c.nome;
  r.uc=HOJE(); LOG[id]=r; salvar();
  esperarVolta(id);
}
/* atualiza só a tarja daquele cliente, sem redesenhar a lista inteira */
function pintarUltContato(id){ atualizarCabecalhoCartao(id); }

/* ----------------------------------------------------------
   5. FILTROS DA LISTA
   ---------------------------------------------------------- */
/* os campos de filtro são lidos UMA vez por redesenho (lerFiltros) — antes cada
   um dos 1.444 clientes consultava a tela 5 vezes. O texto de busca de cada
   cliente também fica guardado (c._busca) e só é refeito se a anotação mudar. */
function lerFiltros(){
  const q=norm($('#q').value.trim());
  return {q, termos:q?q.split(/\s+/):[], L:$('#loc').value, P:$('#prod').value, M:$('#mes').value, T:$('#tempo').value};
}
function textoBusca(c){
  const u=ultReg(c), a=anot(c)+(u&&u.o?' '+u.o:'');
  if(c._busca===undefined||c._buscaAnot!==a){
    const ap=c.aps?c.aps.map(x=>x.nome+' '+x.prod+' '+(x.obs||'')).join(' '):'';
    c._busca=norm(c.nome+' '+c.end+' '+c.loc+' '+c.fc+' '+c.prod+' '+(c.obs||'')+' '+ap+' '+a)+' '+dig(c.tel)+' '+dig(c.tel2);
    c._buscaAnot=a;
  }
  return c._busca;
}
function passa(c,F){
  F=F||lerFiltros();
  const S=sit(c);
  const q=F.q;
  if(q){
    const alvo=textoBusca(c);
    if(!F.termos.every(t=>alvo.includes(t))) return false;
  }
  if(!q){
    if(filtro==='fila'&&(S==='naoquer'||S==='trocou')) return false;
    /* na fila entram os vencimentos próximos/recentes E quem tem retorno vencendo
       hoje (mesmo que o refil tenha vencido há mais tempo) */
    if(filtro==='fila'&&!naFila(c)&&!vencido_retorno(c)) return false;
    /* combinou de voltar a falar numa data que ainda não chegou → sai da fila até lá */
    if(filtro==='fila'&&volta(c)>HOJE()) return false;
    /* já tem serviço agendado → não precisa ligar */
    if(filtro==='fila'&&agDoCliente(c.id).length) return false;
    if(filtro==='b30'&&!(c.dias<=0&&c.dias>=-30)) return false;
    if(filtro==='aberto'&&!(S==='aguardando'||S==='naoatendeu'||S==='pensar')) return false;
    if(filtro==='agendado'&&S!=='agendado') return false;
    if(filtro==='retorno'&&!vencido_retorno(c)) return false;
    if(filtro==='ligar'&&!paraLigar(c)) return false;
    if(filtro==='vencido'&&c.st!=='vencido') return false;
    if(filtro==='breve'&&c.st!=='vence em breve'&&c.st!=='vence este mes') return false;
    if(filtro==='sumido'&&!(c.st==='vencido'&&c.dias>365)) return false;
    if(filtro==='semwa'&&c.tp==='Celular') return false;
    if(filtro==='novos'&&(ultContato(c)||!naFila(c)||S==='naoquer'||S==='trocou'||agDoCliente(c.id).length)) return false;
    if(filtro==='semretorno'&&!(c.ct&&c.st==='vencido')) return false;
  }
  if(tipo!=='todos'&&c.tipo!==tipo) return false;
  const L=F.L,P=F.P,M=F.M,T=F.T;
  if(L&&c.loc!==L) return false;
  if(P&&c.prod!==P&&!(c.aps&&c.aps.some(x=>x.prod===P))) return false;
  if(M&&(c.venc||'').slice(0,7)!==M) return false;
  if(T){
    if(c.st==='sem data') return false;
    const d=c.dias;
    if(T==='v30'&&!(d>0&&d<=30)) return false;
    if(T==='v90'&&!(d>30&&d<=90)) return false;
    if(T==='v180'&&!(d>90&&d<=180)) return false;
    if(T==='v365'&&!(d>180&&d<=365)) return false;
    if(T==='v365+'&&!(d>365)) return false;
    if(T==='v730+'&&!(d>730)) return false;
    if(T==='a30'&&!(d<=0&&d>=-30)) return false;
    if(T==='a60'&&!(d<=0&&d>=-60)) return false;
    if(T==='a90'&&!(d<=0&&d>=-90)) return false;
  }
  return true;
}
/* conta um filtro sem depender da busca/selects da tela (para os contadores) */
const lerFiltrosVazio=()=>({q:'',termos:[],L:'',P:'',M:'',T:''});
function passaFiltro(c,f,F){const fa=filtro,ta=tipo;filtro=f;tipo='todos';const r=passa(c,F);filtro=fa;tipo=ta;return r}
function filtrados(){
  const F=lerFiltros();
  return DADOS.filter(c=>passa(c,F)).sort((a,b)=>{
    const pa=peso(a)-peso(b); if(pa) return pa;
    if(filtro==='breve') return a.dias-b.dias;
    return b.dias-a.dias;
  });
}

/* ----------------------------------------------------------
   6. LISTA DE CLIENTES
   ---------------------------------------------------------- */
/* nome sem as observações entre colchetes — elas viram etiquetas discretas
   (ex.: "Maria [Comprou refil na internet]" → Maria + etiqueta) */
function nomePartes(nome){
  const t=[];
  const n=String(nome||'').replace(/\[([^\]]+)\]/g,(m,x)=>{t.push(x.trim());return ' '}).replace(/\s+/g,' ').trim();
  return {n:n||String(nome||''),t};
}
function cartao(c){
  const [cl,rot]=ST[c.st]||ST['sem data'];
  const aa=anot(c), vv=volta(c), sk=sit(c);
  const prazo=c.st==='vencido'
    ? (c.dias>365?Math.floor(c.dias/365)+(c.dias>730?' anos':' ano'):c.dias+' dias')
    : (c.venc?fmt(c.venc):'');
  const atraso=vv&&vv<=HOJE();
  const NZ=zapNums(c), u=ultReg(c), uc=ultContato(c), np=nomePartes(c.nome);

  /* segunda linha: o último contato do histórico; sem histórico, o registro antigo */
  let lin2='';
  if(u) lin2=`<small class="ult${u.t==='falei'?' falou':''}">${esc(rotUltReg(u))}</small>`;
  else if(uc) lin2=`<small class="ult velho">${esc(rotUltContato(uc))}</small>`;

  const btzap=NZ.length
    ? NZ.map(x=>`<a class="zap" href="${linkZap(c,modeloPadrao(c),x.rot)}" target="_blank" rel="noopener"
        data-zap="${c.id}" title="Abrir conversa com ${esc(x.n)}">WhatsApp${NZ.length>1?' '+x.rot:''}</a>`).join('')
    : '<span class="zap off" title="Sem celular cadastrado">sem zap</span>';

  const ags=agDoCliente(c.id);
  const tp=TROCAS.find(t=>t.cid===c.id&&!t.ok);

  return `<div class="linha${ABERTOS.has(c.id)?' open':''}" data-id="${c.id}">
  <div class="cab">
    <span class="pegar" draggable="true" data-drag="${c.id}" title="Arraste para a Agenda de serviços">⠿</span>
    <div class="pino" style="background:${COR[c.st]}"></div>
    <div class="quem"><b>${esc(np.n)}</b>
      <span>${esc([c.loc,c.prod,c.tel||'sem telefone'].filter(Boolean).join(' · '))}${np.t.map(x=>` <i class="tagnome" title="${esc(x)}">${esc(x)}</i>`).join('')}</span>
      ${lin2}</div>
    <div class="tags">
      ${c.aparelhos?`<span class="mini">${c.aparelhos} aparelhos</span>`:''}
      ${vv?`<span class="mini ${atraso?'t-bad':'t-info'}">${atraso?'⏰ ':''}voltar ${fmtc(vv)}</span>`:''}
      ${paraLigar(c)?'<span class="mini t-warn" title="Mandou mensagem há 3 dias ou mais e não teve resposta">☎️ hora de ligar</span>':''}
      ${sk==='naoquer'?'<span class="mini t-bad">não quer mais</span>':''}
      ${ags.length?`<span class="mini t-ok">📆 ${esc(ags[0].serv)} ${fmtc(ags[0].d)}</span>`:''}
      ${tp?`<span class="mini t-warn">🔧 trocou ${fmtc(tp.d)} · falta planilha</span>`:''}
      ${aa?`<span class="mini so-ic nota" title="${esc(aa)}">📝</span>`:''}
      <span class="selo ${cl}">${rot}${prazo?' · '+prazo:''}</span>
    </div>
    <div class="bts">
      <button class="bag" data-ag="${c.id}" title="Agendar serviço">📆<span> Agendar</span></button>
      <button class="breg" data-reg="${c.id}" title="Registrar como foi o contato">📞<span> Registrar</span></button>
      ${c.tel?`<button class="bcp" data-cptel="${esc(c.tel)}" title="Copiar o telefone ${esc(c.tel)}">📋<span> Copiar</span></button>`:''}
      ${btzap}
    </div>
    <span class="seta">▶</span>
  </div>
  <div class="det"><div class="dentro"><div class="detpad">${ABERTOS.has(c.id)?detalhe(c):''}</div></div></div></div>`;
}

/* nome de cada aparelho: a observação da planilha (ex.: "1º andar"); se as linhas
   têm nomes diferentes (Ótica térreo / 1º andar), o nome; senão "Aparelho 2" */
function apRot(c,a,i){
  if(a.obs) return a.obs;
  if(a.nome&&_n(a.nome)!==_n(c.nome)) return nomePartes(a.nome).n;
  return 'Aparelho '+(i+1);
}
function apsHTML(c){
  if(!c.aps||c.aps.length<2) return '';
  return `<span class="rotulo">Aparelhos (${c.aps.length})</span>
    <div class="aplista">${c.aps.map((a,i)=>{const [cl,rot]=ST[a.st]||ST['sem data'];
      return `<div class="aplin"><i class="pino" style="background:${COR[a.st]}"></i>
        <span class="apq"><b>${esc(apRot(c,a,i))}</b><small>${esc(a.prod||'—')}</small></span>
        <span class="apd"><small>trocou ${fmt(a.ult)}</small><small>vence ${fmt(a.venc)}</small></span>
        <span class="selo ${cl}">${rot}</span></div>`}).join('')}</div>`;
}

/* lista do histórico, do mais novo para o mais velho */
function histHTML(c,max){
  const r=LOG[c.id], H=r&&r.h?r.h.map((x,i)=>({x,i})).sort((a,b)=>(b.x.d||'').localeCompare(a.x.d||'')||b.i-a.i):[];
  if(!H.length) return '<div class="hvazio">Nenhum contato registrado ainda.</div>';
  const L=max?H.slice(0,max):H;
  return '<ol class="hlista">'+L.map(({x,i})=>{
    const T=TCONT[x.t]||{ic:'•',rot:x.t};
    return `<li class="h-${x.t}"><span class="hd">${fmtc(x.d)}${x.d.slice(0,4)!==HOJE().slice(0,4)?'/'+x.d.slice(2,4):''}</span>
      <span class="hic">${T.ic}</span>
      <span class="hq"><b>${esc(T.rot)}</b>${x.cel?' <i class="hcel" title="Registrado no celular">celular</i>':''}${x.o?`<span>“${esc(x.o)}”</span>`:''}${x.v?`<em>voltar em ${fmt(x.v)}</em>`:''}</span>
      <button class="hdel" data-hdel="${c.id}|${i}" title="Apagar este registro">✕</button></li>`}).join('')
    +'</ol>'+(max&&H.length>max?`<div class="hmais">+ ${H.length-max} registros mais antigos</div>`:'');
}

/* O miolo do cartão (campos, histórico, botões) só é montado quando o
   cliente é aberto — antes vinha pronto para os 400 da tela de uma vez,
   o que deixava o navegador pesado à toa. */
function detalhe(c){
  const reg=LOG[c.id], aa=anot(c), vv=volta(c), sk=sit(c), dLig=diasDe(reg&&reg.d);
  const NZ=zapNums(c), uc=ultContato(c), temH=reg&&reg.h&&reg.h.length;
  const ags=agDoCliente(c.id);
  const mods=NZ.length
    ? MODELOS.map(m=>`<button class="mod" data-m="${m.k}" data-id="${c.id}">${esc(m.r)}</button>`).join('')
      +`<button class="mod copiar" data-cp="${c.id}">📋 Copiar o texto padrão</button>`
      +(NZ.length>1?`<div class="mini" style="width:100%;margin-top:6px">Dois números — depois de escolher a mensagem, você decide para qual mandar.</div>`:'')
    : `<span class="mini">Sem WhatsApp cadastrado — telefone marcado como “${esc(c.tp||'sem tipo')}”.</span>`;
  return `
    ${comJanela()?pophead(c):''}
    ${c.sv?`<div class="faixa f-sv"><b>Serviço feito em ${fmt(c.sv.d)}</b> — ${esc(c.sv.o||'—')} · ${esc(c.sv.p||'')}</div>`:''}
    ${c.obs?`<div class="faixa f-obs">${esc(c.obs)}</div>`:''}
    ${c.tambem?`<div class="faixa f-obs">Mesmo telefone de: <b>${esc(c.tambem.join(' · '))}</b></div>`:''}

    <div class="detgrade">
      <div class="detcol">
        <div class="secab"><span class="rotulo">Contatos</span>
          <button class="bt pri peq" data-reg="${c.id}">📞 Registrar contato</button></div>
        ${histHTML(c)}
        ${uc&&uc.fonte!=='crm'&&!temH?`<div class="hvelho">Registro antigo (planilha antiga de contatos, ${esc(uc.aba||'')}): ${fmt(uc.d)}${uc.s?' · '+esc(uc.s):''}${uc.o?' — '+esc(uc.o):''}</div>`:''}

        <div class="anota">
          <div class="cx"><span class="rotulo">Anotação fixa sobre o cliente</span>
            <textarea data-obs="${c.id}" rows="2" placeholder="O que vale sempre — ex.: só atende à tarde; falar com a esposa; portão azul">${esc(aa)}</textarea></div>
          <div class="dt"><span class="rotulo">Voltar a falar em</span>
            <input type="date" data-v="${c.id}" value="${vv}"></div>
        </div>
      </div>

      <div class="detcol">
        <div class="campos">
          <div class="campo"><i>Endereço</i>${esc(c.end||'—')}</div>
          <div class="campo"><i>Bairro / cidade</i>${esc(c.loc||'—')}</div>
          ${c.aps&&c.aps.length>1?'':`<div class="campo"><i>Produto</i>${esc(c.prod||'—')}</div>
          <div class="campo"><i>Última troca</i>${fmt(c.ult)}</div>
          <div class="campo"><i>Vencimento</i>${fmt(c.venc)} (${c.vld} dias)</div>`}
          <div class="campo"><i>Telefone</i>${btTel(c.tel)||'—'} ${c.tp?`<span class="mini">${esc(c.tp)}</span>`:''}</div>
          ${c.tel2?`<div class="campo"><i>Telefone 2</i>${btTel(c.tel2)}</div>`:''}
          ${c.fc?`<div class="campo"><i>Falar com</i>${esc(c.fc)}</div>`:''}
          <div class="campo"><i>Cadastro</i>${esc(c.tipo)} · ${esc(c.aba)} ${linTxt(c)}</div>
        </div>
        ${apsHTML(c)}

        <span class="rotulo">Serviço</span>
        <div class="acoes">
          <button class="forte" data-ag="${c.id}">📆 Agendar serviço</button>
          <button class="forte" data-tr="${c.id}">🔧 Troquei o refil</button>
          ${ags.length?ags.map(a=>`<button data-ed="${a.id}" title="Editar esse agendamento">✏️ ${esc(a.serv)} · ${fmt(a.d)}${a.h?' '+a.h:''}</button>`).join(''):''}
        </div>

        <span class="rotulo" style="margin-top:14px">Mensagem pronta para o WhatsApp</span>
        <div class="modelos">${mods}</div>
        <div class="previa-slot" data-slot="${c.id}"></div>
      </div>
    </div>
    <div class="hist">${reg&&SIT[reg.a]?`Situação: <b>${SIT[reg.a].rot}</b> desde ${fmt(reg.d)}${dLig!==null?' (há '+dLig+(dLig===1?' dia':' dias')+')':''}`:'Sem contato registrado — use “Registrar contato” depois de falar com o cliente.'}</div>
  `;
}

function resumoFiltros(){
  const d=[];
  if($('#tempo').value) d.push($('#tempo').selectedOptions[0].textContent);
  if($('#mes').value) d.push($('#mes').selectedOptions[0].textContent);
  if($('#loc').value) d.push($('#loc').value);
  if($('#prod').value) d.push($('#prod').value);
  if(tipo!=='todos') d.push(tipo);
  $('#fresumo').textContent=d.length?d.join(' · '):'';
}

/* ----------------------------------------------------------
   5b. QUADRO — clientes em 3 colunas por etapa (14/10/2026)
   Mesmos cartões da lista, agrupados pela situação do cliente.
   Clicar no cartão abre as informações numa janela por cima da tela
   (Esc, ✕ ou clicar fora fecha). Arrastar pela alça ⠿ para outra coluna
   abre a janela certa (registrar contato, agendar ou registrar troca).
   O botão "Lista" volta ao formato de antes.
   ---------------------------------------------------------- */
let MODO='quadro';
try{const m=localStorage.getItem('crm_demo_modo'); if(m==='lista'||m==='quadro') MODO=m}catch(e){}
const COLS=[
  {id:'novo', nome:'Para contatar',      dica:'Ninguém falou ainda',                 cor:'var(--sem)'},
  {id:'cont', nome:'Em contato',         dica:'Mandei mensagem, falei ou deixei recado', cor:'var(--info)'},
  {id:'feito',nome:'Agendado ou trocou', dica:'Serviço marcado ou refil já trocado', cor:'var(--ok)'}];
const QPASSO=20, QLIM={};
function colDe(c){
  const k=sit(c);
  if(k==='naoquer') return 'naoquer';
  if(k==='trocou'||trPend().some(t=>t.cid===c.id)) return 'feito';
  if(agDoCliente(c.id).length||k==='agendado') return 'feito';
  if(k) return 'cont';
  return 'novo';
}
/* etiqueta com a etapa exata, já que "Em contato" junta várias */
function etapaRot(c){
  const k=sit(c);
  if(k==='trocou'||trPend().some(t=>t.cid===c.id)) return 'Já trocou';
  if(agDoCliente(c.id).length||k==='agendado') return 'Agendado';
  if(k==='aguardando') return 'Mandei mensagem';
  if(k==='naoatendeu') return 'Não atendeu';
  if(k==='falei') return 'Falei';
  if(k==='pensar') return 'Vai pensar';
  return '';
}
const emQuadro=()=>MODO==='quadro'&&(VISTA==='fila'||VISTA==='base');
/* 23/10/2026: nos dois modos (Quadro e Lista) o cliente abre numa janela */
const comJanela=()=>VISTA==='fila'||VISTA==='base';
function cartaoQ(c){
  const r=etapaRot(c), h=cartao(c);
  return r?h.replace('<div class="tags">',`<div class="tags"><span class="mini t-etapa">${esc(r)}</span>`):h;
}
/* cabeçalho da janela do cliente */
function pophead(c){
  const np=nomePartes(c.nome);
  return `<div class="pophead"><div><b>${esc(np.n)}</b><span>${esc([c.tel,c.loc,c.prod].filter(Boolean).join(' · '))}</span></div>
    <button type="button" class="popx" data-popx title="Fechar (Esc)">✕</button></div>`;
}
function sincPop(){
  const abre=comJanela()&&!!$('#lista .linha.open');
  $('#lista').classList.toggle('tem-pop',abre);
  document.body.style.overflow=abre?'hidden':'';
}
function fecharPop(){
  document.querySelectorAll('#lista .linha.open').forEach(l=>{ABERTOS.delete(l.dataset.id);l.classList.remove('open')});
  sincPop();
}
function renderQuadro(L){
  const g={}; COLS.forEach(k=>g[k.id]=[]); g.naoquer=[];
  L.forEach(c=>g[colDe(c)].push(c));
  const lim=id=>QLIM[id]||QPASSO;
  $('#lista').className='modo-quadro com-janela';
  $('#lista').innerHTML='<div class="qgrade">'+COLS.map(k=>{
    const arr=g[k.id], resto=arr.length-lim(k.id);
    return `<section class="qcol" data-col="${k.id}" style="--c:${k.cor}">
      <div class="qhead"><i></i><h3>${k.nome}</h3><span class="qn">${arr.length}</span><p>${k.dica}</p></div>
      <div class="qcards">${arr.length?arr.slice(0,lim(k.id)).map(cartaoQ).join(''):'<div class="qvazio">Solte um cliente aqui</div>'}
      ${resto>0?`<button class="qmais" data-qmais="${k.id}">Mostrar mais ${Math.min(QPASSO,resto)} (faltam ${resto})</button>`:''}</div>
    </section>`}).join('')+'</div>'
    +(g.naoquer.length?`<div class="qnota">${g.naoquer.length} ${g.naoquer.length===1?'cliente marcado como “Não quer mais” aparece':'clientes marcados como “Não quer mais” aparecem'} só na Lista.</div>`:'');
  sincPop();
}
function pintarModo(){
  $('#vmQ').classList.toggle('on',MODO==='quadro');
  $('#vmL').classList.toggle('on',MODO==='lista');
  if(MODO==='lista'){ $('#lista').className=''; document.body.style.overflow='' }
}
$('#vmQ').addEventListener('click',()=>{MODO='quadro';try{localStorage.setItem('crm_demo_modo',MODO)}catch(e){};pintarModo();render(true)});
$('#vmL').addEventListener('click',()=>{MODO='lista';try{localStorage.setItem('crm_demo_modo',MODO)}catch(e){};pintarModo();render(true)});
$('#lista').addEventListener('click',e=>{
  const m=e.target.closest('[data-qmais]');
  if(m){QLIM[m.dataset.qmais]=(QLIM[m.dataset.qmais]||QPASSO)+QPASSO; render(); return}
  if(e.target.closest('[data-popx]')||e.target===$('#lista')&&$('#lista').classList.contains('tem-pop')){fecharPop()}
});
document.addEventListener('keydown',e=>{
  if(e.key==='Escape'&&$('#fundo').hidden&&$('#lista').classList.contains('tem-pop')) fecharPop();
});
/* arrastar entre colunas (a alça ⠿ já existia, para a agenda de serviços) */
document.addEventListener('dragover',e=>{
  if(!ARRASTA||!emQuadro()) return;
  const col=e.target.closest('.qcol'); if(!col) return;
  e.preventDefault(); e.dataTransfer.dropEffect='copy';
  document.querySelectorAll('.qcol.alvo').forEach(x=>{if(x!==col)x.classList.remove('alvo')});
  col.classList.add('alvo');
});
document.addEventListener('dragend',()=>document.querySelectorAll('.qcol.alvo').forEach(x=>x.classList.remove('alvo')));
document.addEventListener('drop',e=>{
  const col=e.target.closest&&e.target.closest('.qcol'); if(!col||!ARRASTA||!emQuadro()) return;
  e.preventDefault(); col.classList.remove('alvo');
  const id=ARRASTA, c=DADOS.find(x=>x.id===id); if(!c) return;
  const destino=col.dataset.col; if(destino===colDe(c)) return;
  if(destino==='cont') modalRegistrar(id,{});
  else if(destino==='feito'){
    abrirModal(`<h3>O que aconteceu com este cliente?</h3>
      <div class="msub"><b>${esc(nomePartes(c.nome).n)}</b></div>
      <div class="macoes" style="display:flex;gap:8px;flex-wrap:wrap;margin-top:14px">
        <button type="button" class="bt pri" id="mvAg">📆 Agendar serviço</button>
        <button type="button" class="bt" id="mvTr">🔧 Troquei o refil</button>
        <button type="button" class="bt" data-fechar>Cancelar</button></div>`,
      m=>{m.querySelector('#mvAg').onclick=()=>{fecharModal();agendarCliente(id)};
          m.querySelector('#mvTr').onclick=()=>{fecharModal();modalTroca(id)}});
  }
  else toast('Para voltar este cliente a "Para contatar", apague os registros dele no histórico.');
});

/* A lista mostra 60 clientes por vez (antes eram 400 de uma vez, o que pesava
   em computador lento). "Mostrar mais" acrescenta mais 60. Quando a busca, o
   filtro ou o cartão de resumo mudam, volta para os 60 primeiros. */
const BLOCO=60;
let LIMITE=BLOCO, _chaveLista='';
function render(anim){
  const L=filtrados();
  const F=lerFiltros(), chave=[VISTA,filtro,tipo,F.q,F.L,F.P,F.M,F.T].join('|');
  if(chave!==_chaveLista){LIMITE=BLOCO;for(const k in QLIM)delete QLIM[k];_chaveLista=chave}
  resumoFiltros();
  const q=$('#q').value.trim();
  $('#xq').hidden=!q;
  $('#conta').innerHTML = q
    ? `<em>${L.length}</em> de ${DADOS.length} clientes · busca em toda a base, vencidos e em dia`
    : `<em>${L.length}</em> de ${DADOS.length} clientes`;
  if(!L.length){
    $('#lista').className='';
    $('#lista').innerHTML='<div class="vazio"><span class="big">🔍</span>Nenhum cliente com esses filtros.<br>Tente limpar os filtros ou trocar o cartão selecionado acima.</div>';
    return;
  }
  if(emQuadro()){renderQuadro(L);return}
  $('#lista').className='com-janela';
  $('#lista').innerHTML = L.slice(0,LIMITE).map(cartao).join('') + rodapeLista(L.length);
  sincPop();
}
function rodapeLista(total){
  if(total<=LIMITE) return '';
  const falta=total-LIMITE, prox=Math.min(BLOCO,falta);
  return `<div class="maislista" id="maislista">
    <span>Mostrando <b>${LIMITE}</b> de ${total.toLocaleString('pt-BR')}</span>
    <button class="bt pri" data-mais="1">Mostrar mais ${prox}</button>
    ${falta>prox?`<button class="bt" data-mais="tudo">Mostrar todos (${total.toLocaleString('pt-BR')})</button>`:''}
  </div>`;
}
/* acrescenta os próximos cartões no fim, sem refazer os que já estão na tela */
function mostrarMais(tudo){
  const L=filtrados(), de=LIMITE;
  LIMITE=tudo?L.length:LIMITE+BLOCO;
  const r=$('#maislista'); if(r) r.remove();
  $('#lista').insertAdjacentHTML('beforeend', L.slice(de,LIMITE).map(cartao).join('') + rodapeLista(L.length));
}

/* ----------------------------------------------------------
   7. KPIS (com contagem animada)
   ---------------------------------------------------------- */
function contarDe(el,de,ate){
  if(de===ate){el.textContent=ate.toLocaleString('pt-BR');return}
  const dur=520, ini=performance.now();
  const passo=t=>{
    const p=Math.min(1,(t-ini)/dur), e=1-Math.pow(1-p,3);
    el.textContent=Math.round(de+(ate-de)*e).toLocaleString('pt-BR');
    if(p<1) requestAnimationFrame(passo);
  };
  requestAnimationFrame(passo);
}
let KPIANT=null;
function kpis(primeira){
  const n=f=>DADOS.filter(f).length;
  /* 29/09/2026: só os 5 números que mandam no dia a dia — o resto foi para o Painel */
  const F=lerFiltrosVazio();
  const K=[
    ['fila',      n(c=>passaFiltro(c,'fila',F)),  'Para contatar','n'],
    ['retorno',   n(vencido_retorno),              'Retornar hoje','r'],
    ['ligar',     n(paraLigar),                    'Hora de ligar','a'],
    ['novos',     n(c=>passaFiltro(c,'novos',F)),  'Nunca contatados','b'],
    ['b30',       n(c=>c.dias<=0&&c.dias>=-30),    'Vencem em 30 dias','g']];
  const cx=$('#kpis');
  cx.classList.toggle('anim',!!primeira);
  cx.innerHTML=K.map(([k,v,l,c],i)=>
    `<button class="kpi c-${c} ${k===filtro?'on':''}" data-k="${k}" data-v="${v}">
       <div class="v">0</div><div class="l">${l}</div></button>`).join('');
  // conta de onde estava até o novo valor — assim só o número que mudou se move
  cx.querySelectorAll('.kpi').forEach(b=>{
    const k=b.dataset.k, novo=+b.dataset.v;
    const de = primeira||!KPIANT ? 0 : (KPIANT[k]===undefined?novo:KPIANT[k]);
    contarDe(b.querySelector('.v'),de,novo);
  });
  KPIANT={}; K.forEach(([k,v])=>KPIANT[k]=v);

  // contadores da barra lateral
  const fila=K[0][1], ret=K[1][1];
  $('#q-fila').textContent=fila;
  $('#q-base').textContent=DADOS.length;
  const qa=$('#q-agenda'); qa.textContent=ret; qa.classList.toggle('al',ret>0);
  const ab=agAbertos(), atrasado=ab.filter(x=>x.d<HOJE()).length;
  const qs=$('#q-serv'); qs.textContent=ab.length; qs.classList.toggle('al',atrasado>0);
  const qt=$('#q-trocas'), np=trPend().length;
  qt.textContent=np; qt.classList.toggle('al',np>0);
}

/* ----------------------------------------------------------
   8. PAINEL — gráficos
   ---------------------------------------------------------- */
const MESES=['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
const MES3=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];

function svgBarrasMes(){
  const h=new Date(), base=new Date(h.getFullYear(),h.getMonth(),1);
  const cols=[];
  for(let k=-5;k<=6;k++){
    const d=new Date(base.getFullYear(),base.getMonth()+k,1);
    const ch=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
    cols.push({ch,rot:MES3[d.getMonth()]+(d.getMonth()===0||k===-5?'/'+String(d.getFullYear()).slice(2):''),k,v:0});
  }
  const ix={}; cols.forEach((c,i)=>ix[c.ch]=i);
  let fora=0;
  for(const c of DADOS){ const m=(c.venc||'').slice(0,7); if(m in ix) cols[ix[m]].v++; else if(m) fora++; }

  const W=760,H=232,L=8,R=8,T=28,B=30;
  const iw=W-L-R, ih=H-T-B, max=Math.max(1,...cols.map(c=>c.v));
  const passo=iw/cols.length, bw=Math.min(42,passo*.6);
  const cor=k=>k<0?'var(--g-pass)':k===0?'var(--g-atual)':'var(--g-prox)';

  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Vencimentos mês a mês">`;
  s+=`<line class="glin" x1="${L}" y1="${T+ih}" x2="${W-R}" y2="${T+ih}"/>`;
  cols.forEach((c,i)=>{
    const bh=Math.max(c.v?3:0,ih*c.v/max);
    const x=L+i*passo+(passo-bw)/2, y=T+ih-bh;
    s+=`<rect class="barra-c" x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${bh.toFixed(1)}"
         rx="6" style="fill:${cor(c.k)};animation-delay:${(i*.045).toFixed(3)}s;transform-box:fill-box"/>`;
    if(c.v) s+=`<text class="gval" x="${(x+bw/2).toFixed(1)}" y="${(y-7).toFixed(1)}" text-anchor="middle">${c.v}</text>`;
    s+=`<text class="gtxt" x="${(x+bw/2).toFixed(1)}" y="${T+ih+18}" text-anchor="middle"${c.k===0?' style="fill:var(--ink);font-weight:700"':''}>${c.rot}</text>`;
  });
  s+='</svg>';
  return {svg:s,fora};
}

function svgRosca(){
  const ordem=['vencido','vence este mes','vence em breve','em dia','sem data'];
  const rot={'vencido':'Vencido','vence este mes':'Vence este mês','vence em breve':'Vence em breve','em dia':'Em dia','sem data':'Sem data'};
  const c={}; ordem.forEach(k=>c[k]=0);
  for(const x of DADOS) if(x.st in c) c[x.st]++;
  const tot=DADOS.length||1;
  const R=74, C=2*Math.PI*R, cx=100, cy=100;
  let ac=0, s=`<svg viewBox="0 0 200 200" style="max-width:230px;margin:0 auto" role="img" aria-label="Situação da base">`;
  s+=`<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" style="stroke:var(--g-trilha)" stroke-width="26"/>`;
  ordem.forEach((k,i)=>{
    if(!c[k]) return;
    const len=C*c[k]/tot;
    s+=`<circle class="arco" cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke-width="26"
        stroke-dasharray="0 ${C.toFixed(2)}" stroke-dashoffset="${(-ac).toFixed(2)}"
        data-final="${len.toFixed(2)} ${(C-len).toFixed(2)}" style="stroke:${COR[k]};transition-delay:${(i*.11).toFixed(2)}s"
        transform="rotate(-90 ${cx} ${cy})"><title>${rot[k]}: ${c[k]}</title></circle>`;
    ac+=len;
  });
  const venc=c['vencido'];
  s+=`<text x="${cx}" y="${cy-4}" text-anchor="middle" style="font-size:29px;font-weight:750;fill:var(--ink)">${venc}</text>`;
  s+=`<text x="${cx}" y="${cy+16}" text-anchor="middle" class="gtxt">vencidos</text></svg>`;
  const leg='<div class="legenda">'+ordem.filter(k=>c[k]).map(k=>
    `<span><i style="background:${COR[k]}"></i>${rot[k]} · <b>${c[k]}</b></span>`).join('')+'</div>';
  return s+leg;
}

function barrasH(itens,cor){
  const max=Math.max(1,...itens.map(x=>x[1]));
  return '<div class="funil">'+itens.map(([nm,v],i)=>
    `<div class="fase"><span class="nm" title="${esc(nm)}">${esc(nm.length>20?nm.slice(0,20)+'…':nm)}</span>
     <div class="tr"><div class="pr" style="width:${(100*v/max).toFixed(1)}%;background:${cor};animation-delay:${(i*.06).toFixed(2)}s">${v>=max*.18?v:''}</div></div>
     <span class="pc">${v<max*.18?v:''}</span></div>`).join('')+'</div>';
}
function topN(campo,n){
  const m={};
  for(const c of DADOS){const k=_t(c[campo]); if(!k) continue; m[k]=(m[k]||0)+1}
  return Object.entries(m).sort((a,b)=>b[1]-a[1]).slice(0,n);
}

function funilConversa(){
  const fila=DADOS.filter(c=>naFila(c)).length;
  const comSit=DADOS.filter(c=>naFila(c)&&sit(c)).length;
  const conv=DADOS.filter(c=>['aguardando','naoatendeu','pensar','falei'].includes(sit(c))).length;
  const ag=DADOS.filter(c=>sit(c)==='agendado').length;
  const tr=DADOS.filter(c=>sit(c)==='trocou').length;
  const F=[['Na fila',fila,'var(--g-fila)'],['Já abordei',comSit,'var(--g-abord)'],
           ['Em conversa',conv,'var(--g-conv)'],['Agendaram',ag,'var(--g-ag)'],['Trocaram',tr,'var(--g-tr)']];
  const max=Math.max(1,fila);
  return '<div class="funil">'+F.map(([nm,v,cor],i)=>
    `<div class="fase"><span class="nm">${nm}</span>
     <div class="tr"><div class="pr" style="width:${Math.max(4,100*v/max).toFixed(1)}%;background:${cor};animation-delay:${(i*.08).toFixed(2)}s">${v>=max*.14?v:''}</div></div>
     <span class="pc">${v<max*.14?'<b style="color:var(--ink)">'+v+'</b> · ':''}${fila?Math.round(100*v/fila)+'%':'—'}</span></div>`).join('')+'</div>';
}

/* "Meus contatos": o que foi registrado nos últimos dias (histórico LOG[id].h) */
function meusContatos(){
  const hoje=HOJE(), d7=somaDias(hoje,-6), d30=somaDias(hoje,-29);
  const todos=[];
  for(const id in LOG){const r=LOG[id]; if(!r||!r.h) continue;
    const c=DADOS.find(x=>x.id===id);
    r.h.forEach(x=>{if(!x.mig) todos.push({x,c,n:c?nomePartes(c.nome).n:(r.n||id),id})});}
  const conta=(de)=>{const L=todos.filter(o=>o.x.d>=de&&o.x.d<=hoje);
    const k={falei:0,recado:0,msg:0,naoquer:0}; L.forEach(o=>{if(o.x.t in k)k[o.x.t]++});
    return {tot:L.length,k,cli:new Set(L.map(o=>o.id)).size}};
  const s7=conta(d7), s30=conta(d30), sh=conta(hoje);
  const ult=todos.sort((a,b)=>(b.x.d||'').localeCompare(a.x.d||'')).slice(0,12);
  const bloco=(t,o)=>`<div class="mcbl"><div class="mct">${t}</div><div class="mcn">${o.tot}</div>
    <div class="mcd">📞 ${o.k.falei} falei · 📵 ${o.k.recado} não atendeu · 💬 ${o.k.msg} mensagem${o.k.naoquer?' · ✋ '+o.k.naoquer+' não quer':''}</div></div>`;
  return `<div class="mcgrade">${bloco('Hoje',sh)}${bloco('Últimos 7 dias',s7)}${bloco('Últimos 30 dias',s30)}</div>
    ${ult.length?`<div class="mcult"><span class="rotulo" style="margin-top:14px">Últimos registros</span>${ult.map(o=>
      `<div class="mclin" data-ir="${o.id}"><span class="hd">${fmtc(o.x.d)}</span><span>${TCONT[o.x.t]?TCONT[o.x.t].ic:''}</span>
       <b>${esc(o.n)}</b><span class="mco">${o.x.o?'“'+esc(o.x.o)+'”':esc(TCONT[o.x.t]?TCONT[o.x.t].rot:'')}</span></div>`).join('')}</div>`
      :'<div class="hvazio" style="margin-top:10px">Nenhum contato registrado ainda. Use “📞 Registrar” na lista de clientes.</div>'}`;
}

function painel(){
  const bm=svgBarrasMes();
  const prox30=DADOS.filter(c=>c.dias<=0&&c.dias>=-30).length;
  const semZap=DADOS.filter(c=>c.tp!=='Celular').length;
  const comCt=DADOS.filter(c=>ultContato(c)).length;
  const comCtCrm=DADOS.filter(c=>{const r=LOG[c.id];return r&&r.uc}).length;
  const marcados=Object.keys(LOG).length;
  const vencidos=DADOS.filter(c=>c.st==='vencido').length;

  const num=(v,l,cor)=>`<div style="flex:1 1 130px"><div style="font-size:27px;font-weight:750;letter-spacing:-.03em;color:${cor}">${v.toLocaleString('pt-BR')}</div><div style="font-size:12px;color:var(--mut);margin-top:3px;line-height:1.35">${l}</div></div>`;

  $('#paineis').innerHTML=`
  <div class="painel largo" style="animation-delay:.00s">
    <h3>Meus contatos</h3>
    <div class="sub">Quantas pessoas você contatou e como foi — conta o que foi registrado com “📞 Registrar” (no computador ou na lista do celular)</div>
    ${meusContatos()}
  </div>

  <div class="painel largo" style="animation-delay:.00s">
    <h3>Números da base</h3>
    <div class="sub">Um retrato geral, atualizado a cada vez que você carrega a planilha</div>
    <div style="display:flex;gap:26px;flex-wrap:wrap;margin-top:4px">
      ${num(DADOS.length,'clientes na base · '+DADOS.reduce((n,c)=>n+(c.aps?c.aps.length:1),0).toLocaleString('pt-BR')+' aparelhos','var(--ink)')}
      ${num(vencidos,'com o refil vencido','var(--bad)')}
      ${num(prox30,'vencem nos próximos 30 dias','var(--warn)')}
      ${num(comCt,'já contatados alguma vez'+(comCtCrm?' · '+comCtCrm+' aqui no CRM':''),'var(--info)')}
      ${num(marcados,'com situação marcada aqui no CRM','var(--ok)')}
      ${num(semZap,'sem WhatsApp cadastrado','var(--mut)')}
      ${num(DADOS.filter(c=>c.st==='vencido'&&c.dias>365).length,'sumidos há mais de 1 ano','var(--info)')}
      ${num(agAbertos().length,'serviços agendados em aberto','var(--ok)')}
    </div>
  </div>

  <div class="painel largo" style="animation-delay:.00s">
    <h3>Vencimentos mês a mês</h3>
    <div class="sub">Quantos refis vencem em cada mês — 5 meses para trás e 6 para frente${bm.fora?` · ${bm.fora} cadastros vencem fora dessa janela`:''}</div>
    ${bm.svg}
    <div class="legenda">
      <span><i style="background:var(--g-pass)"></i>meses passados</span>
      <span><i style="background:var(--g-atual)"></i>mês atual</span>
      <span><i style="background:var(--g-prox)"></i>próximos meses</span>
    </div>
  </div>

  <div class="painel" style="animation-delay:.06s">
    <h3>Situação da base</h3>
    <div class="sub">Como estão os ${DADOS.length.toLocaleString('pt-BR')} clientes hoje</div>
    ${svgRosca()}
  </div>

  <div class="painel" style="animation-delay:.12s">
    <h3>Funil de conversa</h3>
    <div class="sub">Do total que está na fila, até onde cada um chegou</div>
    ${funilConversa()}
  </div>

  <div class="painel" style="animation-delay:.18s">
    <h3>Bairros com mais clientes</h3>
    <div class="sub">Ajuda a montar a rota do dia</div>
    ${barrasH(topN('loc',8),'var(--g-bairro)')}
  </div>

  <div class="painel" style="animation-delay:.24s">
    <h3>Produtos mais usados</h3>
    <div class="sub">O que mais sai — bom para conferir o estoque</div>
    ${barrasH(topN('prod',8),'var(--g-prod)')}
  </div>`;

  // dispara a animação da rosca
  requestAnimationFrame(()=>{
    $('#paineis').querySelectorAll('.arco').forEach(a=>{
      a.style.transition='stroke-dasharray .85s cubic-bezier(.22,.72,.28,1)';
      a.style.strokeDasharray=a.dataset.final;
    });
  });
}

/* ----------------------------------------------------------
   9. AGENDA
   ---------------------------------------------------------- */
let agAno, agMes, agSel='';
(function(){const h=new Date();agAno=h.getFullYear();agMes=h.getMonth()})();

function retornos(){
  const R=[];
  for(const c of DADOS){
    const v=volta(c);
    if(v) R.push({c,d:v,tipo:'retorno'});
  }
  return R.sort((a,b)=>a.d<b.d?-1:1);
}
function agenda(){
  const R=retornos(), hoje=HOJE();
  const porDia={}; R.forEach(x=>(porDia[x.d]=porDia[x.d]||[]).push(x));
  const atrasados=R.filter(x=>x.d<hoje);
  const deHoje=R.filter(x=>x.d===hoje);

  const prim=new Date(agAno,agMes,1);
  const inicio=new Date(agAno,agMes,1-prim.getDay());
  const SEM=['dom','seg','ter','qua','qui','sex','sáb'];

  let cel='';
  for(let i=0;i<42;i++){
    const d=new Date(inicio.getFullYear(),inicio.getMonth(),inicio.getDate()+i);
    const iso=_iso(d), fora=d.getMonth()!==agMes, lista=porDia[iso]||[];
    const atras=lista.length&&iso<hoje;
    cel+=`<div class="dia${fora?' fora':''}${iso===hoje?' hoje':''}${lista.length?' tem':''}${atras?' atras':''}${agSel===iso?' sel':''}"
      ${lista.length?`data-d="${iso}"`:''}>
      <span class="n">${d.getDate()}</span>
      ${lista.length?`<div class="pts">${lista.slice(0,6).map(()=>'<i></i>').join('')}</div>
        <span class="qtd">${lista.length} ${lista.length===1?'retorno':'retornos'}</span>`:''}
    </div>`;
  }

  const item=(x,cls)=>`<div class="agitem ${cls||''}">
    <span class="d">${fmtc(x.d)}</span>
    <span class="q"><b>${esc(x.c.nome)}</b><span>${esc([x.c.loc,x.c.tel,anot(x.c)].filter(Boolean).join(' · '))||'—'}</span></span>
    ${sit(x.c)?`<span class="mini" style="background:${SIT[sit(x.c)].bg};color:${SIT[sit(x.c)].cor}">${SIT[sit(x.c)].rot}</span>`:''}
    <button class="bt" data-ir="${x.c.id}">Abrir</button></div>`;

  const selLista=agSel?(porDia[agSel]||[]):[];

  $('#agenda').innerHTML=`
  ${atrasados.length?`<div class="agsec"><h4>⏰ Atrasados — ${atrasados.length}</h4>
    ${atrasados.slice(0,40).map(x=>item(x,'atras')).join('')}</div>`:''}
  ${deHoje.length?`<div class="agsec"><h4>📌 Para hoje — ${deHoje.length}</h4>
    ${deHoje.map(x=>item(x)).join('')}</div>`:''}
  ${!R.length?`<div class="vazio"><span class="big">📅</span>Nenhum retorno marcado ainda.<br>
    Abra um cliente na fila e preencha <b>“Voltar a falar em”</b> que ele aparece aqui.</div>`:''}

  <div class="agtopo">
    <h3>${MESES[agMes].charAt(0).toUpperCase()+MESES[agMes].slice(1)} de ${agAno}</h3>
    <div class="agnav">
      <button data-ag="-1" title="Mês anterior">‹</button>
      <button data-ag="0" title="Voltar para hoje">•</button>
      <button data-ag="1" title="Próximo mês">›</button>
    </div>
  </div>
  <div class="cal">
    <div class="calsem">${SEM.map(s=>`<div>${s}</div>`).join('')}</div>
    <div class="caldias">${cel}</div>
  </div>
  ${agSel?`<div class="agsec"><h4>${fmt(agSel)} — ${selLista.length} ${selLista.length===1?'cliente':'clientes'}</h4>
    ${selLista.map(x=>item(x,agSel<hoje?'atras':'')).join('')||'<div class="vazio">Nada marcado nesse dia.</div>'}</div>`:''}`;
}

/* ----------------------------------------------------------
   9b. AGENDA DE SERVIÇOS
   ---------------------------------------------------------- */
let svAno, svMes, svSel='';
(function(){const h=new Date();svAno=h.getFullYear();svMes=h.getMonth()})();

const ICS={'Troca de refil':'💧','Instalação':'🔧','Reinstalação':'📦','Manutenção':'🛠️',
  'Higienização':'🧼','Troca de peça':'⚙️','Orçamento':'💬','Entrega':'🚚','Outro':'📌'};

function itemServ(a){
  const hoje=HOJE(), atras=!a.feito&&a.d<hoje;
  const c=a.cid?DADOS.find(x=>x.id===a.cid):null;
  const sub=[a.h?a.h+'h':'',a.tel||(c&&c.tel)||'',a.end||(c?[c.end,c.loc].filter(Boolean).join(' — '):''),a.obs]
    .filter(Boolean).join(' · ');
  return `<div class="agitem${atras?' atras':''}${a.feito?' feito':''}" data-a="${a.id}">
    <span class="d">${fmtc(a.d)}${a.h?`<small>${esc(a.h)}</small>`:''}</span>
    <span class="q"><b>${esc(a.nome)}</b><span>${esc(sub)||'—'}</span></span>
    <span class="tagserv${a.serv==='Troca de refil'?'':' t-outro'}">${ICS[a.serv]||'📌'} ${esc(a.serv)}</span>
    ${a.cid?`<button class="icb" data-ir="${a.cid}" title="Abrir o cliente na base">👤</button>`:''}
    ${a.serv==='Troca de refil'&&a.cid?`<button class="icb ok" data-trok="${a.cid}" title="Troquei o refil">🔧</button>`:''}
    <button class="icb ok" data-feito="${a.id}" title="${a.feito?'Reabrir':'Marcar como feito'}">${a.feito?'↩':'✓'}</button>
    <button class="icb" data-ed="${a.id}" title="Editar">✏️</button>
    <button class="icb del" data-del="${a.id}" title="Apagar">🗑</button>
  </div>`;
}

function servicos(){
  const hoje=HOJE(), R=agOrd();
  const porDia={}; R.forEach(x=>{if(!x.feito)(porDia[x.d]=porDia[x.d]||[]).push(x)});
  const atrasados=R.filter(x=>!x.feito&&x.d<hoje);
  const deHoje=R.filter(x=>!x.feito&&x.d===hoje);
  const prox=R.filter(x=>!x.feito&&x.d>hoje).slice(0,25);

  /* calendário do mês */
  const prim=new Date(svAno,svMes,1);
  const inicio=new Date(svAno,svMes,1-prim.getDay());
  const SEM=['dom','seg','ter','qua','qui','sex','sáb'];
  let cel='';
  for(let i=0;i<42;i++){
    const d=new Date(inicio.getFullYear(),inicio.getMonth(),inicio.getDate()+i);
    const iso=_iso(d), fora=d.getMonth()!==svMes, lista=porDia[iso]||[];
    const atras=lista.length&&iso<hoje;
    cel+=`<div class="dia livre${fora?' fora':''}${iso===hoje?' hoje':''}${lista.length?' tem':''}${atras?' atras':''}${svSel===iso?' sel':''}"
      data-d="${iso}">
      <span class="n">${d.getDate()}</span>
      ${lista.length?`<div class="pts">${lista.slice(0,6).map(()=>'<i></i>').join('')}</div>
        <span class="qtd">${lista.length} ${lista.length===1?'serviço':'serviços'}</span>`:''}
    </div>`;
  }
  $('#svmes').textContent=MESES[svMes].charAt(0).toUpperCase()+MESES[svMes].slice(1)+' de '+svAno;
  $('#svcal').innerHTML=`<div class="calsem">${SEM.map(s=>`<div>${s}</div>`).join('')}</div>
    <div class="caldias">${cel}</div>`;

  const selLista=svSel?agDoDia(svSel):[];
  const feitos=R.filter(x=>x.feito).slice(-8).reverse();

  $('#svlista').innerHTML=`
  ${svSel?`<div class="agsec"><h4>${fmt(svSel)} — ${selLista.length} ${selLista.length===1?'serviço':'serviços'}
      <button class="bt fantasma" data-novo="${svSel}" style="margin-left:8px">＋ agendar nesse dia</button></h4>
    ${selLista.map(itemServ).join('')||'<div class="vazio">Nada agendado nesse dia. Arraste um cliente para cá ou clique em “agendar nesse dia”.</div>'}</div>`:''}
  ${atrasados.length?`<div class="agsec"><h4>⏰ Atrasados — ${atrasados.length}</h4>
    ${atrasados.map(itemServ).join('')}</div>`:''}
  ${deHoje.length?`<div class="agsec"><h4>📌 Hoje — ${deHoje.length}</h4>
    ${deHoje.map(itemServ).join('')}</div>`:''}
  ${prox.length?`<div class="agsec"><h4>Próximos dias — ${prox.length}</h4>
    ${prox.map(itemServ).join('')}</div>`:''}
  ${feitos.length?`<div class="agsec"><h4>Já feitos (últimos)</h4>${feitos.map(itemServ).join('')}</div>`:''}
  ${!R.length?`<div class="vazio"><span class="big">🛠️</span>Nenhum serviço agendado ainda.<br>
    Clique em <b>Novo agendamento</b>, ou arraste um cliente da lista ao lado para um dia do calendário.</div>`:''}`;

  /* lista lateral de clientes para arrastar */
  const q=norm(($('#qsv').value||'').trim());
  let cand=DADOS.filter(c=>naFila(c)&&!['naoquer','trocou'].includes(sit(c))&&!agDoCliente(c.id).length);
  if(q) cand=DADOS.filter(c=>norm(c.nome+' '+c.loc+' '+c.end).includes(q)||dig(c.tel).includes(q));
  cand=cand.sort((a,b)=>b.dias-a.dias).slice(0,80);
  $('#svfila').innerHTML=cand.length?cand.map(c=>`
    <div class="svcli" draggable="true" data-drag="${c.id}" title="Arraste para um dia do calendário">
      <b>${esc(c.nome)}</b>
      <span>${esc([c.loc,c.prod].filter(Boolean).join(' · '))||'—'}</span>
      <span class="${c.st==='vencido'?'pz':''}">${c.st==='vencido'?'venceu há '+c.dias+' dias':'vence '+fmt(c.venc)}</span>
    </div>`).join('')
    :'<div class="vazio" style="padding:22px 12px;font-size:12.5px">Ninguém encontrado.</div>';
}

/* --- arrastar cliente até um dia do calendário --- */
let ARRASTA='';
document.addEventListener('dragstart',e=>{
  const el=e.target.closest('[data-drag]'); if(!el) return;
  ARRASTA=el.dataset.drag;
  e.dataTransfer.effectAllowed='copy';
  try{e.dataTransfer.setData('text/plain',ARRASTA)}catch(x){}
  const linha=el.closest('.linha'); if(linha) linha.classList.add('arrastando');
});
document.addEventListener('dragend',()=>{
  ARRASTA='';
  document.querySelectorAll('.arrastando').forEach(x=>x.classList.remove('arrastando'));
  document.querySelectorAll('.dia.alvo').forEach(x=>x.classList.remove('alvo'));
  document.querySelectorAll('.nav button.solto').forEach(x=>x.classList.remove('solto'));
});
document.addEventListener('dragover',e=>{
  if(!ARRASTA) return;
  const dia=e.target.closest('#svcal .dia[data-d]');
  const nav=e.target.closest('#nav button[data-v="servicos"]');
  if(dia||nav){e.preventDefault();e.dataTransfer.dropEffect='copy'}
  document.querySelectorAll('.dia.alvo').forEach(x=>{if(x!==dia)x.classList.remove('alvo')});
  if(dia) dia.classList.add('alvo');
  if(nav) nav.classList.add('solto'); else document.querySelectorAll('.nav button.solto').forEach(x=>x.classList.remove('solto'));
});
document.addEventListener('drop',e=>{
  const id=ARRASTA||(e.dataTransfer?e.dataTransfer.getData('text/plain'):'');
  if(!id) return;
  const dia=e.target.closest('#svcal .dia[data-d]');
  const nav=e.target.closest('#nav button[data-v="servicos"]');
  if(dia){e.preventDefault();dia.classList.remove('alvo');agendarCliente(id,dia.dataset.d)}
  else if(nav){e.preventDefault();irVista('servicos');agendarCliente(id,HOJE())}
});

/* eventos da agenda de serviços */
$('#btnovo').addEventListener('click',()=>modalAgendar({d:svSel||HOJE()}));
$('#qsv').addEventListener('input',()=>servicos());
$('#v-servicos').addEventListener('click',e=>{
  const nav=e.target.closest('[data-sv]');
  if(nav){const p=+nav.dataset.sv;
    if(p===0){const h=new Date();svAno=h.getFullYear();svMes=h.getMonth();svSel=''}
    else{svMes+=p;if(svMes>11){svMes=0;svAno++}if(svMes<0){svMes=11;svAno--}}
    servicos();return}
  const nv=e.target.closest('[data-novo]');
  if(nv){modalAgendar({d:nv.dataset.novo});return}
  const dia=e.target.closest('#svcal .dia[data-d]');
  if(dia){svSel=(svSel===dia.dataset.d?'':dia.dataset.d);servicos();return}
  const ed=e.target.closest('[data-ed]');
  if(ed){const a=AGEND.find(x=>x.id===ed.dataset.ed);if(a)modalAgendar(a);return}
  const del=e.target.closest('[data-del]');
  if(del){if(confirm('Apagar esse agendamento?')){apagarAgendamento(del.dataset.del);toast('Agendamento apagado.')}return}
  const fz=e.target.closest('[data-feito]');
  if(fz){const a=AGEND.find(x=>x.id===fz.dataset.feito);
    if(a){a.feito=!a.feito;salvarAg();atualizarTudo();toast(a.feito?'Serviço marcado como feito.':'Serviço reaberto.','ok')}return}
  const tk=e.target.closest('[data-trok]');
  if(tk){modalTroca(tk.dataset.trok);return}
  const ir=e.target.closest('[data-ir]');
  if(ir){abrirCliente(ir.dataset.ir);return}
  const cli=e.target.closest('.svcli[data-drag]');
  if(cli){agendarCliente(cli.dataset.drag,svSel||HOJE());return}
});

/* ----------------------------------------------------------
   9c. TROCAS FEITAS — atualizar a planilha
   ---------------------------------------------------------- */
function trocasView(){
  const P=trPend().sort((a,b)=>a.d<b.d?-1:1);
  const F=TROCAS.filter(x=>x.ok).sort((a,b)=>a.d<b.d?1:-1).slice(0,30);
  const lin=(t,cls)=>`<div class="trlin ${cls}" data-t="${t.id}">
    <span class="d" style="width:64px;flex:none;font-size:12px;font-weight:700;color:var(--mut)">${fmtc(t.d)}</span>
    <span class="q"><b>${esc(t.nome)}</b>
      <span>${esc([t.aps&&t.aps.length?t.aps.length+(t.aps.length===1?' aparelho':' aparelhos')+' — '+t.aps.map(a=>a.obs||a.prod).join(', '):t.prod,
        t.aba+(t.aps&&t.aps.length>1?' linhas '+t.aps.map(a=>a.lin).join(', '):' linha '+t.lin),t.tel].filter(Boolean).join(' · '))}</span></span>
    ${t.ok?'<span class="mini" style="background:var(--ok-b);color:var(--ok)">na planilha</span>'
          :`<button class="icb ok" data-tok="${t.id}" title="Já atualizei essa na planilha">✓</button>`}
    <button class="icb del" data-tdel="${t.id}" title="Tirar da lista">🗑</button>
  </div>`;

  $('#trocas').innerHTML=`
  <div class="trbar">
    <span class="cont"><b>${P.length}</b> ${P.length===1?'troca esperando':'trocas esperando'} para entrar na Planilha de Clientes</span>
    <button class="bt" id="trlimpa">Limpar as já atualizadas</button>
    <button class="bt pri" id="trprompt">🤖 Gerar pedido para o Claude</button>
  </div>
  ${P.length?`<div class="agsec"><h4>Ainda não estão na planilha</h4>${P.map(t=>lin(t,'pend')).join('')}</div>`
    :`<div class="vazio"><span class="big">🔧</span>Nenhuma troca esperando.<br>
      Marque <b>“Troquei o refil”</b> no cliente que você atendeu e ele aparece aqui.</div>`}
  ${F.length?`<div class="agsec"><h4>Já atualizadas</h4>${F.map(t=>lin(t,'pron')).join('')}</div>`:''}`;
}
$('#trocas').addEventListener('click',e=>{
  if(e.target.closest('#trprompt')){modalPrompt();return}
  if(e.target.closest('#trlimpa')){
    const n=TROCAS.filter(x=>x.ok).length;
    if(!n){toast('Não há trocas já atualizadas para limpar.','err');return}
    if(confirm('Tirar da lista as '+n+' trocas já atualizadas?')){
      TROCAS=TROCAS.filter(x=>!x.ok);salvarTr();trocasView();toast('Lista limpa.','ok')}
    return}
  const ok=e.target.closest('[data-tok]');
  if(ok){const t=TROCAS.find(x=>x.id===ok.dataset.tok);
    if(t){t.ok=true;t.okd=HOJE();salvarTr();atualizarTudo();toast('Marcada como já atualizada.','ok')}return}
  const dl=e.target.closest('[data-tdel]');
  if(dl){const i=TROCAS.findIndex(x=>x.id===dl.dataset.tdel);
    if(i>=0){TROCAS.splice(i,1);salvarTr();atualizarTudo();toast('Tirada da lista.')}return}
});

/* ----------------------------------------------------------
   10. LISTA DO DIA PARA IMPRIMIR
   ---------------------------------------------------------- */
function imprimir(){
  const L=filtrados();
  if(!L.length){toast('Não há clientes na lista para imprimir.','err');return}
  // evita mandar 30 páginas pra impressora sem querer
  if(L.length>120){
    const pag=Math.ceil(L.length/34);
    if(!confirm(`São ${L.length} clientes — cerca de ${pag} páginas.\n\n`
      +`Se você quer só a lista do dia, feche isto, filtre a tela primeiro `
      +`(por bairro, por mês de vencimento ou pelo cartão "Para contatar agora") `
      +`e clique de novo.\n\nImprimir mesmo assim?`)) return;
  }
  const q=$('#q').value.trim();
  const desc=[];
  if(q) desc.push('busca “'+q+'”');
  const rotFiltro={fila:'Fila de hoje',retorno:'Retornar hoje',ligar:'Hora de ligar',b30:'Vencem em 30 dias',
    aberto:'Conversas em aberto',agendado:'Agendados',novos:'Na fila, nunca contatados',
    semretorno:'Contatados sem retorno',vencido:'Vencidos',breve:'Vencem em 60 dias',
    sumido:'Sumidos há +1 ano',semwa:'Sem WhatsApp',todos:'Toda a base'};
  if(!q) desc.push(rotFiltro[filtro]||filtro);
  if(tipo!=='todos') desc.push(tipo);
  if($('#loc').value) desc.push('bairro '+$('#loc').value);
  if($('#prod').value) desc.push($('#prod').value);
  if($('#mes').value){const m=$('#mes').value;desc.push('vence em '+MESES[+m.slice(5,7)-1]+' de '+m.slice(0,4))}

  const linhas=L.map(c=>`<tr>
    <td><span class="bx"></span></td>
    <td><b>${esc(c.nome)}</b>${c.fc?'<br><span style="font-size:8pt;color:#555">falar com '+esc(c.fc)+'</span>':''}</td>
    <td>${esc(c.tel||'—')}${c.tel2?'<br>'+esc(c.tel2):''}</td>
    <td>${esc(c.end||'—')}${c.loc?'<br><span style="font-size:8pt;color:#555">'+esc(c.loc)+'</span>':''}</td>
    <td>${esc(c.prod||'—')}</td>
    <td>${fmt(c.venc)}${c.st==='vencido'?'<br><span style="font-size:8pt">venceu há '+c.dias+'d</span>':''}</td>
    <td style="width:96px"></td></tr>`).join('');

  $('#folha').innerHTML=`
   <h1>Fonte Clara — lista do dia</h1>
   <div class="cabf">${new Date().toLocaleDateString('pt-BR',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}
     · ${L.length} ${L.length===1?'cliente':'clientes'}${desc.length?' · '+esc(desc.join(' · ')):''}</div>
   <table>
     <thead><tr><th>✓</th><th>Cliente</th><th>Telefone</th><th>Endereço</th><th>Produto</th><th>Vencimento</th><th>Anotação</th></tr></thead>
     <tbody>${linhas}</tbody>
   </table>
   <div class="pe">Gerado pelo CRM Fonte Clara em ${new Date().toLocaleString('pt-BR')}. Base carregada em ${fmt(BASEDATA)}.</div>`;
  window.print();
}

/* ----------------------------------------------------------
   10b. EXPORTAR LISTAS — texto para copiar, PDF ou .txt
   ---------------------------------------------------------- */
const somaDias=(iso,n)=>{const d=new Date(iso+'T00:00:00');d.setDate(d.getDate()+n);return _iso(d)};
const iniMes=iso=>{const d=new Date(iso+'T00:00:00');return _iso(new Date(d.getFullYear(),d.getMonth(),1))};
const fimMes=iso=>{const d=new Date(iso+'T00:00:00');return _iso(new Date(d.getFullYear(),d.getMonth()+1,0))};
const SEMANA=['domingo','segunda','terça','quarta','quinta','sexta','sábado'];
const diaSem=iso=>SEMANA[new Date(iso+'T00:00:00').getDay()];
const periodoRot=(de,ate)=>de===ate?fmt(de):fmt(de)+' a '+fmt(ate);

/* dados do agendamento completados com os do cliente da base */
function agCheio(a){
  const c=a.cid?DADOS.find(x=>x.id===a.cid):null;
  return {tel:a.tel||(c&&c.tel)||'', end:a.end||(c?[c.end,c.loc].filter(Boolean).join(' — '):''),
          prod:c?c.prod:'', c};
}

/* quais agendamentos entram */
const expAg=(de,ate,incFeitos)=>agOrd().filter(a=>a.d>=de&&a.d<=ate&&(incFeitos||!a.feito));

/* filtros da tela que continuam valendo na exportação de vencimentos
   (o período digitado substitui os filtros de data da tela) */
function passaExp(c){
  const q=norm($('#q').value.trim());
  if(q){
    const alvo=norm(c.nome+' '+c.end+' '+c.loc+' '+c.fc+' '+c.prod+' '+(c.obs||'')+' '+anot(c))+' '+dig(c.tel)+' '+dig(c.tel2);
    if(!q.split(/\s+/).every(t=>alvo.includes(t))) return false;
  }
  if(tipo!=='todos'&&c.tipo!==tipo) return false;
  if($('#loc').value&&c.loc!==$('#loc').value) return false;
  if($('#prod').value&&c.prod!==$('#prod').value&&!(c.aps&&c.aps.some(x=>x.prod===$('#prod').value))) return false;
  return true;
}
function expFiltrosAtivos(){
  const d=[], q=$('#q').value.trim();
  if(q) d.push('busca “'+q+'”');
  if(tipo!=='todos') d.push(tipo);
  if($('#loc').value) d.push('bairro '+$('#loc').value);
  if($('#prod').value) d.push($('#prod').value);
  return d;
}
function expVenc(de,ate,semRecusa){
  return DADOS.filter(c=>c.venc&&c.venc>=de&&c.venc<=ate&&passaExp(c)
      &&!(semRecusa&&sit(c)==='naoquer'))
    .sort((a,b)=>a.venc<b.venc?-1:a.venc>b.venc?1:a.nome.localeCompare(b.nome,'pt-BR'));
}

/* tabelinha de texto em colunas alinhadas (modo completo) */
function mkTabela(cab,linhas,max){
  const larg=cab.map((h,i)=>Math.max(h.length,
    Math.min(max[i],Math.max(0,...linhas.map(l=>String(l[i]==null?'':l[i]).length)))));
  const cx=(t,n)=>{t=String(t==null?'':t);return t.length>n?t.slice(0,n-1)+'…':t.padEnd(n)};
  const linha=l=>l.map((v,i)=>cx(v,larg[i])).join(' | ').replace(/\s+$/,'');
  return [linha(cab),larg.map(n=>'-'.repeat(n)).join('-+-'),...linhas.map(linha)].join('\n');
}
const rodapeTxt=()=>'\nGerado pelo CRM Fonte Clara em '+new Date().toLocaleString('pt-BR')+'.';

/* ---------- textos: agendamentos ---------- */
function txtAgZap(L,de,ate){
  let s='📅 AGENDAMENTOS — '+periodoRot(de,ate)+'\n'+L.length+(L.length===1?' serviço':' serviços')+'\n';
  const porDia={}; L.forEach(a=>(porDia[a.d]=porDia[a.d]||[]).push(a));
  Object.keys(porDia).sort().forEach(d=>{
    s+='\n──── '+fmt(d)+' · '+diaSem(d)+' ────\n';
    porDia[d].forEach((a,i)=>{
      const x=agCheio(a);
      s+='\n'+(i+1)+'. '+a.nome+(a.h?'  ⏰ '+a.h:'')+'\n';
      s+='   '+a.serv+(a.feito?' ✓ feito':'')+'\n';
      if(x.end) s+='   📍 '+x.end+'\n';
      if(x.tel) s+='   📞 '+x.tel+'\n';
      if(a.obs) s+='   📝 '+a.obs+'\n';
    });
  });
  return s.trim();
}
function txtAgFull(L,de,ate){
  const rows=L.map(a=>{const x=agCheio(a);
    return [fmt(a.d),a.h||'—',a.nome,a.serv,x.end||'—',x.tel||'—',
            (a.feito?'feito':'aberto')+(a.obs?' · '+a.obs:'')]});
  return 'AGENDAMENTOS — '+periodoRot(de,ate)+'   ('+L.length+')\n\n'
    +mkTabela(['Data','Hora','Cliente','Serviço','Endereço','Telefone','Situação'],rows,[10,5,30,16,36,17,32])
    +'\n'+rodapeTxt();
}

/* ---------- textos: clientes a vencer ---------- */
function txtVencZap(L,de,ate){
  let s='💧 CLIENTES A VENCER — '+periodoRot(de,ate)+'\n'+L.length+(L.length===1?' cliente':' clientes')+'\n';
  const porMes={}; L.forEach(c=>(porMes[c.venc.slice(0,7)]=porMes[c.venc.slice(0,7)]||[]).push(c));
  Object.keys(porMes).sort().forEach(m=>{
    s+='\n──── '+MESES[+m.slice(5,7)-1].toUpperCase()+' DE '+m.slice(0,4)+' ────\n';
    porMes[m].forEach((c,i)=>{
      s+='\n'+(i+1)+'. '+c.nome+'\n';
      s+='   vence '+fmt(c.venc)+(c.st==='vencido'?' · venceu há '+c.dias+(c.dias===1?' dia':' dias'):'')+'\n';
      if(c.prod) s+='   '+c.prod+'\n';
      const e=[c.end,c.loc].filter(Boolean).join(' — ');
      if(e) s+='   📍 '+e+'\n';
      if(c.tel) s+='   📞 '+c.tel+(c.tel2?' / '+c.tel2:'')+'\n';
    });
  });
  return s.trim();
}
function txtVencFull(L,de,ate){
  const rows=L.map(c=>[fmt(c.venc),
    c.st==='vencido'?'venceu há '+c.dias+'d':'em dia',
    c.nome,c.tipo,c.prod||'—',[c.end,c.loc].filter(Boolean).join(' — ')||'—',
    c.tel||'—',(SIT[sit(c)]||{}).rot||'—',c.aba+' '+(c.aps&&c.aps.length>1?'L'+c.aps.map(a=>a.lin).join('/'):'L'+c.lin)]);
  return 'CLIENTES A VENCER — '+periodoRot(de,ate)+'   ('+L.length+')\n'
    +(expFiltrosAtivos().length?'Filtros: '+expFiltrosAtivos().join(' · ')+'\n':'')+'\n'
    +mkTabela(['Vencimento','Prazo','Cliente','Tipo','Produto','Endereço','Telefone','Situação','Planilha'],
              rows,[12,14,30,13,26,36,17,20,22])
    +'\n'+rodapeTxt();
}

/* ---------- folhas para PDF / impressão ---------- */
function folhaImprimir(titulo,sub,cab,linhas){
  $('#folha').innerHTML=`
   <h1>Fonte Clara — ${esc(titulo)}</h1>
   <div class="cabf">${esc(sub)}</div>
   <table>
     <thead><tr>${cab.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
     <tbody>${linhas}</tbody>
   </table>
   <div class="pe">Gerado pelo CRM Fonte Clara em ${new Date().toLocaleString('pt-BR')}. Base carregada em ${fmt(BASEDATA)}.</div>`;
  window.print();
}
function folhaAg(L,de,ate){
  let atual='', linhas='';
  L.forEach(a=>{
    if(a.d!==atual){atual=a.d;
      linhas+=`<tr class="grupo"><td colspan="7"><b>${fmt(a.d)} · ${diaSem(a.d)}</b></td></tr>`}
    const x=agCheio(a);
    linhas+=`<tr>
      <td><span class="bx"></span></td>
      <td>${esc(a.h||'—')}</td>
      <td><b>${esc(a.nome)}</b>${a.feito?'<br><span style="font-size:8pt;color:#555">já feito</span>':''}</td>
      <td>${esc(a.serv)}</td>
      <td>${esc(x.end||'—')}</td>
      <td>${esc(x.tel||'—')}</td>
      <td>${esc(a.obs||'')}</td></tr>`;
  });
  folhaImprimir('agendamentos',
    periodoRot(de,ate)+' · '+L.length+(L.length===1?' serviço':' serviços'),
    ['✓','Hora','Cliente','Serviço','Endereço','Telefone','Observação'],linhas);
}
function folhaVenc(L,de,ate){
  const f=expFiltrosAtivos();
  const linhas=L.map(c=>`<tr>
    <td><span class="bx"></span></td>
    <td><b>${esc(c.nome)}</b>${c.fc?'<br><span style="font-size:8pt;color:#555">falar com '+esc(c.fc)+'</span>':''}</td>
    <td>${esc(c.tel||'—')}${c.tel2?'<br>'+esc(c.tel2):''}</td>
    <td>${esc([c.end,c.loc].filter(Boolean).join(' — ')||'—')}</td>
    <td>${esc(c.prod||'—')}</td>
    <td>${fmt(c.venc)}${c.st==='vencido'?'<br><span style="font-size:8pt">venceu há '+c.dias+'d</span>':''}</td>
    <td style="width:96px"></td></tr>`).join('');
  folhaImprimir('clientes a vencer',
    periodoRot(de,ate)+' · '+L.length+(L.length===1?' cliente':' clientes')+(f.length?' · '+f.join(' · '):''),
    ['✓','Cliente','Telefone','Endereço','Produto','Vencimento','Anotação'],linhas);
}

/* ---------- copiar com plano B ---------- */
/* telefone clicável no cartão aberto: um clique copia só os números */
function btTel(t){
  if(!t) return '';
  return `<button class="cptel" data-cptel="${esc(t)}" title="Clique para copiar o número">${esc(t)} <span>📋</span></button>`;
}

function copiarTexto(txt,aviso){
  aviso=aviso||'Texto copiado.';
  const plandeB=()=>{
    try{
      const t=document.createElement('textarea');
      t.value=txt; t.style.position='fixed'; t.style.opacity='0';
      document.body.appendChild(t); t.select();
      const ok=document.execCommand('copy'); document.body.removeChild(t);
      toast(ok?aviso:'Não consegui copiar. Selecione o texto e use Ctrl+C.',ok?'ok':'err');
    }catch(e){toast('Não consegui copiar. Selecione o texto e use Ctrl+C.','err')}
  };
  if(navigator.clipboard&&navigator.clipboard.writeText)
    navigator.clipboard.writeText(txt).then(()=>toast(aviso,'ok')).catch(plandeB);
  else plandeB();
}

/* ---------- a janela de exportar ---------- */
function modalExportar(kind){
  const ag=kind==='ag', hoje=HOJE();
  let de = ag ? (svSel||hoje) : hoje;
  let ate= ag ? (svSel||hoje) : somaDias(hoje,30);
  let modo='zap', L=[], txt='';

  const atalhos = ag
    ? [['Hoje',       ()=>[hoje,hoje]],
       ['Amanhã',     ()=>[somaDias(hoje,1),somaDias(hoje,1)]],
       ['Próximos 7 dias',()=>[hoje,somaDias(hoje,6)]],
       ['Este mês',   ()=>[iniMes(hoje),fimMes(hoje)]]]
    : [['Este mês',   ()=>[iniMes(hoje),fimMes(hoje)]],
       ['Próximo mês',()=>{const p=somaDias(fimMes(hoje),1);return[iniMes(p),fimMes(p)]}],
       ['Próximos 30 dias',()=>[hoje,somaDias(hoje,30)]],
       ['Todos os vencidos',()=>['2015-01-01',hoje]]];

  abrirModal(`
    <h3>${ag?'Exportar agendamentos':'Exportar clientes a vencer'}</h3>
    <div class="msub">${ag
      ? 'Os serviços marcados no período — para levar na rua ou mandar pro técnico.'
      : 'Quem tem vencimento dentro do período. Respeita os filtros que estiverem na tela.'}</div>

    <div class="expqp" id="exqp">${atalhos.map((a,i)=>
      `<button type="button" class="chip" data-qp="${i}">${esc(a[0])}</button>`).join('')}</div>

    <div class="mlin">
      <div class="mcampo"><span class="rotulo">De</span><input type="date" id="exde" value="${de}"></div>
      <div class="mcampo"><span class="rotulo">Até</span><input type="date" id="exate" value="${ate}"></div>
    </div>

    <label class="excheck">${ag
      ? '<input type="checkbox" id="exop"> incluir os serviços já marcados como feitos'
      : '<input type="checkbox" id="exop" checked> esconder quem disse que não quer mais'}</label>

    <div class="segm" id="exmodo">
      <button type="button" data-m="zap" class="on">Enxuto (WhatsApp)</button>
      <button type="button" data-m="full">Completo</button>
    </div>

    <div class="exconta" id="exconta"></div>
    <div class="prompt" id="extxt"></div>
    <div class="exnota" id="exnota"></div>

    <div class="mrod">
      <button class="bt" data-fechar>Fechar</button>
      <button class="bt" id="exdl">💾 Baixar .txt</button>
      <button class="bt" id="expdf">🖨️ PDF</button>
      <button class="bt pri" id="excopy">📋 Copiar texto</button>
    </div>`, m=>{

    m.classList.add('largo');

    function pintar(){
      de=m.querySelector('#exde').value||hoje;
      ate=m.querySelector('#exate').value||de;
      if(ate<de){const t=de;de=ate;ate=t;
        m.querySelector('#exde').value=de;m.querySelector('#exate').value=ate}
      const op=m.querySelector('#exop').checked;
      L = ag ? expAg(de,ate,op) : expVenc(de,ate,op);
      txt = L.length
        ? (ag ? (modo==='zap'?txtAgZap(L,de,ate):txtAgFull(L,de,ate))
              : (modo==='zap'?txtVencZap(L,de,ate):txtVencFull(L,de,ate)))
        : '';
      m.querySelector('#exconta').textContent = L.length
        ? L.length+(ag?(L.length===1?' serviço':' serviços'):(L.length===1?' cliente':' clientes'))+' em '+periodoRot(de,ate)
        : 'Nada encontrado nesse período.';
      m.querySelector('#extxt').textContent = txt || '—';
      const f=expFiltrosAtivos();
      m.querySelector('#exnota').textContent = ag
        ? 'O PDF sai pela janela de impressão do Chrome — escolha “Salvar como PDF” no destino.'
        : (f.length?'Filtros da tela aplicados: '+f.join(' · ')+'. ':'')
          +'O PDF sai pela janela de impressão do Chrome — escolha “Salvar como PDF” no destino.';
    }

    m.querySelector('#exqp').addEventListener('click',e=>{
      const b=e.target.closest('[data-qp]'); if(!b) return;
      const [a,z]=atalhos[+b.dataset.qp][1]();
      m.querySelector('#exde').value=a; m.querySelector('#exate').value=z;
      m.querySelectorAll('#exqp .chip').forEach(x=>x.classList.remove('on'));
      b.classList.add('on'); pintar();
    });
    m.querySelector('#exmodo').addEventListener('click',e=>{
      const b=e.target.closest('[data-m]'); if(!b) return;
      modo=b.dataset.m;
      m.querySelectorAll('#exmodo button').forEach(x=>x.classList.toggle('on',x===b));
      pintar();
    });
    m.querySelector('#exde').addEventListener('change',pintar);
    m.querySelector('#exate').addEventListener('change',pintar);
    m.querySelector('#exop').addEventListener('change',pintar);

    m.querySelector('#excopy').addEventListener('click',()=>{
      if(!txt){toast('Não há nada para copiar nesse período.','err');return}
      copiarTexto(txt);
    });
    m.querySelector('#exdl').addEventListener('click',()=>{
      if(!txt){toast('Não há nada para baixar nesse período.','err');return}
      baixar(new Blob(['﻿'+txt],{type:'text/plain;charset=utf-8'}),
        (ag?'agendamentos_':'a_vencer_')+de+(de===ate?'':'_a_'+ate)+'.txt');
      toast('Arquivo .txt salvo na pasta Downloads.','ok');
    });
    m.querySelector('#expdf').addEventListener('click',()=>{
      if(!L.length){toast('Não há nada para gerar nesse período.','err');return}
      if(L.length>150&&!confirm('São '+L.length+' registros — vai dar bastante página.\n\nGerar mesmo assim?')) return;
      if(ag) folhaAg(L,de,ate); else folhaVenc(L,de,ate);
    });

    pintar();
  });
}

/* ----------------------------------------------------------
   10c. MUTIRÃO — falar com vários clientes em sequência
   ----------------------------------------------------------
   O WhatsApp não deixa disparar para várias pessoas de uma vez por link:
   cada conversa precisa ser aberta individualmente (é o que impede spam).
   Então o mutirão faz o caminho curto possível: monta a fila dos clientes
   que estão filtrados na tela, e a cada clique abre a conversa do próximo
   já com a mensagem escrita, marcando o contato sozinho.
   Quem já recebeu mensagem há menos de N dias sai da fila (crm_demo_mut_dias,
   padrão 7; "não pular" = 0 traz todo mundo de volta).
   ---------------------------------------------------------- */
let MUT={base:[],fila:[],i:0,modelo:'aviso',num:'1'};

/* pular quem já recebeu mensagem há menos de N dias (0 = não pular) */
const MUT_DIAS='crm_demo_mut_dias';
const MUT_OPC=[[0,'não pular'],[3,'3 dias'],[7,'7 dias'],[15,'15 dias'],[30,'30 dias']];
function mutDias(){
  let v=null; try{v=localStorage.getItem(MUT_DIAS)}catch(e){}
  const n=v===null?7:parseInt(v,10);
  return isNaN(n)?7:n;
}
function filaMutirao(L,dias){
  if(!dias) return L.slice();
  return L.filter(c=>{
    const uc=ultContato(c); if(!uc) return true;
    const d=diasDe(uc.d);
    return d===null||d>=dias;
  });
}

function modalMutirao(){
  const L=filtrados().filter(temZap);
  if(!L.length){toast('Nenhum cliente com WhatsApp na lista filtrada.','err');return}
  MUT={base:L,fila:filaMutirao(L,mutDias()),i:0,modelo:'aviso',num:'1'};

  abrirModal(`
    <h3>Mutirão de mensagens</h3>
    <div class="msub" id="mutsub"></div>

    <div class="mcampo">
      <span class="rotulo">Pular quem já recebeu mensagem nos últimos</span>
      <div class="mservs" id="mutdias">${MUT_OPC.map(([d,r])=>
        `<button type="button" class="chip${d===mutDias()?' on':''}" data-md="${d}">${r}</button>`).join('')}</div>
    </div>

    <div class="mcampo">
      <span class="rotulo">Mensagem</span>
      <div class="mservs" id="mutmod">${MODELOS.map(m=>
        `<button type="button" class="chip${m.k==='aviso'?' on':''}" data-mm="${m.k}">${esc(m.r)}</button>`).join('')}</div>
    </div>

    <div class="mutcx" id="mutcx"></div>
    <div class="prompt" id="muttxt"></div>

    <div class="mrod">
      <button class="bt" data-fechar>Parar</button>
      <button class="bt" id="mutpula">Pular este</button>
      <button class="bt" id="mutnums" title="Para montar uma lista de transmissão no seu celular">📋 Copiar os números</button>
      <button class="bt pri" id="mutabre">Abrir conversa e ir para o próximo</button>
    </div>`, m=>{

    m.classList.add('largo');

    function atual(){return MUT.fila[MUT.i]}
    function subtitulo(){
      const tot=MUT.base.length, f=MUT.fila.length, pul=tot-f, d=mutDias();
      let t=`<b>${f}</b> ${f===1?'cliente na fila':'clientes na fila'}`;
      if(pul) t+=` · ${pul} ${pul===1?'pulado':'pulados'} por já ter recebido mensagem nos últimos ${d} dias`;
      if(f) t+=' — o WhatsApp exige abrir uma conversa por vez; aqui vai uma atrás da outra, sem você procurar ninguém.';
      m.querySelector('#mutsub').innerHTML=t;
    }
    function pintar(){
      subtitulo();
      const c=atual();
      if(!c){
        const tot=MUT.base.length, f=MUT.fila.length;
        m.querySelector('#mutcx').innerHTML= f
          ? `<div class="mutfim">✅ Fim da fila — ${f} ${f===1?'cliente percorrido':'clientes percorridos'}.</div>`
          : `<div class="mutfim">✅ Todos os ${tot} já receberam mensagem nos últimos ${mutDias()} dias.<br>
             Para falar com eles assim mesmo, escolha <b>não pular</b> ali em cima.</div>`;
        m.querySelector('#muttxt').textContent='';
        m.querySelector('#mutabre').disabled=true;
        m.querySelector('#mutpula').disabled=true;
        return;
      }
      m.querySelector('#mutabre').disabled=false;
      m.querySelector('#mutpula').disabled=false;
      const mod=MODELOS.find(x=>x.k===MUT.modelo)||MODELOS[0];
      const txt=mod.f(c), N=zapNums(c), uc=ultContato(c);
      m.querySelector('#mutcx').innerHTML=`
        <div class="mutpos">${MUT.i+1} de ${MUT.fila.length}</div>
        <div class="mutnome">${esc(c.nome)}</div>
        <div class="mutsub">${esc([c.loc,c.prod].filter(Boolean).join(' · '))}
          · ${c.st==='vencido'?'venceu há '+c.dias+' dias':'vence '+fmt(c.venc)}</div>
        ${uc?`<div class="mutav">${rotUltContato(uc)}</div>`:''}
        ${N.length>1?`<div class="mutnum">Número: ${N.map(x=>
          `<button type="button" class="chip${x.rot===MUT.num?' on':''}" data-mn="${x.rot}">${esc(x.n)}</button>`).join('')}</div>`
          :`<div class="mutnum">Número: <b>${esc(N[0].n)}</b></div>`}`;
      m.querySelector('#muttxt').textContent=txt;
    }

    m.querySelector('#mutdias').addEventListener('click',e=>{
      const b=e.target.closest('[data-md]'); if(!b) return;
      try{localStorage.setItem(MUT_DIAS,b.dataset.md)}catch(e){}
      m.querySelectorAll('#mutdias .chip').forEach(x=>x.classList.toggle('on',x===b));
      MUT.fila=filaMutirao(MUT.base,mutDias()); MUT.i=0; MUT.num='1';
      pintar();
    });
    m.querySelector('#mutmod').addEventListener('click',e=>{
      const b=e.target.closest('[data-mm]'); if(!b) return;
      MUT.modelo=b.dataset.mm;
      m.querySelectorAll('#mutmod .chip').forEach(x=>x.classList.toggle('on',x===b));
      pintar();
    });
    m.querySelector('#mutcx').addEventListener('click',e=>{
      const b=e.target.closest('[data-mn]'); if(!b) return;
      MUT.num=b.dataset.mn; pintar();
    });
    m.querySelector('#mutpula').addEventListener('click',()=>{MUT.i++;MUT.num='1';pintar()});
    m.querySelector('#mutabre').addEventListener('click',()=>{
      const c=atual(); if(!c) return;
      const mod=MODELOS.find(x=>x.k===MUT.modelo)||MODELOS[0];
      /* mutirão = mensagem: entra no histórico como "Mandei mensagem"
         (mantém a data de retorno, se houver uma no futuro) */
      const vv=volta(c);
      registrarContato(c.id,{t:'msg',v:vv>HOJE()?vv:''});
      window.open(linkZap(c,mod.f(c),MUT.num),'_blank','noopener');
      MUT.i++; MUT.num='1'; pintar();
    });
    m.querySelector('#mutnums').addEventListener('click',()=>{
      const nums=MUT.fila.map(c=>{const N=zapNums(c);return N.length?'+55'+N[0].d:''}).filter(Boolean);
      copiarTexto(nums.join('\n'));
    });

    pintar();
  });
}

/* ----------------------------------------------------------
   10d. LISTA PARA LIGAR NO CELULAR (29/09/2026)
   O CRM fica no computador, mas as ligações são feitas pelo
   WhatsApp do celular. Então: o CRM gera um arquivo .html com os
   clientes filtrados → você manda para o celular → liga e registra
   ali mesmo → toca em "Enviar resultado para o computador" → cola
   a mensagem aqui em "Trazer ligações do celular".
   O modelo da página fica em js/celular.js (carregado só quando precisa).
   ---------------------------------------------------------- */
let _celProm=null;
function carregarModeloCelular(){
  if(window.CELULAR_MODELO) return Promise.resolve();
  if(_celProm) return _celProm;
  _celProm=new Promise((ok,falha)=>{
    const sc=document.createElement('script'); sc.src='js/celular.js';
    sc.onload=()=>ok(); sc.onerror=()=>{_celProm=null;falha(new Error('Não achei o arquivo js/celular.js na pasta do CRM.'))};
    document.head.appendChild(sc);
  });
  return _celProm;
}
/* números para o celular: celular ou fixo (fixo de empresa às vezes tem WhatsApp Business) */
function numsCelular(c){
  const o=[],visto=new Set();
  for(const t of [c.tel,c.tel2]){let d=dig(t); if(d.length===8||d.length===9) d='00'+d;
    if(d.length<10||visto.has(d)) continue; visto.add(d); o.push({d,n:t})}
  return o;
}
function descFiltroTela(){
  const q=$('#q').value.trim(), d=[];
  const rot={fila:'Para contatar',retorno:'Retornar hoje',ligar:'Hora de ligar',novos:'Nunca contatados',b30:'Vencem em 30 dias',todos:'Toda a base'};
  if(q) d.push('busca “'+q+'”'); else d.push(rot[filtro]||filtro);
  if(tipo!=='todos') d.push(tipo);
  if($('#loc').value) d.push($('#loc').value);
  if($('#prod').value) d.push($('#prod').value);
  if($('#mes').value){const m=$('#mes').value;d.push('vence em '+MESES[+m.slice(5,7)-1])}
  return d.join(' · ');
}
function modalCelular(){
  const base=filtrados();
  if(!base.length){toast('Não há clientes na lista da tela.','err');return}
  let soComTel=true, lim=0;
  abrirModal(`
    <h3>📱 Lista para ligar no celular</h3>
    <div class="msub">Gera um arquivo com os clientes que estão na tela agora (<b>${esc(descFiltroTela())}</b>).
      Você manda para o celular, liga pelo WhatsApp e registra tudo ali. No fim, o celular devolve o resultado para cá.</div>
    <div class="mcampo"><span class="rotulo">Quantos clientes</span>
      <div class="mservs" id="cellim">${[[0,'Todos'],[15,'15 primeiros'],[30,'30 primeiros'],[50,'50 primeiros']].map(([n,r])=>
        `<button type="button" class="chip${n===lim?' on':''}" data-n="${n}">${r}</button>`).join('')}</div></div>
    <label class="excheck"><input type="checkbox" id="celtel" checked> deixar de fora quem não tem telefone</label>
    <div class="exconta" id="celconta"></div>
    <ol class="passos">
      <li>Clique em <b>Gerar arquivo</b> — ele vai para a pasta Downloads.</li>
      <li>Mande o arquivo para o seu celular (WhatsApp, na conversa com você mesmo) e abra no <b>Chrome</b>.</li>
      <li>Ligue e registre cada um. Ao terminar, toque em <b>Enviar resultado para o computador</b>.</li>
      <li>No computador, copie a mensagem e cole em <b>📥 Trazer ligações do celular</b>.</li>
    </ol>
    <div class="mrod">
      <button class="bt" data-fechar>Cancelar</button>
      <button class="bt pri" id="celok">Gerar arquivo</button>
    </div>`, m=>{
    const lista=()=>{let L=base; if(soComTel) L=L.filter(c=>numsCelular(c).length); if(lim) L=L.slice(0,lim); return L};
    const pinta=()=>{const n=lista().length; m.querySelector('#celconta').textContent=n+(n===1?' cliente':' clientes')+' no arquivo';
      m.querySelector('#celok').disabled=!n};
    m.querySelector('#cellim').addEventListener('click',e=>{const b=e.target.closest('[data-n]');if(!b)return;lim=+b.dataset.n;
      m.querySelectorAll('#cellim .chip').forEach(x=>x.classList.toggle('on',x===b));pinta()});
    m.querySelector('#celtel').addEventListener('change',e=>{soComTel=e.target.checked;pinta()});
    m.querySelector('#celok').addEventListener('click',()=>{
      const L=lista();
      carregarModeloCelular().then(()=>{
        const lid='L'+Date.now().toString(36);
        const D={l:lid,g:HOJE(),ds:descFiltroTela(),fr:FRASES,vo:VOLTAS,c:L.map(c=>{
          const np=nomePartes(c.nome), r=LOG[c.id];
          const H=r&&r.h?r.h.slice().sort((a,b)=>(b.d||'').localeCompare(a.d||'')).slice(0,3).map(x=>({d:x.d,t:x.t,o:x.o||''})):[];
          const tg=np.t.slice(); if(c.aparelhos) tg.push(c.aparelhos+' aparelhos');
          return {id:c.id,nm:c.nome,n:np.n,tg,fc:c.fc||'',loc:c.loc||'',end:c.end||'',prod:c.prod||'',ult:c.ult||'',
            venc:c.venc||'',dias:c.dias,st:c.st,z:numsCelular(c),an:anot(c),ob:c.obs||'',h:H,v:volta(c)};
        })};
        const json=JSON.stringify(D).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
        const html=window.CELULAR_MODELO.split('__DADOS__').join(json);
        baixar(new Blob([html],{type:'text/html;charset=utf-8'}),'Ligacoes_'+HOJE()+'_'+L.length+'clientes.html');
        fecharModal();
        toast('Arquivo salvo em Downloads ('+L.length+' clientes). Mande para o celular e abra no Chrome.','ok');
      }).catch(e=>toast(e.message,'err'));
    });
    pinta();
  });
}

/* ---- trazer o resultado do celular ---- */
function lerPacoteCelular(txt){
  const m=String(txt||'').replace(/\s+/g,'').match(/#CRMPA1#([A-Za-z0-9+/=]+)#FIM#/);
  if(!m) return null;
  try{return JSON.parse(decodeURIComponent(escape(atob(m[1]))))}catch(e){return null}
}
function modalImportarCelular(){
  abrirModal(`
    <h3>📥 Trazer ligações do celular</h3>
    <div class="msub">No WhatsApp Web, abra a mensagem “LIGAÇÕES — Fonte Clara” que o celular mandou,
      copie a mensagem inteira e cole aqui embaixo.</div>
    <textarea id="imptxt" rows="5" placeholder="Cole aqui a mensagem inteira (Ctrl+V)" autofocus></textarea>
    <div id="impprev"></div>
    <div class="mrod">
      <button class="bt" data-fechar>Cancelar</button>
      <button class="bt pri" id="impok" disabled>Trazer para o CRM</button>
    </div>`, m=>{
    m.classList.add('largo');
    let itens=[];
    const porNome=indicePorNome();
    function analisa(){
      const P=lerPacoteCelular(m.querySelector('#imptxt').value);
      const pv=m.querySelector('#impprev'), ok=m.querySelector('#impok');
      itens=[];
      if(!P){pv.innerHTML=m.querySelector('#imptxt').value.trim()
          ?'<div class="aviso err" style="margin-top:10px">Não achei o código do celular nessa mensagem. Copie a mensagem <b>inteira</b>, até o fim (a parte que termina em #FIM#).</div>':'';
        ok.disabled=true;return}
      for(const [id,nm,d,t,o,v,hr] of (P.r||[])){
        let c=achaC(id); if(!c||!nomeBate(c,nm)) c=(porNome[_n(nm)]||[])[0]||null;
        const r=c?LOG[c.id]:null;
        const dup=!!(r&&r.h&&r.h.some(y=>y.d===d&&y.t===t&&(y.o||'')===(o||'')));
        itens.push({c,nm,d,t,o,v,hr,dup});
      }
      const novos=itens.filter(x=>x.c&&!x.dup&&TCONT[x.t]);
      let jaImp=[];try{jaImp=JSON.parse(localStorage.getItem('crm_demo_imp')||'[]')}catch(e){}
      pv.innerHTML=`<div class="exconta">${itens.length} ${itens.length===1?'registro':'registros'} na mensagem · <b>${novos.length} para trazer</b>`
        +`${itens.length-novos.length?' · '+(itens.length-novos.length)+' já estavam aqui ou não foram achados':''}</div>`
        +(jaImp.includes(P.l)?'<div class="exnota" style="margin-top:4px">Essa lista já foi trazida antes — só entra o que ainda não estava aqui.</div>':'')
        +'<div class="implista">'+itens.map(x=>`<div class="impl${x.c&&!x.dup?'':' fora'}">
          <span class="hd">${fmtc(x.d)}${x.hr?' '+esc(x.hr):''}</span><span>${TCONT[x.t]?TCONT[x.t].ic:'?'}</span>
          <b>${esc(x.c?nomePartes(x.c.nome).n:x.nm)}</b>
          <span class="mco">${x.o?'“'+esc(x.o)+'”':esc(TCONT[x.t]?TCONT[x.t].rot:'')}${x.v?' · voltar '+fmtc(x.v):''}</span>
          ${!x.c?'<em>não achei na base</em>':x.dup?'<em>já estava</em>':''}</div>`).join('')+'</div>';
      ok.disabled=!novos.length;
      ok.dataset.lid=P.l||'';
    }
    m.querySelector('#imptxt').addEventListener('input',analisa);
    m.querySelector('#impok').addEventListener('click',e=>{
      let n=0;
      for(const x of itens){ if(!x.c||x.dup||!TCONT[x.t]) continue;
        registrarContato(x.c.id,{t:x.t,o:x.o,v:x.v,d:x.d,hr:x.hr,cel:1}); n++; }
      try{const j=JSON.parse(localStorage.getItem('crm_demo_imp')||'[]');const lid=e.currentTarget.dataset.lid;
        if(lid&&!j.includes(lid)){j.push(lid);localStorage.setItem('crm_demo_imp',JSON.stringify(j.slice(-60)))}}catch(_){}
      fecharModal(); atualizarTudo();
      toast(n+(n===1?' ligação trazida':' ligações trazidas')+' do celular.','ok');
    });
  });
}

/* ----------------------------------------------------------
   11. NAVEGAÇÃO ENTRE VISTAS
   ---------------------------------------------------------- */
pintarModo();
const TITULOS={fila:'Fila de hoje', base:'Toda a base', painel:'Painel',
  agenda:'Retornos marcados', servicos:'Agendamentos de serviço', trocas:'Trocas feitas'};

function irVista(v){
  VISTA=v;
  fecharPop();
  document.querySelectorAll('#nav button').forEach(b=>b.classList.toggle('on',b.dataset.v===v));
  $('#titulo').textContent=TITULOS[v]||v;
  const lista=(v==='fila'||v==='base');
  $('#v-lista').classList.toggle('on',lista);
  $('#v-painel').classList.toggle('on',v==='painel');
  $('#v-agenda').classList.toggle('on',v==='agenda');
  $('#v-servicos').classList.toggle('on',v==='servicos');
  $('#v-trocas').classList.toggle('on',v==='trocas');
  $('#kpis').style.display=lista?'':'none';
  $('#btnkpi').style.display=lista?'':'none';
  $('#btnimp').style.display=lista?'':'none';
  $('#btnexpv').style.display=lista?'':'none';
  $('#btnmut').style.display=lista?'':'none';
  $('#btncel').style.display=lista?'':'none';
  $('#q').parentElement.style.display=lista?'':'none';
  if(lista){
    filtro = v==='fila' ? 'fila' : 'todos';
    kpis(); render(true);
  }
  if(v==='painel') painel();
  if(v==='agenda') agenda();
  if(v==='servicos') servicos();
  if(v==='trocas') trocasView();
  try{localStorage.setItem('crm_demo_vista',v)}catch(e){}
  window.scrollTo({top:0,behavior:'smooth'});
}
$('#nav').addEventListener('click',e=>{const b=e.target.closest('button[data-v]');if(b)irVista(b.dataset.v)});

/* redesenha tudo o que estiver aberto depois de uma mudança */
function atualizarTudo(){
  kpis();
  if(VISTA==='fila'||VISTA==='base') render();
  if(VISTA==='servicos') servicos();
  if(VISTA==='trocas') trocasView();
  if(VISTA==='agenda') agenda();
  if(VISTA==='painel') painel();
}

/* abre um cliente na base, com a linha expandida */
function abrirCliente(id){
  const c=DADOS.find(x=>x.id===id); if(!c) return;
  irVista('base');
  $('#q').value=c.nome; ABERTOS.add(id); render(true);
  const linha=$(`.linha[data-id="${id}"]`);
  if(linha){linha.classList.add('open'); montarDetalhe(linha,id);
    if(linha.scrollIntoView) linha.scrollIntoView({behavior:'smooth',block:'center'})}
}

/* esconder / mostrar os cartões de resumo */
let KPIOCULTO=false;
try{KPIOCULTO=localStorage.getItem('crm_demo_kpi')==='0'}catch(e){}
function pintarKpiBotao(){
  $('#kpis').classList.toggle('oculto',KPIOCULTO);
  $('#btnkpi').textContent=KPIOCULTO?'▾':'▴';
  $('#btnkpi').title=KPIOCULTO?'Mostrar os cartões de resumo':'Esconder os cartões de resumo';
}
$('#btnkpi').addEventListener('click',()=>{
  KPIOCULTO=!KPIOCULTO; try{localStorage.setItem('crm_demo_kpi',KPIOCULTO?'0':'1')}catch(e){}
  pintarKpiBotao();
});
pintarKpiBotao();

/* ----------------------------------------------------------
   12. EVENTOS DA LISTA
   ---------------------------------------------------------- */
$('#kpis').addEventListener('click',e=>{
  const b=e.target.closest('.kpi'); if(!b) return;
  const k=b.dataset.k;
  if(VISTA!=='fila'&&VISTA!=='base') irVista('fila');
  filtro=k;
  kpis(); render(true);
});

$('#tipos').innerHTML=['todos','Pessoa física','Empresa','Escola','Outra cidade']
  .map(t=>`<button class="chip ${t==='todos'?'on':''}" data-t="${t}">${t==='todos'?'Todos os tipos':t}</button>`).join('');
$('#tipos').addEventListener('click',e=>{
  const b=e.target.closest('.chip'); if(!b) return;
  tipo=b.dataset.t;
  [...$('#tipos').children].forEach(x=>x.classList.toggle('on',x===b));
  render(true);
});

const opts=(sel,arr,lbl)=>{const v=sel.value;
  sel.innerHTML=`<option value="">${lbl}</option>`+arr.map(x=>`<option>${esc(x)}</option>`).join('');
  if(arr.includes(v))sel.value=v};
function filtros(){
  opts($('#loc'),[...new Set(DADOS.map(c=>c.loc).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt')),'Todos os bairros');
  opts($('#prod'),[...new Set(DADOS.flatMap(c=>c.aps?c.aps.map(x=>x.prod):[c.prod]).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt')),'Todos os produtos');
  const ms=[...new Set(DADOS.map(c=>(c.venc||'').slice(0,7)).filter(Boolean))].sort();
  const porAno={}; for(const m of ms){(porAno[m.slice(0,4)]=porAno[m.slice(0,4)]||[]).push(m)}
  const atual=$('#mes').value;
  $('#mes').innerHTML='<option value="">Mês do vencimento — todos</option>'+
    Object.keys(porAno).sort().reverse().map(a=>'<optgroup label="'+a+'">'+
      porAno[a].map(m=>`<option value="${m}">${MESES[+m.slice(5,7)-1]} de ${a}</option>`).join('')+'</optgroup>').join('');
  if(ms.includes(atual)) $('#mes').value=atual;
}

/* espera você parar de digitar antes de redesenhar a lista — sem isso,
   cada tecla refazia a tela inteira com os 1.400 cadastros */
let tBusca=null;
$('#q').addEventListener('input',()=>{
  clearTimeout(tBusca);
  tBusca=setTimeout(()=>{tBusca=null;render()},260);
});
$('#q').addEventListener('keydown',e=>{
  if(e.key==='Enter'&&tBusca){clearTimeout(tBusca);tBusca=null;render()}
});
$('#xq').addEventListener('click',()=>{$('#q').value='';$('#q').focus();render(true)});
['#loc','#prod','#mes','#tempo'].forEach(s=>$(s).addEventListener('change',()=>render(true)));
$('#limpar').addEventListener('click',()=>{
  $('#q').value='';$('#loc').value='';$('#prod').value='';$('#mes').value='';$('#tempo').value='';
  render(true); toast('Filtros limpos.');
});
$('#btnimp').addEventListener('click',imprimir);
$('#btnexpv').addEventListener('click',()=>modalExportar('venc'));
$('#btnmut').addEventListener('click',modalMutirao);
$('#btncel').addEventListener('click',modalCelular);
$('#btnimpcel').addEventListener('click',modalImportarCelular);
$('#btnexpag').addEventListener('click',()=>modalExportar('ag'));

$('#lista').addEventListener('click',e=>{
  const mais=e.target.closest('[data-mais]');
  if(mais){e.stopPropagation();mostrarMais(mais.dataset.mais==='tudo');return}
  // registrar contato
  const rg=e.target.closest('[data-reg]');
  if(rg){e.stopPropagation();modalRegistrar(rg.dataset.reg);return}
  // copiar telefone
  const ct=e.target.closest('[data-cptel]');
  if(ct){e.stopPropagation();const n=dig(ct.dataset.cptel)||ct.dataset.cptel;
    copiarTexto(n,'Telefone '+n+' copiado.');
    ct.classList.add('ok');setTimeout(()=>ct.classList.remove('ok'),1200);return}
  // agendar serviço / registrar troca / editar agendamento
  const ag=e.target.closest('[data-ag]');
  if(ag){e.stopPropagation();agendarCliente(ag.dataset.ag);return}
  const tr=e.target.closest('[data-tr]');
  if(tr){e.stopPropagation();modalTroca(tr.dataset.tr);return}
  const ed=e.target.closest('[data-ed]');
  if(ed){e.stopPropagation();const a=AGEND.find(x=>x.id===ed.dataset.ed);if(a)modalAgendar(a);return}
  // não abre/fecha a linha ao pegar a alça de arrastar
  if(e.target.closest('.pegar')){e.stopPropagation();return}
  // botão WhatsApp do cabeçalho — deixa o link abrir, só registra o contato
  const bz=e.target.closest('[data-zap]');
  if(bz){e.stopPropagation();marcarContato(bz.dataset.zap);pintarUltContato(bz.dataset.zap);return}
  // modelo de mensagem
  const m=e.target.closest('.mod[data-m]');
  if(m){
    e.stopPropagation();
    const c=DADOS.find(x=>x.id===m.dataset.id), mod=MODELOS.find(x=>x.k===m.dataset.m);
    if(!c||!mod) return;
    const txt=mod.f(c); MSGPRONTA[c.id]=txt;
    const N=zapNums(c), slot=$(`.previa-slot[data-slot="${c.id}"]`);
    if(slot) slot.innerHTML=`<div class="previa">${esc(txt)}</div>`
      +(N.length>1?`<div class="paraquem"><span>Mandar para qual número?</span>${
        N.map(x=>`<button class="mod" data-envia="${c.id}" data-num="${x.rot}">📱 ${esc(x.n)}</button>`).join('')
        }</div>`:'');
    if(N.length===1){
      marcarContato(c.id); pintarUltContato(c.id);
      window.open(linkZap(c,txt,N[0].rot),'_blank','noopener');
    }
    return;
  }
  // escolheu para qual dos dois números mandar
  const en=e.target.closest('[data-envia]');
  if(en){
    e.stopPropagation();
    const id=en.dataset.envia, c=DADOS.find(x=>x.id===id); if(!c) return;
    const txt=MSGPRONTA[id]||modeloPadrao(c);
    marcarContato(id); pintarUltContato(id);
    window.open(linkZap(c,txt,en.dataset.num),'_blank','noopener');
    return;
  }
  // copiar texto
  const cp=e.target.closest('.mod[data-cp]');
  if(cp){
    e.stopPropagation();
    const c=DADOS.find(x=>x.id===cp.dataset.cp); if(!c) return;
    const txt=modeloPadrao(c); MSGPRONTA[c.id]=txt;
    const slot=$(`.previa-slot[data-slot="${c.id}"]`);
    if(slot) slot.innerHTML=`<div class="previa">${esc(txt)}</div>`;
    marcarContato(c.id); pintarUltContato(c.id);
    copiarTexto(txt);
    return;
  }
  // apagar um item do histórico
  const hd=e.target.closest('[data-hdel]');
  if(hd){
    e.stopPropagation();
    if(!confirm('Apagar este registro do histórico?')) return;
    const [cid,ix]=hd.dataset.hdel.split('|');
    apagarContato(cid,+ix); kpis(); redesenharCartao(cid); toast('Registro apagado.');
    return;
  }
  // abrir/fechar
  const cab=e.target.closest('.cab'); if(!cab) return;
  const linha=cab.parentElement, id=linha.dataset.id;
  if(comJanela()&&!ABERTOS.has(id)){   // janela: só um cliente aberto por vez
    document.querySelectorAll('#lista .linha.open').forEach(l=>{if(l!==linha){ABERTOS.delete(l.dataset.id);l.classList.remove('open')}});
  }
  if(ABERTOS.has(id)) ABERTOS.delete(id); else ABERTOS.add(id);
  linha.classList.toggle('open');
  montarDetalhe(linha,id);
  if(comJanela()) sincPop();
});

/* Troca só o cartão de um cliente, no mesmo lugar da lista. Antes, marcar uma
   situação refazia a lista inteira — lento, e o cliente "pulava" de posição.
   A ordem nova vale na próxima vez que a lista for redesenhada. */
function redesenharCartao(id){
  if(emQuadro()){render();return}   // no quadro o cliente pode mudar de coluna
  const velho=$(`.linha[data-id="${id}"]`), c=DADOS.find(x=>x.id===id);
  if(!velho||!c){render();return}
  const tmp=document.createElement('div'); tmp.innerHTML=cartao(c);
  const novo=tmp.firstElementChild; velho.replaceWith(novo);
}
/* só o cabeçalho (etiquetas) — usado quando se digita a anotação ou a data de
   retorno, para não perder o que está aberto e o foco dentro do cartão */
function atualizarCabecalhoCartao(id){
  const linha=$(`.linha[data-id="${id}"]`), c=DADOS.find(x=>x.id===id);
  if(!linha||!c) return;
  const tmp=document.createElement('div'); tmp.innerHTML=cartao(c);
  const cabNova=tmp.querySelector('.cab'), cab=linha.querySelector('.cab');
  if(cab&&cabNova) cab.replaceWith(cabNova);
  const hist=linha.querySelector('.hist'), dNovo=tmp.querySelector('.hist');
  if(hist&&dNovo) hist.replaceWith(dNovo);
}

/* só liga/desliga o ícone 📝 do cabeçalho — sem trocar botões, para não
   engolir o clique de quem saiu da anotação clicando num botão do cartão */
function pintarNota(id,txt){
  const linha=$(`.linha[data-id="${id}"]`); if(!linha) return;
  let el=linha.querySelector('.tags>.mini.nota');
  if(!txt){if(el)el.remove();return}
  if(!el){
    el=document.createElement('span'); el.className='mini so-ic nota'; el.textContent='📝';
    const selo=linha.querySelector('.tags>.selo'); if(selo) selo.parentElement.insertBefore(el,selo);
  }
  el.title=txt;
}

/* monta o miolo do cartão na hora de abrir, se ainda não tiver sido montado */
function montarDetalhe(linha,id){
  if(!linha||!ABERTOS.has(id)) return;
  const pad=linha.querySelector('.detpad');
  if(!pad||pad.firstElementChild) return;
  const c=DADOS.find(x=>x.id===id); if(!c) return;
  pad.innerHTML=detalhe(c);
}

function gravarCampo(id,campo,valor){
  const c=DADOS.find(x=>x.id===id);
  const r=LOG[id]||{};
  if(valor) r[campo]=valor; else delete r[campo];
  if(!r.n&&c) r.n=c.nome;
  if(!r.a&&!r.obs&&!r.v&&!(r.h&&r.h.length)&&!r.uc){delete LOG[id]} else {LOG[id]=r}
  salvar();
}
let _tmrObs=null;
$('#lista').addEventListener('input',e=>{
  const t=e.target; if(!t||!t.dataset||t.dataset.obs===undefined) return;
  const id=t.dataset.obs, txt=t.value;
  clearTimeout(_tmrObs); _tmrObs=setTimeout(()=>gravarCampo(id,'obs',txt.trim()),400);
});
$('#lista').addEventListener('focusout',e=>{
  const t=e.target; if(!t||!t.dataset||t.dataset.obs===undefined) return;
  clearTimeout(_tmrObs);
  const id=t.dataset.obs, novo=t.value.trim();
  gravarCampo(id,'obs',novo);
  pintarNota(id,novo);
});
$('#lista').addEventListener('change',e=>{
  const t=e.target; if(!t||!t.dataset||t.dataset.v===undefined) return;
  gravarCampo(t.dataset.v,'v',t.value);
  if(t.value) toast('Retorno marcado para '+fmt(t.value)+'. Aparece na Agenda.','ok');
  kpis(); atualizarCabecalhoCartao(t.dataset.v);
});

/* eventos da agenda */
$('#agenda').addEventListener('click',e=>{
  const nav=e.target.closest('[data-ag]');
  if(nav){
    const p=+nav.dataset.ag;
    if(p===0){const h=new Date();agAno=h.getFullYear();agMes=h.getMonth();agSel=''}
    else{agMes+=p; if(agMes>11){agMes=0;agAno++} if(agMes<0){agMes=11;agAno--}}
    agenda(); return;
  }
  const dia=e.target.closest('.dia[data-d]');
  if(dia){agSel=(agSel===dia.dataset.d?'':dia.dataset.d);agenda();return}
  const ir=e.target.closest('[data-ir]');
  if(ir) abrirCliente(ir.dataset.ir);
});

$('#paineis').addEventListener('click',e=>{
  const ir=e.target.closest('[data-ir]'); if(ir) abrirCliente(ir.dataset.ir);
});

/* ----------------------------------------------------------
   13. BACKUP / EXPORTAÇÃO / IMPORTAÇÃO
   ---------------------------------------------------------- */
function baixar(blob,nome){
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=nome; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),4000);
}
$('#btnexp').addEventListener('click',()=>{
  const ids=Object.keys(LOG);
  if(!ids.length){toast('Nenhuma situação marcada ainda.','err');return}
  const lin=ids.map(i=>{
    const c=DADOS.find(x=>x.id===i), r=LOG[i];
    const txt=(r.obs||'').replace(/[;\r\n]+/g,' ');
    const u=ultRegR(r), disse=(u&&u.o?u.o:'').replace(/[;\r\n]+/g,' ');
    return [r.d||'',c?c.nome:(r.n||i),c?c.tel:'',(SIT[r.a]||{}).rot||r.a||'',disse,
      r.v?r.v.split('-').reverse().join('/'):'',txt,(r.h||[]).length,c?c.aba:'',c?c.lin:''].join(';')});
  baixar(new Blob(['﻿Data;Cliente;Telefone;Situação;O que disse (último contato);Voltar em;Anotação fixa;Nº de contatos;Aba;Linha\n'+lin.join('\n')],
    {type:'text/csv;charset=utf-8'}), 'contatos_'+HOJE()+'.csv');
  toast(ids.length+' registros exportados para a pasta Downloads.','ok');
});
function salvarBackup(auto){
  const ids=Object.keys(LOG);
  if(!ids.length&&!AGEND.length&&!TROCAS.length){
    if(!auto) toast('Não há nada marcado para salvar ainda.','err');
    return false;
  }
  const pacote={arquivo:'CRM Fonte Clara — backup de situações',versao:2,quando:HOJE(),total:ids.length,
    situacoes:LOG,agendamentos:AGEND,trocas:TROCAS};
  baixar(new Blob([JSON.stringify(pacote,null,1)],{type:'application/json;charset=utf-8'}),
    'crm_situacoes_'+HOJE()+'.json');
  try{localStorage.setItem('crm_demo_bkp',HOJE())}catch(e){}
  avisos();
  toast(auto
    ? 'Backup do dia salvo sozinho na pasta Downloads ('+ids.length+' situações).'
    : 'Backup com '+ids.length+' situações salvo em Downloads. Guarde uma cópia em Financeiro.','ok');
  return true;
}
$('#btnbkp').addEventListener('click',()=>salvarBackup(false));
$('#btnrst').addEventListener('click',()=>$('#arqbkp').click());
$('#arqbkp').addEventListener('change',ev=>{
  const f=ev.target.files[0]; if(!f) return;
  const fr=new FileReader();
  fr.onload=()=>{
    try{
      const j=JSON.parse(fr.result);
      const S=j&&j.situacoes?j.situacoes:j;
      if(!S||typeof S!=='object'||Array.isArray(S)) throw new Error('formato');
      const porNome=indicePorNome();
      let novos=0,atualizados=0,ignorados=0;
      for(const k in S){
        const r=S[k]; if(!r||(!r.a&&!r.obs&&!r.v&&!(r.h&&r.h.length))){ignorados++;continue}
        const a=r.a?(SIT[r.a]?r.a:(_MIG[r.a]||'')):'';
        if(r.a&&!a){ignorados++;continue}
        let id=ALIAS[k]||k;
        const alvo=r.n?(porNome[_n(r.n)]||[])[0]:null;
        if(alvo) id=alvo.id;
        const cur=LOG[id];
        if(!cur){
          LOG[id]={}; if(a){LOG[id].a=a;LOG[id].d=r.d||''}
          if(r.obs)LOG[id].obs=r.obs; if(r.v)LOG[id].v=r.v; LOG[id].n=r.n||'';
          if(r.uc)LOG[id].uc=r.uc; LOG[id].h=Array.isArray(r.h)?r.h.slice():[]; novos++;
        }else{
          let mudou=false;
          /* histórico: junta os dois, sem repetir o mesmo registro */
          if(Array.isArray(r.h)){cur.h=cur.h||[];
            for(const x of r.h){if(!cur.h.some(y=>y.d===x.d&&y.t===x.t&&(y.o||'')===(x.o||''))){cur.h.push(x);mudou=true}}}
          if(r.uc&&(!cur.uc||r.uc>cur.uc)){cur.uc=r.uc;mudou=true}
          if(a&&(r.d||'')>(cur.d||'')){cur.a=a;cur.d=r.d||'';mudou=true}
          if(r.obs&&r.obs!==cur.obs){cur.obs=cur.obs?(cur.obs+' | '+r.obs):r.obs;mudou=true}
          if(r.v&&r.v!==cur.v&&(!cur.v||r.v>cur.v)){cur.v=r.v;mudou=true}
          if(!cur.n&&r.n)cur.n=r.n;
          if(mudou)atualizados++; else ignorados++;
        }
      }
      // agendamentos e trocas do backup (versão 2)
      let ag=0,tr=0;
      if(Array.isArray(j.agendamentos)) for(const a of j.agendamentos){
        if(!a||!a.id||AGEND.some(x=>x.id===a.id)) continue; AGEND.push(a); ag++;}
      if(Array.isArray(j.trocas)) for(const t of j.trocas){
        if(!t||!t.id||TROCAS.some(x=>x.id===t.id)) continue; TROCAS.push(t); tr++;}
      if(ag) salvarAg(); if(tr) salvarTr();
      migrarIds();
      salvar(); kpis(); render(true); atualizarTudo();
      toast('Backup restaurado: '+novos+' novas, '+atualizados+' atualizadas'
        +(ignorados?', '+ignorados+' ignoradas':'')
        +(ag?' · '+ag+' agendamentos':'')+(tr?' · '+tr+' trocas':'')+'.','ok');
    }catch(e){
      toast('Esse arquivo não parece um backup do CRM (esperado o .json do botão Salvar backup).','err');
    }
    ev.target.value='';
  };
  fr.onerror=()=>toast('Erro ao abrir o arquivo.','err');
  fr.readAsText(f);
});

/* --- carregar planilhas --- */
function lendo(nome){const a=$('#aviso');a.className='aviso carregando';a.textContent='Lendo '+nome+'…'}
/* O leitor de planilhas (js/sheetjs.js, 881 KB) só é carregado quando precisa:
   ao clicar em "Atualizar clientes" — ou sozinho, uns segundos depois de abrir,
   quando o computador já está parado. Antes ele era lido toda vez que o CRM abria. */
let _xlsxProm=null;
function carregarLeitor(){
  if(window.XLSX) return Promise.resolve();
  if(_xlsxProm) return _xlsxProm;
  _xlsxProm=new Promise((ok,falha)=>{
    const sc=document.createElement('script'); sc.src='js/sheetjs.js';
    sc.onload=()=>ok();
    sc.onerror=()=>{_xlsxProm=null;falha(new Error('Não consegui carregar o leitor de planilhas (js/sheetjs.js). Confira se o arquivo está na pasta js do CRM.'))};
    document.head.appendChild(sc);
  });
  return _xlsxProm;
}
$('#btnupd').addEventListener('click',()=>{carregarLeitor().catch(()=>{}); $('#arq').click()});
$('#arq').addEventListener('change',ev=>{
  const f=ev.target.files[0]; if(!f) return;
  irVista(VISTA==='fila'||VISTA==='base'?VISTA:'fila');
  lendo(f.name);
  const fr=new FileReader();
  fr.onload=()=>carregarLeitor().then(()=>{
    const a=$('#aviso');
    try{
      const wb=XLSX.read(new Uint8Array(fr.result),{type:'array',cellDates:true});
      const R=casar(recalcular(construir(wb)),CONTATOS);
      DADOS=R; BASEDATA=HOJE(); reancorar();
      let extra='';
      try{localStorage.setItem('crm_demo_base',JSON.stringify({q:BASEDATA,d:R}))}
      catch(e){extra=' (não coube na memória do navegador — vai precisar carregar de novo na próxima abertura)'}
      filtros(); cabecalho(); avisos(); kpis(true); render(true);
      if(SEMPADRAO.size){
        a.className='aviso err';
        a.textContent='Carreguei '+R.length+' clientes, mas esse arquivo não tem a coluna TIPO DE TELEFONE — os botões de WhatsApp não vão funcionar. Use a Planilha de Clientes.xlsx atual, em Financeiro/01 - Clientes.';
        toast('Planilha sem a coluna TIPO DE TELEFONE.','err');
      }else{
        const nAp=R.reduce((n,c)=>n+(c.aps?c.aps.length:1),0);
        a.className='aviso ok'; a.textContent=R.length+' clientes ('+nAp+' aparelhos) carregados de '+f.name+extra;
        toast(R.length+' clientes carregados.','ok');
      }
    }catch(e){
      a.className='aviso err'; a.textContent='Não consegui ler o arquivo: '+e.message;
      toast('Não consegui ler o arquivo.','err');
    }
    ev.target.value='';
  }).catch(e=>{
    const a=$('#aviso'); a.className='aviso err'; a.textContent=e.message;
    toast('Não consegui abrir o leitor de planilhas.','err'); ev.target.value='';
  });
  fr.onerror=()=>toast('Erro ao abrir o arquivo.','err');
  fr.readAsArrayBuffer(f);
});
/* --- IMPORTAR CONTATOS (desativado em 31/08/2026) ---------------------
   O botão "Atualizar contatos" lia as abas mensais da planilha antiga de contatos
   (Documentos/Nova pasta/wb/lg/2026.xlsx). Essas abas existiam por causa da
   regra antiga de comissão (a venda tinha que sair em até 30 dias do contato)
   e não são mais preenchidas — hoje o próprio CRM guarda as situações.

   O botão e o seletor de arquivo saíram da tela. Os contatos que já foram
   importados continuam salvos (crm_demo_ct) e aparecem normalmente nos cartões.
   A função construirContatos() foi mantida logo acima: se um dia for preciso
   voltar a importar, basta devolver o botão ao index.html e este bloco aqui.
   ------------------------------------------------------------------- */

/* ----------------------------------------------------------
   13b. TEMA CLARO / ESCURO
   O tema é aplicado já no <head> do index.html (sem piscar).
   Aqui fica só o botão 🌙/☀️ e o "seguir o Windows" quando
   não há escolha salva. Chave: crm_demo_tema = 'claro' | 'escuro'.
   ---------------------------------------------------------- */
const temaAtual=()=>document.documentElement.getAttribute('data-tema')==='escuro'?'escuro':'claro';
let _tmrTema=null;
function aplicarTema(t,guardar){
  const h=document.documentElement;
  h.classList.add('trocando-tema');
  h.setAttribute('data-tema',t);
  if(guardar) try{localStorage.setItem('crm_demo_tema',t)}catch(e){}
  $('#tema').title=t==='escuro'?'Mudar para o modo claro':'Mudar para o modo escuro';
  clearTimeout(_tmrTema); _tmrTema=setTimeout(()=>h.classList.remove('trocando-tema'),380);
}
$('#tema').addEventListener('click',()=>aplicarTema(temaAtual()==='escuro'?'claro':'escuro',true));
$('#tema').title=temaAtual()==='escuro'?'Mudar para o modo claro':'Mudar para o modo escuro';
try{
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change',e=>{
    let t=null; try{t=localStorage.getItem('crm_demo_tema')}catch(_){}
    if(t!=='claro'&&t!=='escuro') aplicarTema(e.matches?'escuro':'claro',false);
  });
}catch(e){}

/* ----------------------------------------------------------
   13c. AVISOS — base de clientes velha e backup atrasado
   ---------------------------------------------------------- */
const DIAS_BASE_VELHA=3;   // a partir de quantos dias a base é "velha"
const DIAS_BACKUP=7;       // lembrar do backup depois de tantos dias
function avisos(){
  const A=[];
  const db=diasDe(BASEDATA), velha=db!==null&&db>=DIAS_BASE_VELHA;
  $('#btnupd').classList.toggle('velha',velha);
  if(velha) A.push(`<div class="alerta"><span class="ic">⚠️</span>
    <div class="tx"><b>A base de clientes é de ${fmt(BASEDATA)} — ${db} dias atrás.</b>
    ${(()=>{const n=TROCAS.filter(t=>(t.quando||t.d||'')>BASEDATA).length;
      return n?`Desde então você registrou <b>${n} ${n===1?'troca':'trocas'}</b> aqui no CRM — quem trocou pela loja sem passar pelo CRM também pode estar aparecendo como vencido.`
        :'Trocas lançadas na planilha depois disso não aparecem aqui, e a fila pode estar mostrando vencimentos errados.'})()}</div>
    <button class="bt pri" data-al="upd">⬆️ Atualizar clientes</button></div>`);

  const temAlgo=Object.keys(LOG).length||AGEND.length||TROCAS.length;
  let ult='',adiado='';
  try{ult=localStorage.getItem('crm_demo_bkp')||'';adiado=localStorage.getItem('crm_demo_bkp_adiar')||''}catch(e){}
  const dk=ult?diasDe(ult):null;
  if(temAlgo&&(dk===null||dk>=DIAS_BACKUP)&&adiado!==HOJE()) A.push(`<div class="alerta"><span class="ic">💾</span>
    <div class="tx"><b>${dk===null?'Ainda não há backup registrado.':'Último backup há '+dk+' dias.'}</b>
    O que você marca aqui fica só dentro deste navegador — salve uma cópia e guarde em Financeiro.</div>
    <button class="bt pri" data-al="bkp">💾 Salvar backup agora</button>
    <button class="fechar" data-al="adiar" title="Esconde este aviso até amanhã">Lembrar amanhã</button></div>`);
  $('#avisos').innerHTML=A.join('');
}
$('#avisos').addEventListener('click',e=>{
  const b=e.target.closest('[data-al]'); if(!b) return;
  const k=b.dataset.al;
  if(k==='upd') $('#btnupd').click();
  if(k==='bkp') $('#btnbkp').click();
  if(k==='adiar'){try{localStorage.setItem('crm_demo_bkp_adiar',HOJE())}catch(e){} avisos()}
});

/* tecla / vai direto para a busca (fora de campos de texto e janelas) */
document.addEventListener('keydown',e=>{
  if(e.key!=='/'||e.ctrlKey||e.altKey||e.metaKey) return;
  const t=document.activeElement&&document.activeElement.tagName;
  if(/INPUT|TEXTAREA|SELECT/.test(t)||!$('#fundo').hidden||!$('#q').offsetParent) return;
  e.preventDefault(); $('#q').focus(); $('#q').select();
});

/* ----------------------------------------------------------
   14. INÍCIO
   ---------------------------------------------------------- */
function cabecalho(){
  $('#hoje').textContent=new Date().toLocaleDateString('pt-BR',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  const nc=DADOS.filter(c=>ultContato(c)).length;
  const db=diasDe(BASEDATA), velha=db!==null&&db>=DIAS_BASE_VELHA;
  $('#rodape').innerHTML=`${DADOS.length.toLocaleString('pt-BR')} clientes · ${nc} já contatados<br>
    <span class="${velha?'velha':''}">Base de clientes de ${fmt(BASEDATA)}${velha?' · há '+db+' dias':''}</span>`;
}
/* ----------------------------------------------------------
   14b. BACKUP AUTOMÁTICO — uma cópia por dia, ao abrir o CRM
   ---------------------------------------------------------- */
const BKP_AUTO='crm_demo_bkp_auto';          /* '1' ligado (padrão), '0' desligado */
function autoLigado(){
  try{const v=localStorage.getItem(BKP_AUTO); return v==='1'}catch(e){return false}
}
function backupAutomatico(){
  if(!autoLigado()) return;
  let ult=''; try{ult=localStorage.getItem('crm_demo_bkp')||''}catch(e){}
  if(ult===HOJE()) return;                 /* já salvou hoje */
  salvarBackup(true);
}
const _ckauto=$('#ckbkpauto');
if(_ckauto){
  _ckauto.checked=autoLigado();
  _ckauto.addEventListener('change',()=>{
    try{localStorage.setItem(BKP_AUTO,_ckauto.checked?'1':'0')}catch(e){}
    toast(_ckauto.checked?'Backup automático ligado — uma cópia por dia na pasta Downloads.'
                         :'Backup automático desligado. Salve backup pelo botão de vez em quando.','ok');
  });
}

filtros(); cabecalho(); avisos(); kpis(true); render(true);
setTimeout(()=>{(window.requestIdleCallback||(f=>f()))(()=>{try{backupAutomatico()}catch(e){}})},3000);
setTimeout(()=>{(window.requestIdleCallback||(f=>f()))(()=>carregarLeitor().catch(()=>{}))},6000);
if(BASEDATA!==GERADO||CTDATA!==GERADO){
  $('#aviso').textContent='Clientes de '+fmt(BASEDATA)+' · contatos de '+fmt(CTDATA)+'.';
}
