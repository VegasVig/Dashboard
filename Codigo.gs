/* ═══════════════════════════════════════════════════════════════
   APPS SCRIPT – VEGAS VIGILÂNCIA  v10 · Gestão de Uniformes
   ---------------------------------------------------------------
   v10: estoque por peça + tamanho, reservas, entrega e devolução
        com checklist, ocorrências financeiras, usuários com senha.
   Compatível com os dados anteriores:
     · Página1 (uniformes) e Página2 (materiais) continuam iguais;
       a Página1 só ganha as colunas P a S à direita.
     · salvarUniforme / salvarMaterial (formulários antigos) seguem
       funcionando para a transição.
   PRIMEIRO USO: edite PRIMEIRO_ADMIN abaixo e rode criarPrimeiroAdmin
   uma vez no editor. Depois apague a senha daqui.
═══════════════════════════════════════════════════════════════ */

const PRIMEIRO_ADMIN = { usuario: "admin", senha: "", nome: "Administrador" };

const ABA = {
  PEDIDOS: "Página1", MATERIAIS: "Página2",
  PECAS: "Estoque_Pecas", MOV: "Estoque_Movimentacoes",
  ENT: "Uniforme_Entregas", DEV: "Uniforme_Devolucoes",
  OCO: "Uniforme_Ocorrencias", USR: "Usuarios", LOG: "Auditoria"
};

const COLS = {
  PECAS: ["sku","nome","categoria","tamanho","saldo","reservado","quarentena","minimo","valorUnit","estado","obs","ativo","alertaAtivo","qtdInicial","criadoEm","atualizadoEm"],
  MOV:   ["id","dataHora","sku","nome","tamanho","tipo","qtd","saldoAntes","saldoDepois","reservadoAntes","reservadoDepois","quarentenaDepois","responsavel","motivo","referencia"],
  ENT:   ["id","lote","dataHora","protocolo","cpf","nome","cidade","posto","sku","peca","tamanho","qtdSolicitada","qtdEntregue","estadoEntrega","valorUnit","responsavel","obs","qtdDevolvida","qtdDanificada","qtdFaltante"],
  DEV:   ["id","lote","dataHora","entregaId","protocolo","cpf","nome","sku","peca","tamanho","pendenteAntes","qtdBoa","qtdDesgaste","qtdDanificada","qtdFaltante","condicao","valorUnitEntrega","valorFaltante","valorDanoEstimado","motivoDano","responsavel","obs","foto"],
  OCO:   ["id","dataHora","lote","cpf","nome","resumo","valorFaltante","valorDano","valorTotal","status","calculadoPor","revisadoPor","aprovadoPor","historico","contestacao"],
  USR:   ["usuario","nome","perfil","salt","hash","ativo","criadoEm"],
  LOG:   ["dataHora","usuario","acao","detalhe"]
};

// Página1: colunas A–O são as originais. P–S novas.
const PED = { data:0, nome:1, cpf:2, telefone:3, posto:4, uniformes:5, obs:6, assinatura:7, status:8, protocolo:9, foto:10,
              dataAprov:11, dataSepar:12, dataEntrega:13, cidade:14, funcao:15, itens:16, historico:17, atualizadoEm:18 };
const PED_HEAD = ["Data","Nome","CPF","Telefone","Posto","Uniformes","Observação","Assinatura","Status","Protocolo","Foto",
                  "Data aprovação","Data separação","Data entrega","Cidade","Função","Itens (JSON)","Histórico (JSON)","Atualizado em"];

const ST = {
  AGUARD: "Aguardando aprovação", APROV: "Aprovado", SEPAR: "Separando peças", DISP: "Disponível para retirada",
  ENTPARC: "Entregue parcialmente", ENTREG: "Entregue ao funcionário", DEVPEND: "Devolução pendente",
  DEVPARC: "Parcialmente devolvido", DEVOK: "Devolução concluída", CANC: "Cancelado"
};
const ST_LEGADO = ["EM ANÁLISE","PENDENTE","SEPARADO","ENTREGUE"];
const FIN = ["Em análise","Aguardando aprovação","Aprovado administrativamente","Contestação registrada","Resolvido","Cancelado após revisão"];
const MINIMO_PADRAO = 5;
const SESSAO_SEG = 8 * 3600;

/* ───────────── utilidades ───────────── */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function erro_(msg, extra) { const e = new Error(msg); if (extra) Object.assign(e, extra); return e; }
function digitos_(v) { return String(v == null ? "" : v).replace(/\D/g, ""); }
function int_(v) { const n = Math.floor(Number(v)); return isFinite(n) ? n : 0; }
function num_(v) { const n = Number(String(v == null ? "" : v).replace(",", ".")); return isFinite(n) ? n : 0; }
function txt_(v, max) { return String(v == null ? "" : v).trim().slice(0, max || 500); }
function iso_(v) { return v instanceof Date ? v.toISOString() : (v ? String(v) : ""); }
function norm_(s) { return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/\s+/g, " ").trim(); }
function slug_(s) { return norm_(s).toUpperCase().replace(/[^A-Z0-9]+/g, "-").replace(/^-|-$/g, ""); }
function id_(pref) { return pref + "-" + Utilities.getUuid().replace(/-/g, "").slice(0, 10).toUpperCase(); }
function json_(raw, padrao) { try { return raw ? JSON.parse(raw) : padrao; } catch (e) { return padrao; } }
function resposta(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function comLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(30000); try { return fn(); } finally { l.releaseLock(); } }

function cpfValido_(c) {
  c = digitos_(c);
  if (c.length !== 11 || /^(\d)\1{10}$/.test(c)) return false;
  for (let t = 9; t < 11; t++) {
    let s = 0;
    for (let i = 0; i < t; i++) s += Number(c[i]) * (t + 1 - i);
    if (((s * 10) % 11) % 10 !== Number(c[t])) return false;
  }
  return true;
}

/* ───────────── acesso às abas ───────────── */
function aba_(chave) {
  const nome = ABA[chave];
  let sh = ss_().getSheetByName(nome);
  if (!sh) {
    sh = ss_().insertSheet(nome);
    sh.getRange(1, 1, 1, COLS[chave].length).setValues([COLS[chave]]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function tabela_(chave) {
  const sh = aba_(chave), cols = COLS[chave];
  const v = sh.getDataRange().getValues(); v.shift();
  const rows = v.filter(r => String(r[0]) !== "").map((r, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => o[c] = r[j] instanceof Date ? r[j] : (r[j] === undefined ? "" : r[j]));
    return o;
  });
  return { chave, sh, cols, rows };
}
function linha_(t, o) { return t.cols.map(c => o[c] === undefined || o[c] === null ? "" : o[c]); }
function gravar_(t, o) { t.sh.getRange(o._row, 1, 1, t.cols.length).setValues([linha_(t, o)]); }
function inserir_(t, objs) {
  if (!objs.length) return;
  const ini = t.sh.getLastRow() + 1;
  t.sh.getRange(ini, 1, objs.length, t.cols.length).setValues(objs.map(o => linha_(t, o)));
  objs.forEach((o, i) => { o._row = ini + i; t.rows.push(o); });
}
function plano_(o) {
  const r = {};
  Object.keys(o).forEach(k => { if (k !== "_row") r[k] = o[k] instanceof Date ? o[k].toISOString() : o[k]; });
  return r;
}
function auditar_(usuario, acao, detalhe) {
  aba_("LOG").appendRow([new Date(), usuario || "", acao, typeof detalhe === "string" ? detalhe : JSON.stringify(detalhe)]);
}

/* ───────────── pedidos (Página1) ───────────── */
function abaPedidos_() {
  let sh = ss_().getSheetByName(ABA.PEDIDOS) || ss_().getSheets()[0];
  const head = sh.getRange(1, 1, 1, PED_HEAD.length).getValues()[0];
  for (let i = 15; i < PED_HEAD.length; i++) if (!head[i]) sh.getRange(1, i + 1).setValue(PED_HEAD[i]);
  return sh;
}
function pedidoDeLinha_(r, row) {
  const itens = json_(r[PED.itens], null);
  return {
    _row: row,
    protocolo: String(r[PED.protocolo] || ""), data: iso_(r[PED.data]),
    nome: String(r[PED.nome] || ""), cpf: String(r[PED.cpf] || ""), telefone: String(r[PED.telefone] || ""),
    cidade: String(r[PED.cidade] || ""), posto: String(r[PED.posto] || ""), funcao: String(r[PED.funcao] || ""),
    uniformes: String(r[PED.uniformes] || ""), obs: String(r[PED.obs] || ""), status: String(r[PED.status] || ""),
    itens: Array.isArray(itens) ? itens : null, legado: !Array.isArray(itens),
    historico: json_(r[PED.historico], []),
    dataAprov: iso_(r[PED.dataAprov]), dataSepar: iso_(r[PED.dataSepar]), dataEntrega: iso_(r[PED.dataEntrega]),
  };
}
// Lê sem as colunas H (assinatura) e K (foto): base64 nessas células deixaria cada chamada lenta.
function pedidos_() {
  const sh = abaPedidos_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const A = sh.getRange(2, 1, n, 7).getValues(), B = sh.getRange(2, 9, n, 2).getValues(), C = sh.getRange(2, 12, n, 8).getValues();
  return A.map((r, i) => pedidoDeLinha_(r.concat([""], B[i], [""], C[i]), i + 2)).filter(p => p.nome || p.protocolo);
}
function acharPedido_(protocolo) {
  const p = pedidos_().find(x => x.protocolo === String(protocolo || "").trim());
  if (!p) throw erro_("Pedido " + protocolo + " não encontrado.");
  return p;
}
function salvarPedido_(p, usuario, evento, col) {
  const sh = abaPedidos_(), agora = new Date();
  p.historico.push({ em: agora.toISOString(), por: usuario || "", status: p.status, evento: evento || "" });
  sh.getRange(p._row, PED.status + 1).setValue(p.status);
  if (p.itens) {
    sh.getRange(p._row, PED.itens + 1).setValue(JSON.stringify(p.itens));
    sh.getRange(p._row, PED.uniformes + 1).setValue(resumoItens_(p.itens));
  }
  sh.getRange(p._row, PED.historico + 1).setValue(JSON.stringify(p.historico));
  sh.getRange(p._row, PED.atualizadoEm + 1).setValue(agora);
  if (col) sh.getRange(p._row, PED[col] + 1).setValue(agora);
}
function resumoItens_(itens) {
  return itens.map(i => `${i.peca} ${i.tamanho} x${i.qtd}` + (i.entregue ? ` (entregue ${i.entregue})` : "")).join(", ");
}
function proximoProtocolo_() {
  const props = PropertiesService.getScriptProperties();
  const ano = new Date().getFullYear(), chave = "SEQ_" + ano;
  const n = int_(props.getProperty(chave)) + 1;
  props.setProperty(chave, String(n));
  return `VGS-${ano}-${String(n).padStart(4, "0")}`;
}

/* ───────────── anexos (Drive quando passar do limite da célula) ───────────── */
function pastaAnexos_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("PASTA_ANEXOS_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const nome = "Vegas - Uniformes (anexos)";
  const it = DriveApp.getFoldersByName(nome);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(nome);
  props.setProperty("PASTA_ANEXOS_ID", f.getId());
  return f;
}
function guardarImagem_(dataUrl, nome, inlineMax) {
  dataUrl = String(dataUrl || "");
  if (!dataUrl.startsWith("data:image/")) return dataUrl.startsWith("http") ? dataUrl : "";
  if (inlineMax && dataUrl.length <= inlineMax) return dataUrl;
  const m = dataUrl.match(/^data:(image\/[a-z+]+);base64,(.+)$/);
  if (!m) return "";
  const arq = pastaAnexos_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nome));
  try { arq.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return "https://drive.google.com/thumbnail?id=" + arq.getId() + "&sz=w1200";
}

/* ───────────── usuários e sessão ───────────── */
function hash_(salt, senha) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + "|" + senha, Utilities.Charset.UTF_8));
}
function criarUsuario_(usuario, senha, nome, perfil) {
  usuario = norm_(usuario).replace(/[^a-z0-9._-]/g, "");
  if (!usuario) throw erro_("Usuário inválido.");
  if (String(senha || "").length < 8) throw erro_("A senha precisa ter pelo menos 8 caracteres.");
  if (["estoque", "admin"].indexOf(perfil) < 0) throw erro_("Perfil inválido.");
  const t = tabela_("USR");
  if (t.rows.some(u => u.usuario === usuario)) throw erro_("Usuário já existe.");
  const salt = Utilities.getUuid();
  inserir_(t, [{ usuario, nome: txt_(nome, 80) || usuario, perfil, salt, hash: hash_(salt, senha), ativo: "SIM", criadoEm: new Date() }]);
  return usuario;
}
function criarPrimeiroAdmin() {
  if (!PRIMEIRO_ADMIN.senha) throw new Error("Preencha PRIMEIRO_ADMIN.senha no topo do arquivo antes de rodar.");
  criarUsuario_(PRIMEIRO_ADMIN.usuario, PRIMEIRO_ADMIN.senha, PRIMEIRO_ADMIN.nome, "admin");
  Logger.log("Admin criado: " + PRIMEIRO_ADMIN.usuario + ". Agora apague a senha do topo do arquivo.");
}
function sessao_(token, perfis) {
  const cache = CacheService.getScriptCache();
  const raw = token ? cache.get("tok_" + token) : null;
  if (!raw) throw erro_("Sessão expirada. Entre novamente.", { auth: true });
  const u = JSON.parse(raw);
  if (perfis && perfis.indexOf(u.perfil) < 0) throw erro_("Seu perfil não tem permissão para esta ação.");
  cache.put("tok_" + token, raw, SESSAO_SEG);
  return u;
}
function login_(d) {
  const cache = CacheService.getScriptCache();
  const usuario = norm_(d.usuario), chaveTent = "login_" + usuario;
  const tent = int_(cache.get(chaveTent));
  if (tent >= 8) throw erro_("Muitas tentativas. Aguarde 15 minutos.");
  const u = tabela_("USR").rows.find(x => x.usuario === usuario && x.ativo === "SIM");
  if (!u || hash_(u.salt, String(d.senha || "")) !== u.hash) {
    cache.put(chaveTent, String(tent + 1), 900);
    throw erro_("Usuário ou senha incorretos.");
  }
  cache.remove(chaveTent);
  const token = Utilities.getUuid() + Utilities.getUuid();
  const sess = { usuario: u.usuario, nome: u.nome, perfil: u.perfil };
  cache.put("tok_" + token, JSON.stringify(sess), SESSAO_SEG);
  auditar_(u.usuario, "login", "");
  return { sucesso: true, token, usuario: sess };
}

/* ───────────── estoque: núcleo de movimentação ───────────── */
function ctxEstoque_(u) { return { u, pecas: tabela_("PECAS"), movs: [], sujas: {}, alertas: [] }; }
function peca_(ctx, sku) {
  const p = ctx.pecas.rows.find(x => x.sku === sku);
  if (!p) throw erro_("Peça " + sku + " não cadastrada.");
  return p;
}
function disponivel_(p) { return int_(p.saldo) - int_(p.reservado); }
function mover_(ctx, p, tipo, dSaldo, dRes, dQuar, qtd, motivo, ref) {
  const a = { s: int_(p.saldo), r: int_(p.reservado), q: int_(p.quarentena) };
  const n = { s: a.s + dSaldo, r: a.r + dRes, q: a.q + dQuar };
  const rot = `${p.nome} ${p.tamanho}`;
  if (n.s < 0) throw erro_(`${rot}: saldo insuficiente (saldo ${a.s}).`);
  if (n.r < 0) throw erro_(`${rot}: reserva não pode ficar negativa.`);
  if (n.q < 0) throw erro_(`${rot}: quarentena insuficiente (${a.q}).`);
  if (n.r > n.s) throw erro_(`${rot}: saldo ficaria abaixo das reservas ativas (${n.r}). Disponível hoje: ${a.s - a.r}.`);
  p.saldo = n.s; p.reservado = n.r; p.quarentena = n.q; p.atualizadoEm = new Date();
  ctx.sujas[p.sku] = p;
  ctx.movs.push({ id: id_("MOV"), dataHora: new Date(), sku: p.sku, nome: p.nome, tamanho: p.tamanho, tipo, qtd,
    saldoAntes: a.s, saldoDepois: n.s, reservadoAntes: a.r, reservadoDepois: n.r, quarentenaDepois: n.q,
    responsavel: ctx.u.nome, motivo: motivo || "", referencia: ref || "" });
  checarAlerta_(ctx, p);
}
function checarAlerta_(ctx, p) {
  const disp = disponivel_(p), min = p.minimo === "" ? MINIMO_PADRAO : int_(p.minimo);
  const baixo = p.ativo === "SIM" && disp <= min, ativo = p.alertaAtivo === "SIM";
  if (baixo === ativo) return;
  p.alertaAtivo = baixo ? "SIM" : "";
  ctx.sujas[p.sku] = p;
  ctx.movs.push({ id: id_("MOV"), dataHora: new Date(), sku: p.sku, nome: p.nome, tamanho: p.tamanho,
    tipo: baixo ? (disp <= 0 ? "ALERTA_SEM_ESTOQUE" : "ALERTA_ESTOQUE_BAIXO") : "ALERTA_NORMALIZADO", qtd: 0,
    saldoAntes: int_(p.saldo), saldoDepois: int_(p.saldo), reservadoAntes: int_(p.reservado), reservadoDepois: int_(p.reservado),
    quarentenaDepois: int_(p.quarentena), responsavel: "Sistema", motivo: `Disponível ${disp} / mínimo ${min}`, referencia: "" });
}
function commitEstoque_(ctx) {
  Object.keys(ctx.sujas).forEach(k => { const p = ctx.sujas[k]; if (p._row) gravar_(ctx.pecas, p); });
  const novas = ctx.pecas.rows.filter(p => !p._row);
  if (novas.length) {
    ctx.pecas.rows = ctx.pecas.rows.filter(p => p._row);
    inserir_(ctx.pecas, novas);
  }
  inserir_(tabela_("MOV"), ctx.movs);
}

/* ───────────── ações públicas (funcionário) ───────────── */
function catalogo_() {
  const pecas = tabela_("PECAS").rows.filter(p => p.ativo === "SIM").map(p => ({
    sku: p.sku, nome: p.nome, categoria: p.categoria, tamanho: String(p.tamanho), disponivel: Math.max(0, disponivel_(p))
  }));
  return { sucesso: true, pecas };
}

function novoPedido_(d) {
  const nome = txt_(d.nome, 120), cpf = digitos_(d.cpf), tel = digitos_(d.telefone);
  if (nome.split(" ").filter(Boolean).length < 2) throw erro_("Informe o nome completo.");
  if (!cpfValido_(cpf)) throw erro_("CPF inválido. Confira os números.");
  if (tel.length < 10) throw erro_("Telefone incompleto.");
  ["cidade", "posto", "funcao"].forEach(k => { if (!txt_(d[k])) throw erro_("Preencha cidade, posto e função."); });
  if (String(d.assinatura || "").length < 200) throw erro_("Assine no quadro antes de enviar.");
  const pedidosItens = Array.isArray(d.itens) ? d.itens : [];
  if (!pedidosItens.length) throw erro_("Escolha ao menos uma peça.");

  const cache = CacheService.getScriptCache(), chave = "ped_" + txt_(d.chave, 60);
  if (d.chave) { const ja = cache.get(chave); if (ja) return JSON.parse(ja); }

  return comLock_(() => {
    const aberto = pedidos_().find(p => digitos_(p.cpf) === cpf && (p.status === ST.AGUARD || p.status === "EM ANÁLISE"));
    if (aberto) throw erro_(`Você já tem o pedido ${aberto.protocolo} aguardando aprovação. Consulte em "Meus pedidos" antes de pedir de novo.`, { protocoloAberto: aberto.protocolo });

    const pecas = tabela_("PECAS").rows, falta = [], itens = [], vistos = {};
    pedidosItens.forEach(it => {
      const p = pecas.find(x => x.sku === it.sku && x.ativo === "SIM");
      const qtd = int_(it.qtd);
      if (!p) throw erro_("Uma das peças escolhidas não está mais disponível. Atualize a página.");
      if (qtd < 1 || qtd > 20) throw erro_(`Quantidade inválida para ${p.nome}.`);
      if (vistos[p.sku]) throw erro_(`${p.nome} ${p.tamanho} aparece duas vezes no pedido.`);
      vistos[p.sku] = 1;
      const disp = Math.max(0, disponivel_(p));
      if (qtd > disp) falta.push({ sku: p.sku, peca: p.nome, tamanho: String(p.tamanho), pedido: qtd, disponivel: disp });
      itens.push({ sku: p.sku, peca: p.nome, tamanho: String(p.tamanho), qtd, reservado: 0, entregue: 0 });
    });
    if (falta.length) return { sucesso: false, erro: "Estoque insuficiente para parte das peças.", falta };

    const protocolo = proximoProtocolo_(), agora = new Date();
    const foto = d.foto ? guardarImagem_(d.foto, protocolo + "_foto.jpg", 0) : "";
    const assinatura = guardarImagem_(d.assinatura, protocolo + "_assinatura.png", 45000);
    const linha = new Array(PED_HEAD.length).fill("");
    linha[PED.data] = agora; linha[PED.nome] = nome; linha[PED.cpf] = cpfFmt_(cpf); linha[PED.telefone] = txt_(d.telefone, 20);
    linha[PED.posto] = txt_(d.posto, 120); linha[PED.uniformes] = resumoItens_(itens); linha[PED.obs] = txt_(d.observacao, 1000);
    linha[PED.assinatura] = assinatura; linha[PED.status] = ST.AGUARD; linha[PED.protocolo] = protocolo; linha[PED.foto] = foto;
    linha[PED.cidade] = txt_(d.cidade, 80); linha[PED.funcao] = txt_(d.funcao, 80); linha[PED.itens] = JSON.stringify(itens);
    linha[PED.historico] = JSON.stringify([{ em: agora.toISOString(), por: nome, status: ST.AGUARD, evento: "Solicitação enviada" }]);
    linha[PED.atualizadoEm] = agora;
    abaPedidos_().appendRow(linha);
    const res = { sucesso: true, protocolo, data: agora.toISOString(), status: ST.AGUARD, itens };
    if (d.chave) cache.put(chave, JSON.stringify(res), 21600);
    return res;
  });
}
function cpfFmt_(c) { c = digitos_(c); return c.length === 11 ? c.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") : c; }

function identificar_(d) {
  const cpf = digitos_(d.cpf), prot = txt_(d.protocolo, 40).toUpperCase();
  const cache = CacheService.getScriptCache(), k = "cons_" + cpf;
  if (int_(cache.get(k)) >= 10) throw erro_("Muitas tentativas. Aguarde uma hora.");
  const lista = pedidos_().filter(p => digitos_(p.cpf) === cpf);
  if (!cpf || !lista.some(p => p.protocolo.toUpperCase() === prot)) {
    cache.put(k, String(int_(cache.get(k)) + 1), 3600);
    throw erro_("CPF e protocolo não conferem. Use o protocolo de qualquer pedido seu.");
  }
  return { cpf, lista };
}
function consultar_(d) {
  const { cpf, lista } = identificar_(d);
  const ent = tabela_("ENT").rows.filter(e => digitos_(e.cpf) === cpf);
  const dev = tabela_("DEV").rows.filter(e => digitos_(e.cpf) === cpf);
  const oco = tabela_("OCO").rows.filter(e => digitos_(e.cpf) === cpf);
  return {
    sucesso: true,
    pedidos: lista.map(p => ({ protocolo: p.protocolo, data: p.data, status: p.status, cidade: p.cidade, posto: p.posto,
      itens: p.itens, uniformes: p.uniformes, historico: p.historico.map(h => ({ em: h.em, status: h.status, evento: h.evento })) })),
    entregas: ent.map(e => ({ protocolo: e.protocolo, dataHora: iso_(e.dataHora), peca: e.peca, tamanho: String(e.tamanho),
      qtdEntregue: int_(e.qtdEntregue), pendente: pendenteEnt_(e), valorUnit: num_(e.valorUnit), responsavel: e.responsavel })),
    devolucoes: dev.map(e => ({ lote: e.lote, dataHora: iso_(e.dataHora), peca: e.peca, tamanho: String(e.tamanho),
      qtdBoa: int_(e.qtdBoa), qtdDesgaste: int_(e.qtdDesgaste), qtdDanificada: int_(e.qtdDanificada), qtdFaltante: int_(e.qtdFaltante),
      condicao: e.condicao, valorFaltante: num_(e.valorFaltante), valorDanoEstimado: num_(e.valorDanoEstimado), responsavel: e.responsavel })),
    ocorrencias: oco.map(o => ({ id: o.id, dataHora: iso_(o.dataHora), resumo: o.resumo, valorFaltante: num_(o.valorFaltante),
      valorDano: num_(o.valorDano), valorTotal: num_(o.valorTotal), status: o.status, contestacao: json_(o.contestacao, []) }))
  };
}
function contestar_(d) {
  const { cpf } = identificar_(d);
  const texto = txt_(d.texto, 1000);
  if (texto.length < 10) throw erro_("Descreva o motivo da contestação.");
  return comLock_(() => {
    const t = tabela_("OCO"), o = t.rows.find(x => x.id === d.ocorrenciaId && digitos_(x.cpf) === cpf);
    if (!o) throw erro_("Ocorrência não encontrada.");
    if (["Resolvido", "Cancelado após revisão"].indexOf(o.status) >= 0) throw erro_("Esta ocorrência já foi encerrada.");
    const c = json_(o.contestacao, []); c.push({ em: new Date().toISOString(), por: "Funcionário", texto });
    const h = json_(o.historico, []); h.push({ em: new Date().toISOString(), por: "Funcionário", status: "Contestação registrada", texto });
    o.contestacao = JSON.stringify(c); o.historico = JSON.stringify(h); o.status = "Contestação registrada";
    gravar_(t, o);
    return { sucesso: true };
  });
}

/* ───────────── painel ───────────── */
function painel_(u) {
  const movs = tabela_("MOV").rows;
  const dados = {
    usuario: u,
    pecas: tabela_("PECAS").rows.map(plano_),
    pedidos: pedidos_().map(p => { const o = Object.assign({}, p); delete o._row; return o; }),
    entregas: tabela_("ENT").rows.map(plano_),
    devolucoes: tabela_("DEV").rows.map(plano_),
    ocorrencias: tabela_("OCO").rows.map(o => Object.assign(plano_(o), { historico: json_(o.historico, []), contestacao: json_(o.contestacao, []) })),
    movimentacoes: movs.slice(-3000).map(plano_)
  };
  if (u.perfil === "admin") dados.usuarios = tabela_("USR").rows.map(x => ({ usuario: x.usuario, nome: x.nome, perfil: x.perfil, ativo: x.ativo }));
  return dados;
}
function anexos_(d) {
  const p = acharPedido_(d.protocolo), r = abaPedidos_().getRange(p._row, 1, 1, PED_HEAD.length).getValues()[0];
  const foto = String(r[PED.foto] || "");
  return { sucesso: true, assinatura: String(r[PED.assinatura] || ""), fotos: foto.startsWith("[") ? json_(foto, []) : (foto ? [foto] : []) };
}

/* ───────────── catálogo e saldo ───────────── */
function expandirTamanhos_(lista) {
  const out = [];
  (Array.isArray(lista) ? lista : []).forEach(x => {
    const s = txt_(x.tamanho, 20), m = s.match(/^(\d+)\s*-\s*(\d+)(?:\s*:\s*(\d+))?$/);
    if (m) {
      const passo = Math.max(1, int_(m[3]) || 1);
      if (int_(m[2]) < int_(m[1]) || (int_(m[2]) - int_(m[1])) / passo > 40) throw erro_("Faixa de tamanhos inválida: " + s);
      for (let n = int_(m[1]); n <= int_(m[2]); n += passo) out.push({ tamanho: String(n), qtd: int_(x.qtd) });
    }
    else if (s) out.push({ tamanho: norm_(s) === "unico" ? "Único" : s.toUpperCase(), qtd: int_(x.qtd) });
  });
  return out;
}
function cadastrarPecas_(d, u) {
  const nome = txt_(d.nome, 80), categoria = txt_(d.categoria, 40) || "Outros";
  if (!nome) throw erro_("Informe o nome da peça.");
  const grade = expandirTamanhos_(d.grade);
  if (!grade.length) throw erro_("Informe ao menos um tamanho (use Único para peças sem tamanho).");
  const valor = num_(d.valorUnit), minimo = d.minimo === "" || d.minimo == null ? MINIMO_PADRAO : int_(d.minimo);
  if (valor < 0 || minimo < 0) throw erro_("Valor e mínimo não podem ser negativos.");
  return comLock_(() => {
    const ctx = ctxEstoque_(u), criados = [];
    grade.forEach(g => {
      const sku = slug_(nome) + "-" + slug_(g.tamanho);
      if (ctx.pecas.rows.some(p => p.sku === sku)) throw erro_(`${nome} ${g.tamanho} já está cadastrada (${sku}). Use "Entrada" para somar peças.`);
      if (g.qtd < 0) throw erro_("Quantidade inicial não pode ser negativa.");
      const p = { sku, nome, categoria, tamanho: g.tamanho, saldo: 0, reservado: 0, quarentena: 0, minimo, valorUnit: valor,
        estado: txt_(d.estado, 20) || "Nova", obs: txt_(d.obs, 300), ativo: "SIM", alertaAtivo: "", qtdInicial: g.qtd,
        criadoEm: new Date(), atualizadoEm: new Date() };
      ctx.pecas.rows.push(p);
      if (g.qtd > 0) mover_(ctx, p, "ENTRADA_INICIAL", g.qtd, 0, 0, g.qtd, "Cadastro da peça", "");
      else checarAlerta_(ctx, p);
      criados.push(sku);
    });
    commitEstoque_(ctx);
    auditar_(u.usuario, "cadastrar_peca", { nome, criados });
    return { sucesso: true, criados };
  });
}
function editarPeca_(d, u) {
  return comLock_(() => {
    const ctx = ctxEstoque_(u), p = peca_(ctx, d.sku), antes = plano_(p);
    if (d.valorUnit !== undefined && num_(d.valorUnit) !== num_(p.valorUnit)) {
      if (u.perfil !== "admin") throw erro_("Só o administrador altera preço.");
      if (num_(d.valorUnit) < 0) throw erro_("Valor inválido.");
      p.valorUnit = num_(d.valorUnit);
    }
    if (d.categoria !== undefined) p.categoria = txt_(d.categoria, 40);
    if (d.minimo !== undefined) { if (int_(d.minimo) < 0) throw erro_("Mínimo inválido."); p.minimo = int_(d.minimo); }
    if (d.estado !== undefined) p.estado = txt_(d.estado, 20);
    if (d.obs !== undefined) p.obs = txt_(d.obs, 300);
    if (d.ativo !== undefined) {
      if (!d.ativo && int_(p.reservado) > 0) throw erro_("Esta peça tem reservas ativas. Entregue ou cancele antes de desativar.");
      p.ativo = d.ativo ? "SIM" : "";
    }
    p.atualizadoEm = new Date(); ctx.sujas[p.sku] = p;
    checarAlerta_(ctx, p);
    commitEstoque_(ctx);
    auditar_(u.usuario, "editar_peca", { sku: p.sku, antes, depois: plano_(p) });
    return { sucesso: true };
  });
}
function movimentarEstoque_(d, u) {
  const qtd = int_(d.qtd), motivo = txt_(d.motivo, 300);
  return comLock_(() => {
    const ctx = ctxEstoque_(u), p = peca_(ctx, d.sku);
    switch (d.tipo) {
      case "entrada":
        if (qtd < 1) throw erro_("Quantidade inválida.");
        mover_(ctx, p, "ENTRADA_COMPRA", qtd, 0, 0, qtd, motivo || "Compra", d.referencia); break;
      case "baixa_dano": case "baixa_extravio":
        if (qtd < 1) throw erro_("Quantidade inválida.");
        if (!motivo) throw erro_("Informe o motivo da baixa.");
        mover_(ctx, p, d.tipo === "baixa_dano" ? "BAIXA_DANO" : "BAIXA_EXTRAVIO", -qtd, 0, 0, qtd, motivo, ""); break;
      case "ajuste": {
        if (u.perfil !== "admin") throw erro_("Ajuste manual é exclusivo do administrador.");
        if (motivo.length < 10) throw erro_("Justifique o ajuste (mínimo 10 caracteres).");
        const novo = int_(d.novoSaldo);
        if (novo < 0) throw erro_("Saldo não pode ser negativo.");
        mover_(ctx, p, "AJUSTE_MANUAL", novo - int_(p.saldo), 0, 0, novo - int_(p.saldo), motivo, ""); break;
      }
      case "liberar_quarentena":
        if (qtd < 1) throw erro_("Quantidade inválida.");
        mover_(ctx, p, "QUARENTENA_LIBERADA", qtd, 0, -qtd, qtd, motivo || "Peça avaliada e liberada para uso", ""); break;
      case "inutilizar_quarentena":
        if (qtd < 1) throw erro_("Quantidade inválida.");
        if (!motivo) throw erro_("Informe o motivo do descarte.");
        mover_(ctx, p, "BAIXA_INUTILIZADA", 0, 0, -qtd, qtd, motivo, ""); break;
      default: throw erro_("Movimentação desconhecida.");
    }
    commitEstoque_(ctx);
    return { sucesso: true };
  });
}

/* ───────────── importação das peças do formulário ───────────── */
// Mesmas peças dos kits de posto do formulário do funcionário.
const KIT_PECAS = ["Apito","Baleiro","Blazer","Blusa cinza","Blusa de Frio","Blusa Preta","Blusa preta serviços","Blusa preta vigilância","Boné","Botina",
  "Botom preto vigilante","Calça elástico cinza","Calça jeans","Calça social","Calça Tática","Calça vigilante tática","Camisa ASG","Camisa porteiro",
  "Camisa Preta","Camisa social branca","Camisa social preta","Camisa vigilante bege curta","Camisa vigilante bege longa","Capa para placa balística",
  "Casaco Preto","Casaco preto serviços","Casaco preto vigilância","Casaco vigilância","Cinto","Cinto tático","Cobertura/Boné","Coldre","Coturno","Fiel",
  "Galocha","Gandola","Gandola Cinza","Gandola Preta","Gravata","Sapato social","Touca","Touca vegas vigilância"];

// Mesma regra de tamanhos que o formulário antigo usava.
function gradePadrao_(nome) {
  const n = norm_(nome), faixa = (a, b, p) => { const o = []; for (let i = a; i <= b; i += p) o.push(String(i)); return o; };
  if (/sapato|botina|coturno|galocha/.test(n)) return { cat: "Calçado", tam: faixa(34, 48, 1) };
  if (/camisa social/.test(n)) return { cat: "Camisa", tam: ["N2","N3","N4","N5","N6","N7","N8"] };
  if (/calca elastico/.test(n)) return { cat: "Calça", tam: ["P","M","G","GG","XG","EXG","EXGG"] };
  if (/calca/.test(n)) return { cat: "Calça", tam: faixa(36, 60, 2) };
  if (/camisa/.test(n)) return { cat: "Camisa", tam: ["P","M","G","GG","EG"] };
  if (/blusa|casaco|gandola|blazer/.test(n)) return { cat: "Agasalho", tam: ["P","M","G","GG","EG"] };
  return { cat: "Acessório", tam: ["Único"] };
}

// Cria saldo zerado para cada peça × tamanho que ainda não existe. Pode rodar de novo: não duplica.
// Peças importadas só entram nos alertas depois da primeira movimentação.
function importarPecasDosKits_(u) {
  return comLock_(() => {
    const t = tabela_("PECAS"), existe = {}, novas = [];
    t.rows.forEach(p => existe[p.sku] = 1);
    const vistos = {};
    KIT_PECAS.forEach(nome => {
      if (vistos[norm_(nome)]) return; vistos[norm_(nome)] = 1;
      const g = gradePadrao_(nome);
      g.tam.forEach(tam => {
        const sku = slug_(nome) + "-" + slug_(tam);
        if (existe[sku]) return; existe[sku] = 1;
        novas.push({ sku, nome, categoria: g.cat, tamanho: tam, saldo: 0, reservado: 0, quarentena: 0, minimo: MINIMO_PADRAO, valorUnit: 0,
          estado: "Nova", obs: "Importada do formulário de uniformes", ativo: "SIM", alertaAtivo: "", qtdInicial: 0, criadoEm: new Date(), atualizadoEm: new Date() });
      });
    });
    inserir_(t, novas);
    auditar_(u ? u.usuario : "editor", "importar_pecas_kits", novas.length);
    return { sucesso: true, criadas: novas.length };
  });
}
function importarPecasDosKits() { Logger.log(JSON.stringify(importarPecasDosKits_(null))); }

/* ───────────── fluxo do pedido ───────────── */
const PODE_IR = {};
PODE_IR[ST.APROV] = [ST.SEPAR, ST.DISP];
PODE_IR[ST.SEPAR] = [ST.DISP];
PODE_IR[ST.DISP] = [ST.SEPAR];
PODE_IR[ST.ENTREG] = [ST.DEVPEND];
PODE_IR[ST.ENTPARC] = [ST.DEVPEND];
PODE_IR[ST.DEVPARC] = [ST.DEVPEND];

function aprovar_(d, u) {
  return comLock_(() => {
    const p = acharPedido_(d.protocolo);
    if (p.legado) throw erro_("Pedido antigo sem itens estruturados: use apenas a mudança de status.");
    if (p.status !== ST.AGUARD) throw erro_("Só pedidos aguardando aprovação podem ser aprovados.");
    const ctx = ctxEstoque_(u), falta = [];
    p.itens.forEach(it => {
      const pc = ctx.pecas.rows.find(x => x.sku === it.sku);
      const need = it.qtd - int_(it.reservado) - int_(it.entregue);
      const disp = pc ? Math.max(0, disponivel_(pc)) : 0;
      if (disp < need) falta.push({ sku: it.sku, peca: it.peca, tamanho: it.tamanho, pedido: need, disponivel: disp });
    });
    if (falta.length && !d.parcial) return { sucesso: false, erro: "Estoque não cobre o pedido inteiro.", falta };
    let reservou = 0;
    p.itens.forEach(it => {
      const pc = ctx.pecas.rows.find(x => x.sku === it.sku); if (!pc) return;
      const take = Math.min(it.qtd - int_(it.reservado) - int_(it.entregue), Math.max(0, disponivel_(pc)));
      if (take > 0) { mover_(ctx, pc, "RESERVA", 0, take, 0, take, "Pedido aprovado", p.protocolo); it.reservado = int_(it.reservado) + take; reservou += take; }
    });
    if (!reservou) throw erro_("Nenhuma peça disponível para reservar. Registre entrada no estoque antes de aprovar.");
    commitEstoque_(ctx);
    p.status = ST.APROV;
    salvarPedido_(p, u.nome, falta.length ? "Aprovado com reserva parcial" : "Aprovado e reservado", "dataAprov");
    return { sucesso: true };
  });
}
function mudarStatus_(d, u) {
  return comLock_(() => {
    const p = acharPedido_(d.protocolo), alvo = d.status;
    if (p.legado) {
      if (Object.values(ST).concat(ST_LEGADO).indexOf(alvo) < 0) throw erro_("Status inválido.");
    } else if ((PODE_IR[p.status] || []).indexOf(alvo) < 0) {
      throw erro_(`Não é possível passar de "${p.status}" para "${alvo}".`);
    }
    p.status = alvo;
    salvarPedido_(p, u.nome, txt_(d.motivo, 300) || "Status alterado", alvo === ST.SEPAR ? "dataSepar" : null);
    return { sucesso: true };
  });
}
function cancelar_(d, u) {
  const motivo = txt_(d.motivo, 300);
  if (!motivo) throw erro_("Informe o motivo.");
  return comLock_(() => {
    const p = acharPedido_(d.protocolo);
    if (p.legado) { p.status = ST.CANC; salvarPedido_(p, u.nome, "Cancelado: " + motivo); return { sucesso: true }; }
    const encerrar = p.status === ST.ENTPARC;
    if (!encerrar && [ST.AGUARD, ST.APROV, ST.SEPAR, ST.DISP].indexOf(p.status) < 0) throw erro_("Este pedido não pode mais ser cancelado.");
    const ctx = ctxEstoque_(u);
    p.itens.forEach(it => {
      const r = int_(it.reservado);
      if (r > 0) { mover_(ctx, peca_(ctx, it.sku), "RESERVA_LIBERADA", 0, -r, 0, r, encerrar ? "Saldo do pedido encerrado" : "Pedido cancelado", p.protocolo); it.reservado = 0; }
      if (encerrar) it.qtd = int_(it.entregue);
    });
    commitEstoque_(ctx);
    p.status = encerrar ? ST.ENTREG : ST.CANC;
    salvarPedido_(p, u.nome, (encerrar ? "Restante encerrado: " : "Cancelado: ") + motivo);
    return { sucesso: true };
  });
}
function entregar_(d, u) {
  return comLock_(() => {
    const p = acharPedido_(d.protocolo);
    if (p.legado) throw erro_("Pedido antigo: registre a entrega pela mudança de status.");
    if ([ST.APROV, ST.SEPAR, ST.DISP, ST.ENTPARC].indexOf(p.status) < 0) throw erro_("Aprove o pedido antes de entregar.");
    const linhas = Array.isArray(d.itens) ? d.itens : [];
    const ctx = ctxEstoque_(u), lote = id_("ENT"), agora = new Date(), novas = [];
    linhas.forEach(l => {
      const it = p.itens[int_(l.idx)], q = int_(l.qtd);
      if (!it || q <= 0) return;
      if (!l.conferido) throw erro_(`Marque a conferência de ${it.peca} ${it.tamanho}.`);
      const pend = it.qtd - int_(it.entregue);
      if (q > pend) throw erro_(`${it.peca} ${it.tamanho}: entregando ${q}, mas faltam só ${pend}.`);
      const pc = peca_(ctx, it.sku), daReserva = Math.min(q, int_(it.reservado));
      if (q - daReserva > Math.max(0, disponivel_(pc))) throw erro_(`${it.peca} ${it.tamanho}: sem saldo livre para ${q - daReserva} peça(s) além da reserva.`);
      mover_(ctx, pc, "SAIDA_ENTREGA", -q, -daReserva, 0, q, "Entrega ao funcionário", p.protocolo);
      it.reservado = int_(it.reservado) - daReserva; it.entregue = int_(it.entregue) + q;
      novas.push({ id: id_("IT"), lote, dataHora: agora, protocolo: p.protocolo, cpf: p.cpf, nome: p.nome, cidade: p.cidade, posto: p.posto,
        sku: it.sku, peca: it.peca, tamanho: it.tamanho, qtdSolicitada: it.qtd, qtdEntregue: q, estadoEntrega: txt_(l.estado, 20) || "Nova",
        valorUnit: num_(pc.valorUnit), responsavel: u.nome, obs: txt_(l.obs, 300), qtdDevolvida: 0, qtdDanificada: 0, qtdFaltante: 0 });
    });
    if (!novas.length) throw erro_("Informe a quantidade entregue de ao menos uma peça.");
    commitEstoque_(ctx);
    inserir_(tabela_("ENT"), novas);
    p.status = p.itens.every(i => int_(i.entregue) >= i.qtd) ? ST.ENTREG : ST.ENTPARC;
    salvarPedido_(p, u.nome, `Entrega ${lote}` + (txt_(d.obs) ? ": " + txt_(d.obs, 300) : ""), "dataEntrega");
    return { sucesso: true, lote, status: p.status, itens: novas.map(plano_) };
  });
}

/* ───────────── devolução ───────────── */
function pendenteEnt_(e) { return int_(e.qtdEntregue) - int_(e.qtdDevolvida) - int_(e.qtdDanificada) - int_(e.qtdFaltante); }
function devolver_(d, u) {
  const cpf = digitos_(d.cpf);
  const linhas = (Array.isArray(d.itens) ? d.itens : []).filter(l => int_(l.boa) + int_(l.desgaste) + int_(l.danificada) + int_(l.faltante) > 0);
  if (!linhas.length) throw erro_("Informe as quantidades conferidas de ao menos uma peça.");
  return comLock_(() => {
    const tEnt = tabela_("ENT"), ctx = ctxEstoque_(u), lote = id_("DEV"), agora = new Date(), novas = [];
    const foto = d.foto ? guardarImagem_(d.foto, lote + "_foto.jpg", 0) : "";
    let vFalt = 0, vDano = 0, nome = "";
    const protocolos = {};
    linhas.forEach(l => {
      const e = tEnt.rows.find(x => x.id === l.entregaId);
      if (!e || digitos_(e.cpf) !== cpf) throw erro_("Item de entrega não encontrado para este funcionário.");
      const boa = int_(l.boa), des = int_(l.desgaste), dan = int_(l.danificada), fal = int_(l.faltante);
      if ([boa, des, dan, fal].some(n => n < 0)) throw erro_("Quantidades não podem ser negativas.");
      const pend = pendenteEnt_(e), soma = boa + des + dan + fal;
      if (soma > pend) throw erro_(`${e.peca} ${e.tamanho}: conferindo ${soma}, mas só ${pend} estão pendentes. A peça não pode ser devolvida duas vezes.`);
      const valorDano = dan > 0 ? Math.max(0, num_(l.valorDano)) : 0, motivoDano = txt_(l.motivoDano, 300);
      if (valorDano > 0 && !motivoDano) throw erro_(`${e.peca} ${e.tamanho}: descreva o dano para registrar o valor estimado.`);
      const pc = ctx.pecas.rows.find(x => x.sku === e.sku);
      if (pc) {
        if (boa + des > 0) mover_(ctx, pc, "ENTRADA_DEVOLUCAO", boa + des, 0, 0, boa + des, des ? `${boa} boas, ${des} com desgaste` : "Devolução em boas condições", e.protocolo);
        if (dan > 0) mover_(ctx, pc, "QUARENTENA_DEVOLUCAO", 0, 0, dan, dan, motivoDano || "Devolvida danificada", e.protocolo);
      }
      if (fal > 0) ctx.movs.push({ id: id_("MOV"), dataHora: agora, sku: e.sku, nome: e.peca, tamanho: e.tamanho, tipo: "NAO_DEVOLVIDA", qtd: fal,
        saldoAntes: pc ? int_(pc.saldo) : "", saldoDepois: pc ? int_(pc.saldo) : "", reservadoAntes: pc ? int_(pc.reservado) : "",
        reservadoDepois: pc ? int_(pc.reservado) : "", quarentenaDepois: pc ? int_(pc.quarentena) : "", responsavel: u.nome,
        motivo: `Não devolvida por ${e.nome}`, referencia: e.protocolo });
      const valFal = fal * num_(e.valorUnit);
      vFalt += valFal; vDano += valorDano; nome = e.nome; protocolos[e.protocolo] = 1;
      const cond = fal === soma ? "Não devolvida" : soma < pend ? "Devolução parcial" : dan ? "Devolvida danificada" : des ? "Devolvida com desgaste normal" : "Devolvida em boas condições";
      novas.push({ id: id_("DV"), lote, dataHora: agora, entregaId: e.id, protocolo: e.protocolo, cpf: e.cpf, nome: e.nome, sku: e.sku,
        peca: e.peca, tamanho: e.tamanho, pendenteAntes: pend, qtdBoa: boa, qtdDesgaste: des, qtdDanificada: dan, qtdFaltante: fal,
        condicao: cond, valorUnitEntrega: num_(e.valorUnit), valorFaltante: valFal, valorDanoEstimado: valorDano, motivoDano,
        responsavel: u.nome, obs: txt_(l.obs, 300), foto });
      e.qtdDevolvida = int_(e.qtdDevolvida) + boa + des; e.qtdDanificada = int_(e.qtdDanificada) + dan; e.qtdFaltante = int_(e.qtdFaltante) + fal;
      gravar_(tEnt, e);
    });
    commitEstoque_(ctx);
    inserir_(tabela_("DEV"), novas);

    let ocorrencia = null;
    if (vFalt + vDano > 0) {
      ocorrencia = { id: id_("OC"), dataHora: agora, lote, cpf: cpfFmt_(cpf), nome,
        resumo: novas.filter(n => n.qtdFaltante || n.valorDanoEstimado).map(n => `${n.peca} ${n.tamanho}: ${n.qtdFaltante ? n.qtdFaltante + " faltante(s)" : ""}${n.qtdFaltante && n.valorDanoEstimado ? ", " : ""}${n.valorDanoEstimado ? "dano estimado" : ""}`).join("; "),
        valorFaltante: vFalt, valorDano: vDano, valorTotal: vFalt + vDano, status: "Em análise", calculadoPor: u.nome, revisadoPor: "", aprovadoPor: "",
        historico: JSON.stringify([{ em: agora.toISOString(), por: u.nome, status: "Em análise", texto: "Valor de referência calculado na conferência da devolução" }]),
        contestacao: "[]" };
      inserir_(tabela_("OCO"), [ocorrencia]);
    }

    const pend = {};
    tEnt.rows.forEach(e => { if (protocolos[e.protocolo]) pend[e.protocolo] = (pend[e.protocolo] || 0) + pendenteEnt_(e); });
    pedidos_().filter(p => protocolos[p.protocolo] && [ST.ENTREG, ST.ENTPARC, ST.DEVPEND, ST.DEVPARC].indexOf(p.status) >= 0).forEach(p => {
      const novo = pend[p.protocolo] > 0 ? ST.DEVPARC : ST.DEVOK;
      if (novo !== p.status) { p.status = novo; salvarPedido_(p, u.nome, `Devolução ${lote}`); }
    });
    return { sucesso: true, lote, itens: novas.map(plano_), ocorrencia: ocorrencia ? plano_(ocorrencia) : null };
  });
}

/* ───────────── ocorrências financeiras ───────────── */
function ocorrencia_(d, u) {
  const alvo = d.status, just = txt_(d.justificativa, 1000);
  if (FIN.indexOf(alvo) < 0) throw erro_("Status financeiro inválido.");
  const soAdmin = ["Aprovado administrativamente", "Resolvido", "Cancelado após revisão"];
  if (soAdmin.indexOf(alvo) >= 0 && u.perfil !== "admin") throw erro_("Só o administrador aprova, resolve ou cancela ocorrências.");
  if (alvo !== "Aguardando aprovação" && just.length < 10) throw erro_("Registre a justificativa (mínimo 10 caracteres).");
  return comLock_(() => {
    const t = tabela_("OCO"), o = t.rows.find(x => x.id === d.id);
    if (!o) throw erro_("Ocorrência não encontrada.");
    if (["Resolvido", "Cancelado após revisão"].indexOf(o.status) >= 0) throw erro_("Ocorrência encerrada.");
    const origem = { "Aguardando aprovação": ["Em análise", "Contestação registrada"], "Aprovado administrativamente": ["Aguardando aprovação"], "Resolvido": ["Aprovado administrativamente"] }[alvo];
    if (origem && alvo !== o.status && origem.indexOf(o.status) < 0) throw erro_(`Não é possível passar de "${o.status}" para "${alvo}".`);
    if (d.valorDano !== undefined && d.valorDano !== "") {
      if (["Em análise", "Contestação registrada"].indexOf(o.status) < 0) throw erro_("O valor só pode ser revisto em análise ou contestação.");
      o.valorDano = Math.max(0, num_(d.valorDano)); o.valorTotal = num_(o.valorFaltante) + o.valorDano;
    }
    const h = json_(o.historico, []);
    h.push({ em: new Date().toISOString(), por: u.nome, status: alvo, texto: just, valorTotal: num_(o.valorTotal) });
    o.historico = JSON.stringify(h); o.status = alvo;
    if (alvo === "Aguardando aprovação" || alvo === "Contestação registrada") o.revisadoPor = u.nome;
    if (soAdmin.indexOf(alvo) >= 0) o.aprovadoPor = u.nome;
    if (alvo === "Contestação registrada") { const c = json_(o.contestacao, []); c.push({ em: new Date().toISOString(), por: u.nome, texto: just }); o.contestacao = JSON.stringify(c); }
    gravar_(t, o);
    auditar_(u.usuario, "ocorrencia", { id: o.id, status: alvo });
    return { sucesso: true };
  });
}

/* ───────────── usuários (admin) ───────────── */
function usuarios_(d, u) {
  return comLock_(() => {
    if (d.op === "criar") { criarUsuario_(d.usuario, d.senha, d.nome, d.perfil); auditar_(u.usuario, "criar_usuario", d.usuario); return { sucesso: true }; }
    const t = tabela_("USR"), x = t.rows.find(r => r.usuario === norm_(d.usuario));
    if (!x) throw erro_("Usuário não encontrado.");
    if (x.usuario === u.usuario && (d.op === "desativar" || d.perfil === "estoque")) throw erro_("Você não pode remover o próprio acesso de administrador.");
    if (d.op === "senha") { if (String(d.senha || "").length < 8) throw erro_("Senha com pelo menos 8 caracteres."); x.salt = Utilities.getUuid(); x.hash = hash_(x.salt, d.senha); }
    else if (d.op === "desativar") x.ativo = "";
    else if (d.op === "ativar") x.ativo = "SIM";
    else if (d.op === "perfil") { if (["estoque", "admin"].indexOf(d.perfil) < 0) throw erro_("Perfil inválido."); x.perfil = d.perfil; }
    else throw erro_("Operação inválida.");
    gravar_(t, x);
    auditar_(u.usuario, "usuario_" + d.op, x.usuario);
    return { sucesso: true };
  });
}

/* ───────────── materiais (Página2, fluxo antigo) ───────────── */
function abaMateriais_() { return ss_().getSheetByName(ABA.MATERIAIS) || ss_().getSheets()[1]; }
function materiais_() {
  const sh = abaMateriais_(); if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues(); v.shift();
  return v.map((l, i) => ({ id: i + 2, data: l[0], nome: l[1], cidade: l[2], posto: l[3], material: l[4], quantidade: l[5],
    observacao: l[6], assinatura: l[7], status: l[8], foto: l[9], protocolo: l[10],
    dataPendente: l[11] || "", dataSeparado: l[12] || "", dataEntrega: l[13] || "" }));
}
function statusMaterial_(d, u) {
  const sh = abaMateriais_(), linha = int_(d.id);
  if (linha < 2 || ST_LEGADO.indexOf(d.status) < 0) throw erro_("Dados inválidos.");
  sh.getRange(linha, 9).setValue(d.status);
  const col = { PENDENTE: 12, SEPARADO: 13, ENTREGUE: 14 }[d.status];
  if (col) sh.getRange(linha, col).setValue(new Date());
  auditar_(u.usuario, "status_material", { linha, status: d.status });
  return { sucesso: true };
}
function apagarMateriaisEntregues_(u) {
  if (u.perfil !== "admin") throw erro_("Só o administrador apaga registros.");
  return comLock_(() => {
    const sh = abaMateriais_(), v = sh.getDataRange().getValues(); let n = 0;
    for (let i = v.length - 1; i >= 1; i--) if (v[i][8] === "ENTREGUE") { sh.deleteRow(i + 1); n++; }
    auditar_(u.usuario, "apagar_materiais_entregues", n);
    return { sucesso: true, apagados: n };
  });
}

/* ───────────── roteamento ───────────── */
const ACOES = {
  catalogo:      { fn: () => catalogo_() },
  novoPedido:    { fn: d => novoPedido_(d) },
  consultar:     { fn: d => consultar_(d) },
  contestar:     { fn: d => contestar_(d) },
  login:         { fn: d => login_(d) },
  logout:        { fn: d => { if (d.token) CacheService.getScriptCache().remove("tok_" + d.token); return { sucesso: true }; } },
  painel:        { auth: true, fn: (d, u) => ({ sucesso: true, dados: painel_(u) }) },
  anexos:        { auth: true, fn: d => anexos_(d) },
  cadastrarPecas:{ auth: true, mut: true, fn: cadastrarPecas_ },
  editarPeca:    { auth: true, mut: true, fn: editarPeca_ },
  importarKits:  { auth: true, mut: true, admin: true, fn: (d, u) => importarPecasDosKits_(u) },
  estoque:       { auth: true, mut: true, fn: movimentarEstoque_ },
  aprovar:       { auth: true, mut: true, fn: aprovar_ },
  status:        { auth: true, mut: true, fn: mudarStatus_ },
  cancelar:      { auth: true, mut: true, fn: cancelar_ },
  entregar:      { auth: true, mut: true, fn: entregar_ },
  devolver:      { auth: true, mut: true, fn: devolver_ },
  ocorrencia:    { auth: true, mut: true, fn: ocorrencia_ },
  usuarios:      { auth: true, mut: true, admin: true, fn: usuarios_ },
  materiais:     { auth: true, fn: () => ({ sucesso: true, itens: materiais_() }) },
  statusMaterial:{ auth: true, fn: statusMaterial_ },
  apagarMateriaisEntregues: { auth: true, fn: (d, u) => apagarMateriaisEntregues_(u) }
};

function executar_(d) {
  const a = ACOES[d.acao];
  if (!a) return { sucesso: false, erro: "Ação desconhecida." };
  const u = a.auth ? sessao_(d.token, a.admin ? ["admin"] : null) : null;
  const r = a.fn(d, u);
  if (a.mut && r && r.sucesso) r.dados = painel_(u);
  return r;
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.func === "salvarUniforme") return resposta(salvarUniformeLegado_(d));
    if (d.func === "salvarMaterial") return resposta(salvarMaterialLegado_(d));
    return resposta(executar_(d));
  } catch (err) {
    return resposta(Object.assign({ sucesso: false, erro: err.message || String(err) }, err.auth ? { auth: true } : {}, err.protocoloAberto ? { protocoloAberto: err.protocoloAberto } : {}));
  }
}

// Leitura e mudança de status por GET ficaram fechadas: expunham CPF/telefone e
// permitiam apagar registros sem senha. O painel novo usa POST com sessão.
function doGet() { return resposta({ sucesso: true, servico: "Vegas Uniformes v10" }); }

/* ───────────── formulários antigos (transição) ───────────── */
function salvarUniformeLegado_(d) {
  const foto = Array.isArray(d.fotos) && d.fotos.length ? guardarImagem_(d.fotos[0], (d.protocolo || "legado") + "_foto.jpg", 0) : "";
  const linha = new Array(PED_HEAD.length).fill("");
  linha[PED.data] = new Date(); linha[PED.nome] = d.nome || ""; linha[PED.cpf] = d.cpf || ""; linha[PED.telefone] = d.telefone || "";
  linha[PED.posto] = d.posto || ""; linha[PED.uniformes] = d.uniformes || ""; linha[PED.obs] = d.observacao || "";
  linha[PED.assinatura] = guardarImagem_(d.assinatura, (d.protocolo || "legado") + "_assinatura.png", 45000);
  linha[PED.status] = "EM ANÁLISE"; linha[PED.protocolo] = d.protocolo || ""; linha[PED.foto] = foto; linha[PED.cidade] = d.cidade || "";
  abaPedidos_().appendRow(linha);
  return { sucesso: true };
}
function salvarMaterialLegado_(d) {
  abaMateriais_().appendRow([new Date(), d.nome || "", d.cidade || "", d.posto || "", d.material || "", d.quantidade || "",
    d.obs || "", d.assinatura || "", "EM ANÁLISE", d.foto || "", d.protocolo || ""]);
  return { sucesso: true };
}
