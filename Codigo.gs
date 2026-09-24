/* ═══════════════════════════════════════════════════
   APPS SCRIPT – VEGAS VIGILÂNCIA  v9
   v8: apagarEntregues — remove da planilha os pedidos
       com status ENTREGUE.
   v9: Uniformes passam a gravar a CIDADE na coluna O (15).
       As colunas 12, 13 e 14 (datas de status) continuam
       no mesmo lugar, então os pedidos antigos não mudam.
═══════════════════════════════════════════════════ */

function getAba(tipo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nome = tipo === "materiais" ? "Página2" : "Página1";
  let sheet = ss.getSheetByName(nome);
  if (sheet) return sheet;
  sheet = ss.getSheets()[tipo === "materiais" ? 1 : 0];
  if (sheet) return sheet;
  const nomes = ss.getSheets().map(s => s.getName()).join(" | ");
  throw new Error("Aba não encontrada. Disponíveis: " + nomes);
}

/* ── doPost ── */
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);

    /* ── SALVAR UNIFORME ── */
    if (dados.func === "salvarUniforme") {
      const sheet = getAba("uniformes");

      const foto = Array.isArray(dados.fotos) && dados.fotos.length > 0
        ? dados.fotos[0]
        : "";

      sheet.appendRow([
        new Date(),
        dados.nome       || "",
        dados.cpf        || "",
        dados.telefone   || "",
        dados.posto      || "",
        dados.uniformes  || "",
        dados.observacao || "",
        dados.assinatura || "",
        "EM ANÁLISE",
        dados.protocolo  || "",
        foto,             // coluna K (11): string base64 direta (não JSON)
        "",               // coluna L (12): dataPendente — preenchida na mudança de status
        "",               // coluna M (13): dataSeparado
        "",               // coluna N (14): dataEntrega
        dados.cidade     || ""  // coluna O (15): cidade
      ]);

      return resposta({ sucesso: true });
    }

    /* ── SALVAR MATERIAL ── */
    if (dados.func === "salvarMaterial") {
      const sheet = getAba("materiais");
      sheet.appendRow([
        new Date(),
        dados.nome       || "",
        dados.cidade     || "",
        dados.posto      || "",
        dados.material   || "",
        dados.quantidade || "",
        dados.obs        || "",
        dados.assinatura || "",
        "EM ANÁLISE",
        dados.foto       || "",
        dados.protocolo  || ""
      ]);
      return resposta({ sucesso: true });
    }

    return resposta({ sucesso: false, mensagem: "Função inválida" });
  } catch (err) {
    return resposta({ sucesso: false, erro: err.toString() });
  }
}

/* ── doGet ── */
function doGet(e) {
  try {
    const tipo = e.parameter.tipo || "uniformes";

    if (e.parameter.func === "atualizarStatus") {
      const sheet  = getAba(tipo);
      const linha  = Number(e.parameter.id);
      const status = decodeURIComponent(e.parameter.status || "");

      if (!linha || linha < 2)
        return resposta({ sucesso: false, erro: "ID inválido: " + e.parameter.id });

      const validos = ["EM ANÁLISE","PENDENTE","SEPARADO","ENTREGUE"];
      if (!validos.includes(status))
        return resposta({ sucesso: false, erro: "Status inválido: " + status });

      // Coluna 9 = status
      sheet.getRange(linha, 9).setValue(status);

      // Registra data/hora conforme o novo status (colunas fixas: 12, 13, 14)
      const agora = new Date();
      if (status === "PENDENTE") {
        sheet.getRange(linha, 12).setValue(agora); // dataPendente
      } else if (status === "SEPARADO") {
        sheet.getRange(linha, 13).setValue(agora); // dataSeparado
      } else if (status === "ENTREGUE") {
        sheet.getRange(linha, 14).setValue(agora); // dataEntrega
      }

      return resposta({ sucesso: true });
    }

    if (e.parameter.func === "apagarEntregues") {
      const apagados = apagarEntregues(tipo);
      return resposta({ sucesso: true, apagados: apagados });
    }

    return ContentService
      .createTextOutput(JSON.stringify(obterDados(tipo)))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return resposta({ sucesso: false, erro: err.toString() });
  }
}

/* ── apagarEntregues ──
   Remove da planilha todas as linhas cujo status (coluna 9) seja ENTREGUE.
   Percorre de baixo para cima para não bagunçar os índices ao deletar.
── */
function apagarEntregues(tipo) {
  const sheet = getAba(tipo);
  const dados = sheet.getDataRange().getValues();
  let apagados = 0;

  for (let i = dados.length - 1; i >= 1; i--) { // i=0 é o cabeçalho, pula
    if (dados[i][8] === "ENTREGUE") { // coluna 9 (índice 8) = status
      sheet.deleteRow(i + 1); // linha real na planilha
      apagados++;
    }
  }
  return apagados;
}

/* ── obterDados ── */
function obterDados(tipo) {
  const sheet = getAba(tipo);
  const dados = sheet.getDataRange().getDisplayValues();
  dados.shift();

  return dados.map((linha, i) => {

    // Datas de status — mesma posição para os dois tipos (colunas 12, 13, 14)
    const dataPendente = linha[11] || "";
    const dataSeparado = linha[12] || "";
    const dataEntrega  = linha[13] || "";

    if (tipo === "uniformes") {
      const fotoRaw = linha[10] || "";
      let fotos = [];
      if (fotoRaw.startsWith("data:") || fotoRaw.startsWith("http")) {
        fotos = [fotoRaw];
      } else if (fotoRaw.startsWith("[")) {
        try { fotos = JSON.parse(fotoRaw); } catch { fotos = []; }
      }
      fotos = fotos.filter(f => typeof f === "string" && f.length > 10);

      return {
        id: i + 2,
        data: linha[0], nome: linha[1], cpf: linha[2], telefone: linha[3],
        posto: linha[4], uniformes: linha[5], observacao: linha[6],
        assinatura: linha[7], status: linha[8], protocolo: linha[9],
        fotos: fotos,
        dataPendente: dataPendente,
        dataSeparado: dataSeparado,
        dataEntrega: dataEntrega,
        cidade: linha[14] || ""   // coluna O (15) — vazia nos pedidos antigos
      };
    }

    // Materiais
    return {
      id: i + 2,
      data: linha[0], nome: linha[1], cidade: linha[2], posto: linha[3],
      material: linha[4], quantidade: linha[5], observacao: linha[6],
      assinatura: linha[7], status: linha[8], foto: linha[9],
      protocolo: linha[10],
      dataPendente: dataPendente,
      dataSeparado: dataSeparado,
      dataEntrega: dataEntrega
    };
  });
}

function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
