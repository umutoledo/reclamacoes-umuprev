// ============================================================
// Umuprev — Sistema de Reclamações — Backend (Google Apps Script)
// VERSÃO 2 — com login por usuário, cada supervisor só vê o que é
// dele, painel completo para a Francielle, e alertas por email
// (reclamação crítica nova + prazo vencido).
//
// Cole este código inteiro no editor de Apps Script da sua planilha
// (Extensões > Apps Script), apagando o que já estiver lá antes.
// Depois de colar, veja no final deste arquivo os 2 passos extras
// que você precisa fazer (autorizar e ligar o alerta diário).
// ============================================================

var SHEET_NAME = "Reclamações";
var USERS_SHEET_NAME = "Usuarios";

// Para onde vão os alertas por email:
var ALERT_EMAIL = "umuprev.toledo@umuprev.com.br";

var HEADERS = [
  "id", "numero", "createdAt", "contrato", "nome", "canal", "setor",
  "categoria", "descricao", "gravidade", "supervisorResponsavel", "status",
  "prazo", "responsavelTratativa", "observacoes", "contatoAssociado",
  "resolvedAt", "resultado", "reincidencia", "timelineJson", "updatedAt",
  "atualizadoPor", "registradoPor"
];

var USER_HEADERS = ["nome", "senha", "papel"];
// papel: "admin" (vê tudo + painel gerencial) ou "supervisor" (só vê o que é dele)
var DEFAULT_USERS = [
  ["Francielle", "", "admin"],
  ["Deysi", "", "supervisor"],
  ["Alanah", "", "supervisor"],
  ["Maria Eduarda", "", "supervisor"],
  ["Rodrigo", "", "supervisor"],
  ["Thiago", "", "supervisor"]
];

// ---------- planilha de reclamações ----------

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // migração automática: adiciona no final qualquer coluna nova
    // (ex.: "registradoPor") que ainda não exista na planilha
    var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    HEADERS.forEach(function (h) {
      if (existing.indexOf(h) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      }
    });
  }
  return sheet;
}

function readAll_() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  return rows
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) {
        obj[h] = row[i];
      });
      try {
        obj.timeline = obj.timelineJson ? JSON.parse(obj.timelineJson) : [];
      } catch (err) {
        obj.timeline = [];
      }
      delete obj.timelineJson;
      return obj;
    })
    .filter(function (o) {
      return o.id;
    });
}

// ---------- planilha de usuários / login ----------

function getUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(USER_HEADERS);
    DEFAULT_USERS.forEach(function (u) {
      sheet.appendRow(u);
    });
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findUser_(nome) {
  var sheet = getUsersSheet_();
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim().toLowerCase() === String(nome).trim().toLowerCase()) {
      return { nome: data[r][0], senha: data[r][1], papel: data[r][2] };
    }
  }
  return null;
}

function doLogin_(nome, senha) {
  var u = findUser_(nome);
  if (!u) return { ok: false, error: "usuario-nao-encontrado" };
  if (!u.senha) return { ok: false, error: "senha-nao-configurada" };
  if (String(u.senha) !== String(senha)) return { ok: false, error: "senha-incorreta" };
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put("tok_" + token, JSON.stringify({ nome: u.nome, papel: u.papel }), 21600); // 6h (máximo do CacheService)
  return { ok: true, token: token, nome: u.nome, papel: u.papel };
}

function getSession_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var raw = cache.get("tok_" + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function filterForUser_(list, session) {
  if (session.papel === "admin") return list;
  return list.filter(function (c) {
    return c.supervisorResponsavel === session.nome || c.registradoPor === session.nome;
  });
}

// ---------- helpers HTTP ----------

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function doGet(e) {
  var token = e.parameter && e.parameter.token;
  var session = getSession_(token);
  if (!session) return jsonOut_({ ok: false, error: "sessao-invalida" });
  var list = readAll_();
  return jsonOut_({ ok: true, nome: session.nome, papel: session.papel, data: filterForUser_(list, session) });
}

function doPost(e) {
  var payload = JSON.parse(e.postData.contents);

  if (payload.action === "login") {
    return jsonOut_(doLogin_(payload.nome, payload.senha));
  }

  var session = getSession_(payload.token);
  if (!session) return jsonOut_({ ok: false, error: "sessao-invalida" });

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getSheet_();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = headers.indexOf("id");

    function podeMexer(rowObj) {
      if (session.papel === "admin") return true;
      return rowObj.supervisorResponsavel === session.nome || rowObj.registradoPor === session.nome;
    }

    if (payload.action === "delete") {
      for (var r = 1; r < data.length; r++) {
        if (data[r][idCol] === payload.id) {
          var rowObj = {};
          headers.forEach(function (h, i) { rowObj[h] = data[r][i]; });
          if (!podeMexer(rowObj)) return jsonOut_({ ok: false, error: "sem-permissao" });
          sheet.deleteRow(r + 1);
          break;
        }
      }
      return jsonOut_({ ok: true });
    }

    // action === "save" (padrão): insere nova ou atualiza existente por id
    var record = payload.record || {};

    var foundRow = -1;
    var existing = null;
    for (var r2 = 1; r2 < data.length; r2++) {
      if (data[r2][idCol] === record.id) {
        foundRow = r2 + 1;
        existing = {};
        headers.forEach(function (h, i) { existing[h] = data[r2][i]; });
        break;
      }
    }

    var isNew = foundRow < 0;
    if (!isNew && !podeMexer(existing)) {
      return jsonOut_({ ok: false, error: "sem-permissao" });
    }

    if (isNew) {
      if (!record.numero) {
        var countAtual = sheet.getLastRow() - 1; // menos a linha de cabeçalho
        record.numero = "REC-" + ("0000" + (countAtual + 1)).slice(-4);
      }
      record.registradoPor = session.nome; // quem registrou é sempre definido pelo servidor
    } else {
      record.registradoPor = existing.registradoPor || record.registradoPor || "";
    }

    record.timelineJson = JSON.stringify(record.timeline || []);
    delete record.timeline;

    var rowValues = headers.map(function (h) {
      return record[h] !== undefined && record[h] !== null ? record[h] : "";
    });

    if (foundRow > 0) {
      sheet.getRange(foundRow, 1, 1, headers.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    if (isNew && record.gravidade === "Crítica") {
      enviarAlertaCritica_(record);
    }

    return jsonOut_({ ok: true, numero: record.numero });
  } finally {
    lock.releaseLock();
  }
}

// ---------- alertas por email ----------

function enviarAlertaCritica_(record) {
  try {
    var assunto = "🚨 Reclamação CRÍTICA registrada — " + (record.numero || "");
    var corpo =
      "Uma nova reclamação crítica foi registrada no Sistema de Reclamações Umuprev.\n\n" +
      "Número: " + (record.numero || "") + "\n" +
      "Setor: " + (record.setor || "") + "\n" +
      "Categoria: " + (record.categoria || "") + "\n" +
      "Descrição: " + (record.descricao || "") + "\n" +
      "Supervisor responsável: " + (record.supervisorResponsavel || "") + "\n" +
      "Registrado por: " + (record.registradoPor || "") + "\n" +
      "Prazo: " + (record.prazo || "não definido") + "\n\n" +
      "Acesse o sistema: https://umutoledo.github.io/reclamacoes-umuprev/";
    MailApp.sendEmail(ALERT_EMAIL, assunto, corpo);
  } catch (err) {
    // não deixa um problema no envio de email travar o salvamento da reclamação
  }
}

// Roda 1x por dia (veja o passo de configuração no final deste arquivo).
// Manda um resumo por email de todas as reclamações em aberto com prazo vencido.
function verificarPrazosVencidos() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  var vencidas = [];
  for (var r = 1; r < data.length; r++) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = data[r][i]; });
    if (!obj.id) continue;
    if (obj.status === "Resolvida") continue;
    if (!obj.prazo) continue;
    var prazoDate = new Date(obj.prazo);
    if (isNaN(prazoDate)) continue;
    prazoDate.setHours(0, 0, 0, 0);
    if (prazoDate.getTime() < hoje.getTime()) vencidas.push(obj);
  }
  if (vencidas.length === 0) return;
  var corpo =
    "As reclamações abaixo estão com o prazo de resolução vencido:\n\n" +
    vencidas
      .map(function (c) {
        return "• " + c.numero + " — " + c.setor + " — prazo " + c.prazo + " — supervisor " + c.supervisorResponsavel;
      })
      .join("\n") +
    "\n\nAcesse o sistema: https://umutoledo.github.io/reclamacoes-umuprev/";
  MailApp.sendEmail(ALERT_EMAIL, "⏰ " + vencidas.length + " reclamação(ões) com prazo vencido", corpo);
}

// Rode esta função UMA VEZ (veja instruções) para ligar o alerta diário de prazos vencidos.
function configurarAlertaDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "verificarPrazosVencidos") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("verificarPrazosVencidos").timeBased().everyDays(1).atHour(8).create();
}

// ============================================================
// DEPOIS DE COLAR ESTE CÓDIGO E SALVAR (ícone de disquete):
//
// 1) Rode a função "configurarAlertaDiario" uma vez, para ligar o
//    alerta diário de prazo vencido:
//    - No topo do editor, no menu que mostra o nome de uma função,
//      selecione "configurarAlertaDiario".
//    - Clique em "Executar" (▷).
//    - Vai pedir autorização de novo (agora para enviar email e
//      criar o gatilho) — autorize com sua conta Google, do mesmo
//      jeito que fez da primeira vez.
//
// 2) Faça um novo deploy para publicar esta versão nova do código:
//    - Clique em "Implantar" > "Gerenciar implantações".
//    - Clique no ícone de lápis (editar) na implantação existente.
//    - Em "Versão", selecione "Nova versão".
//    - Clique em "Implantar".
//    - A URL do site continua a mesma de antes — não precisa mudar
//      nada no GitHub.
//
// 3) Abra a planilha, vá na aba nova "Usuarios" (foi criada
//    automaticamente) e digite uma senha para cada pessoa na coluna
//    "senha". Depois é só avisar cada supervisor da senha dela.
// ============================================================
