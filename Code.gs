/**
 * Plataforma COB — Conectando Orquestras Brasileiras
 * Voluntário(a) / Embaixador(a) / Representante de Orquestra
 * Backend: Google Apps Script + Google Sheets + Google Calendar + Gmail
 *
 * Opus AI — revisão estrutural 23/08/2026
 * ------------------------------------------------------------------
 * MUDANÇAS DESTA VERSÃO
 *  [SEG] Sessão administrativa validada no servidor (antes: qualquer
 *        pessoa com a URL podia chamar getDashboardData e ler todos
 *        os dados pessoais — risco LGPD).
 *  [SEG] Link de aprovação por e-mail assinado com HMAC (antes:
 *        bastava adivinhar o ID de 8 caracteres).
 *  [SEG] Throttle no envio de código OTP.
 *  [FIX] Motor de horários reescrito: respeita ano corrente, sem
 *        buraco no primeiro dia útil, uma única leitura de agenda,
 *        antecedência mínima configurável, eventos de dia inteiro.
 *  [FIX] Aba Configuracoes_Agenda passa a ser realmente lida.
 *  [FIX] Agenda do painel não mostra mais aprovações sem data.
 *  [FIX] Checagem de conflito no momento da confirmação (evita
 *        dois candidatos no mesmo horário).
 *  [NOV] Validação e deduplicação de e-mail no servidor.
 *  [NOV] verificarInstalacao() — diagnóstico de escopos.
 * ------------------------------------------------------------------
 */

// ============================================================
// CONFIGURAÇÃO
// ============================================================
/**
 * VALORES PADRÃO (fase de desenvolvimento — conta jr.conductor83@gmail.com).
 * Para transferir a plataforma ao COB, NÃO edite este bloco:
 * rode configurarProprietario(...) uma única vez. Os valores gravados em
 * Script Properties têm prioridade sobre estes padrões.
 */
const DEF_FOLDER_ID = "1jz6vfaFTkRqNP7FvrT4ECDrl5l4DHh2C";
const DEF_SS_ID     = "1NIzil2c1OUhKaYT6jUE1PaqilHWiMb8myHVpYqOuVts";
const DEF_ADMIN     = "jr.conductor83@gmail.com";
const DEF_CONTATO   = "orquestrasbrasileiras@gmail.com";

function P_(chave) { return PropertiesService.getScriptProperties().getProperty(chave) || ""; }
function FOLDER_ID_()  { return P_("COB_FOLDER_ID")   || DEF_FOLDER_ID; }
function SS_ID_()      { return P_("COB_SS_ID")       || DEF_SS_ID; }
function ADMIN_EMAIL_(){ return P_("COB_ADMIN_EMAIL") || DEF_ADMIN; }
function CONTATO_()    { return P_("COB_CONTATO")     || DEF_CONTATO; }

/** Agenda usada para ler disponibilidade e criar eventos. */
function CAL_() {
  const id = P_("COB_CALENDAR_ID");
  if (id) {
    const c = CalendarApp.getCalendarById(id);
    if (c) return c;
    throw new Error("Agenda COB_CALENDAR_ID não encontrada ou sem permissão: " + id);
  }
  return CalendarApp.getDefaultCalendar();
}

/**
 * TRANSFERÊNCIA DE POSSE — rodar no editor, uma vez, já na conta do COB.
 * Ex.: configurarProprietario({
 *        pastaDrive: "1AbC...",           // pasta do Drive do COB
 *        agenda: "orquestrasbrasileiras@gmail.com",
 *        planilha: "1XyZ...",             // opcional, se a planilha mudar
 *        contato: "orquestrasbrasileiras@gmail.com"
 *      });
 */
function configurarProprietario(cfg) {
  cfg = cfg || {};
  const p = PropertiesService.getScriptProperties();
  const gravado = [];
  if (cfg.pastaDrive) { p.setProperty("COB_FOLDER_ID", cfg.pastaDrive);   gravado.push("pasta do Drive"); }
  if (cfg.planilha)   { p.setProperty("COB_SS_ID", cfg.planilha);         gravado.push("planilha"); }
  if (cfg.agenda)     { p.setProperty("COB_CALENDAR_ID", cfg.agenda);     gravado.push("agenda"); }
  if (cfg.contato)    { p.setProperty("COB_CONTATO", cfg.contato);        gravado.push("contato LGPD"); }
  if (cfg.admin)      { p.setProperty("COB_ADMIN_EMAIL", cfg.admin);      gravado.push("admin padrão"); }
  const txt = "Gravado: " + (gravado.join(", ") || "nada") + "\n\n" + verificarInstalacao();
  Logger.log(txt);
  return txt;
}

/** Volta tudo para os padrões de desenvolvimento. */
function reverterProprietario() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  return "Script Properties limpas. A plataforma voltou aos valores padrão.";
}

const CATEGORIAS = {
  voluntario:    { sheet: "Voluntarios",    label: "Voluntário(a)",              folder: "VOLUNTARIOS",    agenda: true  },
  embaixador:    { sheet: "Embaixadores",   label: "Embaixador(a)",              folder: "EMBAIXADORES",   agenda: false },
  representante: { sheet: "Representantes", label: "Representante de Orquestra", folder: "REPRESENTANTES", agenda: false }
};

const HEADERS_BASE = ["ID","DATA","NOME","SOBRENOME","EMAIL","EMAIL_EXTRA","WHATSAPP","FUNCAO","FOTO_URL","STATUS_AGENDA","CV_URL","EVENT_ID","APELIDO","DETALHES","TERMOS"];

// Índices de coluna (0-based) — usar sempre estas constantes, nunca números soltos
const C = { ID:0, DATA:1, NOME:2, SOBRENOME:3, EMAIL:4, EMAIL_EXTRA:5, WHATSAPP:6,
            FUNCAO:7, FOTO:8, STATUS:9, CV:10, EVENT_ID:11, APELIDO:12, DETALHES:13, TERMOS:14 };

const CFG_PADRAO = { inicio: 10, fim: 18, horizonte: 21, antecedencia: 24, sabado: false };

function TZ_() { return Session.getScriptTimeZone() || "America/Sao_Paulo"; }

// ============================================================
// SETUP
// ============================================================
function garantirAbasBase() {
  const ss = SpreadsheetApp.openById(SS_ID_());

  if (!ss.getSheetByName("Admins")) {
    const s = ss.insertSheet("Admins");
    s.appendRow(["E-MAILS AUTORIZADOS"]);
    s.getRange("A1").setFontWeight("bold");
    s.appendRow([ADMIN_EMAIL_()]);
  }

  if (!ss.getSheetByName("Estrutura_COB")) {
    const s = ss.insertSheet("Estrutura_COB");
    s.appendRow(["CARGO / GT", "LIDERANÇA"]);
    s.getRange("A1:B1").setFontWeight("bold");
    s.appendRow(["COORDENAÇÃO GERAL", "Gisely Nascimento"]);
    s.appendRow(["GT ORQUESTRANDO ARQUIVO", "Carô Tenório"]);
    s.appendRow(["GT GESTÃO E COMUNICAÇÃO", "Lideranças Ativas"]);
  }

  _garantirConfigAgenda(ss);

  Object.keys(CATEGORIAS).forEach(function (key) {
    const nome = CATEGORIAS[key].sheet;
    let s = ss.getSheetByName(nome);
    if (!s) {
      s = ss.insertSheet(nome);
      s.appendRow(HEADERS_BASE);
      s.getRange(1, 1, 1, HEADERS_BASE.length).setFontWeight("bold");
      s.setFrozenRows(1);
    } else if (s.getLastColumn() < HEADERS_BASE.length) {
      s.getRange(1, 1, 1, HEADERS_BASE.length).setValues([HEADERS_BASE]).setFontWeight("bold");
    }
  });
}

function _garantirConfigAgenda(ss) {
  let s = ss.getSheetByName("Configuracoes_Agenda");
  if (!s) {
    s = ss.insertSheet("Configuracoes_Agenda");
    s.appendRow(["HORA_INICIO","HORA_FIM","DIAS_HORIZONTE","ANTECEDENCIA_HORAS","INCLUIR_SABADO"]);
    s.getRange(1, 1, 1, 5).setFontWeight("bold");
    s.appendRow([CFG_PADRAO.inicio, CFG_PADRAO.fim, CFG_PADRAO.horizonte, CFG_PADRAO.antecedencia, "NAO"]);
    s.getRange("A4").setValue("Editar apenas a linha 2. INCLUIR_SABADO aceita SIM ou NAO.");
  }
  return s;
}

function _lerConfigAgenda() {
  try {
    const s = _garantirConfigAgenda(SpreadsheetApp.openById(SS_ID_()));
    const v = s.getRange(2, 1, 1, 5).getDisplayValues()[0];
    const num = function (x, pad) { const n = parseInt(String(x).replace(/\D/g, ""), 10); return isNaN(n) ? pad : n; };
    return {
      inicio:       Math.min(23, Math.max(0,  num(v[0], CFG_PADRAO.inicio))),
      fim:          Math.min(24, Math.max(1,  num(v[1], CFG_PADRAO.fim))),
      horizonte:    Math.min(90, Math.max(1,  num(v[2], CFG_PADRAO.horizonte))),
      antecedencia: Math.min(720, Math.max(0, num(v[3], CFG_PADRAO.antecedencia))),
      sabado:       /^s/i.test(String(v[4] || "").trim())
    };
  } catch (e) {
    return { inicio: CFG_PADRAO.inicio, fim: CFG_PADRAO.fim, horizonte: CFG_PADRAO.horizonte,
             antecedencia: CFG_PADRAO.antecedencia, sabado: CFG_PADRAO.sabado };
  }
}

// ============================================================
// SEGURANÇA
// ============================================================
function _segredo() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty("APP_SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty("APP_SECRET", s); }
  return s;
}

/** Token determinístico por candidato — usado no link de aprovação enviado por e-mail. */
function _tokenCandidato(id) {
  const sig = Utilities.computeHmacSha256Signature(String(id), _segredo());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, "").substring(0, 28);
}

function _emailsAutorizados() {
  const ss = SpreadsheetApp.openById(SS_ID_());
  const aba = ss.getSheetByName("Admins");
  if (!aba) return [ADMIN_EMAIL_()];
  return aba.getDataRange().getValues()
    .map(function (l) { return String(l[0] || "").toLowerCase().trim(); })
    .filter(function (e) { return e && e.indexOf("@") > 0; });
}

function _novaSessao(email) {
  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().split("-")[0];
  CacheService.getScriptCache().put("SESS_" + token, email.toLowerCase().trim(), 21600); // 6h (máximo do CacheService)
  return token;
}

/** Lança erro se a sessão não for válida. Toda função privilegiada começa por aqui. */
function _exigirAdmin(sessao) {
  const email = sessao ? CacheService.getScriptCache().get("SESS_" + sessao) : null;
  if (!email) throw new Error("SESSAO_EXPIRADA");
  if (_emailsAutorizados().indexOf(email) === -1) throw new Error("SESSAO_EXPIRADA");
  return email;
}

// ============================================================
// AUTENTICAÇÃO ADMIN (OTP por e-mail)
// ============================================================
function solicitarAcessoAdmin(emailDigitado) {
  try {
    garantirAbasBase();
    const email = String(emailDigitado || "").toLowerCase().trim();
    if (!_emailValido(email)) return { success: false, message: "E-mail inválido." };

    // Resposta idêntica para autorizado e não autorizado, para não enumerar admins.
    const autorizado = _emailsAutorizados().indexOf(email) !== -1;

    const cache = CacheService.getScriptCache();
    if (cache.get("THROTTLE_" + email)) {
      return { success: false, message: "Aguarde 60 segundos para pedir um novo código." };
    }
    cache.put("THROTTLE_" + email, "1", 60);

    if (!autorizado) return { success: false, message: "E-mail não autorizado." };

    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    cache.put("OTP_" + email, codigo, 600);

    const corpo = '<p style="text-align:center;color:#9BA89F;margin:0 0 18px 0;">'
      + 'Use o código abaixo para entrar no painel administrativo.</p>'
      + '<p style="text-align:center;font-size:40px;letter-spacing:14px;color:#E8B4BC;'
      + 'font-family:Georgia,serif;margin:0 0 22px 0;">' + codigo + '</p>'
      + '<p style="max-width:330px;margin:0 auto;text-align:center;color:#7E8C83;font-size:11.5px;'
      + 'line-height:1.85;">O código expira em dez minutos. Se não foi<br>'
      + 'você quem pediu, ignore este e-mail.</p>';
    MailApp.sendEmail({ to: email, subject: "Seu código de acesso — Painel COB",
      htmlBody: criarTemplateEmail("ACESSO AO PAINEL", corpo, { lgpd: false }) });
    return { success: true };
  } catch (e) {
    return { success: false, message: "Falha ao enviar o código. Tente novamente." };
  }
}

function validarCodigoAdmin(emailDigitado, codigoDigitado) {
  const email = String(emailDigitado || "").toLowerCase().trim();
  const cache = CacheService.getScriptCache();
  const salvo = cache.get("OTP_" + email);
  const tentativas = parseInt(cache.get("TRY_" + email) || "0", 10);

  if (tentativas >= 5) return { success: false, message: "Muitas tentativas. Peça um novo código." };
  if (!salvo || salvo !== String(codigoDigitado || "").trim()) {
    cache.put("TRY_" + email, String(tentativas + 1), 600);
    return { success: false, message: "Código inválido." };
  }
  cache.remove("OTP_" + email);
  cache.remove("TRY_" + email);
  return { success: true, sessao: _novaSessao(email), email: email };
}

function encerrarSessao(sessao) {
  if (sessao) CacheService.getScriptCache().remove("SESS_" + sessao);
  return { success: true };
}

// ============================================================
// ROTEAMENTO
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const view = p.view || "user";
  const t = HtmlService.createTemplateFromFile("index");

  t.candidato = null;
  t.candidatoCategoria = "";
  t.candidatoCategoriaLabel = "";
  t.aprovacaoErro = "";
  t.tokenAprovacao = "";

  if (view === "admin") {
    if (!p.id || !p.t || p.t !== _tokenCandidato(p.id)) {
      t.aprovacaoErro = "Link inválido ou expirado. Acesse o painel administrativo para concluir.";
    } else {
      const cand = getCandidatoById(p.id, p.categoria || null);
      if (!cand) {
        t.aprovacaoErro = "Cadastro não encontrado. Ele pode ter sido excluído.";
      } else if (String(cand.opcoes || "").toUpperCase().indexOf("CONFIRMADO") === 0) {
        t.aprovacaoErro = "Este cadastro já foi confirmado: " + cand.opcoes;
      } else {
        t.candidato = cand;
        t.candidatoCategoria = cand.categoria;
        t.candidatoCategoriaLabel = CATEGORIAS[cand.categoria] ? CATEGORIAS[cand.categoria].label : "";
        t.tokenAprovacao = p.t;
      }
    }
  }

  t.appUrl = ScriptApp.getService().getUrl();
  t.viewMode = view;
  t.abrirPainel = (view === "painel");   // link do e-mail cai direto no login do painel
  return t.evaluate()
    .setTitle("Plataforma COB")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// DASHBOARD
// ============================================================
function getDashboardData(sessao) {
  try {
    _exigirAdmin(sessao);
  } catch (err) {
    return { success: false, expirada: true, message: "SESSAO_EXPIRADA" };
  }

  // A criação de abas nunca pode derrubar a leitura do painel.
  const avisos = [];
  try { garantirAbasBase(); }
  catch (err) { avisos.push("Não foi possível verificar as abas: " + err); }

  const ss = SpreadsheetApp.openById(SS_ID_());
  const candidaturas = {};
  const agenda = [];
  const diagnostico = [];

  // Cada categoria é lida de forma isolada: uma aba com problema não
  // pode mais apagar os dados das outras (era o que acontecia antes).
  Object.keys(CATEGORIAS).forEach(function (key) {
    const nomeAba = CATEGORIAS[key].sheet;
    candidaturas[key] = [];
    try {
      const aba = _acharAba(ss, nomeAba);
      if (!aba) {
        diagnostico.push({ categoria: key, aba: nomeAba, linhas: 0, erro: "aba não encontrada" });
        avisos.push('A aba "' + nomeAba + '" não existe na planilha.');
        return;
      }

      const dados = aba.getDataRange().getDisplayValues();
      const lista = [];
      let ignoradas = 0;

      for (let i = 1; i < dados.length; i++) {
        const r = dados[i];
        if (!r[C.ID] && !r[C.EMAIL] && !r[C.NOME]) { ignoradas++; continue; }  // linha realmente vazia
        const status = String(r[C.STATUS] || "PENDENTE").trim() || "PENDENTE";
        const id = String(r[C.ID] || ("LINHA" + (i + 1)));

        lista.push({
          id: id,
          linha: i + 1,
          data: r[C.DATA],
          nome: (String(r[C.NOME] || "") + " " + String(r[C.SOBRENOME] || "")).trim() || "(sem nome)",
          email: r[C.EMAIL], emailExtra: r[C.EMAIL_EXTRA],
          whatsapp: String(r[C.WHATSAPP] || "").replace(/^'/, ""),
          funcao: r[C.FUNCAO], foto: r[C.FOTO], cv: r[C.CV],
          status: status,
          apelido: r[C.APELIDO] || "", detalhes: r[C.DETALHES] || "", termos: r[C.TERMOS] || "",
          categoria: key, categoriaLabel: CATEGORIAS[key].label,
          token: _tokenCandidato(id)
        });

        const m = status.match(/^CONFIRMADO:\s*(.+)$/i);
        if (m) {
          const partes = m[1].split(/\s+às\s+/i);
          agenda.push({
            titulo: "COB · Onboarding " + r[C.NOME] + " (" + CATEGORIAS[key].label + ")",
            data: (partes[0] || m[1]).trim(),
            hora: (partes[1] || "—").trim(),
            convidados: r[C.EMAIL],
            categoria: key
          });
        }
      }

      candidaturas[key] = lista.reverse();
      diagnostico.push({ categoria: key, aba: aba.getName(), linhas: lista.length, vazias: ignoradas, erro: "" });
    } catch (err) {
      diagnostico.push({ categoria: key, aba: nomeAba, linhas: 0, erro: String(err.message || err) });
      avisos.push('Falha ao ler "' + nomeAba + '": ' + err);
    }
  });

  let estrutura = [];
  try {
    const abaEst = ss.getSheetByName("Estrutura_COB");
    if (abaEst) {
      estrutura = abaEst.getDataRange().getDisplayValues().slice(1)
        .filter(function (r) { return r[0]; })
        .map(function (r) { return { cargo: r[0], nome: r[1] }; });
    }
  } catch (err) { avisos.push("Estrutura administrativa: " + err); }

  let admins = [];
  try { admins = _emailsAutorizados(); } catch (err) { avisos.push("Admins: " + err); }

  return {
    success: true,
    candidaturas: candidaturas,
    estrutura: estrutura,
    admins: admins,
    agenda: agenda,
    diagnostico: diagnostico,
    avisos: avisos,
    planilha: ss.getName(),
    planilhaUrl: ss.getUrl(),
    lido: Utilities.formatDate(new Date(), TZ_(), "dd/MM/yyyy 'às' HH:mm")
  };
}

/** Acha a aba pelo nome exato ou, se não achar, por comparação sem acento/caixa. */
function _acharAba(ss, nome) {
  let aba = ss.getSheetByName(nome);
  if (aba) return aba;
  const chave = _normalizar(nome);
  const todas = ss.getSheets();
  for (let i = 0; i < todas.length; i++) {
    if (_normalizar(todas[i].getName()) === chave) return todas[i];
  }
  return null;
}

function _normalizar(s) {
  return String(s || "").toLowerCase().trim()
    .replace(/[áàâãä]/g, "a").replace(/[éèêë]/g, "e").replace(/[íìîï]/g, "i")
    .replace(/[óòôõö]/g, "o").replace(/[úùûü]/g, "u").replace(/ç/g, "c")
    .replace(/[^a-z0-9]/g, "");
}

// ============================================================
// UTILITÁRIOS
// ============================================================
function formatarNome(texto) {
  if (!texto) return "";
  const minusculas = ["de", "da", "do", "dos", "das", "e", "di", "del"];
  return String(texto).toLowerCase().split(/\s+/).filter(String).map(function (w, i) {
    if (i > 0 && minusculas.indexOf(w) !== -1) return w;
    return w.charAt(0).toUpperCase() + w.slice(1); // charAt preserva acentos (não usar /\b\w/)
  }).join(" ");
}

function _emailValido(e) {
  return /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(String(e || "").trim());
}

function _emailJaCadastrado(sheet, email) {
  if (!email) return false;
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][C.EMAIL] || "").toLowerCase().trim() === email) return true;
  }
  return false;
}

function getOrCreateSubfolder(pai, nome) {
  const f = pai.getFoldersByName(nome);
  return f.hasNext() ? f.next() : pai.createFolder(nome);
}

// ============================================================
// PROCESSAMENTO DO FORMULÁRIO (upload em Base64)
// ============================================================
function processFormBase64(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    garantirAbasBase();

    const ss = SpreadsheetApp.openById(SS_ID_());
    const categoria = (payload && CATEGORIAS[payload.categoria]) ? payload.categoria : "voluntario";
    const sheet = ss.getSheetByName(CATEGORIAS[categoria].sheet);

    // ---- Validação de servidor (não confiar no cliente) ----
    const email = String(payload.email || "").toLowerCase().trim();
    const emailExtra = String(payload.email_extra || "").toLowerCase().trim();
    if (!String(payload.nome || "").trim() || !String(payload.sobrenome || "").trim()) {
      return { success: false, message: "Nome e sobrenome são obrigatórios." };
    }
    if (!_emailValido(email)) return { success: false, message: "E-mail principal inválido." };
    if (emailExtra && !_emailValido(emailExtra)) return { success: false, message: "E-mail adicional inválido." };
    if (emailExtra && emailExtra === email) return { success: false, message: "O e-mail adicional deve ser diferente do principal." };
    if (!payload.aceiteTermos) return { success: false, message: "É necessário aceitar o Termo de Adesão." };
    if (_emailJaCadastrado(sheet, email)) {
      return { success: false, duplicado: true,
               message: "Este e-mail já possui cadastro nesta categoria. Fale com a equipe do COB." };
    }

    const uniqueId = Utilities.getUuid().split("-")[0].toUpperCase();
    const raiz = DriveApp.getFolderById(FOLDER_ID_());
    const pastaCat = getOrCreateSubfolder(raiz, CATEGORIAS[categoria].folder);

    const nome = formatarNome(payload.nome);
    const sobrenome = formatarNome(payload.sobrenome);
    const identif = (nome + "_" + sobrenome + "_ID" + uniqueId).toUpperCase().replace(/\s+/g, "_");

    let fotoUrl = "", cvUrl = "";
    try {
      if (payload.foto && payload.foto.data) {
        const pasta = getOrCreateSubfolder(pastaCat, "FOTOS_PERFIL");
        const ext = String(payload.foto.name || "jpg").split(".").pop();
        const blob = Utilities.newBlob(Utilities.base64Decode(payload.foto.data), payload.foto.mimeType, "FOTO_" + identif + "." + ext);
        fotoUrl = pasta.createFile(blob).getUrl();
      }
    } catch (err) { Logger.log("Erro foto: " + err); }

    try {
      if (payload.cv && payload.cv.data) {
        const pasta = getOrCreateSubfolder(pastaCat, "CURRICULOS");
        const ext = String(payload.cv.name || "pdf").split(".").pop();
        const blob = Utilities.newBlob(Utilities.base64Decode(payload.cv.data), payload.cv.mimeType, "CV_" + identif + "." + ext);
        cvUrl = pasta.createFile(blob).getUrl();
      }
    } catch (err) { Logger.log("Erro CV: " + err); }

    const temAgendamento = !!(payload.opt1 && payload.opt2 && payload.opt3);
    const status = temAgendamento ? ("1: " + payload.opt1 + " | 2: " + payload.opt2 + " | 3: " + payload.opt3) : "PENDENTE";
    const detalhes = montarDetalhes(payload);
    const termos = "Aceito em " + Utilities.formatDate(new Date(), TZ_(), "dd/MM/yyyy HH:mm");

    sheet.appendRow([
      uniqueId, new Date(), nome, sobrenome, email, emailExtra,
      "'" + String(payload.whatsapp || ""),       // apóstrofo: senão o Sheets interpreta "+" como fórmula
      String(payload.funcao_cob || ""), fotoUrl, status, cvUrl, "",
      String(payload.apelido || ""), detalhes, termos
    ]);

    try {
      enviarNotificacoes({
        categoria: categoria, categoriaLabel: CATEGORIAS[categoria].label,
        nome: nome, sobrenome: sobrenome, funcao_cob: payload.funcao_cob,
        email: email, email_extra: emailExtra, whatsapp: payload.whatsapp,
        opt1: payload.opt1, opt2: payload.opt2, opt3: payload.opt3, detalhes: detalhes
      }, uniqueId);
    } catch (err) {
      Logger.log("Erro e-mail: " + err); // cadastro já gravado — nunca falhar por causa do e-mail
    }

    return { success: true, id: uniqueId };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function montarDetalhes(payload) {
  if (!payload || !payload.details || !payload.detailsLabels) return "";
  const linhas = [];
  payload.detailsLabels.forEach(function (f) {
    const v = payload.details[f.id];
    if (v && String(v).trim()) linhas.push(f.label + ": " + String(v).trim());
  });
  return linhas.join("\n");
}

// ============================================================
// E-MAILS
// ============================================================
/**
 * Template de e-mail.
 * Corpo alinhado à esquerda; rodapé centralizado em coluna estreita,
 * de modo que as linhas quebrem com extensões parecidas.
 */
function criarTemplateEmail(titulo, conteudo, opcoes) {
  opcoes = opcoes || {};
  const contato = CONTATO_();

  const notaComprovante = opcoes.comprovante
    ? '<p style="max-width:340px;margin:0 auto 16px auto;text-align:center;font-size:12.5px;'
      + 'line-height:1.8;color:#9BA89F;">Este e-mail é o comprovante do seu cadastro.<br>'
      + 'Guarde-o para consulta quando precisar.</p>'
    : "";

  const lgpd = opcoes.lgpd === false ? "" :
      '<p style="max-width:390px;margin:0 auto 14px auto;text-align:center;font-size:11.5px;'
    + 'line-height:1.85;color:#7E8C83;">Seus dados são tratados conforme a Lei nº 13.709/2018 '
    + '(Lei Geral de Proteção de Dados). O COB coleta apenas o necessário para o seu cadastro '
    + 'e não compartilha suas informações com terceiros.</p>'
    + '<p style="max-width:390px;margin:0 auto 14px auto;text-align:center;font-size:11.5px;'
    + 'line-height:1.85;color:#7E8C83;">Para acessar, corrigir ou excluir os seus dados,<br>'
    + 'escreva para <a href="mailto:' + contato + '" style="color:#E8B4BC;text-decoration:none;">'
    + contato + '</a>.</p>';

  const rodape = (notaComprovante || lgpd)
    ? '<div style="margin-top:34px;padding-top:22px;border-top:1px solid #1F3D30;">'
      + notaComprovante + lgpd
      + '<p style="max-width:440px;margin:18px auto 0 auto;text-align:center;font-size:10.5px;'
      + 'letter-spacing:2px;color:#4E5C54;">COB · CONECTANDO ORQUESTRAS BRASILEIRAS</p></div>'
    : "";

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width, initial-scale=1"></head>'
    + '<body style="background:#0A1812;margin:0;padding:30px 10px;">'
    + '<div style="font-family:Helvetica,Arial,sans-serif;background:#122620;color:#F5EFE6;'
    + 'padding:40px 32px;border:1px solid #1F3D30;border-radius:12px;max-width:640px;margin:0 auto;'
    + 'border-top:4px solid #E8B4BC;text-align:left;">'
    + '<h1 style="font-family:Georgia,serif;color:#E8B4BC;text-align:center;font-size:22px;'
    + 'letter-spacing:2px;margin:0 0 8px 0;font-weight:600;">' + titulo + '</h1>'
    + '<p style="max-width:440px;margin:0 auto 30px auto;text-align:center;color:#9BA89F;'
    + 'font-size:10.5px;letter-spacing:2.4px;">CONECTANDO ORQUESTRAS BRASILEIRAS</p>'
    + '<div style="line-height:1.65;font-size:15px;text-align:left;">' + conteudo + '</div>'
    + rodape
    + '</div></body></html>';
}

function _botao(texto, url) {
  return '<div style="text-align:center;margin:32px 0 8px 0;"><a href="' + url
    + '" style="background:#E8B4BC;color:#0A1812;padding:15px 34px;text-decoration:none;'
    + 'border-radius:50px;font-family:Georgia,serif;font-weight:600;font-size:16px;'
    + 'display:inline-block;letter-spacing:1px;">' + texto + '</a></div>';
}

function _linha(rotulo, valor) {
  if (!valor) return "";
  return '<p style="margin:9px 0;text-align:left;"><b style="color:#E8B4BC;">' + rotulo
    + ':</b> <span style="color:#F5EFE6;">' + valor + '</span></p>';
}

function enviarNotificacoes(d, id) {
  const appUrl = ScriptApp.getService().getUrl();
  const agora = Utilities.formatDate(new Date(), TZ_(), "dd/MM/yyyy 'às' HH:mm");
  const temAgendamento = !!(d.opt1 && d.opt2 && d.opt3);
  const caixa = 'style="background:#0F2A20;padding:24px;border-radius:8px;margin:24px 0;'
              + 'border-left:4px solid #E8B4BC;text-align:left;"';

  const horarios = temAgendamento
    ? '<p style="color:#E8B4BC;margin:20px 0 6px 0;"><b>Horários que você sugeriu</b></p>'
      + '<ul style="margin:0;color:#9BA89F;padding-left:20px;line-height:1.9;"><li>'
      + d.opt1 + '</li><li>' + d.opt2 + '</li><li>' + d.opt3 + '</li></ul>'
    : "";

  const _especificos = function (rotulo) {
    if (!d.detalhes) return "";
    return '<p style="color:#E8B4BC;margin:22px 0 6px 0;"><b>' + rotulo + '</b></p>'
      + d.detalhes.split("\n").map(function (l) {
          const i = l.indexOf(":");
          return i > 0 ? _linha(l.substring(0, i), l.substring(i + 1).trim())
                       : '<p style="margin:9px 0;">' + l + '</p>';
        }).join("");
  };

  // ---------- Candidato: comprovante completo do que foi enviado ----------
  const corpoCand = '<p>Olá <b>' + d.nome + ' ' + d.sobrenome + '</b>,</p>'
    + '<p>Recebemos o seu cadastro como <b>' + d.categoriaLabel + '</b> no Conectando Orquestras '
    + 'Brasileiras. Abaixo está o registro exato das informações que você nos enviou.</p>'
    + '<div ' + caixa + '>'
    + '<h3 style="color:#E8B4BC;margin:0 0 14px 0;font-family:Georgia,serif;font-size:16px;'
    + 'letter-spacing:1px;">RESUMO DO SEU CADASTRO</h3>'
    + _linha("Categoria", d.categoriaLabel)
    + _linha("ID de cadastro", id)
    + _linha("Enviado em", agora)
    + _linha("Nome", d.nome + " " + d.sobrenome)
    + _linha("Função", d.funcao_cob)
    + _linha("E-mail", d.email)
    + _linha("E-mail adicional", d.email_extra)
    + _linha("WhatsApp", d.whatsapp)
    + _especificos('Informações que você forneceu')
    + horarios
    + '</div>'
    + (temAgendamento
        ? '<p>A equipe do COB confirma a data e o horário exatos em até 48 horas, por e-mail.</p>'
        : '<p>A equipe do COB analisa o seu cadastro e entra em contato com os próximos passos.</p>')
    + '<p style="color:#9BA89F;font-size:13.5px;">Encontrou algum dado errado? Responda a este '
    + 'e-mail que corrigimos para você.</p>';

  const paraCand = [d.email].concat(d.email_extra ? [d.email_extra] : []).join(",");
  MailApp.sendEmail({
    to: paraCand,
    subject: "Cadastro recebido — COB · " + d.categoriaLabel + " (ID " + id + ")",
    htmlBody: criarTemplateEmail("CADASTRO RECEBIDO", corpoCand, { comprovante: true, lgpd: true })
  });

  // ---------- Administração: ação direta + acesso ao painel ----------
  const linkAprov = appUrl + "?view=admin&id=" + id + "&categoria=" + d.categoria + "&t=" + _tokenCandidato(id);
  const linkPainel = appUrl + "?view=painel";

  const acao = temAgendamento
    ? '<p style="color:#9BA89F;">A pessoa sugeriu três horários. Confirme um deles para criar o '
      + 'evento no Google Calendar e avisar o candidato automaticamente:</p>'
      + _botao("CONFIRMAR HORÁRIO", linkAprov)
      + '<p style="text-align:center;color:#7E8C83;font-size:11.5px;max-width:330px;margin:0 auto;'
      + 'line-height:1.8;">Link pessoal e de uso único.<br>Não encaminhe para terceiros.</p>'
    : '<p style="color:#9BA89F;">Abra o painel para revisar o cadastro completo e aprová-lo.</p>';

  const corpoAdmin = '<p style="font-size:17px;margin:0 0 6px 0;">Novo cadastro: '
    + '<b style="color:#E8B4BC;">' + d.nome + ' ' + d.sobrenome + '</b></p>'
    + '<p style="color:#9BA89F;margin:0 0 20px 0;">' + d.categoriaLabel + ' · recebido em ' + agora + '</p>'
    + '<div ' + caixa + '>'
    + _linha("ID", id)
    + _linha("Função", d.funcao_cob)
    + _linha("E-mail", d.email)
    + _linha("E-mail adicional", d.email_extra)
    + _linha("WhatsApp", d.whatsapp)
    + _especificos('Informações fornecidas no cadastro')
    + horarios
    + '</div>'
    + acao
    + '<hr style="border:none;border-top:1px solid #1F3D30;margin:32px 0 24px 0;">'
    + '<p style="color:#9BA89F;text-align:center;">Painel administrativo do COB</p>'
    + _botao("ACESSAR O PAINEL", linkPainel)
    + '<p style="max-width:360px;margin:0 auto;text-align:center;color:#7E8C83;font-size:11.5px;'
    + 'line-height:1.8;">Informe o seu e-mail e receba um código<br>de seis dígitos para entrar no painel.</p>';

  MailApp.sendEmail({
    to: _destinatariosAdmin(),
    subject: "COB · Novo cadastro (" + d.categoriaLabel + ") — " + d.nome + " " + d.sobrenome,
    htmlBody: criarTemplateEmail("NOVO CADASTRO", corpoAdmin, { lgpd: false })
  });
}

/** Todos os e-mails da aba Admins recebem as notificações. */
function _destinatariosAdmin() {
  try {
    const lista = _emailsAutorizados();
    if (lista.length) return lista.join(",");
  } catch (e) {}
  return ADMIN_EMAIL_();
}

// ============================================================
// CRUD
// ============================================================
function salvarNovaEstrutura(sessao, lista) {
  _exigirAdmin(sessao);
  const sheet = SpreadsheetApp.openById(SS_ID_()).getSheetByName("Estrutura_COB");
  sheet.clear();
  sheet.appendRow(["CARGO / GT", "LIDERANÇA"]);
  sheet.getRange("A1:B1").setFontWeight("bold");
  (lista || []).forEach(function (i) { if (i && i.cargo) sheet.appendRow([i.cargo, i.nome || ""]); });
  return { success: true };
}

function addAdminEmail(sessao, email) {
  _exigirAdmin(sessao);
  const e = String(email || "").toLowerCase().trim();
  if (!_emailValido(e)) return { success: false, message: "E-mail inválido." };
  if (_emailsAutorizados().indexOf(e) !== -1) return { success: false, message: "Este e-mail já tem acesso." };
  SpreadsheetApp.openById(SS_ID_()).getSheetByName("Admins").appendRow([e]);
  return { success: true };
}

function removeAdminEmail(sessao, email) {
  const atual = _exigirAdmin(sessao);
  const e = String(email || "").toLowerCase().trim();
  if (e === atual) return { success: false, message: "Você não pode remover o seu próprio acesso." };
  if (_emailsAutorizados().length <= 1) return { success: false, message: "É preciso manter ao menos um administrador." };
  const s = SpreadsheetApp.openById(SS_ID_()).getSheetByName("Admins");
  const d = s.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) {
    if (String(d[i][0] || "").toLowerCase().trim() === e) s.deleteRow(i + 1);
  }
  return { success: true };
}

function deletarCandidato(sessao, id, categoria) {
  _exigirAdmin(sessao);
  const ss = SpreadsheetApp.openById(SS_ID_());
  const chaves = (categoria && CATEGORIAS[categoria]) ? [categoria] : Object.keys(CATEGORIAS);
  for (let k = 0; k < chaves.length; k++) {
    const sheet = ss.getSheetByName(CATEGORIAS[chaves[k]].sheet);
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][C.ID] == id) {
        const eventId = d[i][C.EVENT_ID];
        if (eventId) {
          try { const ev = CAL_().getEventById(eventId); if (ev) ev.deleteEvent(); }
          catch (err) { Logger.log("Evento não removido: " + err); }
        }
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
  }
  return { success: false, message: "Cadastro não encontrado." };
}

function getCandidatoById(id, categoria) {
  const ss = SpreadsheetApp.openById(SS_ID_());
  const chaves = (categoria && CATEGORIAS[categoria]) ? [categoria] : Object.keys(CATEGORIAS);
  for (let k = 0; k < chaves.length; k++) {
    const sheet = ss.getSheetByName(CATEGORIAS[chaves[k]].sheet);
    if (!sheet) continue;
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][C.ID] == id) {
        return {
          id: String(id), nome: d[i][C.NOME], sobrenome: d[i][C.SOBRENOME],
          email: d[i][C.EMAIL], emailExtra: d[i][C.EMAIL_EXTRA],
          funcao: d[i][C.FUNCAO], opcoes: String(d[i][C.STATUS] || ""),
          linha: i + 1, sheet: sheet.getSheetName(), categoria: chaves[k]
        };
      }
    }
  }
  return null;
}

// ============================================================
// AGENDA — geração de horários
// ============================================================
function getAvailableTimes() {
  try {
    const cfg = _lerConfigAgenda();
    const tz = TZ_();
    const dias = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
    const agora = new Date();
    const minimo = new Date(agora.getTime() + cfg.antecedencia * 3600000);

    const hoje = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate(), 0, 0, 0, 0);
    const fimJanela = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + cfg.horizonte + 1, 0, 0, 0, 0);

    // Uma única leitura da agenda para toda a janela (antes: uma leitura por hora)
    const cal = CAL_();
    const ocupados = cal.getEvents(hoje, fimJanela).map(function (ev) {
      return { i: ev.getStartTime().getTime(), f: ev.getEndTime().getTime() };
    });

    const slots = [];
    for (let d = 0; d <= cfg.horizonte; d++) {
      const dia = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() + d, 0, 0, 0, 0);
      const dow = dia.getDay();
      if (dow === 0) continue;                    // domingo nunca
      if (dow === 6 && !cfg.sabado) continue;     // sábado conforme configuração

      for (let h = cfg.inicio; h < cfg.fim; h++) {
        const ini = new Date(dia.getFullYear(), dia.getMonth(), dia.getDate(), h, 0, 0, 0);
        const fim = new Date(ini.getTime() + 3600000);
        if (ini.getTime() < minimo.getTime()) continue;

        const conflito = ocupados.some(function (o) { return o.i < fim.getTime() && o.f > ini.getTime(); });
        if (conflito) continue;

        slots.push({
          id: Utilities.formatDate(ini, tz, "yyyyMMddHHmm"),
          label: Utilities.formatDate(ini, tz, "dd/MM/yyyy") + " (" + dias[ini.getDay()] + ") às " + Utilities.formatDate(ini, tz, "HH:mm"),
          data: Utilities.formatDate(ini, tz, "dd/MM/yyyy"),
          diaSemana: dias[ini.getDay()],
          hora: Utilities.formatDate(ini, tz, "HH:mm"),
          periodo: h < 12 ? "manha" : (h < 18 ? "tarde" : "noite")
        });
      }
    }
    return slots;
  } catch (e) {
    Logger.log("getAvailableTimes: " + e);
    return [];
  }
}

/** Aceita "dd/MM/yyyy (Ter) às HH:mm" e o formato antigo "dd/MM (Ter) às HH:mm". */
function parseDate(str) {
  const s = String(str || "");
  let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})[^\d]*?(\d{1,2}):(\d{2})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5], 0, 0);

  m = s.match(/(\d{1,2})\/(\d{1,2})[^\d]*?(\d{1,2}):(\d{2})/);
  if (m) {
    const hoje = new Date();
    let ano = hoje.getFullYear();
    const cand = new Date(ano, +m[2] - 1, +m[1], +m[3], +m[4], 0, 0);
    // Sem ano no rótulo: se a data já passou há mais de um mês, é do ano seguinte
    if (cand.getTime() < hoje.getTime() - 30 * 86400000) ano += 1;
    return new Date(ano, +m[2] - 1, +m[1], +m[3], +m[4], 0, 0);
  }
  throw new Error("Horário em formato não reconhecido: " + s);
}

// ============================================================
// AGENDA — confirmação
// ============================================================
function confirmarAgendamento(id, horario, categoria, credencial) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    // Aceita sessão do painel OU token do link de e-mail
    const viaToken = credencial && credencial === _tokenCandidato(id);
    if (!viaToken) _exigirAdmin(credencial);

    const cand = getCandidatoById(id, categoria);
    if (!cand) return { success: false, message: "Cadastro não encontrado." };
    if (/^CONFIRMADO/i.test(cand.opcoes)) return { success: false, message: "Este cadastro já foi confirmado." };

    const inicio = parseDate(horario);
    const fim = new Date(inicio.getTime() + 3600000);
    const cal = CAL_();

    // Revalida: outro candidato pode ter ocupado o horário nesse meio-tempo
    const choque = cal.getEvents(inicio, fim);
    if (choque.length > 0) {
      return { success: false, message: "Esse horário acabou de ser ocupado. Escolha outra opção." };
    }

    const titulo = "COB · Onboarding " + cand.nome + " " + (cand.sobrenome || "");
    const evento = cal.createEvent(titulo.trim(), inicio, fim, {
      guests: cand.email,
      sendInvites: true,
      description: "Sessão de onboarding — " + (CATEGORIAS[cand.categoria] ? CATEGORIAS[cand.categoria].label : "")
        + "\nID: " + cand.id + "\nE-mail: " + cand.email + "\nFunção: " + (cand.funcao || "—")
    });
    try { evento.addPopupReminder(60); evento.addEmailReminder(1440); } catch (e) {}

    const ss = SpreadsheetApp.openById(SS_ID_());
    const sheet = ss.getSheetByName(cand.sheet);
    sheet.getRange(cand.linha, C.STATUS + 1).setValue("CONFIRMADO: " + horario);
    sheet.getRange(cand.linha, C.EVENT_ID + 1).setValue(evento.getId());

    // E-mail próprio ao candidato: o convite do Calendar é silenciosamente
    // ignorado pelo Google quando o convidado é a própria conta dona do script.
    const caixa = 'style="background:#0F2A20;padding:24px;border-radius:8px;margin:24px 0;border-left:4px solid #E8B4BC;text-align:left;"';
    const corpo = '<p>Olá <b>' + cand.nome + '</b>,</p>'
      + '<p>Sua sessão de onboarding com a equipe do COB está confirmada.</p>'
      + '<div ' + caixa + '><p style="margin:0;font-size:17px;font-family:Georgia,serif;">'
      + '<b style="color:#E8B4BC;">Data e hora:</b> ' + horario + '</p></div>'
      + '<p style="color:#9BA89F;">O evento <b>"' + titulo.trim() + '"</b> foi adicionado à sua agenda. '
      + 'Se precisar remarcar, responda a este e-mail.</p>';
    const para = [cand.email].concat(cand.emailExtra ? [cand.emailExtra] : []).join(",");
    MailApp.sendEmail({ to: para, subject: "Onboarding confirmado — COB",
      htmlBody: criarTemplateEmail("ONBOARDING AGENDADO", corpo, { lgpd: true }) });

    const linkAgenda = "https://calendar.google.com/calendar/event?eid="
      + Utilities.base64Encode(evento.getId().split("@")[0] + " " + ADMIN_EMAIL_()).replace(/=/g, "");
    const corpoAdmin = '<p>Onboarding de <b>' + cand.nome + ' ' + (cand.sobrenome || "") + '</b> agendado.</p>'
      + '<div ' + caixa + '>'
      + '<p style="margin:0 0 8px 0;font-size:17px;font-family:Georgia,serif;">'
      + '<b style="color:#E8B4BC;">Data e hora:</b> ' + horario + '</p>'
      + _linha("E-mail", cand.email) + _linha("Função", cand.funcao)
      + '</div>'
      + _botao("ABRIR NO GOOGLE CALENDAR", linkAgenda)
      + '<hr style="border:none;border-top:1px solid #1F3D30;margin:32px 0 24px 0;">'
      + _botao("ACESSAR O PAINEL", ScriptApp.getService().getUrl() + "?view=painel");
    MailApp.sendEmail({ to: _destinatariosAdmin(), subject: "COB · Onboarding agendado — " + cand.nome,
      htmlBody: criarTemplateEmail("AGENDAMENTO CONCLUÍDO", corpoAdmin, { lgpd: false }) });

    return { success: true, message: "Agendado para " + horario };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function aprovarCandidato(sessao, id, categoria) {
  try {
    _exigirAdmin(sessao);
    const cand = getCandidatoById(id, categoria);
    if (!cand) return { success: false, message: "Cadastro não encontrado." };

    const sheet = SpreadsheetApp.openById(SS_ID_()).getSheetByName(cand.sheet);
    sheet.getRange(cand.linha, C.STATUS + 1).setValue("CONFIRMADO");

    const label = CATEGORIAS[cand.categoria] ? CATEGORIAS[cand.categoria].label : "membro";
    const corpo = '<p>Olá <b>' + cand.nome + '</b>,</p>'
      + '<p>Seu cadastro como <b>' + label + '</b> foi aprovado pela equipe do COB.</p>'
      + '<p style="color:#9BA89F;">Entramos em contato em breve com os próximos passos.</p>';
    const para = [cand.email].concat(cand.emailExtra ? [cand.emailExtra] : []).join(",");
    MailApp.sendEmail({ to: para, subject: "Cadastro aprovado — COB",
      htmlBody: criarTemplateEmail("CADASTRO APROVADO", corpo, { lgpd: true }) });

    return { success: true, message: "Cadastro aprovado." };
  } catch (e) {
    return { success: false, message: String(e.message || e) };
  }
}

// ============================================================
// DIAGNÓSTICO — rodar no editor antes de publicar
// ============================================================
function verificarInstalacao() {
  const r = [];
  const ok = function (n, f) { try { f(); r.push("OK    · " + n); } catch (e) { r.push("FALHA · " + n + " → " + e); } };

  ok("Planilha (leitura)", function () { SpreadsheetApp.openById(SS_ID_()).getName(); });
  ok("Planilha (escrita/abas)", function () { garantirAbasBase(); });
  ok("Drive (pasta raiz)", function () { DriveApp.getFolderById(FOLDER_ID_()).getName(); });
  ok("Calendar (leitura)", function () { CAL_().getEvents(new Date(), new Date(Date.now() + 86400000)); });
  ok("Calendar (escrita)", function () {
    const cal = CAL_();
    const t = new Date(Date.now() + 400 * 86400000);
    const ev = cal.createEvent("TESTE OPUS AI — pode apagar", t, new Date(t.getTime() + 1800000));
    ev.deleteEvent();
  });
  ok("Gmail / MailApp", function () {
    if (MailApp.getRemainingDailyQuota() <= 0) throw new Error("cota diária esgotada");
  });
  ok("Segredo HMAC", function () { if (!_tokenCandidato("TESTE")) throw new Error("token vazio"); });

  const slots = getAvailableTimes();
  r.push("INFO  · Horários disponíveis nos próximos dias: " + slots.length);
  if (slots.length) r.push("INFO  · Primeiro horário: " + slots[0].label);
  r.push("INFO  · Cota de e-mails restante hoje: " + MailApp.getRemainingDailyQuota());
  r.push("INFO  · Fuso do script: " + TZ_());
  try { r.push("INFO  · Agenda em uso: " + CAL_().getName() + " (" + CAL_().getId() + ")"); }
  catch (e) { r.push("FALHA · Agenda em uso → " + e); }
  r.push("INFO  · Pasta do Drive: " + FOLDER_ID_());
  r.push("INFO  · Notificações vão para: " + _destinatariosAdmin());
  r.push("INFO  · URL publicada: " + ScriptApp.getService().getUrl());

  const txt = r.join("\n");
  Logger.log(txt);
  return txt;
}

/**
 * DIAGNÓSTICO DO PAINEL — rodar no editor quando um cadastro não aparecer.
 * Mostra exatamente o que existe na planilha e o que o painel consegue ler.
 */
function diagnosticarPainel() {
  const r = [];
  const ss = SpreadsheetApp.openById(SS_ID_());
  r.push("Planilha: " + ss.getName());
  r.push("URL: " + ss.getUrl());
  r.push("Abas existentes: " + ss.getSheets().map(function (a) { return a.getName(); }).join(" | "));
  r.push("");

  Object.keys(CATEGORIAS).forEach(function (key) {
    const esperada = CATEGORIAS[key].sheet;
    const aba = _acharAba(ss, esperada);
    if (!aba) { r.push("[" + key + "] aba \"" + esperada + "\" NÃO ENCONTRADA"); return; }

    const dados = aba.getDataRange().getDisplayValues();
    r.push("[" + key + "] aba \"" + aba.getName() + "\" · " + dados.length + " linha(s) incluindo cabeçalho"
           + " · " + aba.getLastColumn() + " coluna(s)");
    r.push("        cabeçalho: " + (dados[0] || []).slice(0, 15).join(" | "));

    let validas = 0;
    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      if (!linha[C.ID] && !linha[C.EMAIL] && !linha[C.NOME]) continue;
      validas++;
      if (validas <= 5) {
        r.push("        linha " + (i + 1) + ": ID=" + (linha[C.ID] || "(vazio)")
               + " · " + (linha[C.NOME] || "?") + " " + (linha[C.SOBRENOME] || "")
               + " · " + (linha[C.EMAIL] || "?")
               + " · status=" + (linha[C.STATUS] || "(vazio)"));
      }
    }
    r.push("        >>> o painel exibiria " + validas + " cadastro(s) nesta categoria");
    r.push("");
  });

  r.push("Admins autorizados: " + _emailsAutorizados().join(", "));
  r.push("Notificações vão para: " + _destinatariosAdmin());
  r.push("Pasta do Drive em uso: " + FOLDER_ID_());
  const idCal = P_("COB_CALENDAR_ID");
  r.push("Agenda em uso: " + (idCal ? idCal : "agenda padrão da conta que autorizou o script"));
  r.push("URL publicada: " + ScriptApp.getService().getUrl());

  const txt = r.join("\n");
  Logger.log(txt);
  return txt;
}

/** Envia um e-mail de teste para o admin — confirma que a entrega funciona. */
function TESTE_email_de_notificacao() {
  MailApp.sendEmail({
    to: ADMIN_EMAIL_(),
    subject: "COB · Teste de entrega de e-mail",
    htmlBody: criarTemplateEmail("TESTE DE ENTREGA", "<p>Se você está lendo isto, o envio de e-mails da plataforma está funcionando.</p>")
  });
  return "Enviado para " + ADMIN_EMAIL_();
}
