/**
 * Plataforma COB - Conectando Orquestras Brasileiras
 * Versão Multi-Categoria (Voluntário / Embaixador / Representante de Orquestra)
 * Identidade visual: Verde Escuro + Rosa Claro
 * Admin: jr.conductor83@gmail.com
 */

const MAIN_FOLDER_ID = "1jz6vfaFTkRqNP7FvrT4ECDrl5l4DHh2C";
const SS_ID = "1NIzil2c1OUhKaYT6jUE1PaqilHWiMb8myHVpYqOuVts";
const ADMIN_EMAIL = "jr.conductor83@gmail.com";

const CATEGORIAS = {
  voluntario:    { sheet: "Voluntarios",    label: "Voluntário(a)",                folder: "VOLUNTARIOS" },
  embaixador:    { sheet: "Embaixadores",   label: "Embaixador(a)",                folder: "EMBAIXADORES" },
  representante: { sheet: "Representantes", label: "Representante de Orquestra",   folder: "REPRESENTANTES" }
};

const HEADERS_BASE = ["ID","DATA","NOME","SOBRENOME","EMAIL","EMAIL_EXTRA","WHATSAPP","FUNCAO","FOTO_URL","STATUS_AGENDA","CV_URL","EVENT_ID"];

// ============================================================
// SETUP - Garante todas as abas necessárias
// ============================================================
function garantirAbasBase() {
  const ss = SpreadsheetApp.openById(SS_ID);

  if (!ss.getSheetByName("Admins")) {
    let s = ss.insertSheet("Admins");
    s.appendRow(["E-MAILS AUTORIZADOS"]);
    s.getRange("A1").setFontWeight("bold");
    s.appendRow([ADMIN_EMAIL]);
  }

  if (!ss.getSheetByName("Estrutura_COB")) {
    let s = ss.insertSheet("Estrutura_COB");
    s.appendRow(["CARGO / GT", "LIDERANÇA"]);
    s.getRange("A1:B1").setFontWeight("bold");
    s.appendRow(["COORDENAÇÃO GERAL", "Gisely Nascimento"]);
    s.appendRow(["GT ORQUESTRANDO ARQUIVO", "Carô Tenório"]);
    s.appendRow(["GT GESTÃO E COMUNICAÇÃO", "Lideranças Ativas"]);
  }

  // Cria as 3 abas de categorias se não existirem
  Object.keys(CATEGORIAS).forEach(key => {
    const sheetName = CATEGORIAS[key].sheet;
    if (!ss.getSheetByName(sheetName)) {
      let s = ss.insertSheet(sheetName);
      s.appendRow(HEADERS_BASE);
      s.getRange(1, 1, 1, HEADERS_BASE.length).setFontWeight("bold");
    }
  });
}

// ============================================================
// AUTENTICAÇÃO ADMIN (OTP por e-mail)
// ============================================================
function solicitarAcessoAdmin(emailDigitado) {
  garantirAbasBase();
  const emailLimpo = emailDigitado.toLowerCase().trim();
  const ss = SpreadsheetApp.openById(SS_ID);
  const emailsAutorizados = ss.getSheetByName("Admins").getDataRange().getValues()
    .map(linha => linha[0].toString().toLowerCase().trim());
  if (!emailsAutorizados.includes(emailLimpo)) return { success: false, message: "E-mail não autorizado." };

  const codigoOTP = Math.floor(100000 + Math.random() * 900000).toString();
  CacheService.getScriptCache().put("OTP_" + emailLimpo, codigoOTP, 600);

  const htmlBody = `<div style="background:#0F2A20;color:#F5EFE6;padding:40px;border-top:4px solid #E8B4BC;font-family:sans-serif;text-align:center;border-radius:8px;"><h3 style="color:#E8B4BC;margin:0 0 20px 0;">Seu código de acesso COB:</h3><h1 style="letter-spacing:10px;color:#F5EFE6;margin:0;">${codigoOTP}</h1><p style="color:#9BA89F;font-size:0.85rem;margin-top:30px;">Este código expira em 10 minutos.</p></div>`;
  MailApp.sendEmail({to: emailLimpo, subject: "Acesso Plataforma COB", htmlBody: htmlBody});
  return { success: true };
}

function validarCodigoAdmin(emailDigitado, codigoDigitado) {
  const emailLimpo = emailDigitado.toLowerCase().trim();
  const codigoSalvo = CacheService.getScriptCache().get("OTP_" + emailLimpo);
  if (codigoSalvo && codigoSalvo === codigoDigitado.trim()) {
    CacheService.getScriptCache().remove("OTP_" + emailLimpo);
    return { success: true };
  }
  return { success: false, message: "Código inválido." };
}

// ============================================================
// ROTEAMENTO DA WEB APP
// ============================================================
function doGet(e) {
  let view = (e && e.parameter && e.parameter.view) ? e.parameter.view : 'user';
  let template = HtmlService.createTemplateFromFile('index');
  if (view === 'admin' && e.parameter.id) {
    const cat = e.parameter.categoria || null;
    template.candidato = getCandidatoById(e.parameter.id, cat);
    template.candidatoCategoria = (template.candidato && template.candidato.categoria) || cat || "voluntario";
    template.candidatoCategoriaLabel = template.candidatoCategoria && CATEGORIAS[template.candidatoCategoria] ? CATEGORIAS[template.candidatoCategoria].label : "";
  }
  template.appUrl = ScriptApp.getService().getUrl();
  template.viewMode = view;
  return template.evaluate()
    .setTitle('Plataforma COB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// DASHBOARD - Dados consolidados das 3 categorias
// ============================================================
function getDashboardData() {
  try {
    garantirAbasBase();
    const ss = SpreadsheetApp.openById(SS_ID);
    const candidaturas = {};
    const agenda = [];

    Object.keys(CATEGORIAS).forEach(key => {
      const sheet = ss.getSheetByName(CATEGORIAS[key].sheet);
      const data = sheet.getDataRange().getDisplayValues();
      const lista = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        let status = data[i][9] ? data[i][9].toString() : "PENDENTE";
        lista.push({
          id: data[i][0],
          data: data[i][1],
          nome: (data[i][2] || "") + " " + (data[i][3] || ""),
          email: data[i][4],
          whatsapp: data[i][6],
          funcao: data[i][7],
          foto: data[i][8],
          status: status,
          cv: data[i][10],
          categoria: key,
          categoriaLabel: CATEGORIAS[key].label
        });
        if (status.toUpperCase().includes("CONFIRMADO")) {
          let strData = status.replace(/CONFIRMADO:?/i, "").trim();
          let partes = strData.split(/ [aàAÀ]s /);
          agenda.push({
            titulo: "COB: Onboarding " + data[i][2] + " (" + CATEGORIAS[key].label + ")",
            data: partes[0] ? partes[0].trim() : strData,
            hora: partes[1] ? partes[1].trim() : "-",
            convidados: data[i][4],
            categoria: key
          });
        }
      }
      candidaturas[key] = lista.reverse();
    });

    const estrutura = ss.getSheetByName("Estrutura_COB").getDataRange().getDisplayValues().slice(1)
      .map(r => ({ cargo: r[0], nome: r[1] }));
    const admins = ss.getSheetByName("Admins").getDataRange().getDisplayValues().slice(1)
      .map(r => r[0]);

    return { success: true, candidaturas: candidaturas, estrutura: estrutura, admins: admins, agenda: agenda.reverse() };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// FORMATAÇÃO
// ============================================================
function formatarNome(texto) {
  if (!texto) return "";
  return texto.toLowerCase().split(/\s+/).map(word => {
    if (['de', 'da', 'do', 'dos', 'das', 'e'].includes(word)) return word;
    return word.charAt(0).toUpperCase() + word.slice(1);
  }).join(' ');
}

// ============================================================
// PROCESSAMENTO DO FORMULÁRIO (Base64)
// ============================================================
function processFormBase64(payload) {
  try {
    garantirAbasBase();
    const ss = SpreadsheetApp.openById(SS_ID);

    const categoria = (payload.categoria && CATEGORIAS[payload.categoria]) ? payload.categoria : "voluntario";
    const sheet = ss.getSheetByName(CATEGORIAS[categoria].sheet);

    const uniqueId = Utilities.getUuid().split('-')[0].toUpperCase();
    const rootFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    const fCat = getOrCreateSubfolder(rootFolder, CATEGORIAS[categoria].folder);
    const fFotos = getOrCreateSubfolder(fCat, "FOTOS_PERFIL");
    const fCvs = getOrCreateSubfolder(fCat, "CURRICULOS");

    const nomeLimpo = formatarNome(payload.nome ? payload.nome.trim() : "Candidato");
    const sobrenomeLimpo = formatarNome(payload.sobrenome ? payload.sobrenome.trim() : "");
    const identif = `${nomeLimpo}_${sobrenomeLimpo}_ID${uniqueId}`.toUpperCase().replace(/\s+/g, '_');

    let fotoUrl = "", cvUrl = "";

    try {
      if (payload.foto && payload.foto.data) {
        const ext = payload.foto.name.split('.').pop();
        const blob = Utilities.newBlob(Utilities.base64Decode(payload.foto.data), payload.foto.mimeType, `FOTO_${identif}.${ext}`);
        fotoUrl = fFotos.createFile(blob).getUrl();
      }
    } catch(e) { Logger.log("Erro Foto: " + e); }

    try {
      if (payload.cv && payload.cv.data) {
        const ext = payload.cv.name.split('.').pop();
        const blob = Utilities.newBlob(Utilities.base64Decode(payload.cv.data), payload.cv.mimeType, `CV_${identif}.${ext}`);
        cvUrl = fCvs.createFile(blob).getUrl();
      }
    } catch(e) { Logger.log("Erro CV: " + e); }

    const emailSafe = payload.email ? payload.email.toLowerCase() : "";
    const emailExtraSafe = payload.email_extra ? payload.email_extra.toLowerCase() : "";

    const temAgendamento = !!(payload.opt1 && payload.opt2 && payload.opt3);
    const statusInicial = temAgendamento
      ? `1: ${payload.opt1} | 2: ${payload.opt2} | 3: ${payload.opt3}`
      : "PENDENTE";

    sheet.appendRow([
      uniqueId,
      new Date(),
      nomeLimpo,
      sobrenomeLimpo,
      emailSafe,
      emailExtraSafe,
      "'" + payload.whatsapp,
      payload.funcao_cob,
      fotoUrl,
      statusInicial,
      cvUrl,
      ""
    ]);

    enviarNotificacoes({
      categoria: categoria,
      categoriaLabel: CATEGORIAS[categoria].label,
      nome: nomeLimpo,
      sobrenome: sobrenomeLimpo,
      funcao_cob: payload.funcao_cob,
      email: emailSafe,
      email_extra: emailExtraSafe,
      whatsapp: payload.whatsapp,
      opt1: payload.opt1,
      opt2: payload.opt2,
      opt3: payload.opt3
    }, uniqueId);

    return { success: true, id: uniqueId };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================================
// E-MAILS
// ============================================================
function criarTemplateEmail(titulo, conteudo) {
  return `<!DOCTYPE html><html><head><link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;600;700&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet"></head><body style="background:#0A1812;margin:0;padding:30px 10px;"><div style="font-family:'Inter',Arial,sans-serif;background:#122620;color:#F5EFE6;padding:40px;border:1px solid #1F3D30;border-radius:12px;max-width:800px;margin:0 auto;border-top:4px solid #E8B4BC;overflow-x:auto;"><h1 style="font-family:'Cormorant Garamond',serif;color:#E8B4BC;text-align:center;font-size:1.8rem;letter-spacing:2px;margin:0 0 10px 0;font-weight:600;">${titulo}</h1><div style="text-align:center;color:#9BA89F;font-size:0.7rem;letter-spacing:3px;margin-bottom:30px;">CONECTANDO ORQUESTRAS BRASILEIRAS</div><div style="line-height:1.6;font-size:15px;">${conteudo}</div></div></body></html>`;
}

function enviarNotificacoes(dados, id) {
  const appUrl = ScriptApp.getService().getUrl();
  const dataHoraAtual = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy 'às' HH:mm");
  const emailExtraHTML = dados.email_extra ? `<p style="margin:8px 0;"><b>E-mail Adicional:</b> ${dados.email_extra}</p>` : "";
  const temAgendamento = !!(dados.opt1 && dados.opt2 && dados.opt3);
  const funcaoHTML = dados.funcao_cob ? `<p style="margin:8px 0;"><b>Função Pretendida:</b> ${dados.funcao_cob}</p>` : "";
  const horariosHTML = temAgendamento
    ? `<p style="color:#E8B4BC;margin-bottom:5px;margin-top:20px;"><b>Horários Sugeridos:</b></p><ul style="margin-top:0;color:#9BA89F;padding-left:20px;"><li style="margin-bottom:5px;">${dados.opt1}</li><li style="margin-bottom:5px;">${dados.opt2}</li><li style="margin-bottom:5px;">${dados.opt3}</li></ul>`
    : "";
  const rodapeCand = temAgendamento
    ? `<p>A nossa equipe analisará a sua disponibilidade e enviará a confirmação da data e horário exatos nas próximas 48h.</p>`
    : `<p>A nossa equipe entrará em contato em breve com os próximos passos.</p>`;

  const conteudoCand = `<p>Olá <b>${dados.nome} ${dados.sobrenome}</b>,</p><p>Este é o comprovante oficial de que recebemos os seus dados como <b>${dados.categoriaLabel}</b> no Conectando Orquestras Brasileiras (COB).</p><div style="background:#0F2A20;padding:25px;border-radius:8px;margin:25px 0;border-left:4px solid #E8B4BC;"><h3 style="color:#E8B4BC;margin-top:0;font-family:'Cormorant Garamond',serif;font-size:1.2rem;">RESUMO DOS DADOS ENVIADOS</h3><p style="margin:8px 0;"><b>Categoria:</b> ${dados.categoriaLabel}</p><p style="margin:8px 0;"><b>ID de Cadastro:</b> ${id}</p><p style="margin:8px 0;"><b>Data/Hora do Envio:</b> ${dataHoraAtual}</p>${funcaoHTML}<p style="margin:8px 0;"><b>E-mail Principal:</b> ${dados.email}</p>${emailExtraHTML}<p style="margin:8px 0;"><b>WhatsApp:</b> ${dados.whatsapp}</p>${horariosHTML}</div>${rodapeCand}`;
  MailApp.sendEmail({to: dados.email, subject: "Recibo de Cadastro COB - " + dados.categoriaLabel + " (ID: " + id + ")", htmlBody: criarTemplateEmail("CADASTRO REALIZADO!", conteudoCand)});

  const acaoAdmin = temAgendamento
    ? `<p style="font-size:1rem;color:#9BA89F;">O candidato sugeriu 3 opções de horários. Para confirmar o agendamento no Google Calendar, acesse o painel:</p><div style="text-align:center;margin:35px 0 20px 0;"><a href="${appUrl}?view=admin&id=${id}&categoria=${dados.categoria}" style="background:#E8B4BC;color:#0A1812;padding:15px 35px;text-decoration:none;border-radius:50px;font-family:'Cormorant Garamond',serif;font-weight:600;font-size:16px;display:inline-block;letter-spacing:1px;">APROVAR AGENDAMENTO</a></div>`
    : `<p style="font-size:1rem;color:#9BA89F;">Acesse o painel administrativo para mais detalhes do cadastro.</p>`;
  const conteudoAdmin = `<h2 style="color:#F5EFE6;font-family:'Inter',sans-serif;margin:0 0 15px 0;">Novo cadastro: <span style="color:#E8B4BC;">${dados.nome} ${dados.sobrenome}</span></h2><p style="font-size:1rem;color:#9BA89F;margin-bottom:20px;">Categoria: <b style="color:#F5EFE6;">${dados.categoriaLabel}</b></p>${funcaoHTML}<hr style="border-top:1px solid #1F3D30;margin:25px 0;">${acaoAdmin}`;
  MailApp.sendEmail({to: ADMIN_EMAIL, subject: "AÇÃO REQUERIDA: Novo cadastro COB (" + dados.categoriaLabel + ") - " + dados.nome, htmlBody: criarTemplateEmail("APROVAÇÃO PENDENTE", conteudoAdmin)});
}

// ============================================================
// CRUD - Admins / Estrutura / Candidatos
// ============================================================
function salvarNovaEstrutura(arr) {
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName("Estrutura_COB");
  sheet.clear().appendRow(["CARGO / GT", "LIDERANÇA"]);
  arr.forEach(item => { if(item.cargo) sheet.appendRow([item.cargo, item.nome]); });
  return "Sucesso";
}

function addAdminEmail(e) {
  SpreadsheetApp.openById(SS_ID).getSheetByName("Admins").appendRow([e.toLowerCase().trim()]);
  return "OK";
}

function removeAdminEmail(e) {
  const s = SpreadsheetApp.openById(SS_ID).getSheetByName("Admins");
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] == e) s.deleteRow(i+1);
  return "OK";
}

function deletarCandidato(id, categoria) {
  const ss = SpreadsheetApp.openById(SS_ID);
  if (categoria && CATEGORIAS[categoria]) {
    return _deletarEm(ss.getSheetByName(CATEGORIAS[categoria].sheet), id);
  }
  // Busca em todas as abas
  for (let key of Object.keys(CATEGORIAS)) {
    let r = _deletarEm(ss.getSheetByName(CATEGORIAS[key].sheet), id);
    if (r === "OK") return "OK";
  }
  return "NOT_FOUND";
}

function _deletarEm(sheet, id) {
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] == id) { sheet.deleteRow(i+1); return "OK"; }
  }
  return null;
}

// ============================================================
// AGENDA - Slots disponíveis no Google Calendar
// ============================================================
function getAvailableTimes() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let config = ss.getSheetByName("Configuracoes_Agenda");
  if (!config) {
    config = ss.insertSheet("Configuracoes_Agenda");
    config.appendRow(["10:00","18:00"]);
  }
  const start = parseInt(config.getRange("A2").getDisplayValue()) || 10;
  const end = parseInt(config.getRange("B2").getDisplayValue()) || 18;
  const cal = CalendarApp.getDefaultCalendar();
  const slots = [];
  const dias = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  for (let i = 1; i <= 14; i++) {
    let d = new Date(new Date().getTime() + (i * 24 * 60 * 60 * 1000));
    if (d.getDay() !== 0 && d.getDay() !== 6) {
      for (let h = start; h < end; h++) {
        let s = new Date(d.setHours(h, 0, 0, 0));
        let evEnd = new Date(d.setHours(h + 1, 0, 0, 0));
        if (cal.getEvents(s, evEnd).length === 0) {
          slots.push({
            id: Utilities.formatDate(s, "GMT-3", "yyyyMMddHH"),
            label: `${Utilities.formatDate(s, "GMT-3", "dd/MM")} (${dias[s.getDay()]}) às ${Utilities.formatDate(s, "GMT-3", "HH:mm")}`,
            data: Utilities.formatDate(s, "GMT-3", "dd/MM"),
            diaSemana: dias[s.getDay()],
            hora: Utilities.formatDate(s, "GMT-3", "HH:mm"),
            periodo: h < 12 ? "manha" : (h < 18 ? "tarde" : "noite")
          });
        }
      }
    }
  }
  return slots;
}

function parseDate(str) {
  const partes = str.split(' às ');
  const dP = partes[0].trim().split(' ')[0].split('/');
  const hP = partes[1].trim().split(':');
  return new Date(new Date().getFullYear(), parseInt(dP[1], 10) - 1, parseInt(dP[0], 10), parseInt(hP[0], 10), parseInt(hP[1], 10));
}

function getCandidatoById(id, categoria) {
  const ss = SpreadsheetApp.openById(SS_ID);
  if (categoria && CATEGORIAS[categoria]) {
    return _buscarEm(ss.getSheetByName(CATEGORIAS[categoria].sheet), id, categoria);
  }
  for (let key of Object.keys(CATEGORIAS)) {
    let r = _buscarEm(ss.getSheetByName(CATEGORIAS[key].sheet), id, key);
    if (r) return r;
  }
  return null;
}

function _buscarEm(sheet, id, categoria) {
  if (!sheet) return null;
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] == id) {
      return {
        id: id, nome: d[i][2], email: d[i][4], funcao: d[i][7],
        opcoes: d[i][9], linha: i+1, sheet: sheet.getSheetName(), categoria: categoria
      };
    }
  }
  return null;
}

function confirmarAgendamento(id, horario, categoria) {
  const cand = getCandidatoById(id, categoria);
  if (!cand) return "Candidato não encontrado.";

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(cand.sheet);
  const dataInicio = parseDate(horario);
  const tituloEvento = `COB: Onboarding ${cand.nome}`;

  const event = CalendarApp.getDefaultCalendar().createEvent(
    tituloEvento,
    dataInicio,
    new Date(dataInicio.getTime() + 3600000),
    { guests: cand.email, sendInvites: true }
  );

  const linkCalendarAdmin = `https://calendar.google.com/calendar/event?eid=${Utilities.base64Encode(event.getId().split('@')[0] + " " + ADMIN_EMAIL).replace(/=/g, '')}`;

  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] == id) {
      sheet.getRange(i+1, 10).setValue("CONFIRMADO: " + horario);
      sheet.getRange(i+1, 12).setValue(event.getId());
    }
  }

  const confCand = `<p>Olá <b>${cand.nome}</b>,</p><p>Sua sessão de onboarding com a equipe do COB foi confirmada com sucesso!</p><div style="background:#0F2A20;padding:25px;border-radius:8px;margin:25px 0;border-left:4px solid #E8B4BC;"><p style="margin:0;font-size:1.1rem;font-family:'Cormorant Garamond',serif;"><b>Data e Hora:</b> <span style="color:#E8B4BC;">${horario}</span></p></div><p style="color:#9BA89F;">O evento oficial já foi adicionado à sua agenda com o título <b>"${tituloEvento}"</b>.</p>`;
  MailApp.sendEmail({to: cand.email, subject: "Onboarding Confirmado - COB", htmlBody: criarTemplateEmail("VÍDEO-CHAMADA AGENDADA!", confCand)});

  const confAdmin = `<p>O onboarding de <b>${cand.nome}</b> foi agendado na sua agenda.</p><div style="background:#0F2A20;padding:25px;border-radius:8px;margin:25px 0;border-left:4px solid #E8B4BC;"><p style="margin:0;font-size:1.1rem;font-family:'Cormorant Garamond',serif;"><b>Data e Hora:</b> <span style="color:#E8B4BC;">${horario}</span></p></div><div style="text-align:center;margin:35px 0 15px 0;"><a href="${linkCalendarAdmin}" style="background:#E8B4BC;color:#0A1812;padding:15px 35px;text-decoration:none;border-radius:50px;font-family:'Cormorant Garamond',serif;font-weight:600;display:inline-block;letter-spacing:1px;">VER NO MEU GOOGLE CALENDAR</a></div>`;
  MailApp.sendEmail({to: ADMIN_EMAIL, subject: "Onboarding Agendado: " + cand.nome, htmlBody: criarTemplateEmail("AGENDAMENTO CONCLUÍDO!", confAdmin)});

  return "Agendado com sucesso para " + horario;
}

function aprovarCandidato(id, categoria) {
  const cand = getCandidatoById(id, categoria);
  if (!cand) return "Candidato não encontrado.";

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(cand.sheet);
  const d = sheet.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] == id) {
      sheet.getRange(i+1, 10).setValue("CONFIRMADO");
    }
  }

  const confCand = `<p>Olá <b>${cand.nome}</b>,</p><p>Seu cadastro como <b>${CATEGORIAS[categoria].label}</b> foi aprovado pela equipe do COB!</p><p style="color:#9BA89F;">Em breve entraremos em contato com os próximos passos.</p>`;
  MailApp.sendEmail({to: cand.email, subject: "Cadastro Aprovado - COB", htmlBody: criarTemplateEmail("CADASTRO APROVADO!", confCand)});

  return "Aprovado com sucesso";
}

function getOrCreateSubfolder(p, n) {
  let f = p.getFoldersByName(n);
  return f.hasNext() ? f.next() : p.createFolder(n);
}