/* global Excel, Office */

// ============================================================
// HANDLERS GLOBAIS DE ERRO
// ============================================================

window.addEventListener("error", function (event) {
  console.error("[GLOBAL] Erro não capturado:", event.error || event.message);
});

window.addEventListener("unhandledrejection", function (event) {
  console.error("[GLOBAL] Promise rejeitada sem catch:", event.reason);
});

// ============================================================
// ESTADO GLOBAL
// ============================================================

var varasData = [];
var competenciasData = [];
var vitimasData = [];

var entrevistadoresData = [];
var todosPolos = [];

var vitimasCache = [];

var audienciasCache = [];

var sugestoesTimeout = null;
var nomeSugestoesTimeout = null;
var modoEdicaoAudiencia = false;
var modoEdicaoEntrevistador = false;
var audienciaEmEdicao = null;
var entrevistadorEmEdicao = null;

var polosSelecionados = [];
var diasEspecificosSelecionados = [];

// Entrevistadores selecionados para a audiência atual
var entrevistadoresSelecionadosAudiencia = [];

var ATIVAR_DISTRIBUICAO = true;
var PROCESSO_SEPARADORES = [7, 10, 15, 17];
var SENHA_ABA = "321bob321";

// ============================================================
// INICIALIZAÇÃO
// ============================================================

Office.onReady(async function () {
  configurarTabs();
  configurarListenersAudiencia();
  configurarListenersEntrevistador();

  await carregarTudoInicial();
  atualizarEstadoBotao();
  atualizarEstadoBotaoEntrevistador();
});

// ============================================================
// TABS
// ============================================================

function configurarTabs() {
  var tabs = document.querySelectorAll(".tab");
  tabs.forEach(function (tab) {
    tab.addEventListener("click", function () {
      var alvo = tab.getAttribute("data-tab");
      if (!alvo) return;
      tabs.forEach(function (t) {
        t.classList.remove("active");
      });
      tab.classList.add("active");
      document.getElementById("panelAudiencias").classList.add("hidden");
      document.getElementById("panelEntrevistadores").classList.add("hidden");
      document
        .getElementById("panel" + (alvo === "audiencias" ? "Audiencias" : "Entrevistadores"))
        .classList.remove("hidden");
    });
  });
}

// ============================================================
// CONFIGURAÇÃO DE LISTENERS
// ============================================================

function configurarListenersAudiencia() {
  var btnSalvar = document.getElementById("btnSalvar");
  btnSalvar.addEventListener("click", function () {
    tryCatch(salvarData);
  });

  document.getElementById("audienciaJaExiste").addEventListener("change", onAudienciaJaExisteChange);
  document.getElementById("audienciaEditSelect").addEventListener("change", onAudienciaEditSelectChange);

  var timeInput = document.getElementById("timeInput");
  timeInput.addEventListener("input", function () {
    timeInput.value = formatarHoraInput(timeInput.value);
    atualizarPreview();
    atualizarEstadoBotao();
  });

  var dataInput = document.getElementById("dataInput");
  dataInput.addEventListener("change", function () {
    onDataOuPoloChange();
  });

  var poloSelect = document.getElementById("poloSelect");
  poloSelect.addEventListener("change", function () {
    onDataOuPoloChange();
  });

  document.getElementById("varaSelect").addEventListener("change", onVaraSelectChange);
  document.getElementById("novaVaraCheck").addEventListener("change", onNovaVaraChange);
  document.getElementById("varaNomeInput").addEventListener("input", atualizarEstadoBotao);
  document.getElementById("varaContatoInput").addEventListener("input", atualizarEstadoBotao);
  document.getElementById("competenciaSelect").addEventListener("change", atualizarEstadoBotao);
  document.getElementById("novaCompetenciaCheck").addEventListener("change", onNovaCompetenciaChange);
  document.getElementById("competenciaNomeInput").addEventListener("input", atualizarEstadoBotao);

  var processoInput = document.getElementById("processoInput");
  processoInput.addEventListener("input", function () {
    processoInput.value = formatarProcesso(processoInput.value);
    atualizarEstadoBotao();
  });

  document.getElementById("vitimaNomesInput").addEventListener("input", function () {
    agendarAtualizacaoSugestoes();
    atualizarEstadoBotao();
  });
  document.getElementById("vitimaIdadesInput").addEventListener("input", atualizarEstadoBotao);

  var adaptacaoRadios = document.querySelectorAll('input[name="adaptacaoRadio"]');
  adaptacaoRadios.forEach(function (r) {
    r.addEventListener("change", onAdaptacaoRadioChange);
  });
  document.getElementById("adaptacaoDescricaoInput").addEventListener("input", atualizarEstadoBotao);

  document.getElementById("statusSelect").addEventListener("change", onStatusChange);
  document.getElementById("statusMotivoInput").addEventListener("input", atualizarEstadoBotao);

  document.getElementById("entrevistadorSelect").addEventListener("change", onEntrevistadorSelectChange);

  var btnAddEnt = document.getElementById("btnAddEntrevistador");
  if (btnAddEnt) btnAddEnt.addEventListener("click", onAddEntrevistadorClick);
}

function configurarListenersEntrevistador() {
  document.getElementById("btnSalvarEntrevistador").addEventListener("click", function () {
    tryCatch(salvarEntrevistador);
  });

  document.getElementById("entrevistadorJaExiste").addEventListener("change", onEntrevistadorJaExisteChange);
  document.getElementById("entrevistadorEditSelect").addEventListener("change", onEntrevistadorEditSelectChange);

  var nomeInput = document.getElementById("entNomeInput");
  nomeInput.addEventListener("input", function () {
    agendarSugestoesNome();
    atualizarEstadoBotaoEntrevistador();
  });

  document.getElementById("entEmailInput").addEventListener("input", atualizarEstadoBotaoEntrevistador);
  document.getElementById("entTipoSelect").addEventListener("change", atualizarEstadoBotaoEntrevistador);
  document.getElementById("entPrioridadeSelect").addEventListener("change", atualizarEstadoBotaoEntrevistador);
  document.getElementById("entDisponibilidadeSelect").addEventListener("change", atualizarEstadoBotaoEntrevistador);
  document.getElementById("entEspecialidadeInput").addEventListener("input", atualizarEstadoBotaoEntrevistador);
  document.getElementById("entLotacaoInput").addEventListener("input", atualizarEstadoBotaoEntrevistador);

  document.getElementById("btnAddPolo").addEventListener("click", onAddPoloClick);

  var diasSemana = document.querySelectorAll("#entDiasSemana input");
  diasSemana.forEach(function (cb) {
    cb.addEventListener("change", function () {
      if (cb.id === "entTodosDiasCheck") {
        var todos = cb.checked;
        var outros = document.querySelectorAll("#entDiasSemana input:not(#entTodosDiasCheck)");
        outros.forEach(function (o) {
          o.checked = todos;
        });
      }
      atualizarEstadoBotaoEntrevistador();
    });
  });

  document.getElementById("btnAddDiaEspecifico").addEventListener("click", onAddDiaEspecificoClick);

  document.getElementById("entObsInput").addEventListener("input", atualizarEstadoBotaoEntrevistador);
}

// ============================================================
// HELPERS DE FORMATAÇÃO
// ============================================================

function padLeft(valor, tamanho) {
  var r = valor;
  while (r.length < tamanho) r = "0" + r;
  return r;
}

function pad2(n) {
  var s = String(n);
  return s.length >= 2 ? s : "0" + s;
}

function converterParaBR(dataISO) {
  var p = dataISO.split("-");
  if (p.length !== 3) return dataISO;
  return p[2] + "/" + p[1] + "/" + p[0];
}

function converterParaISO(dataBR) {
  var p = dataBR.split("/");
  if (p.length !== 3) return "";
  return p[2] + "-" + p[1] + "-" + p[0];
}

function dateObjFromISO(iso) {
  var p = iso.split("-");
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function normalizarId(valor) {
  var s = String(valor).trim();
  if (/^\d+$/.test(s)) return String(parseInt(s, 10));
  return s;
}

function dateParaSerialExcel(d) {
  var utcMs = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round(utcMs / 86400000) + 25569;
}

function excelSerialParaBR(serial) {
  var d = new Date((serial - 25569) * 86400000);
  return pad2(d.getUTCDate()) + "/" + pad2(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear();
}

function normalizarDataBR(raw) {
  if (raw === null || raw === undefined || raw === "") return "";
  if (raw instanceof Date) return pad2(raw.getDate()) + "/" + pad2(raw.getMonth() + 1) + "/" + raw.getFullYear();
  if (typeof raw === "number") return excelSerialParaBR(raw);
  return String(raw).trim();
}

function formatarHoraInput(valor) {
  var d = valor.replace(/\D/g, "").substring(0, 4);
  if (d.length <= 2) return d;
  return d.substring(0, 2) + ":" + d.substring(2);
}

function calcularIntervalo(horaFormatada) {
  var p = horaFormatada.split(":");
  if (p.length !== 2 || p[0].length !== 2 || p[1].length !== 2) return "";
  var h = parseInt(p[0], 10),
    m = parseInt(p[1], 10);
  if (isNaN(h) || isNaN(m) || h < 1 || h > 23 || m < 0 || m > 59) return "";
  return (
    padLeft(String(h - 1), 2) + ":" + padLeft(String(m), 2) + "/" + padLeft(String(h), 2) + ":" + padLeft(String(m), 2)
  );
}

function formatarProcesso(valor) {
  var d = valor.replace(/\D/g, "").substring(0, 20);
  var r = "";
  for (var i = 0; i < d.length; i++) {
    if (PROCESSO_SEPARADORES.indexOf(i) !== -1) r += i === 7 ? "-" : ".";
    r += d[i];
  }
  return r;
}

function extrairDigitosProcesso(valor) {
  return valor.replace(/\D/g, "");
}

function calcularDiaSemana(dia, mes, ano) {
  var data = new Date(ano, mes - 1, dia);
  var nomes = ["domingo", "2ª feira", "3ª feira", "4ª feira", "5ª feira", "6ª feira", "sábado"];
  return nomes[data.getDay()];
}

function converterStringParaData(dataStr) {
  if (!dataStr) return null;
  var p = dataStr.split("/");
  if (p.length !== 3) return null;
  return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
}

function isEntrevistadorCadExterno(tipo) {
  if (!tipo) return false;
  return (
    tipo
      .toUpperCase()
      .replace(/[\s\-_]+/g, "")
      .trim() === "CADEXTERNO"
  );
}

// ============================================================
// HELPERS DE DIAS
// ============================================================

function parsearDias(diasStr) {
  var semana = new Set();
  var especificos = [];
  var semanaSet = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"];

  var partes = (diasStr || "")
    .split(/[,;]/)
    .map(function (s) {
      return s.trim();
    })
    .filter(Boolean);

  for (var i = 0; i < partes.length; i++) {
    var p = partes[i];
    var normalizado = p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
    if (semanaSet.indexOf(normalizado) !== -1) {
      semana.add(normalizado);
    } else if (p.toLowerCase() === "todos" || p.toLowerCase() === "a combinar") {
      semanaSet.slice(0, 5).forEach(function (d) {
        semana.add(d);
      });
    } else if (p.toLowerCase().match(/^(1ª|2ª|3ª|4ª|5ª|última)\s+(segunda|terça|quarta|quinta|sexta)-feira$/i)) {
      especificos.push(p.toLowerCase());
    } else {
      especificos.push("?" + p);
    }
  }
  return { semana: semana, especificos: especificos };
}

function verificarDiaBate(diasStr, dataStr) {
  var parsed = parsearDias(diasStr);
  var semana = parsed.semana;
  var especificos = parsed.especificos;
  if (semana.size === 0 && especificos.length === 0) return { ok: true, motivo: "" };

  var data = converterStringParaData(dataStr);
  if (!data) return { ok: true, motivo: "" };

  var nomes = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  var diaSemana = nomes[data.getDay()];

  if (semana.has(diaSemana)) return { ok: true, motivo: "" };

  var diaDoMes = data.getDate();
  var ocorrencia = Math.ceil(diaDoMes / 7);
  var ultimoDiaDoMes = new Date(data.getFullYear(), data.getMonth() + 1, 0).getDate();
  var isUltima = diaDoMes + 7 > ultimoDiaDoMes;

  for (var i = 0; i < especificos.length; i++) {
    var esp = especificos[i];
    if (esp.startsWith("?")) continue;
    var match = esp.match(/^(1ª|2ª|3ª|4ª|5ª|última)\s+(segunda|terça|quarta|quinta|sexta)-feira$/i);
    if (!match) continue;
    var ordem = match[1].toLowerCase();
    var diaNome = match[2].toLowerCase();
    var diaNomeCapitalizado = diaNome.charAt(0).toUpperCase() + diaNome.slice(1);
    if (diaNomeCapitalizado !== diaSemana) continue;

    if (ordem === "última") {
      if (isUltima) return { ok: true, motivo: "" };
    } else {
      var alvo = parseInt(ordem.charAt(0));
      if (ocorrencia === alvo) return { ok: true, motivo: "" };
    }
  }

  return { ok: false, motivo: "Não trabalha neste dia (" + diaSemana + ", " + dataStr + ")" };
}

// ============================================================
// HELPERS DE ADAPTAÇÃO
// ============================================================

function coletarAdaptacao() {
  var radios = document.querySelectorAll('input[name="adaptacaoRadio"]');
  var tipo = "";
  radios.forEach(function (r) {
    if (r.checked) tipo = r.value;
  });
  if (tipo === "") return null;
  var desc = document.getElementById("adaptacaoDescricaoInput").value.trim();
  return { tipo: tipo, descricao: desc };
}

function formatarAdaptacaoFinal(obj) {
  return obj.tipo === "SIM" ? "SIM, " + obj.descricao : "NÃO";
}

function preencherAdaptacaoRadio(valor) {
  var radios = document.querySelectorAll('input[name="adaptacaoRadio"]');
  var inp = document.getElementById("adaptacaoDescricaoInput");
  var v = (valor || "").trim();
  var tipo = "",
    desc = "";
  if (v === "" || v.toUpperCase() === "NÃO" || v.toUpperCase() === "NAO") tipo = v === "" ? "" : "NÃO";
  else if (v.toUpperCase() === "SIM") tipo = "SIM";
  else if (v.toUpperCase().indexOf("SIM,") === 0) {
    tipo = "SIM";
    desc = v.substring(4).trim();
  } else if (v.toUpperCase().indexOf("SIM ") === 0) {
    tipo = "SIM";
    desc = v.substring(4).trim();
  } else {
    tipo = "SIM";
    desc = v;
  }
  radios.forEach(function (r) {
    r.checked = r.value.toUpperCase() === tipo;
  });
  if (inp) {
    if (tipo === "SIM") {
      inp.value = desc;
      inp.classList.remove("hidden");
    } else {
      inp.value = "";
      inp.classList.add("hidden");
    }
  }
}

function limparAdaptacaoRadio() {
  document.querySelectorAll('input[name="adaptacaoRadio"]').forEach(function (r) {
    r.checked = false;
  });
  var inp = document.getElementById("adaptacaoDescricaoInput");
  if (inp) {
    inp.value = "";
    inp.classList.add("hidden");
  }
}

function onAdaptacaoRadioChange() {
  var radios = document.querySelectorAll('input[name="adaptacaoRadio"]');
  var inp = document.getElementById("adaptacaoDescricaoInput");
  if (!inp) return;
  var tipo = "";
  radios.forEach(function (r) {
    if (r.checked) tipo = r.value;
  });
  if (tipo === "SIM") inp.classList.remove("hidden");
  else {
    inp.classList.add("hidden");
    inp.value = "";
  }
  atualizarEstadoBotao();
}

// ============================================================
// HELPERS DE STATUS
// ============================================================

function coletarStatus() {
  var s = document.getElementById("statusSelect");
  var m = document.getElementById("statusMotivoInput");
  return { tipo: s ? s.value : "PENDENTE", motivo: m ? m.value.trim() : "" };
}

function formatarStatusFinal(obj) {
  if (obj.tipo === "CANCELADO" || obj.tipo === "NÃO REALIZADO") return obj.tipo + ", " + obj.motivo;
  return obj.tipo;
}

function preencherStatusSelect(valor) {
  var sel = document.getElementById("statusSelect");
  var mot = document.getElementById("statusMotivoInput");
  if (!sel || !mot) return;
  var v = (valor || "").trim().toUpperCase();
  var tipo = "PENDENTE",
    motivo = "";
  var limparMotivo = function (s) {
    var r = s.trim();
    if (r.indexOf(",") === 0) r = r.substring(1).trim();
    else if (r.indexOf(":") === 0) r = r.substring(1).trim();
    return r;
  };
  if (v.indexOf("CANCELADO") === 0) {
    tipo = "CANCELADO";
    motivo = limparMotivo(v.substring("CANCELADO".length));
  } else if (v.indexOf("NÃO REALIZADO") === 0) {
    tipo = "NÃO REALIZADO";
    motivo = limparMotivo(v.substring("NÃO REALIZADO".length));
  } else if (v.indexOf("NAO REALIZADO") === 0) {
    tipo = "NÃO REALIZADO";
    motivo = limparMotivo(v.substring("NAO REALIZADO".length));
  } else if (v === "REALIZADO") {
    tipo = "REALIZADO";
  } else if (
    v === "CONFIRMADO" ||
    v.indexOf("ENVIAR CONFIRMAÇÃO") === 0 ||
    v === "ENVIAR CONFIRMACAO" ||
    v === "CONVOCAR"
  ) {
    tipo = "CONFIRMADO";
  } else {
    tipo = "PENDENTE";
  }
  sel.value = tipo;
  if (tipo === "CANCELADO" || tipo === "NÃO REALIZADO") {
    mot.value = motivo;
    mot.classList.remove("hidden");
  } else {
    mot.value = "";
    mot.classList.add("hidden");
  }
}

function onStatusChange() {
  var sel = document.getElementById("statusSelect");
  var mot = document.getElementById("statusMotivoInput");
  if (!sel || !mot) return;
  if (sel.value === "CANCELADO" || sel.value === "NÃO REALIZADO") mot.classList.remove("hidden");
  else {
    mot.classList.add("hidden");
    mot.value = "";
  }
  atualizarEstadoBotao();
}

// ============================================================
// RANKING DE ENTREVISTADORES
// ============================================================

function calcularScore(e, poloAudi, dataAudi) {
  var polos = e.polos
    .split(/[,;]/)
    .map(function (s) {
      return s.trim();
    })
    .filter(Boolean);
  var poloBate = polos.some(function (p) {
    return p.toUpperCase() === poloAudi.toUpperCase();
  });

  var diaResult = verificarDiaBate(e.dias, dataAudi);
  var diaBate = diaResult.ok;

  var disponivel = e.disponibilidade.toUpperCase() === "DISPONÍVEL";

  if (!poloBate) {
    return { tipo: "red", score: -999, motivo: "Polo não bate (atua em: " + (polos.join(", ") || "—") + ")" };
  }
  if (!diaBate) {
    return { tipo: "red", score: -999, motivo: diaResult.motivo };
  }
  if (!disponivel) {
    return { tipo: "red", score: -999, motivo: "Disponibilidade: " + e.disponibilidade };
  }

  var score = 0;
  score += Math.min(e.qtdCancelado * 10, 30);
  score -= e.qtdMarcado * 4;
  score -= e.qtdRealizado * 2;

  if (e.tipo === "INTERNO" || e.tipo === "RESIDENTE") score += 10;
  if (e.prioridade === "OBRIGATÓRIO") score += 5;

  return { tipo: "green", score: score, motivo: "" };
}

function atualizarSelectEntrevistador() {
  var sel = document.getElementById("entrevistadorSelect");
  if (!sel) return;

  var valorAtual = sel.value;
  var dataISO = document.getElementById("dataInput").value;
  var polo = document.getElementById("poloSelect").value;

  sel.innerHTML = "";
  var def = document.createElement("option");
  def.value = "";
  def.textContent = "Selecione o entrevistador...";
  sel.appendChild(def);

  if (!dataISO || !polo) {
    var ordenados = entrevistadoresData.slice().sort(function (a, b) {
      return a.nome.localeCompare(b.nome);
    });
    for (var i = 0; i < ordenados.length; i++) {
      var e = ordenados[i];
      var o = document.createElement("option");
      o.value = e.nome;
      o.textContent = e.nome;
      sel.appendChild(o);
    }
  } else {
    var dataBR = converterParaBR(dataISO);
    var avaliados = entrevistadoresData.map(function (e) {
      return { ent: e, res: calcularScore(e, polo, dataBR) };
    });

    var verdes = avaliados
      .filter(function (x) {
        return x.res.tipo === "green";
      })
      .sort(function (a, b) {
        return b.res.score - a.res.score;
      })
      .slice(0, 5);

    var idsVerdes = new Set(
      verdes.map(function (x) {
        return x.ent.nome;
      })
    );

    var restantes = avaliados
      .filter(function (x) {
        return !idsVerdes.has(x.ent.nome);
      })
      .sort(function (a, b) {
        return a.ent.nome.localeCompare(b.ent.nome);
      });

    for (var i2 = 0; i2 < verdes.length; i2++) {
      var x = verdes[i2];
      var o2 = document.createElement("option");
      o2.value = x.ent.nome;
      o2.textContent = "✓ " + x.ent.nome;
      o2.style.color = "#2e7d32";
      o2.style.fontWeight = "600";
      sel.appendChild(o2);
    }

    if (verdes.length > 0) {
      var sep = document.createElement("option");
      sep.disabled = true;
      sep.textContent = "────── OUTROS ──────";
      sep.style.color = "#999";
      sel.appendChild(sep);
    }

    for (var i3 = 0; i3 < restantes.length; i3++) {
      var x2 = restantes[i3];
      var o3 = document.createElement("option");
      o3.value = x2.ent.nome;
      o3.textContent = (x2.res.tipo === "red" ? "⚠ " : "") + x2.ent.nome;
      if (x2.res.tipo === "red") o3.style.color = "#c62828";
      sel.appendChild(o3);
    }
  }

  if (valorAtual) {
    var ainda = Array.from(sel.options).some(function (o) {
      return o.value === valorAtual;
    });
    if (ainda) sel.value = valorAtual;
  }
}

// ============================================================
// ATUALIZAÇÕES DE UI - AUDIÊNCIA
// ============================================================

function atualizarPreview() {
  var t = document.getElementById("timeInput");
  var p = document.getElementById("timePreview");
  if (!t || !p) return;
  p.className = "time-preview";
  var v = t.value;
  if (!v || v.length < 5) {
    p.textContent = "";
    return;
  }
  var i = calcularIntervalo(v);
  if (i === "") {
    p.textContent = "⚠️ Hora inválida";
    p.className = "time-preview erro";
    return;
  }
  p.textContent = "→ " + i;
}

function atualizarEstadoBotao() {
  var btn = document.getElementById("btnSalvar");
  var dataInput = document.getElementById("dataInput");
  var timeInput = document.getElementById("timeInput");
  var poloSelect = document.getElementById("poloSelect");
  var novaVaraCheck = document.getElementById("novaVaraCheck");
  var varaSelect = document.getElementById("varaSelect");
  var varaNomeInput = document.getElementById("varaNomeInput");
  var varaContatoInput = document.getElementById("varaContatoInput");
  var novaCompetenciaCheck = document.getElementById("novaCompetenciaCheck");
  var competenciaSelect = document.getElementById("competenciaSelect");
  var competenciaNomeInput = document.getElementById("competenciaNomeInput");
  var processoInput = document.getElementById("processoInput");

  if (!btn || !dataInput || !timeInput || !poloSelect) return;

  var temData = !!dataInput.value;
  var temHora = !!timeInput.value && calcularIntervalo(timeInput.value) !== "";
  var temPolo = !!poloSelect.value;

  var temVara = false;
  if (novaVaraCheck && novaVaraCheck.checked) {
    temVara = !!(varaNomeInput && varaNomeInput.value.trim()) && !!(varaContatoInput && varaContatoInput.value.trim());
  } else {
    temVara = !!(varaSelect && varaSelect.value);
  }

  var temCompetencia = false;
  if (novaCompetenciaCheck && novaCompetenciaCheck.checked) {
    temCompetencia = !!(competenciaNomeInput && competenciaNomeInput.value.trim());
  } else {
    temCompetencia = !!(competenciaSelect && competenciaSelect.value);
  }

  var digitosProcesso = processoInput ? extrairDigitosProcesso(processoInput.value) : "";
  var temProcesso = digitosProcesso.length === 20;

  var vitimas = coletarVitimas();
  var temVitima = vitimas !== null;
  var adaptacao = coletarAdaptacao();
  var temAdaptacao =
    adaptacao !== null && (adaptacao.tipo === "NÃO" || (adaptacao.tipo === "SIM" && adaptacao.descricao !== ""));

  var temEntrevistador = entrevistadoresSelecionadosAudiencia.length > 0;

  var temStatus = true;
  if (modoEdicaoAudiencia) {
    var status = coletarStatus();
    var exigeMotivo = status.tipo === "CANCELADO" || status.tipo === "NÃO REALIZADO";
    temStatus = !exigeMotivo || status.motivo !== "";
  }

  btn.disabled = !(
    temData &&
    temHora &&
    temPolo &&
    temVara &&
    temCompetencia &&
    temProcesso &&
    temVitima &&
    temAdaptacao &&
    temEntrevistador &&
    temStatus
  );
}

function definirProcessando(ativo) {
  var overlay = document.getElementById("overlay");
  if (overlay) {
    if (ativo) overlay.classList.remove("hidden");
    else overlay.classList.add("hidden");
  }

  var ids = [
    "dataInput",
    "timeInput",
    "poloSelect",
    "novaVaraCheck",
    "varaSelect",
    "varaNomeInput",
    "varaContatoInput",
    "novaCompetenciaCheck",
    "competenciaSelect",
    "competenciaNomeInput",
    "processoInput",
    "vitimaNomesInput",
    "vitimaIdadesInput",
    "btnSalvar",
    "btnSalvarEntrevistador",
    "entNomeInput",
    "entEmailInput",
    "entTipoSelect",
    "entPrioridadeSelect",
    "entDisponibilidadeSelect",
    "entEspecialidadeInput",
    "entLotacaoInput",
    "entObsInput",
    "audienciaEditSelect",
    "entrevistadorEditSelect",
    "audienciaJaExiste",
    "entrevistadorJaExiste",
    "adaptacaoDescricaoInput",
    "statusSelect",
    "statusMotivoInput",
    "entPoloSelect",
    "btnAddPolo",
    "entDiaEspOrdem",
    "entDiaEspDia",
    "btnAddDiaEspecifico",
    "btnAddEntrevistador",
  ];
  for (var i = 0; i < ids.length; i++) {
    var el = document.getElementById(ids[i]);
    if (el) el.disabled = ativo;
  }

  document.querySelectorAll("#entDiasSemana input").forEach(function (cb) {
    cb.disabled = ativo;
  });
  document.querySelectorAll('input[name="adaptacaoRadio"]').forEach(function (r) {
    r.disabled = ativo;
  });
  document.querySelectorAll(".chip-remove").forEach(function (s) {
    if (ativo) s.style.pointerEvents = "none";
    else s.style.pointerEvents = "";
  });
}

// ============================================================
// CARREGAMENTO CONSOLIDADO
// ============================================================

async function carregarTudoInicial() {
  try {
    await Excel.run(async function (context) {
      var formSheet = context.workbook.worksheets.getItem("Formulário");
      var cadSheet = context.workbook.worksheets.getItem("CADASTROS");

      var rangePolos = formSheet.getRange("AD2:AD30");
      var rangeVaras = cadSheet.getRange("AA2:AB1000");
      var rangeComp = cadSheet.getRange("AC2:AC1000");
      var rangeEnt = cadSheet.getRange("A2:J1000");

      rangePolos.load("values");
      rangeVaras.load("values");
      rangeComp.load("values");
      rangeEnt.load("values");

      var tabela = formSheet.tables.getItem("Solicitacoes");

      var colK = tabela.columns.getItemAt(10).getDataBodyRange();
      var colL = tabela.columns.getItemAt(11).getDataBodyRange();
      if (colK) colK.numberFormat = [["@"]];
      if (colL) colL.numberFormat = [["@"]];

      var rangeSol = tabela.getRange();
      rangeSol.load("values, text");

      await context.sync();

      todosPolos = [];
      for (var i = 0; i < rangePolos.values.length; i++) {
        var v = rangePolos.values[i][0];
        if (v !== "" && v !== null && v !== undefined) todosPolos.push(String(v).trim());
      }

      var selectPolo = document.getElementById("poloSelect");
      var selPoloEnt = document.getElementById("entPoloSelect");

      if (selectPolo) {
        selectPolo.innerHTML = "";
        var d = document.createElement("option");
        d.value = "";
        d.textContent = "Selecione o polo...";
        selectPolo.appendChild(d);
        for (var ip = 0; ip < todosPolos.length; ip++) {
          var p = todosPolos[ip];
          var o = document.createElement("option");
          o.value = p;
          o.textContent = p;
          selectPolo.appendChild(o);
        }
      }

      if (selPoloEnt) {
        selPoloEnt.innerHTML = "";
        var d2 = document.createElement("option");
        d2.value = "";
        d2.textContent = "Selecione o polo...";
        selPoloEnt.appendChild(d2);
        for (var ip2 = 0; ip2 < todosPolos.length; ip2++) {
          var p2 = todosPolos[ip2];
          var o2 = document.createElement("option");
          o2.value = p2;
          o2.textContent = p2;
          selPoloEnt.appendChild(o2);
        }
      }

      varasData = [];
      for (var iv = 0; iv < rangeVaras.values.length; iv++) {
        var nomeV = String(rangeVaras.values[iv][0]).trim();
        var contatoV = String(rangeVaras.values[iv][1]).trim();
        if (nomeV !== "" && nomeV !== "null" && nomeV !== "undefined")
          varasData.push({ nome: nomeV, contato: contatoV });
      }

      var selectVara = document.getElementById("varaSelect");
      if (selectVara) {
        selectVara.innerHTML = "";
        var dv = document.createElement("option");
        dv.value = "";
        dv.textContent = "Selecione a vara...";
        selectVara.appendChild(dv);
        for (var iv2 = 0; iv2 < varasData.length; iv2++) {
          var vv = varasData[iv2];
          var ov = document.createElement("option");
          ov.value = vv.nome;
          ov.textContent = vv.nome;
          selectVara.appendChild(ov);
        }
      }

      competenciasData = [];
      for (var ic = 0; ic < rangeComp.values.length; ic++) {
        var vc = String(rangeComp.values[ic][0]).trim();
        if (vc !== "" && vc !== "null" && vc !== "undefined") competenciasData.push(vc);
      }

      var selectComp = document.getElementById("competenciaSelect");
      if (selectComp) {
        selectComp.innerHTML = "";
        var dc = document.createElement("option");
        dc.value = "";
        dc.textContent = "Selecione a competência...";
        selectComp.appendChild(dc);
        for (var ic2 = 0; ic2 < competenciasData.length; ic2++) {
          var cc = competenciasData[ic2];
          var oc = document.createElement("option");
          oc.value = cc;
          oc.textContent = cc;
          selectComp.appendChild(oc);
        }
      }

      audienciasCache = [];
      for (var ia = 1; ia < rangeSol.values.length; ia++) {
        var row = rangeSol.values[ia];
        var id = String(row[0]).trim();
        if (!id) continue;
        var dataNormalizada = normalizarDataBR(row[3]);
        if (!dataNormalizada && rangeSol.text && rangeSol.text[ia]) {
          dataNormalizada = String(rangeSol.text[ia][3] || "").trim();
        }

        var vitimasRaw = row[10];
        var idadesRaw = row[11];
        var vitimasTexto =
          typeof vitimasRaw === "number"
            ? String(rangeSol.text[ia][10] || "").trim()
            : String(vitimasRaw || "").trim();
        var idadesTexto =
          typeof idadesRaw === "number"
            ? String(rangeSol.text[ia][11] || "").trim()
            : String(idadesRaw || "").trim();

        audienciasCache.push({
          id: id,
          rowIndex: ia + 1,
          data: dataNormalizada,
          hora: String(row[4]).trim(),
          polo: String(row[5]).trim(),
          vara: String(row[6]).trim(),
          contatoVara: String(row[7]).trim(),
          competencia: String(row[8]).trim(),
          processo: String(row[9]).trim(),
          vitimas: vitimasTexto,
          idades: idadesTexto,
          adaptacao: String(row[12]).trim(),
          entrevistador: String(row[13]).trim(),
          emailEntrevistador: String(row[14]).trim(),
          status: String(row[15]).trim(),
        });
      }

      entrevistadoresData = [];
      for (var ie = 0; ie < rangeEnt.values.length; ie++) {
        var rowE = rangeEnt.values[ie];
        var nome = String(rowE[0]).trim();
        if (nome === "" || nome === "null" || nome === "undefined") continue;

        var prioridade = String(rowE[3]).trim();
        if (prioridade === "Regular") prioridade = "OBRIGATÓRIO";
        else if (prioridade === "A priorizar") prioridade = "NEGOCIÁVEL";

        var disponibilidade = String(rowE[4]).trim();
        if (disponibilidade === "Entrevistador disponível") disponibilidade = "DISPONÍVEL";
        else if (disponibilidade === "Entrevistador não disponível") disponibilidade = "INATIVO";

        entrevistadoresData.push({
          nome: nome,
          email: String(rowE[1]).trim(),
          tipo: String(rowE[2]).trim(),
          prioridade: prioridade,
          disponibilidade: disponibilidade,
          especialidade: String(rowE[5]).trim(),
          lotacao: String(rowE[6]).trim(),
          polos: String(rowE[7]).trim(),
          dias: String(rowE[8]).trim(),
          obs: String(rowE[9]).trim(),
          rowIndex: ie + 2,
          qtdMarcado: 0,
          qtdCancelado: 0,
          qtdRealizado: 0,
        });
      }

      for (var ie2 = 0; ie2 < entrevistadoresData.length; ie2++) {
        var e2 = entrevistadoresData[ie2];
        for (var ia2 = 0; ia2 < audienciasCache.length; ia2++) {
          var a2 = audienciasCache[ia2];
          if (a2.entrevistador.indexOf(e2.nome) === -1) continue;
          var s = a2.status.toUpperCase();
          if (s === "CONFIRMADO") {
            e2.qtdMarcado++;
          } else if (s.indexOf("CANCELADO") === 0) {
            e2.qtdCancelado++;
          } else if (s.indexOf("NÃO REALIZADO") === 0 || s.indexOf("NAO REALIZADO") === 0) {
            e2.qtdCancelado++;
          } else if (s === "REALIZADO") {
            e2.qtdRealizado++;
          }
        }
      }

      var editSelect = document.getElementById("entrevistadorEditSelect");
      if (editSelect) {
        editSelect.innerHTML = "";
        var de = document.createElement("option");
        de.value = "";
        de.textContent = "Selecione o entrevistador...";
        editSelect.appendChild(de);
        var ordenados = entrevistadoresData.slice().sort(function (a, b) {
          return a.nome.localeCompare(b.nome);
        });
        for (var io = 0; io < ordenados.length; io++) {
          var eo = ordenados[io];
          var oo = document.createElement("option");
          oo.value = eo.nome;
          oo.textContent = eo.nome + " — " + eo.email;
          editSelect.appendChild(oo);
        }
      }

      vitimasCache = [];
      for (var it = 0; it < rangeSol.text.length; it++) {
        var linha = rangeSol.text[it];
        var idT = String(linha[0]).trim();
        if (idT === "") continue;
        var poloT = String(linha[5]).trim();
        var processoT = String(linha[9]).trim();
        var vitimasTexto2 = String(linha[10]).trim();
        if (!vitimasTexto2) continue;
        var nomes = vitimasTexto2
          .split("/")
          .map(function (s2) {
            return s2.trim();
          })
          .filter(function (s2) {
            return s2 !== "";
          });
        for (var in2 = 0; in2 < nomes.length; in2++) {
          vitimasCache.push({ nome: nomes[in2], polo: poloT, processo: processoT });
        }
      }

      var selectAud = document.getElementById("audienciaEditSelect");
      if (selectAud) {
        selectAud.innerHTML = "";
        var da = document.createElement("option");
        da.value = "";
        da.textContent = "Selecione a audiência (ID)...";
        selectAud.appendChild(da);
        for (var ia3 = 0; ia3 < audienciasCache.length; ia3++) {
          var aa = audienciasCache[ia3];
          var oa = document.createElement("option");
          oa.value = aa.id;
          oa.textContent = aa.id + " · " + aa.processo + " · " + aa.data;
          selectAud.appendChild(oa);
        }
      }

      atualizarSelectEntrevistador();
      renderizarChipsEntrevistadores();
    });
  } catch (error) {
    console.error("[INIT] Falha no carregamento:", error);
  }
}

// ============================================================
// HANDLERS DE SELEÇÃO
// ============================================================

function onDataOuPoloChange() {
  var sel = document.getElementById("entrevistadorSelect");
  if (sel) sel.value = "";
  var email = document.getElementById("entrevistadorEmailDisplay");
  if (email) email.value = "";
  var extra = document.getElementById("entrevistadorExtraDisplay");
  if (extra) extra.value = "";
  var div = document.getElementById("entrevistadorDivergencia");
  if (div) {
    div.classList.add("hidden");
    div.textContent = "";
  }

  atualizarSelectEntrevistador();
  atualizarEstadoBotao();
}

function onEntrevistadorSelectChange() {
  var sel = document.getElementById("entrevistadorSelect");
  var email = document.getElementById("entrevistadorEmailDisplay");
  var extra = document.getElementById("entrevistadorExtraDisplay");
  var div = document.getElementById("entrevistadorDivergencia");
  if (!sel || !email || !extra || !div) return;

  var nome = sel.value;
  if (!nome) {
    email.value = "";
    extra.value = "";
    div.classList.add("hidden");
    div.textContent = "";
    atualizarEstadoBotao();
    return;
  }

  var e = entrevistadoresData.find(function (x) {
    return x.nome === nome;
  });
  if (!e) return;

  email.value = e.email;
  extra.value = e.tipo + " · " + e.disponibilidade + " · " + e.polos;

  var dataISO = document.getElementById("dataInput").value;
  var polo = document.getElementById("poloSelect").value;

  if (dataISO && polo) {
    var res = calcularScore(e, polo, converterParaBR(dataISO));
    if (res.tipo === "red") {
      div.textContent = "⚠️ " + res.motivo;
      div.classList.remove("hidden");
    } else {
      div.classList.add("hidden");
      div.textContent = "";
    }
  } else {
    div.classList.add("hidden");
    div.textContent = "";
  }

  atualizarEstadoBotao();
}

function onVaraSelectChange() {
  var s = document.getElementById("varaSelect");
  var d = document.getElementById("varaContatoDisplay");
  if (!s || !d) return;
  if (!s.value) {
    d.value = "";
    atualizarEstadoBotao();
    return;
  }
  var v = varasData.find(function (x) {
    return x.nome === s.value;
  });
  d.value = v ? v.contato : "";
  atualizarEstadoBotao();
}

function alternarModoVara() {
  var c = document.getElementById("novaVaraCheck");
  var g = document.getElementById("varaSelectGroup");
  var ng = document.getElementById("novaVaraGroup");
  var fg = document.getElementById("varaFieldGroup");
  var av = document.getElementById("avisoNovaVara");
  if (!c || !g || !ng) return;
  if (c.checked) {
    g.classList.add("hidden");
    ng.classList.remove("hidden");
    if (fg) fg.classList.add("nova-vara-ativa");
    if (av) av.classList.remove("hidden");
  } else {
    g.classList.remove("hidden");
    ng.classList.add("hidden");
    if (fg) fg.classList.remove("nova-vara-ativa");
    if (av) av.classList.add("hidden");
  }
}

function onNovaVaraChange() {
  var c = document.getElementById("novaVaraCheck");
  if (!c) return;
  if (c.checked) {
    document.getElementById("varaSelect").value = "";
    document.getElementById("varaContatoDisplay").value = "";
  } else {
    document.getElementById("varaNomeInput").value = "";
    document.getElementById("varaContatoInput").value = "";
  }
  alternarModoVara();
  atualizarEstadoBotao();
}

function alternarModoCompetencia() {
  var c = document.getElementById("novaCompetenciaCheck");
  var g = document.getElementById("competenciaSelectGroup");
  var ng = document.getElementById("novaCompetenciaGroup");
  var fg = document.getElementById("competenciaFieldGroup");
  var av = document.getElementById("avisoNovaCompetencia");
  if (!c || !g || !ng) return;
  if (c.checked) {
    g.classList.add("hidden");
    ng.classList.remove("hidden");
    if (fg) fg.classList.add("nova-competencia-ativa");
    if (av) av.classList.remove("hidden");
  } else {
    g.classList.remove("hidden");
    ng.classList.add("hidden");
    if (fg) fg.classList.remove("nova-competencia-ativa");
    if (av) av.classList.add("hidden");
  }
}

function onNovaCompetenciaChange() {
  var c = document.getElementById("novaCompetenciaCheck");
  if (!c) return;
  if (c.checked) {
    document.getElementById("competenciaSelect").value = "";
  } else {
    document.getElementById("competenciaNomeInput").value = "";
  }
  alternarModoCompetencia();
  atualizarEstadoBotao();
}

function coletarVitimas() {
  var n = document
    .getElementById("vitimaNomesInput")
    .value.split("/")
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return s !== "";
    });
  var i = document
    .getElementById("vitimaIdadesInput")
    .value.split("/")
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return s !== "";
    });
  if (n.length === 0 || n.length !== i.length) return null;
  for (var k = 0; k < i.length; k++) {
    if (i[k].length > 2 || isNaN(parseInt(i[k], 10))) return null;
  }
  return { nomes: n, idades: i };
}

function agendarAtualizacaoSugestoes() {
  if (sugestoesTimeout) clearTimeout(sugestoesTimeout);
  sugestoesTimeout = setTimeout(function () {
    atualizarSugestoesVitima();
  }, 400);
}

function atualizarSugestoesVitima() {
  var input = document.getElementById("vitimaNomesInput");
  var listDiv = document.getElementById("sugestoesVitimaList");
  var avisoDiv = document.getElementById("avisoRepeticaoVitima");
  if (!input || !listDiv || !avisoDiv) return;

  var nomes = input.value
    .trim()
    .split("/")
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return s !== "";
    });
  listDiv.innerHTML = "";
  listDiv.classList.add("hidden");
  avisoDiv.classList.add("hidden");
  avisoDiv.textContent = "";
  if (nomes.length === 0) return;

  var ultimo = nomes[nomes.length - 1].toLowerCase();
  if (ultimo.length < 3) return;

  var encontrados = vitimasCache.filter(function (v) {
    return v.nome.toLowerCase().indexOf(ultimo) !== -1 && v.nome.toLowerCase() !== ultimo;
  });
  if (encontrados.length > 0) {
    for (var i = 0; i < encontrados.length && i < 5; i++) {
      var item = encontrados[i];
      var d = document.createElement("div");
      d.className = "sugestao-item";
      d.textContent = item.nome + " — " + item.polo + " — Proc: " + item.processo;
      (function (item) {
        d.addEventListener("click", function () {
          nomes[nomes.length - 1] = item.nome;
          input.value = nomes.join("/");
          listDiv.classList.add("hidden");
          listDiv.innerHTML = "";
          atualizarEstadoBotao();
          atualizarSugestoesVitima();
        });
      })(item);
      listDiv.appendChild(d);
    }
    listDiv.classList.remove("hidden");
  }

  var repetidos = [];
  for (var j = 0; j < nomes.length; j++) {
    var nome = nomes[j];
    var ms = vitimasCache.filter(function (v) {
      return v.nome.toLowerCase() === nome.toLowerCase();
    });
    for (var k = 0; k < ms.length; k++) repetidos.push(ms[k]);
  }
  if (repetidos.length === 0) {
    avisoDiv.className = "aviso-nova-vara hidden";
    avisoDiv.textContent = "";
  } else {
    var html = "<strong>⚠️ Atenção!</strong> Já existe(m) audiência(s) cadastrada(s) para:<br>";
    for (var r = 0; r < repetidos.length; r++) {
      html += "• " + repetidos[r].nome + " (Polo: " + repetidos[r].polo + ", Proc: " + repetidos[r].processo + ")<br>";
    }
    html += "<br>Você pode continuar, mas verifique se não é duplicidade.";
    avisoDiv.innerHTML = html;
    avisoDiv.classList.remove("hidden");
  }
}

// ============================================================
// MODO EDIÇÃO — AUDIÊNCIA
// ============================================================

function onAudienciaJaExisteChange() {
  var check = document.getElementById("audienciaJaExiste");
  var selector = document.getElementById("audienciaEditSelector");
  var formPanel = document.getElementById("audienciaFormPanel");
  var select = document.getElementById("audienciaEditSelect");
  var btn = document.getElementById("btnSalvar");
  var statusGroup = document.getElementById("statusFieldGroup");

  modoEdicaoAudiencia = check.checked;
  audienciaEmEdicao = null;

  if (check.checked) {
    selector.classList.remove("hidden");
    formPanel.classList.add("hidden");
    select.value = "";
    btn.textContent = "Atualizar";
    if (statusGroup) statusGroup.classList.remove("hidden");
  } else {
    selector.classList.add("hidden");
    formPanel.classList.remove("hidden");
    limparFormularioAudiencia();
    btn.textContent = "Cadastrar";
    if (statusGroup) statusGroup.classList.add("hidden");
    atualizarEstadoBotao();
  }
}

async function onAudienciaEditSelectChange() {
  var sel = document.getElementById("audienciaEditSelect");
  var formPanel = document.getElementById("audienciaFormPanel");
  var id = sel.value;
  if (!id) {
    formPanel.classList.add("hidden");
    audienciaEmEdicao = null;
    return;
  }
  var cache = audienciasCache.find(function (a) {
    return a.id === id;
  });
  if (!cache) return;
  audienciaEmEdicao = cache;
  preencherFormularioAudiencia(cache, cache.idades);
  formPanel.classList.remove("hidden");
  atualizarPreview();
  atualizarEstadoBotao();
}

function preencherFormularioAudiencia(c, idades) {
  document.getElementById("dataInput").value = converterParaISO(c.data);
  var hp = c.hora.split("/");
  document.getElementById("timeInput").value = (hp.length > 1 ? hp[1].trim() : c.hora.trim()).substring(0, 5);
  document.getElementById("poloSelect").value = c.polo;

  var nv = document.getElementById("novaVaraCheck");
  if (
    varasData.some(function (v) {
      return v.nome === c.vara;
    })
  ) {
    nv.checked = false;
    document.getElementById("varaSelect").value = c.vara;
    document.getElementById("varaContatoDisplay").value = c.contatoVara;
  } else {
    nv.checked = true;
    document.getElementById("varaNomeInput").value = c.vara;
    document.getElementById("varaContatoInput").value = c.contatoVara;
  }
  alternarModoVara();

  var nc = document.getElementById("novaCompetenciaCheck");
  if (competenciasData.indexOf(c.competencia) !== -1) {
    nc.checked = false;
    document.getElementById("competenciaSelect").value = c.competencia;
  } else {
    nc.checked = true;
    document.getElementById("competenciaNomeInput").value = c.competencia;
  }
  alternarModoCompetencia();

  document.getElementById("processoInput").value = formatarProcesso(c.processo);
  document.getElementById("vitimaNomesInput").value = c.vitimas;
  document.getElementById("vitimaIdadesInput").value = idades;
  preencherAdaptacaoRadio(c.adaptacao);

  entrevistadoresSelecionadosAudiencia = String(c.entrevistador || "")
    .split("/")
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return s !== "";
    });
  renderizarChipsEntrevistadores();

  atualizarSelectEntrevistador();
  document.getElementById("entrevistadorSelect").value = "";
  document.getElementById("entrevistadorEmailDisplay").value = "";
  document.getElementById("entrevistadorExtraDisplay").value = "";

  preencherStatusSelect(c.status);
  atualizarPreview();
}

function limparFormularioAudiencia() {
  document.getElementById("dataInput").value = "";
  document.getElementById("timeInput").value = "";
  document.getElementById("poloSelect").value = "";
  document.getElementById("novaVaraCheck").checked = false;
  document.getElementById("varaSelect").value = "";
  document.getElementById("varaContatoDisplay").value = "";
  document.getElementById("varaNomeInput").value = "";
  document.getElementById("varaContatoInput").value = "";
  document.getElementById("novaCompetenciaCheck").checked = false;
  document.getElementById("competenciaSelect").value = "";
  document.getElementById("competenciaNomeInput").value = "";
  document.getElementById("processoInput").value = "";
  document.getElementById("vitimaNomesInput").value = "";
  document.getElementById("vitimaIdadesInput").value = "";
  limparAdaptacaoRadio();

  document.getElementById("entrevistadorSelect").value = "";
  document.getElementById("entrevistadorEmailDisplay").value = "";
  document.getElementById("entrevistadorExtraDisplay").value = "";
  var div = document.getElementById("entrevistadorDivergencia");
  if (div) {
    div.classList.add("hidden");
    div.textContent = "";
  }

  entrevistadoresSelecionadosAudiencia = [];
  renderizarChipsEntrevistadores();

  var sr = document.getElementById("statusSelect");
  if (sr) sr.value = "PENDENTE";
  var sm = document.getElementById("statusMotivoInput");
  if (sm) {
    sm.value = "";
    sm.classList.add("hidden");
  }

  atualizarSelectEntrevistador();
  alternarModoVara();
  alternarModoCompetencia();
  atualizarPreview();
}

// ============================================================
// MODO EDIÇÃO — ENTREVISTADOR
// ============================================================

function onEntrevistadorJaExisteChange() {
  var check = document.getElementById("entrevistadorJaExiste");
  var selector = document.getElementById("entrevistadorEditSelector");
  var formPanel = document.getElementById("entrevistadorFormPanel");
  var sel = document.getElementById("entrevistadorEditSelect");
  var btn = document.getElementById("btnSalvarEntrevistador");
  modoEdicaoEntrevistador = check.checked;
  entrevistadorEmEdicao = null;
  if (check.checked) {
    selector.classList.remove("hidden");
    formPanel.classList.add("hidden");
    sel.value = "";
    btn.textContent = "Atualizar";
  } else {
    selector.classList.add("hidden");
    formPanel.classList.remove("hidden");
    limparFormularioEntrevistador();
    btn.textContent = "Cadastrar";
    atualizarEstadoBotaoEntrevistador();
  }
}

function onEntrevistadorEditSelectChange() {
  var sel = document.getElementById("entrevistadorEditSelect");
  var formPanel = document.getElementById("entrevistadorFormPanel");
  var nome = sel.value;
  if (!nome) {
    formPanel.classList.add("hidden");
    entrevistadorEmEdicao = null;
    return;
  }
  var e = entrevistadoresData.find(function (x) {
    return x.nome === nome;
  });
  if (!e) return;
  entrevistadorEmEdicao = e;

  document.getElementById("entNomeInput").value = e.nome;
  document.getElementById("entEmailInput").value = e.email;
  document.getElementById("entTipoSelect").value = e.tipo;
  document.getElementById("entPrioridadeSelect").value = e.prioridade;
  document.getElementById("entDisponibilidadeSelect").value = e.disponibilidade;
  document.getElementById("entEspecialidadeInput").value = e.especialidade;
  document.getElementById("entLotacaoInput").value = e.lotacao;
  document.getElementById("entObsInput").value = e.obs;

  polosSelecionados = e.polos
    .split(/[,;]/)
    .map(function (s) {
      return s.trim();
    })
    .filter(Boolean);
  renderizarChipsPolos();

  var parsed = parsearDias(e.dias);
  var semana = parsed.semana;
  var especificos = parsed.especificos;
  var checkboxes = document.querySelectorAll("#entDiasSemana input");
  checkboxes.forEach(function (cb) {
    if (cb.id === "entTodosDiasCheck") {
      cb.checked = semana.size >= 5;
    } else {
      cb.checked = semana.has(cb.value);
    }
  });
  diasEspecificosSelecionados = especificos.filter(function (s) {
    return !s.startsWith("?");
  });
  renderizarChipsDiasEspecificos();

  formPanel.classList.remove("hidden");
  atualizarEstadoBotaoEntrevistador();
}

function limparFormularioEntrevistador() {
  document.getElementById("entNomeInput").value = "";
  document.getElementById("entEmailInput").value = "";
  document.getElementById("entTipoSelect").value = "";
  document.getElementById("entPrioridadeSelect").value = "";
  document.getElementById("entDisponibilidadeSelect").value = "";
  document.getElementById("entEspecialidadeInput").value = "";
  document.getElementById("entLotacaoInput").value = "";
  document.getElementById("entObsInput").value = "";
  document.querySelectorAll("#entDiasSemana input").forEach(function (cb) {
    cb.checked = false;
  });
  polosSelecionados = [];
  diasEspecificosSelecionados = [];
  renderizarChipsPolos();
  renderizarChipsDiasEspecificos();

  var erroEmail = document.getElementById("entEmailErro");
  if (erroEmail) {
    erroEmail.classList.add("hidden");
    erroEmail.textContent = "";
  }
  var erroNome = document.getElementById("entNomeErro");
  if (erroNome) {
    erroNome.classList.add("hidden");
    erroNome.textContent = "";
  }
}

// ============================================================
// CHIPS: POLOS
// ============================================================

function renderizarChipsPolos() {
  var cont = document.getElementById("entPolosChips");
  if (!cont) return;
  cont.innerHTML = "";
  for (var i = 0; i < polosSelecionados.length; i++) {
    var p = polosSelecionados[i];
    var chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = p + " ";
    var x = document.createElement("span");
    x.className = "chip-remove";
    x.textContent = "×";
    (function (p) {
      x.addEventListener("click", function () {
        polosSelecionados = polosSelecionados.filter(function (q) {
          return q !== p;
        });
        renderizarChipsPolos();
        atualizarEstadoBotaoEntrevistador();
      });
    })(p);
    chip.appendChild(x);
    cont.appendChild(chip);
  }
}

function onAddPoloClick() {
  var sel = document.getElementById("entPoloSelect");
  var p = sel.value;
  if (!p) return;
  if (polosSelecionados.indexOf(p) === -1) {
    polosSelecionados.push(p);
    renderizarChipsPolos();
  }
  sel.value = "";
  atualizarEstadoBotaoEntrevistador();
}

// ============================================================
// CHIPS: DIAS ESPECÍFICOS
// ============================================================

function renderizarChipsDiasEspecificos() {
  var cont = document.getElementById("entDiasEspecificosChips");
  if (!cont) return;
  cont.innerHTML = "";
  for (var i = 0; i < diasEspecificosSelecionados.length; i++) {
    var d = diasEspecificosSelecionados[i];
    var chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = d + " do mês ";
    var x = document.createElement("span");
    x.className = "chip-remove";
    x.textContent = "×";
    (function (d) {
      x.addEventListener("click", function () {
        diasEspecificosSelecionados = diasEspecificosSelecionados.filter(function (q) {
          return q !== d;
        });
        renderizarChipsDiasEspecificos();
        atualizarEstadoBotaoEntrevistador();
      });
    })(d);
    chip.appendChild(x);
    cont.appendChild(chip);
  }
}

function onAddDiaEspecificoClick() {
  var ordem = document.getElementById("entDiaEspOrdem").value;
  var dia = document.getElementById("entDiaEspDia").value;
  if (!ordem || !dia) return;
  var texto = ordem + " " + dia;
  if (diasEspecificosSelecionados.indexOf(texto) === -1) {
    diasEspecificosSelecionados.push(texto);
    renderizarChipsDiasEspecificos();
  }
  document.getElementById("entDiaEspOrdem").value = "";
  document.getElementById("entDiaEspDia").value = "";
  atualizarEstadoBotaoEntrevistador();
}

// ============================================================
// CHIPS: ENTREVISTADORES DA AUDIÊNCIA
// ============================================================

function renderizarChipsEntrevistadores() {
  var cont = document.getElementById("entrevistadoresChips");
  if (!cont) return;
  cont.innerHTML = "";
  for (var i = 0; i < entrevistadoresSelecionadosAudiencia.length; i++) {
    var nome = entrevistadoresSelecionadosAudiencia[i];
    var chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = nome + " ";
    var x = document.createElement("span");
    x.className = "chip-remove";
    x.textContent = "×";
    (function (nome) {
      x.addEventListener("click", function () {
        entrevistadoresSelecionadosAudiencia = entrevistadoresSelecionadosAudiencia.filter(function (q) {
          return q !== nome;
        });
        renderizarChipsEntrevistadores();
        atualizarEstadoBotao();
      });
    })(nome);
    chip.appendChild(x);
    cont.appendChild(chip);
  }
}

function onAddEntrevistadorClick() {
  var sel = document.getElementById("entrevistadorSelect");
  var nome = sel.value;
  if (!nome) return;
  if (entrevistadoresSelecionadosAudiencia.indexOf(nome) !== -1) {
    exibirMensagem("Este entrevistador já está na lista.", "erro");
    return;
  }
  entrevistadoresSelecionadosAudiencia.push(nome);
  renderizarChipsEntrevistadores();
  sel.value = "";

  var email = document.getElementById("entrevistadorEmailDisplay");
  var extra = document.getElementById("entrevistadorExtraDisplay");
  var div = document.getElementById("entrevistadorDivergencia");
  if (email) email.value = "";
  if (extra) extra.value = "";
  if (div) {
    div.classList.add("hidden");
    div.textContent = "";
  }

  atualizarEstadoBotao();
}

function tiposEntrevistadoresSelecionados() {
  var tipos = [];
  for (var i = 0; i < entrevistadoresSelecionadosAudiencia.length; i++) {
    var nome = entrevistadoresSelecionadosAudiencia[i];
    var e = entrevistadoresData.find(function (x) {
      return x.nome === nome;
    });
    tipos.push(e ? e.tipo : "");
  }
  return tipos;
}

function algumEntrevistadorExterno() {
  var tipos = tiposEntrevistadoresSelecionados();
  for (var i = 0; i < tipos.length; i++) {
    if (isEntrevistadorCadExterno(tipos[i])) return true;
  }
  return false;
}

function separarEntrevistadoresPorTipo() {
  var internos = [];
  var externos = [];

  for (var i = 0; i < entrevistadoresSelecionadosAudiencia.length; i++) {
    var nome = entrevistadoresSelecionadosAudiencia[i];
    var e = entrevistadoresData.find(function (x) {
      return x.nome === nome;
    });
    var tipo = e ? e.tipo : "";
    if (isEntrevistadorCadExterno(tipo)) {
      externos.push(nome);
    } else {
      internos.push(nome);
    }
  }

  return {
    internos: internos.join("/"),
    externos: externos.join("/"),
  };
}

function calcularStatusInicial(nomesEntrevistadores, diffDias) {
  for (var i = 0; i < nomesEntrevistadores.length; i++) {
    var nome = nomesEntrevistadores[i];
    var e = entrevistadoresData.find(function (x) {
      return x.nome === nome;
    });
    var tipoNorm = e ? e.tipo.toUpperCase().trim() : "";
    var isInterno = tipoNorm === "INTERNO" || tipoNorm === "RESIDENTE";
    var precisaConfirmarEste = !isInterno || diffDias <= 9;
    if (precisaConfirmarEste) return "PENDENTE";
  }
  return "CONFIRMADO";
}

function emailsEntrevistadoresSelecionados() {
  var emails = [];
  for (var i = 0; i < entrevistadoresSelecionadosAudiencia.length; i++) {
    var nome = entrevistadoresSelecionadosAudiencia[i];
    var e = entrevistadoresData.find(function (x) {
      return x.nome === nome;
    });
    if (e && e.email) emails.push(e.email);
  }
  return emails.join("/");
}

// ============================================================
// COLETA DE DIAS DO FORM
// ============================================================

function coletarDiasSelecionados() {
  var partes = [];
  var checkboxes = document.querySelectorAll("#entDiasSemana input");
  checkboxes.forEach(function (cb) {
    if (cb.checked && cb.id !== "entTodosDiasCheck") partes.push(cb.value);
  });
  for (var i = 0; i < diasEspecificosSelecionados.length; i++) partes.push(diasEspecificosSelecionados[i]);
  return partes.join(", ");
}

// ============================================================
// VALIDAÇÃO DO BOTÃO DO ENTREVISTADOR
// ============================================================

function atualizarEstadoBotaoEntrevistador() {
  var btn = document.getElementById("btnSalvarEntrevistador");
  if (!btn) return;

  var nome = document.getElementById("entNomeInput").value.trim();
  var email = document.getElementById("entEmailInput").value.trim();
  var tipo = document.getElementById("entTipoSelect").value;
  var prioridade = document.getElementById("entPrioridadeSelect").value;
  var disp = document.getElementById("entDisponibilidadeSelect").value;
  var esp = document.getElementById("entEspecialidadeInput").value.trim();

  var emailValido = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  var temPolos = polosSelecionados.length > 0;
  var temDias = coletarDiasSelecionados().length > 0;

  var nomeDuplicado = (function () {
    if (!nome) return false;
    var n = nome.toLowerCase();
    return entrevistadoresData.some(function (e) {
      if (modoEdicaoEntrevistador && entrevistadorEmEdicao && e.nome === entrevistadorEmEdicao.nome) return false;
      return e.nome.toLowerCase() === n;
    });
  })();

  var valido = !!(nome && emailValido && tipo && prioridade && disp && esp && temPolos && temDias) && !nomeDuplicado;
  btn.disabled = !valido;

  var erroEmail = document.getElementById("entEmailErro");
  if (erroEmail) {
    if (email && !emailValido) {
      erroEmail.textContent = "⚠️ E-mail inválido";
      erroEmail.classList.remove("hidden");
    } else {
      erroEmail.classList.add("hidden");
      erroEmail.textContent = "";
    }
  }

  var erroNome = document.getElementById("entNomeErro");
  if (erroNome) {
    if (nomeDuplicado) {
      erroNome.textContent = "⚠️ Já existe um entrevistador cadastrado com este nome.";
      erroNome.classList.remove("hidden");
    } else {
      erroNome.classList.add("hidden");
      erroNome.textContent = "";
    }
  }
}

// ============================================================
// AUTOCOMPLETE DO NOME (Entrevistador)
// ============================================================

function agendarSugestoesNome() {
  if (nomeSugestoesTimeout) clearTimeout(nomeSugestoesTimeout);
  nomeSugestoesTimeout = setTimeout(mostrarSugestoesNome, 300);
}

function mostrarSugestoesNome() {
  var input = document.getElementById("entNomeInput");
  var listDiv = document.getElementById("entNomeSugestoes");
  if (!input || !listDiv) return;

  var texto = input.value.trim().toLowerCase();
  listDiv.innerHTML = "";
  listDiv.classList.add("hidden");
  if (texto.length < 3) return;

  var encontrados = entrevistadoresData.filter(function (e) {
    if (modoEdicaoEntrevistador && entrevistadorEmEdicao && e.nome === entrevistadorEmEdicao.nome) return false;
    return e.nome.toLowerCase().indexOf(texto) !== -1;
  });

  if (encontrados.length === 0) return;

  for (var i = 0; i < encontrados.length && i < 5; i++) {
    var e = encontrados[i];
    var div = document.createElement("div");
    div.className = "sugestao-info";
    div.innerHTML = "🔒 <strong>" + e.nome + "</strong> — " + e.email;
    listDiv.appendChild(div);
  }
  listDiv.classList.remove("hidden");
}

// ============================================================
// HELPERS DE PLANILHA
// ============================================================

async function obterAutor(context, sheet) {
  try {
    var cellTemp = sheet.getCell(0, 77);
    var comentarioTemp = sheet.comments.add(cellTemp, "temp");
    comentarioTemp.load("authorName, authorEmail");
    await context.sync();
    var nome = comentarioTemp.authorName || "Desconhecido";
    var email = comentarioTemp.authorEmail || "";
    comentarioTemp.delete();
    await context.sync();
    return nome + " (" + email + ")";
  } catch (e) {
    return "Desconhecido";
  }
}

async function registrarLog(context, acao, entidade, referencia, autor, detalhes, payloadComentario) {
  if (payloadComentario === undefined) payloadComentario = null;
  try {
    var s = context.workbook.worksheets.getItemOrNullObject("Logs");
    await context.sync();
    if (s.isNullObject) return;
    s.load("protection/protected");
    await context.sync();
    var estava = s.protection.protected;
    if (estava) {
      s.protection.unprotect(SENHA_ABA);
      await context.sync();
    }
    var used = s.getRange("A:A").getUsedRangeOrNullObject(true);
    used.load("values");
    await context.sync();
    var isEmpty = used.isNullObject || used.values.length === 0;
    if (isEmpty) {
      s.getRange("A1:F1").values = [["Data/Hora", "Ação", "Entidade", "Referência", "Autor", "Detalhes"]];
      s.getRange("A1:F1").format.font.bold = true;
      await context.sync();
    } else if (used.values.length > 1) {
      s.getRange("2:2").insert(Excel.InsertShiftDirection.down);
      await context.sync();
    }
    var agora = new Date().toLocaleString("pt-BR");
    s.getRange("A2:F2").values = [[agora, acao, entidade, referencia, autor, detalhes]];
    await context.sync();

    if (payloadComentario) {
      var idNovo = payloadComentario.id ? String(payloadComentario.id) : "";
      if (idNovo) {
        var comments = s.comments;
        comments.load("items/content");
        await context.sync();

        var apagou = false;
        for (var i = 0; i < comments.items.length; i++) {
          var c = comments.items[i];
          var conteudo = String(c.content || "").trim();
          if (conteudo.indexOf("{") !== 0) continue;
          try {
            var parsed = JSON.parse(conteudo);
            if (parsed && String(parsed.id) === idNovo) {
              c.delete();
              apagou = true;
            }
          } catch (e) {
            // comentário malformado
          }
        }
        if (apagou) await context.sync();
      }

      s.comments.add("A2", JSON.stringify(payloadComentario));
      await context.sync();
    }

    if (estava) {
      s.protection.protect({}, SENHA_ABA);
      await context.sync();
    }
  } catch (error) {
    console.error("Erro ao registrar log:", error);
  }
}

async function encontrarFimTabela(context, sheet) {
  var r = sheet.getRange("J2:J200");
  r.load("text");
  await context.sync();
  for (var i = 0; i < r.text.length; i++) {
    if (String(r.text[i][0]).trim().toUpperCase() === "DEPOIMENTOS AGENDADOS") return i + 2;
  }
  return -1;
}

async function processoJaExiste(context, sheet, digitos, ignorarId) {
  var rj = sheet.getRange("J2:J1000");
  var ra = sheet.getRange("A2:A1000");
  rj.load("text");
  ra.load("text");
  await context.sync();
  for (var i = 0; i < rj.text.length; i++) {
    var valor = String(rj.text[i][0]).trim();
    if (valor === "") continue;
    var id = String(ra.text[i][0]).trim();
    if (ignorarId && id === ignorarId) continue;
    if (extrairDigitosProcesso(valor) === digitos) return true;
  }
  return false;
}

async function aplicarZebrado(context, sheet, linhaInicial, linhaFinal, colunaFinal) {
  var rd = sheet.getRange("B" + linhaInicial + ":L" + linhaFinal);
  rd.load("text");
  await context.sync();
  for (var i = 0; i < rd.text.length; i++) {
    var linha = linhaInicial + i;
    var dataTexto = String(rd.text[i][0]).trim();
    var statusTexto = String(rd.text[i][10] || "")
      .trim()
      .toUpperCase();
    var dia = 0;
    if (dataTexto && dataTexto.indexOf("/") !== -1) dia = parseInt(dataTexto.split("/")[0], 10);
    var rl = sheet.getRange("A" + linha + ":" + colunaFinal + linha);
    rl.format.horizontalAlignment = "Center";
    rl.format.verticalAlignment = "Center";
    rl.format.wrapText = true;
    var isPar = dia % 2 === 0;
    if (isPar && dia !== 0) rl.format.fill.color = "#FCD5B4";
    else rl.format.fill.clear();
    rl.format.font.strikethrough = false;

    if (statusTexto.indexOf("CANCELADO") === 0) {
      rl.format.fill.color = "#FFCDD2";
      rl.format.font.strikethrough = true;
    } else if (statusTexto.indexOf("NÃO REALIZADO") === 0 || statusTexto.indexOf("NAO REALIZADO") === 0) {
      rl.format.fill.color = "#FFCC80";
    } else if (statusTexto === "REALIZADO") {
      rl.format.fill.color = "#C8E6C9";
    } else if (
      statusTexto === "CONFIRMADO" ||
      statusTexto.indexOf("ENVIAR CONFIRMAÇÃO") === 0 ||
      statusTexto === "ENVIAR CONFIRMACAO" ||
      statusTexto === "CONVOCAR"
    ) {
      rl.format.fill.color = "#81D4FA";
    } else {
      rl.format.fill.color = "#FFF9C4";
    }
  }
  await context.sync();
}

async function aplicarFormatacaoLinhaUnica(context, sheet, linha, colunaFinal) {
  var rb = sheet.getRange("B" + linha);
  var rl = sheet.getRange("L" + linha);
  rb.load("text");
  rl.load("text");
  await context.sync();
  var dataTexto = String(rb.text[0][0]).trim();
  var statusTexto = String(rl.text[0][0]).trim().toUpperCase();
  var dia = 0;
  if (dataTexto && dataTexto.indexOf("/") !== -1) dia = parseInt(dataTexto.split("/")[0], 10);
  var rangeLinha = sheet.getRange("A" + linha + ":" + colunaFinal + linha);
  rangeLinha.format.horizontalAlignment = "Center";
  rangeLinha.format.verticalAlignment = "Center";
  rangeLinha.format.wrapText = true;
  var isPar = dia % 2 === 0;
  if (isPar && dia !== 0) rangeLinha.format.fill.color = "#FCD5B4";
  else rangeLinha.format.fill.clear();
  rangeLinha.format.font.strikethrough = false;

  if (statusTexto.indexOf("CANCELADO") === 0) {
    rangeLinha.format.fill.color = "#FFCDD2";
    rangeLinha.format.font.strikethrough = true;
  } else if (statusTexto.indexOf("NÃO REALIZADO") === 0 || statusTexto.indexOf("NAO REALIZADO") === 0) {
    rangeLinha.format.fill.color = "#FFCC80";
  } else if (statusTexto === "REALIZADO") {
    rangeLinha.format.fill.color = "#C8E6C9";
  } else if (
    statusTexto === "CONFIRMADO" ||
    statusTexto.indexOf("ENVIAR CONFIRMAÇÃO") === 0 ||
    statusTexto === "ENVIAR CONFIRMACAO" ||
    statusTexto === "CONVOCAR"
  ) {
    rangeLinha.format.fill.color = "#81D4FA";
  } else {
    rangeLinha.format.fill.color = "#FFF9C4";
  }
  await context.sync();
}

async function inserirNaAbaPolo(
  context,
  nomeAba,
  dados,
  dataTexto,
  horaTexto,
  varaTexto,
  contatoVaraTexto,
  competenciaTexto,
  adaptacaoTexto,
  processoTexto,
  vitimasTexto,
  idadesVitimasTexto,
  entrevistadoresInternosTexto,
  entrevistadoresExternosTexto,
  statusTexto
) {
  var aba = context.workbook.worksheets.getItemOrNullObject(nomeAba);
  aba.load("name");
  await context.sync();
  if (aba.isNullObject) return false;
  aba.load("protection/protected");
  await context.sync();
  var estava = aba.protection.protected;
  if (estava) {
    aba.protection.unprotect(SENHA_ABA);
    await context.sync();
  }

  var linhaFim = await encontrarFimTabela(context, aba);
  if (linhaFim === -1) {
    if (estava) {
      aba.protection.protect({}, SENHA_ABA);
      await context.sync();
    }
    return false;
  }

  var prim = 2,
    ult = linhaFim - 1;
  var pd = dataTexto.split("/");
  var diaSem = calcularDiaSemana(parseInt(pd[0]), parseInt(pd[1]), parseInt(pd[2]));
  var hi = horaTexto.split("/")[0].trim();
  var ph = hi.split(":");
  var minNovo = parseInt(ph[0]) * 60 + parseInt(ph[1]);
  var pos = ult + 1;

  if (ult >= prim) {
    var rb = aba.getRange("B" + prim + ":B" + ult);
    var rd = aba.getRange("D" + prim + ":D" + ult);
    rb.load("text");
    rd.load("text");
    await context.sync();
    var dataNova = converterStringParaData(dataTexto);
    if (dataNova) {
      for (var i = 0; i < rb.text.length; i++) {
        var dL = String(rb.text[i][0]).trim();
        var hL = String(rd.text[i][0]).trim();
        var dObj = converterStringParaData(dL);
        if (!dObj) continue;
        if (dataNova < dObj) {
          pos = prim + i;
          break;
        } else if (dataNova.getTime() === dObj.getTime()) {
          var hI = hL.split("/")[0].trim();
          var phL = hI.split(":");
          if (phL.length >= 2) {
            var minL = parseInt(phL[0]) * 60 + parseInt(phL[1]);
            if (minNovo < minL) {
              pos = prim + i;
              break;
            }
          }
        }
      }
    }
  }

  aba.getRange(pos + ":" + pos).insert(Excel.InsertShiftDirection.down);
  var rnl = aba.getRange("A" + pos + ":AE" + pos);
  rnl.clear("Formats");
  await context.sync();

  aba.getCell(pos - 1, 0).values = [[dados[0]]];
  aba.getCell(pos - 1, 1).values = [["'" + dataTexto]];
  aba.getCell(pos - 1, 2).values = [[diaSem]];
  aba.getCell(pos - 1, 3).values = [[horaTexto]];
  aba.getCell(pos - 1, 4).values = [[dados[5]]];
  aba.getCell(pos - 1, 5).values = [[varaTexto]];
  aba.getCell(pos - 1, 6).values = [[contatoVaraTexto]];
  aba.getCell(pos - 1, 7).values = [[competenciaTexto]];
  aba.getCell(pos - 1, 8).values = [[adaptacaoTexto]];
  aba.getCell(pos - 1, 9).values = [["'" + processoTexto]];
  aba.getCell(pos - 1, 10).values = [[vitimasTexto]];
  aba.getCell(pos - 1, 11).values = [[statusTexto]];
  aba.getCell(pos - 1, 12).values = [[entrevistadoresInternosTexto]];
  aba.getCell(pos - 1, 13).values = [[entrevistadoresExternosTexto]];
  aba.getCell(pos - 1, 14).values = [[idadesVitimasTexto]];
  await context.sync();

  await aplicarZebrado(context, aba, prim, ult + 1, "AE");
  if (estava) {
    aba.protection.protect({}, SENHA_ABA);
    await context.sync();
  }
  return true;
}

async function removerDaAbaPolo(context, nomeAba, id) {
  var s = context.workbook.worksheets.getItemOrNullObject(nomeAba);
  await context.sync();
  if (s.isNullObject) return false;
  s.load("protection/protected");
  await context.sync();
  var estava = s.protection.protected;
  if (estava) {
    s.protection.unprotect(SENHA_ABA);
    await context.sync();
  }
  var ra = s.getRange("A2:A200");
  ra.load("values");
  await context.sync();
  var idAlvo = normalizarId(id);
  var enc = -1;
  for (var i = 0; i < ra.values.length; i++) {
    if (normalizarId(ra.values[i][0]) === idAlvo) {
      enc = i + 2;
      break;
    }
  }
  if (enc !== -1) {
    s.getRange(enc + ":" + enc).delete(Excel.DeleteShiftDirection.up);
    await context.sync();
  }
  if (estava) {
    s.protection.protect({}, SENHA_ABA);
    await context.sync();
  }
  return enc !== -1;
}

// ============================================================
// SALVAR AUDIÊNCIA (NOVO)
// ============================================================

async function salvarData() {
  if (modoEdicaoAudiencia) return salvarAudienciaEdicao();

  var inputData = document.getElementById("dataInput");
  var timeInput = document.getElementById("timeInput");
  var poloSelect = document.getElementById("poloSelect");
  var novaVaraCheck = document.getElementById("novaVaraCheck");
  var varaSelect = document.getElementById("varaSelect");
  var varaContatoDisplay = document.getElementById("varaContatoDisplay");
  var varaNomeInput = document.getElementById("varaNomeInput");
  var varaContatoInput = document.getElementById("varaContatoInput");
  var novaCompetenciaCheck = document.getElementById("novaCompetenciaCheck");
  var competenciaSelect = document.getElementById("competenciaSelect");
  var competenciaNomeInput = document.getElementById("competenciaNomeInput");
  var processoInput = document.getElementById("processoInput");
  var erroProcessoDiv = document.getElementById("erroProcesso");
  var msgDiv = document.getElementById("mensagem");

  msgDiv.className = "mensagem hidden";
  msgDiv.textContent = "";
  if (erroProcessoDiv) erroProcessoDiv.classList.add("hidden");

  if (!inputData.value) {
    exibirMensagem("Selecione uma data.", "erro");
    return;
  }
  if (!timeInput.value || timeInput.value.length !== 5) {
    exibirMensagem("Informe o horário.", "erro");
    return;
  }
  if (!poloSelect.value) {
    exibirMensagem("Selecione o polo.", "erro");
    return;
  }

  var ehNovaVara = novaVaraCheck.checked;
  var varaFinal = "",
    contatoVaraFinal = "";
  if (ehNovaVara) {
    varaFinal = varaNomeInput.value.trim();
    contatoVaraFinal = varaContatoInput.value.trim();
    if (!varaFinal || !contatoVaraFinal) {
      exibirMensagem("Preencha nome e contato da nova Vara.", "erro");
      return;
    }
  } else {
    varaFinal = varaSelect.value;
    contatoVaraFinal = varaContatoDisplay.value.trim();
    if (!varaFinal) {
      exibirMensagem("Selecione uma Vara.", "erro");
      return;
    }
  }

  var ehNovaCompetencia = novaCompetenciaCheck.checked;
  var competenciaFinal = "";
  if (ehNovaCompetencia) {
    competenciaFinal = competenciaNomeInput.value.trim();
    if (!competenciaFinal) {
      exibirMensagem("Preencha o nome da nova Competência.", "erro");
      return;
    }
  } else {
    competenciaFinal = competenciaSelect.value;
    if (!competenciaFinal) {
      exibirMensagem("Selecione uma Competência.", "erro");
      return;
    }
  }

  var digitos = extrairDigitosProcesso(processoInput.value);
  if (digitos.length !== 20) {
    exibirMensagem("Processo precisa ter 20 dígitos.", "erro");
    return;
  }
  var processoFormatado = formatarProcesso(digitos);

  var vitimas = coletarVitimas();
  if (vitimas === null) {
    exibirMensagem("Preencha Vítima(s) corretamente.", "erro");
    return;
  }
  var nomesVitimasFinal = vitimas.nomes.join("/");
  var idadesVitimasFinal = vitimas.idades.join("/");

  var adaptacaoObj = coletarAdaptacao();
  if (!adaptacaoObj) {
    exibirMensagem("Selecione a adaptação (SIM ou NÃO).", "erro");
    return;
  }
  if (adaptacaoObj.tipo === "SIM" && adaptacaoObj.descricao === "") {
    exibirMensagem("Descreva a adaptação (obrigatório).", "erro");
    return;
  }
  var adaptacaoFinal = formatarAdaptacaoFinal(adaptacaoObj);

  if (entrevistadoresSelecionadosAudiencia.length === 0) {
    exibirMensagem("Adicione pelo menos um entrevistador.", "erro");
    return;
  }
  var entrevistadorFinal = entrevistadoresSelecionadosAudiencia.join("/");
  var emailEntrevistadorFinal = emailsEntrevistadoresSelecionados();
  var tiposSelecionados = tiposEntrevistadoresSelecionados();
  var algumExterno = algumEntrevistadorExterno();
  var tipoParaColunaPolo = algumExterno ? "CAD EXTERNO" : "INTERNO";

  var separados = separarEntrevistadoresPorTipo();
  var internosParaPolo = separados.internos;
  var externosParaPolo = separados.externos;

  var dataFormatada = converterParaBR(inputData.value);
  var dataObj = dateObjFromISO(inputData.value);
  var poloSelecionado = poloSelect.value;

  var dataSel = new Date(inputData.value + "T00:00:00");
  if (isNaN(dataSel.getTime())) {
    exibirMensagem("Data inválida.", "erro");
    return;
  }
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  if (dataSel < hoje) {
    exibirMensagem("A data não pode ser anterior a hoje.", "erro");
    return;
  }
  var diaSem = dataSel.getDay();
  if (diaSem === 0 || diaSem === 6) {
    exibirMensagem("A data não pode ser fim de semana.", "erro");
    return;
  }

  var intervalo = calcularIntervalo(timeInput.value);
  if (intervalo === "") {
    exibirMensagem("Horário inválido.", "erro");
    return;
  }

  var novoID = "";
  var inseridoNaAba = false;
  var statusInicial = "PENDENTE";
  var salvou = false;

  definirProcessando(true);
  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItem("Formulário");

      var rangeF = sheet.getRange("AB2:AB30");
      rangeF.load("text");
      await context.sync();
      var feriados = [];
      for (var i = 0; i < rangeF.text.length; i++) {
        var v = rangeF.text[i][0];
        if (v !== "" && v !== null && v !== undefined) feriados.push(String(v).trim());
      }
      if (feriados.indexOf(dataFormatada) !== -1) {
        exibirMensagem("A data " + dataFormatada + " é feriado.", "erro");
        definirProcessando(false);
        atualizarEstadoBotao();
        return;
      }

      if (await processoJaExiste(context, sheet, digitos, null)) {
        exibirMensagem("Este processo já está cadastrado.", "erro");
        if (erroProcessoDiv) {
          erroProcessoDiv.textContent = "⚠️ Este processo já existe.";
          erroProcessoDiv.classList.remove("hidden");
        }
        definirProcessando(false);
        atualizarEstadoBotao();
        return;
      }

      sheet.load("protection/protected");
      await context.sync();
      var estava = sheet.protection.protected;
      if (estava) {
        sheet.protection.unprotect(SENHA_ABA);
        await context.sync();
      }

      var autor = await obterAutor(context, sheet);

      var interval = sheet.getRange("A:A").getUsedRangeOrNullObject(true);
      interval.load("values");
      await context.sync();
      var ultimoID = 0;
      if (!interval.isNullObject) {
        for (var i2 = 0; i2 < interval.values.length; i2++) {
          var c = interval.values[i2][0];
          if (c !== "" && c !== null && c !== undefined) {
            var n = parseInt(String(c).replace(/\D/g, ""), 10);
            if (!isNaN(n) && n > ultimoID) ultimoID = n;
          }
        }
      }

      novoID = padLeft(String(ultimoID + 1), 3);
      var dataHoraEnvio = new Date().toLocaleString("pt-BR");

      var diffDias = Math.round((dataObj.getTime() - hoje.getTime()) / 86400000);
      statusInicial = calcularStatusInicial(entrevistadoresSelecionadosAudiencia, diffDias);

      var tabela = sheet.tables.getItem("Solicitacoes");
      tabela.rows.add(null, [
        [
          novoID,
          dataHoraEnvio,
          autor,
          dateParaSerialExcel(dataObj),
          intervalo,
          poloSelecionado,
          varaFinal,
          contatoVaraFinal,
          competenciaFinal,
          processoFormatado,
          nomesVitimasFinal,
          idadesVitimasFinal,
          adaptacaoFinal,
          entrevistadorFinal,
          emailEntrevistadorFinal,
          statusInicial,
        ],
      ]);
      await context.sync();

      var rangeTabela = tabela.getRange();
      rangeTabela.load("rowIndex, rowCount");
      await context.sync();
      var ultimaLinha = rangeTabela.rowIndex + rangeTabela.rowCount - 1;
      var rangeKL = sheet.getRange("K" + ultimaLinha + ":L" + ultimaLinha);
      rangeKL.numberFormat = [["@", "@"]];
      rangeKL.values = [[nomesVitimasFinal, idadesVitimasFinal]];
      await context.sync();

      var colD = tabela.columns.getItemAt(3).getDataBodyRange();
      if (colD) {
        colD.numberFormat = [["dd/mm/yyyy"]];
        await context.sync();
      }

      if (estava) {
        sheet.protection.protect({}, SENHA_ABA);
        await context.sync();
      }

      if (ATIVAR_DISTRIBUICAO) {
        if (ehNovaVara) {
          var cs = context.workbook.worksheets.getItem("CADASTROS");
          cs.load("protection/protected");
          await context.sync();
          var p = cs.protection.protected;
          if (p) {
            cs.protection.unprotect(SENHA_ABA);
            await context.sync();
          }
          var ra = cs.getRange("AA2:AA1000");
          ra.load("values");
          await context.sync();
          var pr = 2;
          for (var i3 = 0; i3 < ra.values.length; i3++) {
            var v3 = String(ra.values[i3][0]).trim();
            if (v3 === "" || v3 === "null" || v3 === "undefined") {
              pr = i3 + 2;
              break;
            }
          }
          cs.getCell(pr - 1, 26).values = [[varaFinal]];
          cs.getCell(pr - 1, 27).values = [[contatoVaraFinal]];
          await context.sync();
          if (p) {
            cs.protection.protect({}, SENHA_ABA);
            await context.sync();
          }
        }

        if (ehNovaCompetencia) {
          var cs2 = context.workbook.worksheets.getItem("CADASTROS");
          cs2.load("protection/protected");
          await context.sync();
          var p2 = cs2.protection.protected;
          if (p2) {
            cs2.protection.unprotect(SENHA_ABA);
            await context.sync();
          }
          var ra2 = cs2.getRange("AC2:AC1000");
          ra2.load("values");
          await context.sync();
          var pr2 = 2;
          for (var i4 = 0; i4 < ra2.values.length; i4++) {
            var v4 = String(ra2.values[i4][0]).trim();
            if (v4 === "" || v4 === "null" || v4 === "undefined") {
              pr2 = i4 + 2;
              break;
            }
          }
          cs2.getCell(pr2 - 1, 28).values = [[competenciaFinal]];
          await context.sync();
          if (p2) {
            cs2.protection.protect({}, SENHA_ABA);
            await context.sync();
          }
        }

        var cs3 = context.workbook.worksheets.getItemOrNullObject("CADASTROS");
        await context.sync();
        if (!cs3.isNullObject) {
          cs3.load("protection/protected");
          await context.sync();
          var p3 = cs3.protection.protected;
          if (p3) {
            cs3.protection.unprotect(SENHA_ABA);
            await context.sync();
          }

          var ra3 = cs3.getRange("AD2:AE1000");
          ra3.load("values");
          await context.sync();

          var vitimasExistentes = new Set();
          var pr3 = 2;

          for (var i5 = 0; i5 < ra3.values.length; i5++) {
            var nomeExist = String(ra3.values[i5][0] || "").trim();
            var idadeExist = String(ra3.values[i5][1] || "")
              .trim()
              .replace(/^'/, "");

            if (nomeExist !== "" && nomeExist !== "null" && nomeExist !== "undefined") {
              vitimasExistentes.add(nomeExist.toLowerCase() + "|" + idadeExist.toLowerCase());
              pr3 = i5 + 3;
            }
          }

          for (var i6 = 0; i6 < vitimas.nomes.length; i6++) {
            var nomeLimpo = vitimas.nomes[i6].trim();
            var idadeLimpa = vitimas.idades[i6].trim();
            var chave = nomeLimpo.toLowerCase() + "|" + idadeLimpa.toLowerCase();

            if (!vitimasExistentes.has(chave)) {
              cs3.getCell(pr3 - 1, 29).values = [[nomeLimpo]];
              cs3.getCell(pr3 - 1, 30).values = [["'" + idadeLimpa]];
              vitimasExistentes.add(chave);
              pr3++;
            }
          }

          await context.sync();
          if (p3) {
            cs3.protection.protect({}, SENHA_ABA);
            await context.sync();
          }
        }

        var dp = [novoID, dataHoraEnvio, autor, dataFormatada, intervalo, poloSelecionado];
        var ok = await inserirNaAbaPolo(
          context,
          poloSelecionado,
          dp,
          dataFormatada,
          intervalo,
          varaFinal,
          contatoVaraFinal,
          competenciaFinal,
          adaptacaoFinal,
          processoFormatado,
          nomesVitimasFinal,
          idadesVitimasFinal,
          internosParaPolo,
          externosParaPolo,
          statusInicial
        );
        inseridoNaAba = ok;
        if (ok) exibirMensagem("Audiência " + novoID + " salva.", "sucesso");
        else exibirMensagem("Audiência " + novoID + " salva, mas aba " + poloSelecionado + " não encontrada.", "erro");
      } else {
        exibirMensagem("Audiência " + novoID + " salva apenas na Formulário.", "sucesso");
      }

      await registrarLog(
        context,
        "CRIAR",
        "Audiência",
        novoID + " · Proc " + processoFormatado,
        autor,
        "Polo: " +
          poloSelecionado +
          " · Data: " +
          dataFormatada +
          " · Hora: " +
          intervalo +
          " · Vítimas: " +
          nomesVitimasFinal +
          " · Adaptação: " +
          adaptacaoFinal +
          " · Entrevistador(es): " +
          entrevistadorFinal,
        {
          evento: "CRIAR_AUDIENCIA",
          id: novoID,
          processo: processoFormatado,
          polo: poloSelecionado,
          data: dataFormatada,
          hora: intervalo,
          vara: varaFinal,
          contatoVara: contatoVaraFinal,
          competencia: competenciaFinal,
          vitimas: nomesVitimasFinal,
          idades: idadesVitimasFinal,
          adaptacao: adaptacaoFinal,
          entrevistador: entrevistadorFinal,
          emailEntrevistador: emailEntrevistadorFinal,
          tipoEntrevistador: tipoParaColunaPolo,
          entrevistadores: entrevistadoresSelecionadosAudiencia.slice(),
          emailsEntrevistadores: entrevistadoresSelecionadosAudiencia.map(function (n2) {
            var e2 = entrevistadoresData.find(function (x) {
              return x.nome === n2;
            });
            return e2 ? e2.email : "";
          }),
          tiposEntrevistadores: tiposSelecionados,
          status: statusInicial,
          timestamp: new Date().toISOString(),
        }
      );

      salvou = true;
    });

    if (!salvou) return;

    await carregarTudoInicial();

    exibirResumoAcao(
      "Audiência " + novoID + " cadastrada",
      [
        "<strong>Data:</strong> " + dataFormatada,
        "<strong>Horário:</strong> " + intervalo,
        "<strong>Polo:</strong> " + poloSelecionado,
        "<strong>Vara:</strong> " + varaFinal + " (" + contatoVaraFinal + ")",
        "<strong>Competência:</strong> " + competenciaFinal,
        "<strong>Processo:</strong> " + processoFormatado,
        "<strong>Vítima(s):</strong> " + nomesVitimasFinal + " — " + idadesVitimasFinal,
        "<strong>Adaptação:</strong> " + adaptacaoFinal,
        "<strong>Entrevistador(es):</strong> " + entrevistadorFinal,
        "<strong>Status:</strong> " + statusInicial,
      ],
      inseridoNaAba ? "Salvo em Formulário e aba " + poloSelecionado + "." : "Salvo só na Formulário.",
      "sucesso"
    );
  } catch (error) {
    console.error(error);
    exibirMensagem("Erro ao acessar a planilha.", "erro");
  } finally {
    definirProcessando(false);
    atualizarEstadoBotao();
  }
}

async function salvarAudienciaEdicao() {
  if (!audienciaEmEdicao) return;

  var inputData = document.getElementById("dataInput");
  var timeInput = document.getElementById("timeInput");
  var poloSelect = document.getElementById("poloSelect");
  var novaVaraCheck = document.getElementById("novaVaraCheck");
  var varaSelect = document.getElementById("varaSelect");
  var varaContatoDisplay = document.getElementById("varaContatoDisplay");
  var varaNomeInput = document.getElementById("varaNomeInput");
  var varaContatoInput = document.getElementById("varaContatoInput");
  var novaCompetenciaCheck = document.getElementById("novaCompetenciaCheck");
  var competenciaSelect = document.getElementById("competenciaSelect");
  var competenciaNomeInput = document.getElementById("competenciaNomeInput");
  var processoInput = document.getElementById("processoInput");

  if (!inputData.value || !timeInput.value || !poloSelect.value) {
    exibirMensagem("Preencha Data, Horário e Polo.", "erro");
    return;
  }

  var ehNovaVara = novaVaraCheck.checked;
  var varaFinal = ehNovaVara ? varaNomeInput.value.trim() : varaSelect.value;
  var contatoVaraFinal = ehNovaVara ? varaContatoInput.value.trim() : varaContatoDisplay.value.trim();
  if (!varaFinal) {
    exibirMensagem("Vara inválida.", "erro");
    return;
  }

  var ehNovaCompetencia = novaCompetenciaCheck.checked;
  var competenciaFinal = ehNovaCompetencia ? competenciaNomeInput.value.trim() : competenciaSelect.value;
  if (!competenciaFinal) {
    exibirMensagem("Competência inválida.", "erro");
    return;
  }

  var digitos = extrairDigitosProcesso(processoInput.value);
  if (digitos.length !== 20) {
    exibirMensagem("Processo precisa ter 20 dígitos.", "erro");
    return;
  }

  var vitimas = coletarVitimas();
  if (vitimas === null) {
    exibirMensagem("Vítimas inválidas.", "erro");
    return;
  }

  var adaptacaoObj = coletarAdaptacao();
  if (!adaptacaoObj) {
    exibirMensagem("Selecione a adaptação.", "erro");
    return;
  }
  if (adaptacaoObj.tipo === "SIM" && adaptacaoObj.descricao === "") {
    exibirMensagem("Descreva a adaptação.", "erro");
    return;
  }
  var adaptacaoFinal = formatarAdaptacaoFinal(adaptacaoObj);

  var statusObj = coletarStatus();
  if ((statusObj.tipo === "CANCELADO" || statusObj.tipo === "NÃO REALIZADO") && statusObj.motivo === "") {
    exibirMensagem(
      "Descreva o motivo do " + (statusObj.tipo === "CANCELADO" ? "cancelamento" : "não realizado") + ".",
      "erro"
    );
    return;
  }
  var statusFinal = formatarStatusFinal(statusObj);

  var intervalo = calcularIntervalo(timeInput.value);
  if (intervalo === "") {
    exibirMensagem("Horário inválido.", "erro");
    return;
  }

  if (entrevistadoresSelecionadosAudiencia.length === 0) {
    exibirMensagem("Adicione pelo menos um entrevistador.", "erro");
    return;
  }

  var dataFormatada = converterParaBR(inputData.value);
  var dataObj = dateObjFromISO(inputData.value);
  var processoFormatado = formatarProcesso(digitos);
  var poloSelecionado = poloSelect.value;
  var nomesVitimasFinal = vitimas.nomes.join("/");
  var idadesVitimasFinal = vitimas.idades.join("/");
  var entrevistadorFinal = entrevistadoresSelecionadosAudiencia.join("/");
  var emailEntrevistadorFinal = emailsEntrevistadoresSelecionados();
  var tiposSelecionados = tiposEntrevistadoresSelecionados();
  var algumExterno = algumEntrevistadorExterno();
  var tipoParaColunaPolo = algumExterno ? "CAD EXTERNO" : "INTERNO";

  var separados2 = separarEntrevistadoresPorTipo();
  var internosParaPolo2 = separados2.internos;
  var externosParaPolo2 = separados2.externos;

  var dataSel = new Date(inputData.value + "T00:00:00");
  if (isNaN(dataSel.getTime())) {
    exibirMensagem("Data inválida.", "erro");
    return;
  }
  var diaSem = dataSel.getDay();
  if (diaSem === 0 || diaSem === 6) {
    exibirMensagem("A data não pode ser fim de semana.", "erro");
    return;
  }

  var idEditado = audienciaEmEdicao.id;
  var salvou = false;

  definirProcessando(true);
  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItem("Formulário");

      if (await processoJaExiste(context, sheet, digitos, idEditado)) {
        exibirMensagem("Processo já está em outra audiência.", "erro");
        definirProcessando(false);
        atualizarEstadoBotao();
        return;
      }

      sheet.load("protection/protected");
      await context.sync();
      var estava = sheet.protection.protected;
      if (estava) {
        sheet.protection.unprotect(SENHA_ABA);
        await context.sync();
      }

      var autor = await obterAutor(context, sheet);
      var tabela = sheet.tables.getItem("Solicitacoes");
      var hr = tabela.getHeaderRowRange();
      var range = tabela.getRange();
      hr.load("rowIndex");
      range.load("values");
      await context.sync();

      var linhaAlvo = -1;
      for (var i = 1; i < range.values.length; i++) {
        if (String(range.values[i][0]).trim() === idEditado) {
          linhaAlvo = i;
          break;
        }
      }
      if (linhaAlvo === -1) {
        exibirMensagem("Audiência não encontrada.", "erro");
        if (estava) {
          sheet.protection.protect({}, SENHA_ABA);
          await context.sync();
        }
        definirProcessando(false);
        atualizarEstadoBotao();
        return;
      }

      var linhaReal = hr.rowIndex + 1 + linhaAlvo;
      var row = sheet.getRange("A" + linhaReal + ":P" + linhaReal);
      row.values = [
        [
          idEditado,
          String(range.values[linhaAlvo][1]),
          String(range.values[linhaAlvo][2]),
          dateParaSerialExcel(dataObj),
          intervalo,
          poloSelecionado,
          varaFinal,
          contatoVaraFinal,
          competenciaFinal,
          processoFormatado,
          nomesVitimasFinal,
          idadesVitimasFinal,
          adaptacaoFinal,
          entrevistadorFinal,
          emailEntrevistadorFinal,
          statusFinal,
        ],
      ];
      await context.sync();

      var rangeKL = sheet.getRange("K" + linhaReal + ":L" + linhaReal);
      rangeKL.numberFormat = [["@", "@"]];
      rangeKL.values = [[nomesVitimasFinal, idadesVitimasFinal]];
      await context.sync();

      var colD = tabela.columns.getItemAt(3).getDataBodyRange();
      if (colD) {
        colD.numberFormat = [["dd/mm/yyyy"]];
        await context.sync();
      }

      if (estava) {
        sheet.protection.protect({}, SENHA_ABA);
        await context.sync();
      }

      if (ATIVAR_DISTRIBUICAO) {
        var poloAntigo = audienciaEmEdicao.polo;
        var poloNovo = poloSelecionado;
        var statusAtual = statusFinal;
        var isExt = algumExterno;

        if (poloAntigo === poloNovo) {
          var aba = context.workbook.worksheets.getItemOrNullObject(poloNovo);
          aba.load("name");
          await context.sync();
          if (!aba.isNullObject) {
            aba.load("protection/protected");
            await context.sync();
            var prot = aba.protection.protected;
            if (prot) {
              aba.protection.unprotect(SENHA_ABA);
              await context.sync();
            }
            var ra = aba.getRange("A2:A200");
            ra.load("values");
            await context.sync();
            var idAlvo = normalizarId(idEditado);
            var linha = -1;
            for (var i2 = 0; i2 < ra.values.length; i2++) {
              if (normalizarId(ra.values[i2][0]) === idAlvo) {
                linha = i2 + 2;
                break;
              }
            }
            if (linha !== -1) {
              var pd = dataFormatada.split("/");
              var dSem = calcularDiaSemana(parseInt(pd[0]), parseInt(pd[1]), parseInt(pd[2]));
              aba.getCell(linha - 1, 1).values = [["'" + dataFormatada]];
              aba.getCell(linha - 1, 2).values = [[dSem]];
              aba.getCell(linha - 1, 3).values = [[intervalo]];
              aba.getCell(linha - 1, 5).values = [[varaFinal]];
              aba.getCell(linha - 1, 6).values = [[contatoVaraFinal]];
              aba.getCell(linha - 1, 7).values = [[competenciaFinal]];
              aba.getCell(linha - 1, 8).values = [[adaptacaoFinal]];
              aba.getCell(linha - 1, 9).values = [["'" + processoFormatado]];
              aba.getCell(linha - 1, 10).values = [[nomesVitimasFinal]];
              aba.getCell(linha - 1, 11).values = [[statusAtual]];
              aba.getCell(linha - 1, 12).values = [[internosParaPolo2]];
              aba.getCell(linha - 1, 13).values = [[externosParaPolo2]];
              aba.getCell(linha - 1, 14).values = [["'" + idadesVitimasFinal]];
              await context.sync();
              await aplicarFormatacaoLinhaUnica(context, aba, linha, "AE");
            }
            if (prot) {
              aba.protection.protect({}, SENHA_ABA);
              await context.sync();
            }
          }
        } else {
          await removerDaAbaPolo(context, poloAntigo, idEditado);
          var dp = [idEditado, new Date().toLocaleString("pt-BR"), autor, dataFormatada, intervalo, poloNovo];
          await inserirNaAbaPolo(
            context,
            poloNovo,
            dp,
            dataFormatada,
            intervalo,
            varaFinal,
            contatoVaraFinal,
            competenciaFinal,
            adaptacaoFinal,
            processoFormatado,
            nomesVitimasFinal,
            idadesVitimasFinal,
            internosParaPolo2,
            externosParaPolo2,
            statusAtual
          );
        }

        if (ehNovaVara) {
          var cs = context.workbook.worksheets.getItem("CADASTROS");
          cs.load("protection/protected");
          await context.sync();
          var p = cs.protection.protected;
          if (p) {
            cs.protection.unprotect(SENHA_ABA);
            await context.sync();
          }
          var ra = cs.getRange("AA2:AA1000");
          ra.load("values");
          await context.sync();
          var pr = 2;
          for (var i3 = 0; i3 < ra.values.length; i3++) {
            var v = String(ra.values[i3][0]).trim();
            if (v === "" || v === "null" || v === "undefined") {
              pr = i3 + 2;
              break;
            }
          }
          cs.getCell(pr - 1, 26).values = [[varaFinal]];
          cs.getCell(pr - 1, 27).values = [[contatoVaraFinal]];
          await context.sync();
          if (p) {
            cs.protection.protect({}, SENHA_ABA);
            await context.sync();
          }
        }
        if (ehNovaCompetencia) {
          var cs2 = context.workbook.worksheets.getItem("CADASTROS");
          cs2.load("protection/protected");
          await context.sync();
          var p2 = cs2.protection.protected;
          if (p2) {
            cs2.protection.unprotect(SENHA_ABA);
            await context.sync();
          }
          var ra2 = cs2.getRange("AC2:AC1000");
          ra2.load("values");
          await context.sync();
          var pr2 = 2;
          for (var i4 = 0; i4 < ra2.values.length; i4++) {
            var v4 = String(ra2.values[i4][0]).trim();
            if (v4 === "" || v4 === "null" || v4 === "undefined") {
              pr2 = i4 + 2;
              break;
            }
          }
          cs2.getCell(pr2 - 1, 28).values = [[competenciaFinal]];
          await context.sync();
          if (p2) {
            cs2.protection.protect({}, SENHA_ABA);
            await context.sync();
          }
        }
      }

      var mudancas = {};
      var registrar = function (campo, de, para) {
        if (String(de || "").trim() !== String(para || "").trim()) {
          mudancas[campo] = { de: String(de || ""), para: String(para || "") };
        }
      };
      registrar("polo", audienciaEmEdicao.polo, poloSelecionado);
      registrar("data", audienciaEmEdicao.data, dataFormatada);
      registrar("hora", audienciaEmEdicao.hora, intervalo);
      registrar("vara", audienciaEmEdicao.vara, varaFinal);
      registrar("contatoVara", audienciaEmEdicao.contatoVara, contatoVaraFinal);
      registrar("competencia", audienciaEmEdicao.competencia, competenciaFinal);
      registrar("processo", audienciaEmEdicao.processo, processoFormatado);
      registrar("vitimas", audienciaEmEdicao.vitimas, nomesVitimasFinal);
      registrar("idades", audienciaEmEdicao.idades, idadesVitimasFinal);
      registrar("adaptacao", audienciaEmEdicao.adaptacao, adaptacaoFinal);
      registrar("entrevistador", audienciaEmEdicao.entrevistador, entrevistadorFinal);
      registrar("status", audienciaEmEdicao.status, statusFinal);

      var entTipoAtual = algumExterno ? "CAD EXTERNO" : "INTERNO";

      var payloadEdicao =
        Object.keys(mudancas).length > 0
          ? {
              evento: "EDITAR_AUDIENCIA",
              id: idEditado,
              processo: processoFormatado,
              polo: poloSelecionado,
              mudancas: mudancas,
              trocaEntrevistador: mudancas["entrevistador"] ? mudancas["entrevistador"] : null,
              trocaPolo: mudancas["polo"] ? mudancas["polo"] : null,
              emailDe: audienciaEmEdicao.emailEntrevistador,
              emailPara: emailEntrevistadorFinal,
              tipoEntrevistador: entTipoAtual,
              entrevistadores: entrevistadoresSelecionadosAudiencia.slice(),
              emailsEntrevistadores: entrevistadoresSelecionadosAudiencia.map(function (n2) {
                var e2 = entrevistadoresData.find(function (x) {
                  return x.nome === n2;
                });
                return e2 ? e2.email : "";
              }),
              tiposEntrevistadores: tiposSelecionados,
              statusAtual: statusFinal,
              timestamp: new Date().toISOString(),
            }
          : null;

      await registrarLog(
        context,
        "EDITAR",
        "Audiência",
        idEditado + " · Proc " + processoFormatado,
        autor,
        "Polo: " +
          poloSelecionado +
          " · Data: " +
          dataFormatada +
          " · Hora: " +
          intervalo +
          " · Vítimas: " +
          nomesVitimasFinal +
          " · Adaptação: " +
          adaptacaoFinal +
          " · Entrevistador(es): " +
          entrevistadorFinal +
          " · Status: " +
          statusFinal,
        payloadEdicao
      );

      salvou = true;
    });

    if (!salvou) return;

    await carregarTudoInicial();

    exibirResumoAcao(
      "Audiência " + idEditado + " atualizada",
      [
        "<strong>Data:</strong> " + dataFormatada,
        "<strong>Horário:</strong> " + intervalo,
        "<strong>Polo:</strong> " + poloSelecionado,
        "<strong>Vara:</strong> " + varaFinal + " (" + contatoVaraFinal + ")",
        "<strong>Competência:</strong> " + competenciaFinal,
        "<strong>Processo:</strong> " + processoFormatado,
        "<strong>Vítima(s):</strong> " + nomesVitimasFinal + " — " + idadesVitimasFinal,
        "<strong>Adaptação:</strong> " + adaptacaoFinal,
        "<strong>Entrevistador(es):</strong> " + entrevistadorFinal,
        "<strong>Status:</strong> " + statusFinal,
      ],
      "Alterações gravadas.",
      "sucesso"
    );
  } catch (error) {
    console.error(error);
    exibirMensagem("Erro ao atualizar audiência.", "erro");
  } finally {
    definirProcessando(false);
    atualizarEstadoBotao();
  }
}

// ============================================================
// SALVAR ENTREVISTADOR
// ============================================================

async function salvarEntrevistador() {
  var nome = document.getElementById("entNomeInput").value.trim();
  var email = document.getElementById("entEmailInput").value.trim();
  var tipo = document.getElementById("entTipoSelect").value;
  var prioridade = document.getElementById("entPrioridadeSelect").value;
  var disponibilidade = document.getElementById("entDisponibilidadeSelect").value;
  var especialidade = document.getElementById("entEspecialidadeInput").value.trim();
  var lotacao = document.getElementById("entLotacaoInput").value.trim();
  var obs = document.getElementById("entObsInput").value.trim();
  var dias = coletarDiasSelecionados();
  var polosStr = polosSelecionados.join(", ");

  var msg = document.getElementById("entMensagem");
  msg.className = "mensagem hidden";
  msg.textContent = "";

  if (!nome) {
    exibirMensagemEnt("Preencha o nome.", "erro");
    return;
  }
  if (!email) {
    exibirMensagemEnt("Preencha o e-mail.", "erro");
    return;
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    exibirMensagemEnt("E-mail inválido.", "erro");
    return;
  }
  if (!tipo) {
    exibirMensagemEnt("Selecione o tipo.", "erro");
    return;
  }
  if (!prioridade) {
    exibirMensagemEnt("Selecione a prioridade.", "erro");
    return;
  }
  if (!disponibilidade) {
    exibirMensagemEnt("Selecione a disponibilidade.", "erro");
    return;
  }
  if (!especialidade) {
    exibirMensagemEnt("Preencha a especialidade.", "erro");
    return;
  }
  if (polosSelecionados.length === 0) {
    exibirMensagemEnt("Adicione pelo menos um polo.", "erro");
    return;
  }
  if (!dias) {
    exibirMensagemEnt("Selecione ao menos um dia.", "erro");
    return;
  }

  var estavaEmEdicao = modoEdicaoEntrevistador;

  definirProcessando(true);
  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getItem("CADASTROS");
      sheet.load("protection/protected");
      await context.sync();
      var estava = sheet.protection.protected;
      if (estava) {
        sheet.protection.unprotect(SENHA_ABA);
        await context.sync();
      }

      var autor = await obterAutor(context, sheet);

      var colA = sheet.getRange("A2:A1000");
      colA.load("values");
      await context.sync();
      for (var i = 0; i < colA.values.length; i++) {
        var v = String(colA.values[i][0]).trim().toLowerCase();
        if (v === nome.toLowerCase()) {
          if (
            estavaEmEdicao &&
            entrevistadorEmEdicao &&
            entrevistadorEmEdicao.nome.toLowerCase() === nome.toLowerCase()
          )
            break;
          if (!estavaEmEdicao) {
            exibirMensagemEnt("Já existe um entrevistador com esse nome.", "erro");
            if (estava) {
              sheet.protection.protect({}, SENHA_ABA);
              await context.sync();
            }
            definirProcessando(false);
            return;
          }
        }
      }

      var linhaDados = [nome, email, tipo, prioridade, disponibilidade, especialidade, lotacao, polosStr, dias, obs];

      if (estavaEmEdicao) {
        if (!entrevistadorEmEdicao) {
          exibirMensagemEnt("Nenhum entrevistador selecionado para edição.", "erro");
          if (estava) {
            sheet.protection.protect({}, SENHA_ABA);
            await context.sync();
          }
          definirProcessando(false);
          return;
        }
        var linha = entrevistadorEmEdicao.rowIndex;
        sheet.getRange("A" + linha + ":J" + linha).values = [linhaDados];
        await context.sync();
        await registrarLog(
          context,
          "EDITAR",
          "Entrevistador",
          nome,
          autor,
          "Tipo: " +
            tipo +
            " · Prioridade: " +
            prioridade +
            " · Disponibilidade: " +
            disponibilidade +
            " · Especialidade: " +
            especialidade +
            " · Polos: " +
            polosStr +
            " · Dias: " +
            dias
        );
        if (estava) {
          sheet.protection.protect({}, SENHA_ABA);
          await context.sync();
        }
      } else {
        var colAN = sheet.getRange("A2:A1000");
        colAN.load("values");
        await context.sync();
        var pr = 2;
        for (var i2 = 0; i2 < colAN.values.length; i2++) {
          var v2 = String(colAN.values[i2][0]).trim();
          if (v2 === "" || v2 === "null" || v2 === "undefined") {
            pr = i2 + 2;
            break;
          }
          pr = i2 + 3;
        }
        sheet.getRange("A" + pr + ":J" + pr).values = [linhaDados];
        await context.sync();
        await registrarLog(
          context,
          "CRIAR",
          "Entrevistador",
          nome,
          autor,
          "Tipo: " +
            tipo +
            " · Prioridade: " +
            prioridade +
            " · Disponibilidade: " +
            disponibilidade +
            " · Especialidade: " +
            especialidade +
            " · Polos: " +
            polosStr +
            " · Dias: " +
            dias
        );
        if (estava) {
          sheet.protection.protect({}, SENHA_ABA);
          await context.sync();
        }
      }
    });

    await carregarTudoInicial();

    exibirResumoAcao(
      estavaEmEdicao
        ? "Entrevistador(a) '" + nome + "' atualizado(a)"
        : "Entrevistador(a) '" + nome + "' cadastrado(a)",
      [
        "<strong>Nome:</strong> " + nome,
        "<strong>E-mail:</strong> " + email,
        "<strong>Tipo:</strong> " + tipo,
        "<strong>Prioridade:</strong> " + prioridade,
        "<strong>Disponibilidade:</strong> " + disponibilidade,
        "<strong>Especialidade:</strong> " + especialidade,
        "<strong>Lotação:</strong> " + (lotacao || "—"),
        "<strong>Polos:</strong> " + polosStr,
        "<strong>Dias:</strong> " + dias,
      ],
      "Alterações gravadas na aba CADASTROS.",
      "sucesso"
    );
  } catch (error) {
    console.error(error);
    exibirMensagemEnt("Erro ao salvar entrevistador.", "erro");
  } finally {
    definirProcessando(false);
  }
}

// ============================================================
// UTILITÁRIOS DE MENSAGEM
// ============================================================

function exibirMensagem(texto, tipo) {
  var d = document.getElementById("mensagem");
  d.textContent = texto;
  d.className = "mensagem " + tipo;
}
function exibirMensagemEnt(texto, tipo) {
  var d = document.getElementById("entMensagem");
  d.textContent = texto;
  d.className = "mensagem " + tipo;
}

function exibirResumoAcao(titulo, linhas, nota, tipo) {
  var painel = document.getElementById("resultadoPanel");
  var t = document.getElementById("resultadoTitulo");
  var d = document.getElementById("resultadoDetalhes");
  var b = document.getElementById("btnReiniciar");
  if (!painel || !t || !d || !b) return;
  t.textContent = titulo;
  t.className = tipo;
  var html = "<ul>";
  for (var i = 0; i < linhas.length; i++) html += "<li>" + linhas[i] + "</li>";
  html += "</ul>";
  if (nota) html += "<em>" + nota + "</em>";
  d.innerHTML = html;
  b.onclick = function () {
    reiniciarFormulario();
  };
  painel.classList.remove("hidden");
}

async function reiniciarFormulario() {
  var p = document.getElementById("resultadoPanel");
  if (p) p.classList.add("hidden");

  modoEdicaoAudiencia = false;
  audienciaEmEdicao = null;

  var audCheck = document.getElementById("audienciaJaExiste");
  if (audCheck) audCheck.checked = false;
  var sg = document.getElementById("statusFieldGroup");
  if (sg) sg.classList.add("hidden");
  var as = document.getElementById("audienciaEditSelector");
  if (as) as.classList.add("hidden");
  var af = document.getElementById("audienciaFormPanel");
  if (af) af.classList.remove("hidden");
  var bs = document.getElementById("btnSalvar");
  if (bs) bs.textContent = "Cadastrar";

  limparFormularioAudiencia();

  var ma = document.getElementById("mensagem");
  if (ma) {
    ma.className = "mensagem hidden";
    ma.textContent = "";
  }
  var ep = document.getElementById("erroProcesso");
  if (ep) ep.classList.add("hidden");
  var sv = document.getElementById("sugestoesVitimaList");
  if (sv) {
    sv.innerHTML = "";
    sv.classList.add("hidden");
  }
  var ar = document.getElementById("avisoRepeticaoVitima");
  if (ar) {
    ar.classList.add("hidden");
    ar.textContent = "";
  }

  modoEdicaoEntrevistador = false;
  entrevistadorEmEdicao = null;
  var ec = document.getElementById("entrevistadorJaExiste");
  if (ec) ec.checked = false;
  var es = document.getElementById("entrevistadorEditSelector");
  if (es) es.classList.add("hidden");
  var ef = document.getElementById("entrevistadorFormPanel");
  if (ef) ef.classList.remove("hidden");
  var be = document.getElementById("btnSalvarEntrevistador");
  if (be) be.textContent = "Cadastrar";

  limparFormularioEntrevistador();

  var me = document.getElementById("entMensagem");
  if (me) {
    me.className = "mensagem hidden";
    me.textContent = "";
  }

  await carregarTudoInicial();
  atualizarEstadoBotao();
  atualizarEstadoBotaoEntrevistador();
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
    exibirMensagem("Ocorreu um erro inesperado.", "erro");
  }
}