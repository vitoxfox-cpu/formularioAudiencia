const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

function carregarScript() {
  const contexto = {
    window: { addEventListener() {} },
    Office: { onReady() {} },
    console,
  };
  vm.createContext(contexto);
  vm.runInContext(fs.readFileSync(__dirname + "/script.js", "utf8"), contexto);
  return contexto;
}

test("calendário distingue internos disponíveis de dias negociáveis", () => {
  const app = carregarScript();
  app.dadosCalendarioCarregados = true;
  app.entrevistadoresData = [{
    nome: "Marina", tipo: "INTERNO", disponibilidade: "DISPONÍVEL",
    polos: "A", dias: "Segunda", prioridade: "OBRIGATÓRIO",
    qtdCancelado: 0, qtdMarcado: 0, qtdRealizado: 0,
  }];
  assert.equal(app.tipoDiaCalendario("2099-01-05", "A"), "disponivel");
  assert.equal(app.tipoDiaCalendario("2099-01-06", "A"), "negociavel");
  assert.equal(app.tipoDiaCalendario("2099-01-05", "B"), "negociavel");
  app.entrevistadoresData[0].disponibilidade = "FÉRIAS";
  assert.equal(app.tipoDiaCalendario("2099-01-05", "A"), "negociavel");
  assert.equal(app.tipoDiaCalendario("2099-01-05", ""), "indisponivel");
});

test("fins de semana, datas passadas e datas da planilha são bloqueados", () => {
  const app = carregarScript();
  app.dadosCalendarioCarregados = true;
  app.diasNaoUteis = app.datasNaoUteisDoRange({
    values: [["6/1/2099"], [72688]],
    text: [["6/1/2099"], [""]],
  });
  assert.equal(app.diasNaoUteis[0], "06/01/2099");
  assert.equal(app.tipoDiaCalendario("2099-01-06", "A"), "indisponivel");
  assert.equal(app.tipoDiaCalendario("2099-01-03", "A"), "indisponivel");
  assert.equal(app.tipoDiaCalendario("2000-01-03", "A"), "indisponivel");
  assert.match(app.diasNaoUteis[1], /^\d{2}\/\d{2}\/\d{4}$/);
});

test("contagem considera apenas audiências agendadas no polo e na data", () => {
  const app = carregarScript();
  app.audienciasCache = [
    { data: "05/01/2099", polo: "A", hora: "15:00/16:00", status: "PENDENTE" },
    { data: "05/01/2099", polo: "A", hora: "13:00/14:00", status: "CONFIRMADO" },
    { data: "05/01/2099", polo: "A", hora: "14:00/15:00", status: "CANCELADO, motivo" },
    { data: "05/01/2099", polo: "A", hora: "14:00/15:00", status: "NÃO REALIZADO, motivo" },
    { data: "05/01/2099", polo: "B", hora: "14:00/15:00", status: "CONFIRMADO" },
    { data: "06/01/2099", polo: "A", hora: "14:00/15:00", status: "CONFIRMADO" },
  ];
  const audiencias = app.audienciasAgendadas("05/01/2099", "A");
  assert.equal(audiencias.length, 2);
  assert.equal(audiencias[0].hora, "13:00/14:00");
});
