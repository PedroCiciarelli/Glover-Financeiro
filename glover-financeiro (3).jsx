import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import {
  ResponsiveContainer, LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
} from "recharts";
import * as XLSX from "xlsx";
import {
  LayoutGrid, Receipt, Wallet, Building2, Users, Plus, Search, Pencil, Trash2,
  Copy, X, ArrowLeftRight, Download, ChevronRight, ArrowUpRight, ArrowDownRight,
  TrendingUp, TrendingDown, Filter, ChevronLeft, Check, Scale, BarChart3,
  CalendarClock, Layers, Undo2, AlertTriangle, CircleDollarSign, ShieldCheck, Lock, LogOut,
} from "lucide-react";

const GLOVER_LOGO_DATA_URI = "data:image/png;base64,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";

/* ---------------------------------------------------------------
   Constantes e utilitários
--------------------------------------------------------------- */

const DEFAULT_APROPRIACOES = [
  "Material", "Mão de obra", "Equipamentos", "Subempreitada", "Serviços",
  "Combustível", "Administrativo", "Impostos", "Aluguel", "Energia elétrica", "Outros",
];

const OBRA_STATUS = ["Em planejamento", "Em execução", "Concluída", "Cancelada"];
const SEM_OBRA = "__sem_obra__";
const STORAGE_KEY = "glover-financeiro-v1";

const DEFAULT_SOCIOS = [
  { id: "pedro", nome: "Pedro", participacao: 60 },
  { id: "fabio", nome: "Fabio", participacao: 40 },
];

/* ---------------------------------------------------------------
   Controle de acesso e permissões
--------------------------------------------------------------- */

const PERMISSION_GROUPS = [
  { group: "Lançamentos", keys: [
    { key: "lancamentos_criar", label: "Criar lançamentos" },
    { key: "lancamentos_editar", label: "Editar lançamentos" },
    { key: "lancamentos_excluir", label: "Excluir lançamentos" },
  ] },
  { group: "Contas a Pagar", keys: [
    { key: "contaspagar_criar", label: "Criar contas a pagar" },
    { key: "contaspagar_editar", label: "Editar contas a pagar" },
    { key: "contaspagar_ratear", label: "Ratear entre centros de custo" },
    { key: "contaspagar_baixar", label: "Registrar pagamentos (baixa)" },
    { key: "contaspagar_excluir", label: "Excluir contas a pagar" },
  ] },
  { group: "Contas Bancárias", keys: [
    { key: "contasbancarias_visualizar", label: "Visualizar saldos das contas" },
    { key: "contasbancarias_cadastrar", label: "Cadastrar/editar contas bancárias" },
    { key: "contasbancarias_saldoinicial", label: "Alterar saldo inicial" },
    { key: "contasbancarias_transferir", label: "Transferir entre contas" },
  ] },
  { group: "Cadastros", keys: [
    { key: "fornecedores_cadastrar", label: "Cadastrar fornecedores" },
    { key: "centroscusto_gerenciar", label: "Cadastrar/alterar centros de custo (obras)" },
    { key: "apropriacoes_gerenciar", label: "Criar/alterar apropriações" },
  ] },
  { group: "Sócios", keys: [
    { key: "socios_acessar", label: "Acessar a página de Sócios" },
  ] },
  { group: "Administração", keys: [
    { key: "relatorios_acessar", label: "Acessar relatórios e dashboards" },
    { key: "config_alterar", label: "Alterar configurações do sistema" },
    { key: "usuarios_gerenciar", label: "Gerenciar usuários e permissões" },
  ] },
];
const ALL_PERM_KEYS = PERMISSION_GROUPS.flatMap((g) => g.keys.map((k) => k.key));
const permLabel = (key) => PERMISSION_GROUPS.flatMap((g) => g.keys).find((k) => k.key === key)?.label || key;

const permsAllTrue = () => Object.fromEntries(ALL_PERM_KEYS.map((k) => [k, true]));
const permsAllFalse = () => Object.fromEntries(ALL_PERM_KEYS.map((k) => [k, false]));

const ASSISTENTE_PERMS = {
  ...permsAllFalse(),
  lancamentos_criar: true, lancamentos_editar: true,
  contaspagar_criar: true, contaspagar_editar: true, contaspagar_ratear: true, contaspagar_baixar: true,
  contasbancarias_visualizar: true, contasbancarias_transferir: true,
  fornecedores_cadastrar: true, apropriacoes_gerenciar: true,
  relatorios_acessar: true,
};

const DEFAULT_USUARIOS = [
  { id: "pedro", nome: "Pedro", papel: "admin", ativo: true, perms: permsAllTrue() },
  { id: "fabio", nome: "Fabio", papel: "admin", ativo: true, perms: permsAllTrue() },
  { id: "assistente", nome: "Assistente Administrativo", papel: "operacional", ativo: true, perms: ASSISTENTE_PERMS },
  { id: "perfil4", nome: "4º usuário (a definir)", papel: "customizavel", ativo: true, perms: permsAllFalse() },
];

const PAPEL_LABEL = { admin: "Administrador", operacional: "Operacional", customizavel: "Personalizado" };

function can(user, key) {
  if (!user || !user.ativo) return false;
  return !!(user.perms && user.perms[key]);
}

function nowStamp() {
  const d = new Date();
  return `${fmtDate(d.toISOString().slice(0, 10))} ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

const uid = (p) => `${p}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
const todayISO = () => new Date().toISOString().slice(0, 10);

const BRL = (v) =>
  (Number(v) || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

const fmtDate = (iso) => {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
};

const monthKey = (iso) => (iso ? iso.slice(0, 7) : "");
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  const names = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  return `${names[Number(m) - 1]}/${y.slice(2)}`;
};

const PERIODICIDADES_PARCELA = [
  { id: "mensal", label: "Mensal" },
  { id: "quinzenal", label: "Quinzenal" },
  { id: "semanal", label: "Semanal" },
  { id: "personalizada", label: "Personalizada" },
];
const PERIODICIDADES_RECORRENTE = [
  { id: "semanal", label: "Semanal" },
  { id: "quinzenal", label: "Quinzenal" },
  { id: "mensal", label: "Mensal" },
  { id: "bimestral", label: "Bimestral" },
  { id: "trimestral", label: "Trimestral" },
  { id: "semestral", label: "Semestral" },
  { id: "anual", label: "Anual" },
];

function addPeriodo(dataISO, periodicidade, customDias) {
  const d = new Date(dataISO + "T00:00:00");
  switch (periodicidade) {
    case "semanal": d.setDate(d.getDate() + 7); break;
    case "quinzenal": d.setDate(d.getDate() + 15); break;
    case "mensal": d.setMonth(d.getMonth() + 1); break;
    case "bimestral": d.setMonth(d.getMonth() + 2); break;
    case "trimestral": d.setMonth(d.getMonth() + 3); break;
    case "semestral": d.setMonth(d.getMonth() + 6); break;
    case "anual": d.setFullYear(d.getFullYear() + 1); break;
    case "personalizada": d.setDate(d.getDate() + (Number(customDias) || 30)); break;
    default: d.setMonth(d.getMonth() + 1);
  }
  return d.toISOString().slice(0, 10);
}

function deriveStatusConta(c) {
  if (c.cancelada) return "Cancelada";
  if (c.pagamento) return "Paga";
  const today = todayISO();
  if (c.dataVencimento < today) return "Vencida";
  if (c.dataVencimento === today) return "Vencendo hoje";
  return "A vencer";
}

// Distribui um valor de referência entre centros de custo conforme percentuais,
// garantindo que a soma feche exatamente (o último item absorve o arredondamento).
function computeRateioValores(percentuais, valorRef) {
  const n = percentuais.length;
  let acumulado = 0;
  return percentuais.map((p, i) => {
    let valor;
    if (i === n - 1) valor = Math.round((valorRef - acumulado) * 100) / 100;
    else { valor = Math.round(((Number(p.percentual) || 0) / 100) * valorRef * 100) / 100; acumulado += valor; }
    return { centroCusto: p.centroCusto, percentual: Number(p.percentual) || 0, valor };
  });
}

// Rateia proporcionalmente um acréscimo/desconto total entre os itens de rateio já existentes.
function rateiaAcrescimo(rateio, acrescimoTotal) {
  const n = rateio.length;
  let acumulado = 0;
  return rateio.map((r, i) => {
    let acrescimo;
    if (i === n - 1) acrescimo = Math.round((acrescimoTotal - acumulado) * 100) / 100;
    else { acrescimo = Math.round(((r.percentual || 0) / 100) * acrescimoTotal * 100) / 100; acumulado += acrescimo; }
    return { ...r, acrescimo, valorFinal: Math.round((r.valor + acrescimo) * 100) / 100 };
  });
}

const STATUS_TONE = {
  "A vencer": "blue", "Vencendo hoje": "orange", "Vencida": "red",
  "Paga": "green", "Cancelada": "gray", "Parcialmente paga": "orange",
};

/* ---------------------------------------------------------------
   Controle orçamentário (Painel de Obras)
--------------------------------------------------------------- */

const ORCAMENTO_STATUS_TONE = { verde: "green", amarelo: "orange", vermelho: "red" };
function orcamentoStatusTone(pct) {
  if (pct > 100) return "vermelho";
  if (pct >= 80) return "amarelo";
  return "verde";
}

// Calcula, para uma obra, o orçado/comprometido/realizado por apropriação.
// Comprometido = Realizado (já pago) + Pendente (contas a pagar em aberto, não canceladas).
// Realizado vem dos lançamentos reais (entries), o que cobre tanto baixas de contas a pagar
// quanto lançamentos diretos feitos na página Lançamentos.
function computeOrcamentoObra(obraId, orcamentoItens, contasPagar, entries) {
  const orcadoPorApropriacao = {};
  orcamentoItens.filter((o) => o.obraId === obraId).forEach((o) => {
    orcadoPorApropriacao[o.apropriacao] = (orcadoPorApropriacao[o.apropriacao] || 0) + (Number(o.valorTotal) || 0);
  });

  const realizadoPorApropriacao = {};
  entries.filter((e) => e.tipo === "saida" && e.obraId === obraId).forEach((e) => {
    realizadoPorApropriacao[e.apropriacao] = (realizadoPorApropriacao[e.apropriacao] || 0) + (Number(e.valor) || 0);
  });

  const pendentePorApropriacao = {};
  contasPagar.filter((c) => !c.pagamento && !c.cancelada).forEach((c) => {
    (c.rateio || []).filter((r) => r.centroCusto === obraId).forEach((r) => {
      pendentePorApropriacao[c.apropriacao] = (pendentePorApropriacao[c.apropriacao] || 0) + (Number(r.valor) || 0);
    });
  });

  const apropriacoesEnvolvidas = new Set([
    ...Object.keys(orcadoPorApropriacao), ...Object.keys(realizadoPorApropriacao), ...Object.keys(pendentePorApropriacao),
  ]);

  const linhas = [...apropriacoesEnvolvidas].map((apropriacao) => {
    const orcado = orcadoPorApropriacao[apropriacao] || 0;
    const realizado = realizadoPorApropriacao[apropriacao] || 0;
    const pendente = pendentePorApropriacao[apropriacao] || 0;
    const comprometido = realizado + pendente;
    const saldo = orcado - comprometido;
    const pct = orcado > 0 ? (comprometido / orcado) * 100 : (comprometido > 0 ? 999 : 0);
    return { apropriacao, orcado, realizado, pendente, comprometido, saldo, pct, status: orcamentoStatusTone(pct) };
  }).sort((a, b) => b.orcado - a.orcado || b.comprometido - a.comprometido);

  const totais = linhas.reduce((acc, l) => ({
    orcado: acc.orcado + l.orcado, realizado: acc.realizado + l.realizado,
    pendente: acc.pendente + l.pendente, comprometido: acc.comprometido + l.comprometido,
  }), { orcado: 0, realizado: 0, pendente: 0, comprometido: 0 });
  totais.saldo = totais.orcado - totais.comprometido;
  totais.pct = totais.orcado > 0 ? (totais.comprometido / totais.orcado) * 100 : 0;

  return { linhas, totais, temOrcamento: Object.keys(orcadoPorApropriacao).length > 0 };
}

// Verifica, para um centro de custo + apropriação específicos, o que aconteceria
// se um novo valor fosse comprometido (usado no alerta ao lançar uma conta a pagar).
function checarAlertaOrcamento(obraId, apropriacao, valorNovo, orcamentoItens, contasPagar, entries) {
  if (!obraId || obraId === SEM_OBRA) return null;
  const { linhas } = computeOrcamentoObra(obraId, orcamentoItens, contasPagar, entries);
  const linha = linhas.find((l) => l.apropriacao === apropriacao);
  const orcado = linha?.orcado || 0;
  if (orcado <= 0) return null; // sem orçamento importado para esta apropriação
  const comprometidoAtual = linha?.comprometido || 0;
  const novoComprometido = comprometidoAtual + valorNovo;
  const pctNovo = (novoComprometido / orcado) * 100;
  const estouro = novoComprometido - orcado;
  return { orcado, comprometidoAtual, novoComprometido, pctNovo, estouro, ultrapassa: estouro > 0, atencao: pctNovo >= 80 };
}

/* ---------------------------------------------------------------
   Dívidas e Investimentos
--------------------------------------------------------------- */

const TIPOS_DIVIDA = ["Empréstimo", "Financiamento", "Cheque especial", "Cartão de crédito", "Outro"];
const SISTEMAS_AMORTIZACAO = ["SAC", "PRICE", "Não se aplica", "Outro"];
const TIPOS_INVESTIMENTO = ["CDB", "Fundo de investimento", "Poupança", "Tesouro Direto", "Ações", "Outro"];
const LIQUIDEZ_OPCOES = ["Diária", "No vencimento", "Outra"];

function computeDividaStats(divida, entries) {
  const pagamentos = entries.filter((e) => e.dividaId === divida.id && e.dividaSubtipo === "pagamento");
  const totalAmortizado = pagamentos.reduce((s, e) => s + (Number(e.amortizacao) || 0), 0);
  const jurosPagos = pagamentos.reduce((s, e) => s + (Number(e.juros) || 0), 0);
  const encargosPagos = pagamentos.reduce((s, e) => s + (Number(e.encargos) || 0), 0);
  const parcelasPagas = pagamentos.length;
  const saldoDevedor = Math.max(0, (Number(divida.valorContratado) || 0) - totalAmortizado);
  const parcelasRestantes = Math.max(0, (Number(divida.qtdParcelas) || 0) - parcelasPagas);
  const totalDesembolsado = totalAmortizado + jurosPagos + encargosPagos;
  return { totalAmortizado, jurosPagos, encargosPagos, parcelasPagas, parcelasRestantes, saldoDevedor, totalDesembolsado, pagamentos };
}

function computeInvestimentoStats(investimento, entries) {
  const relacionados = entries.filter((e) => e.investimentoId === investimento.id);
  const aportes = relacionados.filter((e) => e.investimentoSubtipo === "aporte").reduce((s, e) => s + (Number(e.valor) || 0), 0);
  const resgates = relacionados.filter((e) => e.investimentoSubtipo === "resgate").reduce((s, e) => s + (Number(e.valor) || 0), 0);
  const rendimentos = relacionados.filter((e) => e.investimentoSubtipo === "rendimento").reduce((s, e) => s + (Number(e.valor) || 0), 0);
  const principalLiquido = (Number(investimento.valorAplicado) || 0) + aportes - resgates;
  const valorAtual = principalLiquido + rendimentos;
  const rentabilidade = principalLiquido > 0 ? (rendimentos / principalLiquido) * 100 : 0;
  return { aportes, resgates, rendimentos, principalLiquido, valorAtual, rentabilidade, relacionados };
}

function weekRange() {
  const now = new Date();
  const day = now.getDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;
  const monday = new Date(now); monday.setDate(now.getDate() + diffToMonday); monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);
  return { monday: monday.toISOString().slice(0, 10), sunday: sunday.toISOString().slice(0, 10) };
}

function addDaysISO(iso, days) {
  const d = new Date(iso + "T00:00:00");
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function seedData() {
  const acc1 = uid("acc"), acc2 = uid("acc"), acc3 = uid("acc");
  const obra1 = uid("obra"), obra2 = uid("obra"), obra3 = uid("obra");
  const accounts = [
    { id: acc1, banco: "Itaú", nome: "Itaú - Conta Principal", agencia: "0921", numero: "18342-5", saldoInicial: 45000, dataSaldoInicial: "2026-01-01", status: "Ativa" },
    { id: acc2, banco: "Bradesco", nome: "Bradesco - Obras", agencia: "3310", numero: "77120-1", saldoInicial: 12000, dataSaldoInicial: "2026-01-01", status: "Ativa" },
    { id: acc3, banco: "Caixa Econômica", nome: "Caixa - Reserva", agencia: "1223", numero: "00452-9", saldoInicial: 8000, dataSaldoInicial: "2026-01-01", status: "Ativa" },
  ];
  const obras = [
    { id: obra1, nome: "Residencial Bela Vista", codigo: "OB-001", cliente: "Bela Vista Empreend.", localizacao: "Ribeirão Preto/SP", dataInicio: "2026-02-01", dataFim: "2026-12-20", orcamento: 850000, status: "Em execução" },
    { id: obra2, nome: "Galpão Industrial XPTO", codigo: "OB-002", cliente: "XPTO Logística", localizacao: "Sertãozinho/SP", dataInicio: "2026-03-10", dataFim: "2026-09-30", orcamento: 420000, status: "Em execução" },
    { id: obra3, nome: "Reforma Sede Administrativa", codigo: "OB-003", cliente: "Glover Engenharia", localizacao: "Ribeirão Preto/SP", dataInicio: "2026-01-15", dataFim: "2026-04-30", orcamento: 60000, status: "Concluída" },
  ];
  const fornecedores = ["Cimento Forte Materiais", "Elétrica Silva & Cia", "Locadora JR Equipamentos", "Madeireira Santa Fé", "Posto Rota Sul"];
  const rows = [
    [obra1, "saida", acc1, "Cimento Forte Materiais", "Material", -1, 18500, "Compra de cimento e agregados"],
    [obra1, "saida", acc2, "Elétrica Silva & Cia", "Mão de obra", -1, 9200, "Instalação elétrica - 1a etapa"],
    [obra1, "entrada", acc1, "Bela Vista Empreend.", "Não se aplica", -2, 120000, "Medição 1 - recebimento cliente"],
    [obra2, "saida", acc2, "Locadora JR Equipamentos", "Equipamentos", -3, 14300, "Locação de betoneira e andaimes"],
    [obra2, "saida", acc1, "Posto Rota Sul", "Combustível", -5, 2100, "Abastecimento de veículos da obra"],
    [obra2, "entrada", acc2, "XPTO Logística", "Não se aplica", -6, 80000, "Adiantamento contratual"],
    [SEM_OBRA, "saida", acc1, "", "Administrativo", -8, 4200, "Despesas administrativas do mês"],
    [obra1, "saida", acc1, "Madeireira Santa Fé", "Material", -10, 6700, "Compra de madeira para forma"],
    [obra3, "saida", acc3, "Elétrica Silva & Cia", "Serviços", -35, 3100, "Manutenção elétrica sede"],
    [obra3, "entrada", acc3, "Glover Engenharia", "Não se aplica", -33, 60000, "Aporte para reforma da sede"],
    [obra1, "saida", acc2, "Cimento Forte Materiais", "Impostos", -15, 3400, "Guia de impostos da obra"],
    [SEM_OBRA, "saida", acc3, "", "Aluguel", -20, 3800, "Aluguel do escritório central"],
    [obra1, "saida", acc2, "Cimento Forte Materiais", "Fundação", -55, 55000, "Escavação e fundação - 1a parte (baixa de conta a pagar)"],
    [obra1, "saida", acc2, "Madeireira Santa Fé", "Estrutura", -45, 100000, "Concreto e armação - etapa 1 (baixa de conta a pagar)"],
    [obra1, "saida", acc2, "Cimento Forte Materiais", "Alvenaria", -18, 40000, "Tijolos e assentamento (baixa de conta a pagar)"],
    [obra1, "saida", acc2, "Elétrica Silva & Cia", "Instalações elétricas", -10, 50000, "Instalação elétrica - fase 1 (baixa de conta a pagar)"],
  ];
  const entries = rows.map(([obraId, tipo, contaId, fornecedor, apropriacao, dOffset, valor, descricao]) => {
    const d = new Date();
    d.setDate(d.getDate() + dOffset);
    return {
      id: uid("lan"), tipo, contaId, obraId, data: d.toISOString().slice(0, 10),
      valor, fornecedor, apropriacao, descricao, createdAt: Date.now(),
    };
  });

  const socioRows = [
    ["pedro", "aporte", acc1, -60, 90000, "Aporte inicial de capital"],
    ["fabio", "aporte", acc1, -60, 60000, "Aporte inicial de capital"],
    ["pedro", "pro-labore", acc1, -30, 10000, "Pró-labore do mês"],
    ["fabio", "pro-labore", acc1, -30, 10000, "Pró-labore do mês"],
    ["pedro", "retirada", acc2, -7, 9000, "Retirada de lucros"],
    ["fabio", "retirada", acc2, -7, 2000, "Retirada de lucros"],
  ];
  const movimentosSocios = socioRows.map(([socioId, tipo, contaId, dOffset, valor, descricao]) => {
    const d = new Date();
    d.setDate(d.getDate() + dOffset);
    return { id: uid("soc"), socioId, tipo, contaId, data: d.toISOString().slice(0, 10), valor, descricao, createdAt: Date.now() };
  });

  const apropriacoesComOrcamento = [...DEFAULT_APROPRIACOES, "Fundação", "Estrutura", "Alvenaria", "Instalações elétricas", "Instalações hidráulicas"];

  const orcamentoItens = [
    { id: uid("orc"), obraId: obra1, apropriacao: "Fundação", descricao: "Fundação e terraplenagem", unidade: "", quantidade: "", valorUnitario: "", valorTotal: 80000, importadoEm: Date.now() },
    { id: uid("orc"), obraId: obra1, apropriacao: "Estrutura", descricao: "Estrutura de concreto armado", unidade: "", quantidade: "", valorUnitario: "", valorTotal: 150000, importadoEm: Date.now() },
    { id: uid("orc"), obraId: obra1, apropriacao: "Alvenaria", descricao: "Alvenaria de vedação", unidade: "", quantidade: "", valorUnitario: "", valorTotal: 60000, importadoEm: Date.now() },
    { id: uid("orc"), obraId: obra1, apropriacao: "Instalações elétricas", descricao: "Instalações elétricas completas", unidade: "", quantidade: "", valorUnitario: "", valorTotal: 90000, importadoEm: Date.now() },
    { id: uid("orc"), obraId: obra1, apropriacao: "Instalações hidráulicas", descricao: "Instalações hidráulicas completas", unidade: "", quantidade: "", valorUnitario: "", valorTotal: 40000, importadoEm: Date.now() },
    { id: uid("orc"), obraId: obra1, apropriacao: "Pintura", descricao: "Pintura interna e externa", unidade: "", quantidade: "", valorUnitario: "", valorTotal: 35000, importadoEm: Date.now() },
  ];

  const divida1 = uid("div");
  const investimento1 = uid("inv");

  const dividaEntries = [
    { id: uid("lan"), tipo: "entrada", contaId: acc1, obraId: obra1, data: addDaysISO(todayISO(), -120), valor: 500000, fornecedor: "Banco X", apropriacao: "Não se aplica", descricao: "Contratação de financiamento - Banco X", dividaId: divida1, dividaSubtipo: "contratacao", createdAt: Date.now() },
    { id: uid("lan"), tipo: "saida", contaId: acc2, obraId: obra1, data: addDaysISO(todayISO(), -90), valor: 16400, fornecedor: "Banco X", apropriacao: "Impostos", descricao: "Parcela 1/48 - Financiamento Banco X", dividaId: divida1, dividaSubtipo: "pagamento", amortizacao: 10000, juros: 6200, encargos: 200, createdAt: Date.now() },
    { id: uid("lan"), tipo: "saida", contaId: acc2, obraId: obra1, data: addDaysISO(todayISO(), -60), valor: 16400, fornecedor: "Banco X", apropriacao: "Impostos", descricao: "Parcela 2/48 - Financiamento Banco X", dividaId: divida1, dividaSubtipo: "pagamento", amortizacao: 10100, juros: 6100, encargos: 200, createdAt: Date.now() },
    { id: uid("lan"), tipo: "saida", contaId: acc2, obraId: obra1, data: addDaysISO(todayISO(), -30), valor: 16400, fornecedor: "Banco X", apropriacao: "Impostos", descricao: "Parcela 3/48 - Financiamento Banco X", dividaId: divida1, dividaSubtipo: "pagamento", amortizacao: 10200, juros: 6000, encargos: 200, createdAt: Date.now() },
  ];

  const investimentoEntries = [
    { id: uid("lan"), tipo: "saida", contaId: acc3, obraId: SEM_OBRA, data: addDaysISO(todayISO(), -90), valor: 100000, fornecedor: "Banco X", apropriacao: "Não se aplica", descricao: "Aplicação financeira - CDB Banco X", investimentoId: investimento1, investimentoSubtipo: "aplicacao", createdAt: Date.now() },
    { id: uid("lan"), tipo: "saida", contaId: acc3, obraId: SEM_OBRA, data: addDaysISO(todayISO(), -40), valor: 20000, fornecedor: "Banco X", apropriacao: "Não se aplica", descricao: "Aporte adicional - CDB Banco X", investimentoId: investimento1, investimentoSubtipo: "aporte", createdAt: Date.now() },
    { id: uid("lan"), tipo: "entrada", contaId: acc3, obraId: SEM_OBRA, data: addDaysISO(todayISO(), -20), valor: 10000, fornecedor: "Banco X", apropriacao: "Não se aplica", descricao: "Resgate parcial - CDB Banco X", investimentoId: investimento1, investimentoSubtipo: "resgate", createdAt: Date.now() },
    { id: uid("lan"), tipo: "entrada", contaId: acc3, obraId: SEM_OBRA, data: addDaysISO(todayISO(), -5), valor: 7500, fornecedor: "Banco X", apropriacao: "Não se aplica", descricao: "Rendimento - CDB Banco X", investimentoId: investimento1, investimentoSubtipo: "rendimento", createdAt: Date.now() },
  ];

  const dividas = [{
    id: divida1, instituicao: "Banco X", tipoOperacao: "Financiamento", dataContratacao: addDaysISO(todayISO(), -120),
    valorContratado: 500000, taxaJuros: "1,2% a.m.", sistemaAmortizacao: "SAC", qtdParcelas: 48,
    valorParcelaEstimado: 16400, proximoVencimento: addDaysISO(todayISO(), 5), obraId: obra1, status: "Ativa", createdAt: Date.now(),
  }];

  const investimentos = [{
    id: investimento1, instituicao: "Banco X", tipoInvestimento: "CDB", dataAplicacao: addDaysISO(todayISO(), -90),
    valorAplicado: 100000, dataVencimento: addDaysISO(todayISO(), 275), liquidez: "No vencimento", obraId: SEM_OBRA, status: "Ativo", createdAt: Date.now(),
  }];

  return {
    accounts, obras, apropriacoes: apropriacoesComOrcamento, entries: [...entries, ...dividaEntries, ...investimentoEntries], transfers: [],
    fornecedoresSeed: fornecedores, socios: DEFAULT_SOCIOS, movimentosSocios,
    contasPagar: seedContasPagar({ obra1, obra2, fornecedores }),
    orcamentoItens, dividas, investimentos,
  };
}

function seedContasPagar({ obra1, obra2, fornecedores }) {
  const list = [];
  const push = (c) => {
    const rateio = c.rateio || [{ centroCusto: c.centroCusto, percentual: 100, valor: c.valorPrevisto }];
    list.push({
      id: uid("cp"), groupId: c.groupId || uid("grp"), seq: c.seq || 1, total: c.total || 1,
      tipo: c.tipo || "unica", centroCusto: rateio.length === 1 ? rateio[0].centroCusto : null,
      rateio, fornecedor: c.fornecedor,
      apropriacao: c.apropriacao, descricao: c.descricao, numeroDocumento: c.numeroDocumento || "",
      observacao: c.observacao || "", dataCompetencia: c.dataCompetencia, dataVencimento: c.dataVencimento,
      valorPrevisto: c.valorPrevisto, pagamento: c.pagamento || null, cancelada: false, createdAt: Date.now(),
    });
  };

  // conta única já vencida (para demonstrar status "Vencida")
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[0], apropriacao: "Material",
    descricao: "Compra de tijolos e blocos", dataCompetencia: addDaysISO(todayISO(), -15),
    dataVencimento: addDaysISO(todayISO(), -3), valorPrevisto: 5400,
  });

  // conta vencendo hoje
  push({
    tipo: "unica", centroCusto: obra2, fornecedor: fornecedores[2], apropriacao: "Equipamentos",
    descricao: "Locação mensal de andaimes", dataCompetencia: addDaysISO(todayISO(), -10),
    dataVencimento: todayISO(), valorPrevisto: 3200,
  });

  // conta paga
  push({
    tipo: "unica", centroCusto: SEM_OBRA, fornecedor: fornecedores[1], apropriacao: "Serviços",
    descricao: "Manutenção elétrica", dataCompetencia: addDaysISO(todayISO(), -25),
    dataVencimento: addDaysISO(todayISO(), -12), valorPrevisto: 1800,
    pagamento: { data: addDaysISO(todayISO(), -12), contaId: null, valorOriginal: 1800, multa: 0, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 1800 },
  });

  // conta rateada entre duas obras (80% / 20%) - aberta, para demonstrar o rateio
  push({
    tipo: "unica", fornecedor: fornecedores[3], apropriacao: "Material",
    descricao: "Fornecimento de madeira - rateado entre obras", dataCompetencia: addDaysISO(todayISO(), -2),
    dataVencimento: addDaysISO(todayISO(), 10), valorPrevisto: 1000,
    rateio: [
      { centroCusto: obra1, percentual: 80, valor: 800 },
      { centroCusto: obra2, percentual: 20, valor: 200 },
    ],
  });

  // conta rateada já paga com acréscimo (multa), mostrando o rateio proporcional do acréscimo
  push({
    tipo: "unica", fornecedor: fornecedores[3], apropriacao: "Material",
    descricao: "Boleto de material - rateado entre obras (pago com multa)", dataCompetencia: addDaysISO(todayISO(), -18),
    dataVencimento: addDaysISO(todayISO(), -8), valorPrevisto: 1000,
    rateio: [
      { centroCusto: obra1, percentual: 80, valor: 800 },
      { centroCusto: obra2, percentual: 20, valor: 200 },
    ],
    pagamento: {
      data: addDaysISO(todayISO(), -5), contaId: null, valorOriginal: 1000, multa: 50, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 1050,
      rateioPago: [
        { centroCusto: obra1, percentual: 80, valorOriginal: 800, acrescimo: 40, valorFinal: 840 },
        { centroCusto: obra2, percentual: 20, valorOriginal: 200, acrescimo: 10, valorFinal: 210 },
      ],
    },
  });

  // conta parcelada (locação de equipamento) - 4 parcelas mensais
  const grp = uid("grp");
  let vencParc = addDaysISO(todayISO(), 5);
  for (let i = 1; i <= 4; i++) {
    push({
      groupId: grp, seq: i, total: 4, tipo: "parcelada", centroCusto: obra2, fornecedor: fornecedores[2],
      apropriacao: "Equipamentos", descricao: "Locação de betoneira - contrato trimestral",
      dataCompetencia: vencParc, dataVencimento: vencParc, valorPrevisto: 4500,
    });
    vencParc = addPeriodo(vencParc, "mensal");
  }

  // contas ligadas ao orçamento da obra1, para demonstrar o Controle Orçamentário
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[0], apropriacao: "Fundação",
    descricao: "Escavação e fundação - 1a parte", dataCompetencia: addDaysISO(todayISO(), -60),
    dataVencimento: addDaysISO(todayISO(), -55), valorPrevisto: 55000,
    pagamento: { data: addDaysISO(todayISO(), -55), contaId: null, valorOriginal: 55000, multa: 0, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 55000, rateioPago: [{ centroCusto: obra1, percentual: 100, valorOriginal: 55000, acrescimo: 0, valorFinal: 55000 }] },
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[0], apropriacao: "Fundação",
    descricao: "Fundação - complemento", dataCompetencia: addDaysISO(todayISO(), -5),
    dataVencimento: addDaysISO(todayISO(), 12), valorPrevisto: 10000,
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[3], apropriacao: "Estrutura",
    descricao: "Concreto e armação - etapa 1", dataCompetencia: addDaysISO(todayISO(), -50),
    dataVencimento: addDaysISO(todayISO(), -45), valorPrevisto: 100000,
    pagamento: { data: addDaysISO(todayISO(), -45), contaId: null, valorOriginal: 100000, multa: 0, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 100000, rateioPago: [{ centroCusto: obra1, percentual: 100, valorOriginal: 100000, acrescimo: 0, valorFinal: 100000 }] },
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[3], apropriacao: "Estrutura",
    descricao: "Concreto e armação - etapa 2", dataCompetencia: addDaysISO(todayISO(), -8),
    dataVencimento: addDaysISO(todayISO(), 20), valorPrevisto: 45000,
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[0], apropriacao: "Alvenaria",
    descricao: "Tijolos e assentamento", dataCompetencia: addDaysISO(todayISO(), -22),
    dataVencimento: addDaysISO(todayISO(), -18), valorPrevisto: 40000,
    pagamento: { data: addDaysISO(todayISO(), -18), contaId: null, valorOriginal: 40000, multa: 0, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 40000, rateioPago: [{ centroCusto: obra1, percentual: 100, valorOriginal: 40000, acrescimo: 0, valorFinal: 40000 }] },
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[0], apropriacao: "Alvenaria",
    descricao: "Tijolos - complemento (aditivo)", dataCompetencia: addDaysISO(todayISO(), -3),
    dataVencimento: addDaysISO(todayISO(), 15), valorPrevisto: 30000,
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[1], apropriacao: "Instalações elétricas",
    descricao: "Instalação elétrica - fase 1", dataCompetencia: addDaysISO(todayISO(), -15),
    dataVencimento: addDaysISO(todayISO(), -10), valorPrevisto: 50000,
    pagamento: { data: addDaysISO(todayISO(), -10), contaId: null, valorOriginal: 50000, multa: 0, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 50000, rateioPago: [{ centroCusto: obra1, percentual: 100, valorOriginal: 50000, acrescimo: 0, valorFinal: 50000 }] },
  });
  push({
    tipo: "unica", centroCusto: obra1, fornecedor: fornecedores[1], apropriacao: "Instalações elétricas",
    descricao: "Instalação elétrica - fase 2", dataCompetencia: addDaysISO(todayISO(), -2),
    dataVencimento: addDaysISO(todayISO(), 25), valorPrevisto: 18000,
  });

  // conta recorrente (aluguel do escritório) - 6 ocorrências, 1 já paga
  const grp2 = uid("grp");
  let venc2 = addDaysISO(todayISO(), -20);
  for (let i = 1; i <= 6; i++) {
    push({
      groupId: grp2, seq: i, total: 6, tipo: "recorrente", centroCusto: SEM_OBRA, fornecedor: "",
      apropriacao: "Aluguel", descricao: "Aluguel do escritório central",
      dataCompetencia: venc2, dataVencimento: venc2, valorPrevisto: 3800,
      pagamento: i === 1 ? { data: venc2, contaId: null, valorOriginal: 3800, multa: 0, juros: 0, outrosAcrescimos: 0, descontos: 0, valorFinal: 3800 } : null,
    });
    venc2 = addPeriodo(venc2, "mensal");
  }

  return list;
}

/* ---------------------------------------------------------------
   Componentes utilitários de UI
--------------------------------------------------------------- */

function Modal({ title, onClose, children, width = 520 }) {
  return (
    <div className="gf-modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="gf-modal" style={{ maxWidth: width }}>
        <div className="gf-modal-head">
          <h3>{title}</h3>
          <button className="gf-icon-btn" onClick={onClose} aria-label="Fechar"><X size={18} /></button>
        </div>
        <div className="gf-modal-body">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children, hint, className = "" }) {
  return (
    <label className={`gf-field ${className}`}>
      <span className="gf-field-label">{label}</span>
      {children}
      {hint && <span className="gf-field-hint">{hint}</span>}
    </label>
  );
}

function Badge({ tone = "neutral", children }) {
  return <span className={`gf-badge gf-badge-${tone}`}>{children}</span>;
}

function EmptyState({ icon: Icon, title, body }) {
  return (
    <div className="gf-empty">
      <Icon size={28} strokeWidth={1.4} />
      <strong>{title}</strong>
      <p>{body}</p>
    </div>
  );
}

/* ---------------------------------------------------------------
   App principal
--------------------------------------------------------------- */

export default function App() {
  const [ready, setReady] = useState(false);
  const [page, setPage] = useState("lancamentos");
  const [accounts, setAccounts] = useState([]);
  const [obras, setObras] = useState([]);
  const [apropriacoes, setApropriacoes] = useState(DEFAULT_APROPRIACOES);
  const [entries, setEntries] = useState([]);
  const [transfers, setTransfers] = useState([]);
  const [fornecedoresSeed, setFornecedoresSeed] = useState([]);
  const [socios, setSocios] = useState(DEFAULT_SOCIOS);
  const [movimentosSocios, setMovimentosSocios] = useState([]);
  const [contasPagar, setContasPagar] = useState([]);
  const [orcamentoItens, setOrcamentoItens] = useState([]);
  const [dividas, setDividas] = useState([]);
  const [investimentos, setInvestimentos] = useState([]);
  const [usuarios, setUsuarios] = useState(DEFAULT_USUARIOS);
  const [currentUserId, setCurrentUserId] = useState(null);
  const [auditLog, setAuditLog] = useState([]);
  const [savedTick, setSavedTick] = useState(null);
  const firstLoad = useRef(true);

  // Carregar dados salvos
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          const d = JSON.parse(res.value);
          setAccounts(d.accounts || []);
          setObras(d.obras || []);
          setApropriacoes(d.apropriacoes && d.apropriacoes.length ? d.apropriacoes : DEFAULT_APROPRIACOES);
          setEntries(d.entries || []);
          setTransfers(d.transfers || []);
          setFornecedoresSeed(d.fornecedoresSeed || []);
          setSocios(d.socios && d.socios.length ? d.socios : DEFAULT_SOCIOS);
          setMovimentosSocios(d.movimentosSocios || []);
          setContasPagar(d.contasPagar || []);
          setOrcamentoItens(d.orcamentoItens || []);
          setDividas(d.dividas || []);
          setInvestimentos(d.investimentos || []);
          setUsuarios(d.usuarios && d.usuarios.length ? d.usuarios : DEFAULT_USUARIOS);
          setAuditLog(d.auditLog || []);
          setCurrentUserId(d.currentUserId || null);
        } else {
          const seed = seedData();
          setAccounts(seed.accounts);
          setObras(seed.obras);
          setApropriacoes(seed.apropriacoes);
          setEntries(seed.entries);
          setTransfers(seed.transfers);
          setFornecedoresSeed(seed.fornecedoresSeed);
          setSocios(seed.socios);
          setMovimentosSocios(seed.movimentosSocios);
          setContasPagar(seed.contasPagar || []);
          setOrcamentoItens(seed.orcamentoItens || []);
          setDividas(seed.dividas || []);
          setInvestimentos(seed.investimentos || []);
          setUsuarios(DEFAULT_USUARIOS);
        }
      } catch (e) {
        const seed = seedData();
        setAccounts(seed.accounts);
        setObras(seed.obras);
        setApropriacoes(seed.apropriacoes);
        setEntries(seed.entries);
        setTransfers(seed.transfers);
        setFornecedoresSeed(seed.fornecedoresSeed);
        setSocios(seed.socios);
        setMovimentosSocios(seed.movimentosSocios);
        setContasPagar(seed.contasPagar || []);
        setOrcamentoItens(seed.orcamentoItens || []);
        setDividas(seed.dividas || []);
        setInvestimentos(seed.investimentos || []);
        setUsuarios(DEFAULT_USUARIOS);
      }
      setReady(true);
    })();
  }, []);

  // Salvar dados (debounced)
  useEffect(() => {
    if (!ready) return;
    if (firstLoad.current) { firstLoad.current = false; return; }
    const t = setTimeout(async () => {
      try {
        await window.storage.set(
          STORAGE_KEY,
          JSON.stringify({ accounts, obras, apropriacoes, entries, transfers, fornecedoresSeed, socios, movimentosSocios, contasPagar, orcamentoItens, dividas, investimentos, usuarios, auditLog, currentUserId }),
          false
        );
        setSavedTick(Date.now());
      } catch (e) { /* silencioso */ }
    }, 350);
    return () => clearTimeout(t);
  }, [accounts, obras, apropriacoes, entries, transfers, fornecedoresSeed, socios, movimentosSocios, contasPagar, orcamentoItens, dividas, investimentos, usuarios, auditLog, currentUserId, ready]);

  const currentUser = usuarios.find((u) => u.id === currentUserId) || null;

  const logAudit = (entry) => {
    setAuditLog((prev) => [{
      id: uid("aud"), data: nowStamp(), usuario: currentUser?.nome || "—", ...entry,
    }, ...prev].slice(0, 500));
  };

  const resetDemo = async () => {
    if (!window.confirm("Isso vai apagar todos os dados atuais e recarregar os dados de exemplo. Continuar?")) return;
    const seed = seedData();
    setAccounts(seed.accounts); setObras(seed.obras); setApropriacoes(seed.apropriacoes);
    setEntries(seed.entries); setTransfers(seed.transfers); setFornecedoresSeed(seed.fornecedoresSeed);
    setSocios(seed.socios); setMovimentosSocios(seed.movimentosSocios); setContasPagar(seed.contasPagar || []);
    setOrcamentoItens(seed.orcamentoItens || []);
    setDividas(seed.dividas || []); setInvestimentos(seed.investimentos || []);
  };

  const clearAll = async () => {
    if (!window.confirm("Isso vai apagar TODOS os dados (contas, obras, lançamentos). Continuar?")) return;
    setAccounts([]); setObras([]); setEntries([]); setTransfers([]); setFornecedoresSeed([]);
    setApropriacoes(DEFAULT_APROPRIACOES);
    setMovimentosSocios([]); setSocios(DEFAULT_SOCIOS); setContasPagar([]); setOrcamentoItens([]);
    setDividas([]); setInvestimentos([]);
  };

  /* --------------------- cálculos derivados --------------------- */

  const accountStats = useMemo(() => {
    const map = {};
    accounts.forEach((a) => (map[a.id] = { entradas: 0, saidas: 0 }));
    entries.forEach((e) => {
      if (!map[e.contaId]) return;
      if (e.tipo === "entrada") map[e.contaId].entradas += Number(e.valor) || 0;
      else map[e.contaId].saidas += Number(e.valor) || 0;
    });
    const transfIn = {}, transfOut = {};
    transfers.forEach((t) => {
      transfIn[t.destinoId] = (transfIn[t.destinoId] || 0) + Number(t.valor);
      transfOut[t.origemId] = (transfOut[t.origemId] || 0) + Number(t.valor);
    });
    const socioIn = {}, socioOut = {};
    movimentosSocios.forEach((m) => {
      if (m.tipo === "aporte") socioIn[m.contaId] = (socioIn[m.contaId] || 0) + Number(m.valor);
      else socioOut[m.contaId] = (socioOut[m.contaId] || 0) + Number(m.valor); // retirada e pró-labore saem do caixa
    });
    accounts.forEach((a) => {
      const s = map[a.id];
      s.transfIn = transfIn[a.id] || 0;
      s.transfOut = transfOut[a.id] || 0;
      s.socioIn = socioIn[a.id] || 0;
      s.socioOut = socioOut[a.id] || 0;
      s.saldoAtual = (Number(a.saldoInicial) || 0) + s.entradas - s.saidas + s.transfIn - s.transfOut + s.socioIn - s.socioOut;
    });
    return map;
  }, [accounts, entries, transfers, movimentosSocios]);

  const obraStats = useMemo(() => {
    const map = {};
    obras.forEach((o) => (map[o.id] = { receitas: 0, despesas: 0, count: 0 }));
    entries.forEach((e) => {
      if (!map[e.obraId]) return;
      map[e.obraId].count += 1;
      if (e.tipo === "entrada") map[e.obraId].receitas += Number(e.valor) || 0;
      else map[e.obraId].despesas += Number(e.valor) || 0;
    });
    obras.forEach((o) => {
      const s = map[o.id];
      s.resultado = s.receitas - s.despesas;
      s.margem = s.receitas > 0 ? (s.resultado / s.receitas) * 100 : null;
    });
    return map;
  }, [obras, entries]);

  const obraSemObra = useMemo(() => {
    let receitas = 0, despesas = 0, count = 0;
    entries.forEach((e) => {
      if (e.obraId !== SEM_OBRA) return;
      count += 1;
      if (e.tipo === "entrada") receitas += Number(e.valor) || 0; else despesas += Number(e.valor) || 0;
    });
    return { receitas, despesas, count, resultado: receitas - despesas };
  }, [entries]);

  const socioStats = useMemo(() => {
    const map = {};
    socios.forEach((s) => (map[s.id] = { aportes: 0, retiradas: 0, proLabore: 0 }));
    movimentosSocios.forEach((m) => {
      if (!map[m.socioId]) map[m.socioId] = { aportes: 0, retiradas: 0, proLabore: 0 };
      if (m.tipo === "aporte") map[m.socioId].aportes += Number(m.valor) || 0;
      else if (m.tipo === "pro-labore") map[m.socioId].proLabore += Number(m.valor) || 0;
      else map[m.socioId].retiradas += Number(m.valor) || 0;
    });
    socios.forEach((s) => { map[s.id].saldo = map[s.id].aportes - map[s.id].retiradas; });
    return map;
  }, [socios, movimentosSocios]);

  const supplierNames = useMemo(() => {
    const set = new Set(fornecedoresSeed);
    entries.forEach((e) => { if (e.fornecedor && e.fornecedor !== "Não se aplica") set.add(e.fornecedor); });
    return [...set].sort((a, b) => a.localeCompare(b, "pt-BR"));
  }, [entries, fornecedoresSeed]);

  const supplierStats = useMemo(() => {
    const map = {};
    entries.forEach((e) => {
      if (e.tipo !== "saida" || !e.fornecedor || e.fornecedor === "Não se aplica") return;
      if (!map[e.fornecedor]) map[e.fornecedor] = { total: 0, count: 0, obras: new Set(), last: null };
      const s = map[e.fornecedor];
      s.total += Number(e.valor) || 0;
      s.count += 1;
      if (e.obraId && e.obraId !== SEM_OBRA) s.obras.add(e.obraId);
      if (!s.last || e.data > s.last) s.last = e.data;
    });
    return map;
  }, [entries]);

  const now = new Date();
  const curMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const prevDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevMonthKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, "0")}`;

  const monthTotals = (key) => {
    let entradas = 0, saidas = 0;
    entries.forEach((e) => {
      if (monthKey(e.data) !== key) return;
      if (e.tipo === "entrada") entradas += Number(e.valor) || 0; else saidas += Number(e.valor) || 0;
    });
    return { entradas, saidas, resultado: entradas - saidas };
  };
  const curMonth = monthTotals(curMonthKey);
  const prevMonth = monthTotals(prevMonthKey);

  const dashboard = {
    saldoTotal: accounts.filter((a) => a.status === "Ativa").reduce((s, a) => s + (accountStats[a.id]?.saldoAtual || 0), 0),
    totalGastoObras: obras.reduce((s, o) => s + (obraStats[o.id]?.despesas || 0), 0) + obraSemObra.despesas,
    obrasAndamento: obras.filter((o) => o.status === "Em execução").length,
  };

  const pct = (cur, prev) => {
    if (!prev) return null;
    return ((cur - prev) / Math.abs(prev)) * 100;
  };

  if (!ready) {
    return <div className="gf-app gf-loading">Carregando…</div>;
  }

  if (!currentUser) {
    return (
      <div className="gf-app">
        <style>{CSS}</style>
        <LoginScreen usuarios={usuarios} onLogin={(id) => setCurrentUserId(id)} />
      </div>
    );
  }

  const doLogout = () => setCurrentUserId(null);

  const NAV_PERMS = {
    lancamentos: null, saldos: "contasbancarias_visualizar", obras: "relatorios_acessar",
    fornecedores: "relatorios_acessar", contaspagar: null, custosinternos: "relatorios_acessar",
    dividasinvestimentos: "relatorios_acessar",
    socios: "socios_acessar", usuarios: "usuarios_gerenciar",
  };
  if (NAV_PERMS[page] && !can(currentUser, NAV_PERMS[page])) {
    return (
      <div className="gf-app">
        <style>{CSS}</style>
        <Sidebar page={page} setPage={setPage} currentUser={currentUser} />
        <main className="gf-main">
          <div className="gf-page"><div className="gf-card"><EmptyState icon={Scale} title="Acesso restrito" body="Seu usuário não tem permissão para acessar esta página." /></div></div>
        </main>
      </div>
    );
  }

  return (
    <div className="gf-app">
      <style>{CSS}</style>
      <Sidebar page={page} setPage={setPage} currentUser={currentUser} />
      <main className="gf-main">
        <TopBar
          curMonth={curMonth} prevMonth={prevMonth} pct={pct}
          dashboard={dashboard} savedTick={savedTick}
          onReset={resetDemo} onClear={clearAll}
          currentUser={currentUser} onLogout={doLogout}
        />
        <div className="gf-page">
          {page === "lancamentos" && (
            <LancamentosPage
              accounts={accounts} obras={obras} apropriacoes={apropriacoes} setApropriacoes={setApropriacoes}
              entries={entries} setEntries={setEntries} supplierNames={supplierNames}
              dividas={dividas} setDividas={setDividas} investimentos={investimentos} setInvestimentos={setInvestimentos}
              currentUser={currentUser} logAudit={logAudit}
            />
          )}
          {page === "saldos" && (
            <SaldosPage
              accounts={accounts} setAccounts={setAccounts} accountStats={accountStats}
              entries={entries} transfers={transfers} setTransfers={setTransfers}
              movimentosSocios={movimentosSocios}
              currentUser={currentUser} logAudit={logAudit}
            />
          )}
          {page === "obras" && (
            <ObrasPage
              obras={obras} setObras={setObras} obraStats={obraStats} obraSemObra={obraSemObra}
              entries={entries} apropriacoes={apropriacoes} setApropriacoes={setApropriacoes}
              contasPagar={contasPagar} orcamentoItens={orcamentoItens} setOrcamentoItens={setOrcamentoItens}
              currentUser={currentUser}
            />
          )}
          {page === "fornecedores" && (
            <FornecedoresPage
              supplierNames={supplierNames} supplierStats={supplierStats} entries={entries} obras={obras}
            />
          )}
          {page === "contaspagar" && (
            <ContasPagarPage
              contasPagar={contasPagar} setContasPagar={setContasPagar}
              entries={entries} setEntries={setEntries}
              accounts={accounts.filter((a) => a.status === "Ativa")}
              obras={obras} apropriacoes={apropriacoes} supplierNames={supplierNames}
              saldoAtualTotal={dashboard.saldoTotal}
              orcamentoItens={orcamentoItens}
              currentUser={currentUser} logAudit={logAudit}
            />
          )}
          {page === "custosinternos" && (
            <CustosInternosPage entries={entries} apropriacoes={apropriacoes} accounts={accounts} />
          )}
          {page === "dividasinvestimentos" && (
            <DividasInvestimentosPage
              dividas={dividas} investimentos={investimentos} entries={entries}
              accounts={accounts} obras={obras}
            />
          )}
          {page === "socios" && (
            <SociosPage
              socios={socios} setSocios={setSocios} socioStats={socioStats}
              movimentos={movimentosSocios} setMovimentos={setMovimentosSocios}
              accounts={accounts.filter((a) => a.status === "Ativa")}
            />
          )}
          {page === "usuarios" && (
            <UsuariosPage usuarios={usuarios} setUsuarios={setUsuarios} currentUser={currentUser} auditLog={auditLog} logAudit={logAudit} />
          )}
        </div>
      </main>
    </div>
  );
}

/* ---------------------------------------------------------------
   Sidebar / TopBar
--------------------------------------------------------------- */

function Sidebar({ page, setPage, currentUser }) {
  const items = [
    { id: "lancamentos", label: "Lançamentos", icon: Receipt },
    { id: "saldos", label: "Saldos Bancários", icon: Wallet, perm: "contasbancarias_visualizar" },
    { id: "obras", label: "Painel de Obras", icon: Building2, perm: "relatorios_acessar" },
    { id: "fornecedores", label: "Fornecedores", icon: Users, perm: "relatorios_acessar" },
    { id: "contaspagar", label: "Contas a Pagar", icon: CalendarClock },
    { id: "custosinternos", label: "Custos Internos", icon: BarChart3, perm: "relatorios_acessar" },
    { id: "dividasinvestimentos", label: "Dívidas e Investimentos", icon: TrendingUp, perm: "relatorios_acessar" },
    { id: "socios", label: "Sócios", icon: Scale, perm: "socios_acessar" },
    { id: "usuarios", label: "Usuários", icon: ShieldCheck, perm: "usuarios_gerenciar" },
  ].filter((it) => !it.perm || can(currentUser, it.perm));
  return (
    <nav className="gf-sidebar">
      <div className="gf-brand">
        <img src={GLOVER_LOGO_DATA_URI} alt="Glover Engenharia" className="gf-brand-logo" />
        <div className="gf-brand-sub">financeiro</div>
      </div>
      <div className="gf-nav">
        {items.map((it) => (
          <button
            key={it.id}
            className={`gf-nav-item ${page === it.id ? "is-active" : ""}`}
            onClick={() => setPage(it.id)}
          >
            <it.icon size={17} strokeWidth={1.8} />
            <span>{it.label}</span>
          </button>
        ))}
      </div>
      <div className="gf-sidebar-user">
        <div className="gf-sidebar-avatar">{(currentUser?.nome || "?").slice(0, 1)}</div>
        <div>
          <div className="gf-sidebar-user-name">{currentUser?.nome}</div>
          <div className="gf-sidebar-user-role">{PAPEL_LABEL[currentUser?.papel] || currentUser?.papel}</div>
        </div>
      </div>
      <div className="gf-sidebar-foot">Sistema com controle de acesso por perfil de usuário.</div>
    </nav>
  );
}

function TopBar({ curMonth, prevMonth, pct, dashboard, savedTick, onReset, onClear, currentUser, onLogout }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const deltaEntradas = pct(curMonth.entradas, prevMonth.entradas);
  const deltaSaidas = pct(curMonth.saidas, prevMonth.saidas);
  const deltaResultado = pct(curMonth.resultado, prevMonth.resultado);
  const isAdmin = can(currentUser, "config_alterar");

  const Delta = ({ v }) => {
    if (v === null || !isFinite(v)) return null;
    const up = v >= 0;
    return (
      <span className={`gf-delta ${up ? "up" : "down"}`}>
        {up ? <ArrowUpRight size={12} /> : <ArrowDownRight size={12} />}
        {Math.abs(v).toFixed(0)}% vs mês ant.
      </span>
    );
  };

  return (
    <header className="gf-topbar">
      <div className="gf-topbar-row">
        <div>
          <div className="gf-eyebrow">Resumo geral</div>
          <h1>Painel financeiro</h1>
        </div>
        <div className="gf-topbar-actions">
          <span className="gf-saved">{savedTick ? "Dados salvos automaticamente" : "Dados carregados"}</span>
          {isAdmin && (
            <div className="gf-menu-wrap">
              <button className="gf-btn gf-btn-ghost" onClick={() => setMenuOpen((v) => !v)}>Opções</button>
              {menuOpen && (
                <div className="gf-menu" onMouseLeave={() => setMenuOpen(false)}>
                  <button onClick={() => { onReset(); setMenuOpen(false); }}>Recarregar dados de exemplo</button>
                  <button onClick={() => { onClear(); setMenuOpen(false); }}>Apagar todos os dados</button>
                </div>
              )}
            </div>
          )}
          <div className="gf-topbar-user">
            <span className="gf-topbar-user-name">{currentUser?.nome}</span>
            <button className="gf-icon-btn" title="Sair" onClick={onLogout}><LogOut size={14} /></button>
          </div>
        </div>
      </div>
      <div className="gf-tiles">
        <div className="gf-tile gf-tile-primary">
          <span className="gf-tile-label">Saldo total em bancos</span>
          <span className="gf-tile-value">{BRL(dashboard.saldoTotal)}</span>
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Entradas no mês</span>
          <span className="gf-tile-value pos">{BRL(curMonth.entradas)}</span>
          <Delta v={deltaEntradas} />
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Saídas no mês</span>
          <span className="gf-tile-value neg">{BRL(curMonth.saidas)}</span>
          <Delta v={deltaSaidas} />
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Resultado do mês</span>
          <span className={`gf-tile-value ${curMonth.resultado >= 0 ? "pos" : "neg"}`}>{BRL(curMonth.resultado)}</span>
          <Delta v={deltaResultado} />
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Total gasto nas obras</span>
          <span className="gf-tile-value">{BRL(dashboard.totalGastoObras)}</span>
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Obras em andamento</span>
          <span className="gf-tile-value">{dashboard.obrasAndamento}</span>
        </div>
      </div>
    </header>
  );
}

/* ---------------------------------------------------------------
   Tela de login
--------------------------------------------------------------- */

function LoginScreen({ usuarios, onLogin }) {
  const [selecionado, setSelecionado] = useState(null);

  return (
    <div className="gf-login">
      <div className="gf-login-card">
        <div className="gf-login-logo-tile">
          <img src={GLOVER_LOGO_DATA_URI} alt="Glover Engenharia" className="gf-login-logo-img" />
        </div>
        <div className="gf-brand-sub" style={{ color: "var(--gf-muted)", marginBottom: 18 }}>financeiro</div>
        <h2 style={{ fontFamily: "var(--gf-font-display)", fontSize: 18, margin: "0 0 4px", color: "var(--gf-primary)" }}>Quem está acessando?</h2>
        <p className="gf-field-hint" style={{ marginBottom: 16 }}>Selecione seu usuário para entrar com o seu nível de permissão.</p>

        <div className="gf-login-users">
          {usuarios.filter((u) => u.ativo).map((u) => (
            <button key={u.id} className={`gf-login-user ${selecionado === u.id ? "is-selected" : ""}`} onClick={() => setSelecionado(u.id)}>
              <div className="gf-sidebar-avatar" style={{ background: "var(--gf-primary)" }}>{u.nome.slice(0, 1)}</div>
              <div>
                <div className="gf-login-user-name">{u.nome}</div>
                <div className="gf-login-user-role">{PAPEL_LABEL[u.papel] || u.papel}</div>
              </div>
              {selecionado === u.id && <Check size={16} className="pos" />}
            </button>
          ))}
        </div>

        <button className="gf-btn gf-btn-primary" style={{ width: "100%", justifyContent: "center", marginTop: 16 }}
          disabled={!selecionado} onClick={() => selecionado && onLogin(selecionado)}>
          Entrar
        </button>
        <p className="gf-field-hint" style={{ marginTop: 14, textAlign: "center" }}>
          Ambiente de demonstração: a seleção de usuário aqui não usa senha real. Em produção isso seria substituído por um login autenticado no servidor.
        </p>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Página: Lançamentos
--------------------------------------------------------------- */

function LancamentosPage({ accounts, obras, apropriacoes, setApropriacoes, entries, setEntries, supplierNames, dividas, setDividas, investimentos, setInvestimentos, currentUser, logAudit }) {
  const blankForm = {
    id: null, tipo: "saida", contaId: accounts.find((a) => a.status === "Ativa")?.id || "",
    obraId: SEM_OBRA, data: todayISO(), valor: "", fornecedor: "", apropriacao: apropriacoes[0] || "", descricao: "",
    classificacao: "",
    divInstituicao: "", divTipoOperacao: TIPOS_DIVIDA[0], divTaxaJuros: "", divSistemaAmortizacao: SISTEMAS_AMORTIZACAO[0],
    divQtdParcelas: "", divValorParcela: "", divProximoVencimento: "",
    dividaSelecionada: "", divAmortizacao: "", divJuros: "", divEncargos: "",
    invModo: "nova", investimentoSelecionado: "",
    invInstituicao: "", invTipoInvestimento: TIPOS_INVESTIMENTO[0], invDataVencimento: "", invLiquidez: LIQUIDEZ_OPCOES[0],
    invModoEntrada: "resgate",
  };
  const [form, setForm] = useState(blankForm);
  const [editingId, setEditingId] = useState(null);
  const [editingClassificacao, setEditingClassificacao] = useState(null);
  const [addingApropriacao, setAddingApropriacao] = useState(false);
  const [novaApropriacao, setNovaApropriacao] = useState("");
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const valorRef = useRef(null);
  const podeCriar = can(currentUser, "lancamentos_criar");
  const podeEditar = can(currentUser, "lancamentos_editar");
  const podeExcluir = can(currentUser, "lancamentos_excluir");

  const [filters, setFilters] = useState({
    q: "", de: "", ate: "", obraId: "", contaId: "", fornecedor: "", apropriacao: "", tipo: "",
  });

  const activeAccounts = accounts.filter((a) => a.status === "Ativa");

  const updateForm = (patch) => setForm((f) => ({ ...f, ...patch }));

  const submit = (e) => {
    e.preventDefault();
    if (!form.contaId || !form.valor || Number(form.valor) <= 0 || !form.data) return;
    if (editingId && !can(currentUser, "lancamentos_editar")) return;
    if (!editingId && !can(currentUser, "lancamentos_criar")) return;

    // Validações da classificação especial (apenas em novos lançamentos)
    let extra = {};
    if (!editingId && form.classificacao === "divida") {
      if (form.tipo === "entrada") {
        if (!form.divInstituicao.trim() || !form.divQtdParcelas) return;
      } else {
        if (!form.dividaSelecionada) return;
        const soma = (Number(form.divAmortizacao) || 0) + (Number(form.divJuros) || 0) + (Number(form.divEncargos) || 0);
        if (Math.abs(soma - Number(form.valor)) > 0.01) return;
      }
    }
    if (!editingId && form.classificacao === "investimento") {
      if (form.tipo === "saida") {
        if (form.invModo === "nova" && !form.invInstituicao.trim()) return;
        if (form.invModo === "aporte" && !form.investimentoSelecionado) return;
      } else {
        if (!form.investimentoSelecionado) return;
      }
    }

    const payload = {
      id: editingId || uid("lan"),
      tipo: form.tipo, contaId: form.contaId, obraId: form.obraId || SEM_OBRA,
      data: form.data, valor: Number(form.valor), fornecedor: form.fornecedor.trim(),
      apropriacao: form.apropriacao, descricao: form.descricao.trim(),
      createdAt: editingId ? entries.find((e2) => e2.id === editingId)?.createdAt || Date.now() : Date.now(),
    };

    if (!editingId && form.classificacao === "divida") {
      if (form.tipo === "entrada") {
        const novaDividaId = uid("div");
        setDividas((prev) => [...prev, {
          id: novaDividaId, instituicao: form.divInstituicao.trim(), tipoOperacao: form.divTipoOperacao,
          dataContratacao: form.data, valorContratado: Number(form.valor), taxaJuros: form.divTaxaJuros,
          sistemaAmortizacao: form.divSistemaAmortizacao, qtdParcelas: Number(form.divQtdParcelas) || 0,
          valorParcelaEstimado: Number(form.divValorParcela) || 0, proximoVencimento: form.divProximoVencimento || "",
          obraId: form.obraId || SEM_OBRA, status: "Ativa", createdAt: Date.now(),
        }]);
        extra = { dividaId: novaDividaId, dividaSubtipo: "contratacao" };
        payload.fornecedor = payload.fornecedor || form.divInstituicao.trim();
      } else {
        extra = {
          dividaId: form.dividaSelecionada, dividaSubtipo: "pagamento",
          amortizacao: Number(form.divAmortizacao) || 0, juros: Number(form.divJuros) || 0, encargos: Number(form.divEncargos) || 0,
        };
      }
    }
    if (!editingId && form.classificacao === "investimento") {
      if (form.tipo === "saida") {
        if (form.invModo === "nova") {
          const novoInvId = uid("inv");
          setInvestimentos((prev) => [...prev, {
            id: novoInvId, instituicao: form.invInstituicao.trim(), tipoInvestimento: form.invTipoInvestimento,
            dataAplicacao: form.data, valorAplicado: Number(form.valor), dataVencimento: form.invDataVencimento || "",
            liquidez: form.invLiquidez, obraId: form.obraId || SEM_OBRA, status: "Ativo", createdAt: Date.now(),
          }]);
          extra = { investimentoId: novoInvId, investimentoSubtipo: "aplicacao" };
          payload.fornecedor = payload.fornecedor || form.invInstituicao.trim();
        } else {
          extra = { investimentoId: form.investimentoSelecionado, investimentoSubtipo: "aporte" };
        }
      } else {
        extra = { investimentoId: form.investimentoSelecionado, investimentoSubtipo: form.invModoEntrada };
      }
    }
    Object.assign(payload, extra);

    if (editingId) {
      const anterior = entries.find((e2) => e2.id === editingId);
      setEntries((prev) => prev.map((e2) => (e2.id === editingId ? { ...anterior, ...payload } : e2)));
      if (anterior && anterior.valor !== payload.valor) {
        logAudit?.({
          tipoOperacao: "Alteração de valor de lançamento", registroAlterado: payload.descricao || payload.fornecedor || "Lançamento",
          valorAnterior: BRL(anterior.valor), valorNovo: BRL(payload.valor), centroCusto: payload.obraId,
        });
      }
    } else {
      setEntries((prev) => [payload, ...prev]);
    }
    setEditingId(null); setEditingClassificacao(null);
    setForm({ ...blankForm, contaId: form.contaId, obraId: form.obraId, tipo: form.tipo });
    setTimeout(() => valorRef.current?.focus(), 0);
  };

  const startEdit = (entry) => {
    if (!can(currentUser, "lancamentos_editar")) return;
    setEditingId(entry.id);
    setEditingClassificacao(entry.classificacao || (entry.dividaId ? "divida" : entry.investimentoId ? "investimento" : null));
    setForm({
      ...blankForm,
      id: entry.id, tipo: entry.tipo, contaId: entry.contaId, obraId: entry.obraId,
      data: entry.data, valor: String(entry.valor), fornecedor: entry.fornecedor, apropriacao: entry.apropriacao, descricao: entry.descricao,
    });
    window.scrollTo?.({ top: 0, behavior: "smooth" });
  };

  const cancelEdit = () => { setEditingId(null); setEditingClassificacao(null); setForm(blankForm); };

  const duplicate = (entry) => {
    if (!can(currentUser, "lancamentos_criar")) return;
    setEntries((prev) => [{ ...entry, id: uid("lan"), data: todayISO(), createdAt: Date.now() }, ...prev]);
  };

  const remove = (id) => {
    if (!can(currentUser, "lancamentos_excluir")) return;
    const entry = entries.find((e) => e.id === id);
    setEntries((prev) => prev.filter((e) => e.id !== id));
    setConfirmDeleteId(null);
    if (entry) {
      logAudit?.({
        tipoOperacao: "Exclusão de lançamento", registroAlterado: entry.descricao || entry.fornecedor || "Lançamento",
        valorAnterior: BRL(entry.valor), valorNovo: "—", centroCusto: entry.obraId,
      });
    }
  };

  const addApropriacao = () => {
    const v = novaApropriacao.trim();
    if (!v) return;
    if (!apropriacoes.includes(v)) setApropriacoes((prev) => [...prev, v]);
    updateForm({ apropriacao: v });
    setNovaApropriacao(""); setAddingApropriacao(false);
  };

  const accountName = (id) => accounts.find((a) => a.id === id)?.nome || "—";
  const obraName = (id) => (id === SEM_OBRA ? "Administrativo / Custos internos" : obras.find((o) => o.id === id)?.nome || "—");

  const filtered = useMemo(() => {
    return entries
      .filter((e) => (filters.tipo ? e.tipo === filters.tipo : true))
      .filter((e) => (filters.obraId ? e.obraId === filters.obraId : true))
      .filter((e) => (filters.contaId ? e.contaId === filters.contaId : true))
      .filter((e) => (filters.fornecedor ? e.fornecedor === filters.fornecedor : true))
      .filter((e) => (filters.apropriacao ? e.apropriacao === filters.apropriacao : true))
      .filter((e) => (filters.de ? e.data >= filters.de : true))
      .filter((e) => (filters.ate ? e.data <= filters.ate : true))
      .filter((e) => {
        if (!filters.q) return true;
        const q = filters.q.toLowerCase();
        return (
          (e.descricao || "").toLowerCase().includes(q) ||
          (e.fornecedor || "").toLowerCase().includes(q) ||
          accountName(e.contaId).toLowerCase().includes(q) ||
          obraName(e.obraId).toLowerCase().includes(q)
        );
      })
      .sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : b.createdAt - a.createdAt));
  }, [entries, filters, accounts, obras]);

  const totals = useMemo(() => {
    let entradas = 0, saidas = 0;
    filtered.forEach((e) => { if (e.tipo === "entrada") entradas += e.valor; else saidas += e.valor; });
    return { entradas, saidas, resultado: entradas - saidas };
  }, [filtered]);

  const datalistOptions = form.tipo === "entrada" ? ["Não se aplica", ...supplierNames] : supplierNames;

  return (
    <div className="gf-stack">
      {!podeCriar && !podeEditar ? (
        <section className="gf-card">
          <EmptyState icon={Lock} title="Sem permissão para lançar" body="Seu usuário não tem permissão para criar ou editar lançamentos. Fale com um administrador." />
        </section>
      ) : (
      <section className="gf-card gf-entry-card">
        <div className="gf-card-head">
          <h2>{editingId ? "Editar lançamento" : "Novo lançamento"}</h2>
          {editingId && <button className="gf-btn gf-btn-ghost" onClick={cancelEdit}>Cancelar edição</button>}
        </div>
        <form onSubmit={submit} className="gf-entry-form">
          <div className="gf-segmented">
            <button type="button" className={form.tipo === "entrada" ? "is-active pos" : ""} onClick={() => updateForm({ tipo: "entrada" })}>
              <ArrowUpRight size={14} /> Entrada
            </button>
            <button type="button" className={form.tipo === "saida" ? "is-active neg" : ""} onClick={() => updateForm({ tipo: "saida" })}>
              <ArrowDownRight size={14} /> Saída
            </button>
          </div>

          <div className="gf-form-grid">
            <Field label="Conta bancária">
              <select required value={form.contaId} onChange={(e) => updateForm({ contaId: e.target.value })}>
                <option value="" disabled>Selecione…</option>
                {activeAccounts.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
              </select>
            </Field>
            <Field label="Centro de custo" hint="Obra vinculada ou custo interno (salários, escritório, etc.)">
              <select value={form.obraId} onChange={(e) => updateForm({ obraId: e.target.value })}>
                <option value={SEM_OBRA}>Administrativo / Custos internos</option>
                {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
              </select>
            </Field>
            <Field label="Data">
              <input type="date" required value={form.data} onChange={(e) => updateForm({ data: e.target.value })} />
            </Field>
            <Field label="Valor (R$)">
              <input ref={valorRef} type="number" min="0.01" step="0.01" required placeholder="0,00"
                value={form.valor} onChange={(e) => updateForm({ valor: e.target.value })} />
            </Field>
            <Field label={form.tipo === "entrada" ? "Cliente / origem" : "Fornecedor"}>
              <input list="gf-fornecedores" placeholder="Digite ou selecione…" value={form.fornecedor}
                onChange={(e) => updateForm({ fornecedor: e.target.value })} />
              <datalist id="gf-fornecedores">
                {datalistOptions.map((f) => <option key={f} value={f} />)}
              </datalist>
            </Field>
            <Field label="Apropriação">
              {!addingApropriacao ? (
                <div className="gf-inline-select">
                  <select value={form.apropriacao} onChange={(e) => updateForm({ apropriacao: e.target.value })}>
                    {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
                  </select>
                  <button type="button" className="gf-icon-btn" title="Nova apropriação" onClick={() => setAddingApropriacao(true)}>
                    <Plus size={15} />
                  </button>
                </div>
              ) : (
                <div className="gf-inline-select">
                  <input autoFocus placeholder="Nova apropriação" value={novaApropriacao}
                    onChange={(e) => setNovaApropriacao(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addApropriacao())} />
                  <button type="button" className="gf-icon-btn" onClick={addApropriacao}><Check size={15} /></button>
                  <button type="button" className="gf-icon-btn" onClick={() => setAddingApropriacao(false)}><X size={15} /></button>
                </div>
              )}
            </Field>
            <Field label="Descrição" className="gf-span-2">
              <input placeholder="Ex: Compra de cimento e agregados" value={form.descricao}
                onChange={(e) => updateForm({ descricao: e.target.value })} />
            </Field>

            {editingId && editingClassificacao && (
              <div className="gf-span-2 gf-hint">
                Este lançamento está vinculado a {editingClassificacao === "divida" ? "uma dívida/crédito" : "um investimento"} — a classificação não pode ser alterada aqui. Ajustes de {editingClassificacao === "divida" ? "dívidas" : "investimentos"} são feitos na página "Dívidas e Investimentos".
              </div>
            )}

            {!editingId && (
              <Field label="Classificação especial" className="gf-span-2" hint="Opcional — vincula este lançamento ao painel Dívidas e Investimentos">
                <select value={form.classificacao} onChange={(e) => updateForm({ classificacao: e.target.value })}>
                  <option value="">Nenhuma</option>
                  <option value="divida">Dívida / Crédito</option>
                  <option value="investimento">Investimento</option>
                </select>
              </Field>
            )}

            {!editingId && form.classificacao === "divida" && form.tipo === "entrada" && (
              <div className="gf-span-2 gf-classificacao-box">
                <span className="gf-field-label" style={{ display: "block", marginBottom: 8 }}>Nova dívida / contratação</span>
                <div className="gf-form-grid" style={{ gridTemplateColumns: "repeat(3,1fr)" }}>
                  <Field label="Instituição financeira/credor">
                    <input value={form.divInstituicao} onChange={(e) => updateForm({ divInstituicao: e.target.value })} placeholder="Ex: Banco X" />
                  </Field>
                  <Field label="Tipo de operação">
                    <select value={form.divTipoOperacao} onChange={(e) => updateForm({ divTipoOperacao: e.target.value })}>
                      {TIPOS_DIVIDA.map((t) => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </Field>
                  <Field label="Taxa de juros">
                    <input value={form.divTaxaJuros} onChange={(e) => updateForm({ divTaxaJuros: e.target.value })} placeholder="Ex: 1,2% a.m." />
                  </Field>
                  <Field label="Sistema de amortização">
                    <select value={form.divSistemaAmortizacao} onChange={(e) => updateForm({ divSistemaAmortizacao: e.target.value })}>
                      {SISTEMAS_AMORTIZACAO.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </Field>
                  <Field label="Quantidade de parcelas">
                    <input type="number" min="1" value={form.divQtdParcelas} onChange={(e) => updateForm({ divQtdParcelas: e.target.value })} />
                  </Field>
                  <Field label="Valor estimado da parcela (R$)" hint="Opcional">
                    <input type="number" step="0.01" value={form.divValorParcela} onChange={(e) => updateForm({ divValorParcela: e.target.value })} />
                  </Field>
                  <Field label="Próximo vencimento" hint="Opcional" className="gf-span-2">
                    <input type="date" value={form.divProximoVencimento} onChange={(e) => updateForm({ divProximoVencimento: e.target.value })} />
                  </Field>
                </div>
                <p className="gf-field-hint">O valor lançado acima (entrada) será registrado como o valor originalmente contratado desta dívida.</p>
              </div>
            )}

            {!editingId && form.classificacao === "divida" && form.tipo === "saida" && (
              <div className="gf-span-2 gf-classificacao-box">
                <span className="gf-field-label" style={{ display: "block", marginBottom: 8 }}>Pagamento de parcela</span>
                <Field label="Dívida" className="gf-span-2">
                  <select value={form.dividaSelecionada} onChange={(e) => updateForm({ dividaSelecionada: e.target.value })}>
                    <option value="">Selecione a dívida…</option>
                    {dividas.filter((d) => d.status === "Ativa").map((d) => <option key={d.id} value={d.id}>{d.instituicao} — {d.tipoOperacao}</option>)}
                  </select>
                </Field>
                {dividas.length === 0 && <p className="gf-field-hint">Nenhuma dívida cadastrada ainda. Registre a contratação primeiro (lançamento de entrada).</p>}
                <div className="gf-form-grid" style={{ gridTemplateColumns: "repeat(3,1fr)", marginTop: 10 }}>
                  <Field label="Amortização (R$)"><input type="number" step="0.01" value={form.divAmortizacao} onChange={(e) => updateForm({ divAmortizacao: e.target.value })} /></Field>
                  <Field label="Juros (R$)"><input type="number" step="0.01" value={form.divJuros} onChange={(e) => updateForm({ divJuros: e.target.value })} /></Field>
                  <Field label="Encargos/IOF (R$)"><input type="number" step="0.01" value={form.divEncargos} onChange={(e) => updateForm({ divEncargos: e.target.value })} /></Field>
                </div>
                {(() => {
                  const soma = (Number(form.divAmortizacao) || 0) + (Number(form.divJuros) || 0) + (Number(form.divEncargos) || 0);
                  const dif = Math.round((Number(form.valor || 0) - soma) * 100) / 100;
                  const fecha = Math.abs(dif) < 0.01;
                  return (
                    <p className={`gf-field-hint ${fecha ? "" : "gf-warn"}`} style={{ marginTop: 8 }}>
                      {fecha ? "✓ A soma bate com o valor do lançamento." : `Falta detalhar ${BRL(dif)} para fechar com o valor total da parcela.`}
                    </p>
                  );
                })()}
              </div>
            )}

            {!editingId && form.classificacao === "investimento" && form.tipo === "saida" && (
              <div className="gf-span-2 gf-classificacao-box">
                <div className="gf-segmented" style={{ marginBottom: 10 }}>
                  <button type="button" className={form.invModo === "nova" ? "is-active" : ""} onClick={() => updateForm({ invModo: "nova" })}>Novo investimento</button>
                  <button type="button" className={form.invModo === "aporte" ? "is-active" : ""} onClick={() => updateForm({ invModo: "aporte" })}>Aporte em investimento existente</button>
                </div>
                {form.invModo === "nova" ? (
                  <div className="gf-form-grid" style={{ gridTemplateColumns: "repeat(2,1fr)" }}>
                    <Field label="Instituição"><input value={form.invInstituicao} onChange={(e) => updateForm({ invInstituicao: e.target.value })} placeholder="Ex: Banco X" /></Field>
                    <Field label="Tipo de investimento">
                      <select value={form.invTipoInvestimento} onChange={(e) => updateForm({ invTipoInvestimento: e.target.value })}>
                        {TIPOS_INVESTIMENTO.map((t) => <option key={t} value={t}>{t}</option>)}
                      </select>
                    </Field>
                    <Field label="Data de vencimento" hint="Opcional"><input type="date" value={form.invDataVencimento} onChange={(e) => updateForm({ invDataVencimento: e.target.value })} /></Field>
                    <Field label="Liquidez">
                      <select value={form.invLiquidez} onChange={(e) => updateForm({ invLiquidez: e.target.value })}>
                        {LIQUIDEZ_OPCOES.map((l) => <option key={l} value={l}>{l}</option>)}
                      </select>
                    </Field>
                  </div>
                ) : (
                  <Field label="Investimento">
                    <select value={form.investimentoSelecionado} onChange={(e) => updateForm({ investimentoSelecionado: e.target.value })}>
                      <option value="">Selecione…</option>
                      {investimentos.filter((i) => i.status === "Ativo").map((i) => <option key={i.id} value={i.id}>{i.instituicao} — {i.tipoInvestimento}</option>)}
                    </select>
                  </Field>
                )}
              </div>
            )}

            {!editingId && form.classificacao === "investimento" && form.tipo === "entrada" && (
              <div className="gf-span-2 gf-classificacao-box">
                <div className="gf-segmented" style={{ marginBottom: 10 }}>
                  <button type="button" className={form.invModoEntrada === "resgate" ? "is-active" : ""} onClick={() => updateForm({ invModoEntrada: "resgate" })}>Resgate</button>
                  <button type="button" className={form.invModoEntrada === "rendimento" ? "is-active" : ""} onClick={() => updateForm({ invModoEntrada: "rendimento" })}>Rendimento</button>
                </div>
                <Field label="Investimento">
                  <select value={form.investimentoSelecionado} onChange={(e) => updateForm({ investimentoSelecionado: e.target.value })}>
                    <option value="">Selecione…</option>
                    {investimentos.filter((i) => i.status === "Ativo").map((i) => <option key={i.id} value={i.id}>{i.instituicao} — {i.tipoInvestimento}</option>)}
                  </select>
                </Field>
                {investimentos.length === 0 && <p className="gf-field-hint">Nenhum investimento cadastrado ainda. Registre a aplicação inicial primeiro (lançamento de saída).</p>}
              </div>
            )}
          </div>

          <div className="gf-form-actions">
            <button type="submit" className="gf-btn gf-btn-primary" disabled={editingId ? !podeEditar : !podeCriar}>
              {editingId ? "Salvar alterações" : "Lançar movimento"}
            </button>
            {!editingId && <span className="gf-hint">Ao salvar, o formulário limpa e fica pronto para o próximo lançamento.</span>}
          </div>
        </form>
      </section>
      )}

      <section className="gf-card">
        <div className="gf-card-head">
          <h2>Lançamentos</h2>
          <div className="gf-summary-inline">
            <span>Entradas <b className="pos">{BRL(totals.entradas)}</b></span>
            <span>Saídas <b className="neg">{BRL(totals.saidas)}</b></span>
            <span>Resultado <b className={totals.resultado >= 0 ? "pos" : "neg"}>{BRL(totals.resultado)}</b></span>
          </div>
        </div>

        <div className="gf-filters">
          <div className="gf-search">
            <Search size={15} />
            <input placeholder="Pesquisar por descrição, fornecedor, conta, obra…" value={filters.q}
              onChange={(e) => setFilters((f) => ({ ...f, q: e.target.value }))} />
          </div>
          <div className="gf-filters-grid">
            <input type="date" title="De" value={filters.de} onChange={(e) => setFilters((f) => ({ ...f, de: e.target.value }))} />
            <input type="date" title="Até" value={filters.ate} onChange={(e) => setFilters((f) => ({ ...f, ate: e.target.value }))} />
            <select value={filters.tipo} onChange={(e) => setFilters((f) => ({ ...f, tipo: e.target.value }))}>
              <option value="">Tipo (todos)</option><option value="entrada">Entrada</option><option value="saida">Saída</option>
            </select>
            <select value={filters.obraId} onChange={(e) => setFilters((f) => ({ ...f, obraId: e.target.value }))}>
              <option value="">Centro de custo (todos)</option>
              <option value={SEM_OBRA}>Administrativo / Custos internos</option>
              {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
            </select>
            <select value={filters.contaId} onChange={(e) => setFilters((f) => ({ ...f, contaId: e.target.value }))}>
              <option value="">Conta (todas)</option>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
            </select>
            <select value={filters.fornecedor} onChange={(e) => setFilters((f) => ({ ...f, fornecedor: e.target.value }))}>
              <option value="">Fornecedor (todos)</option>
              {supplierNames.map((f) => <option key={f} value={f}>{f}</option>)}
            </select>
            <select value={filters.apropriacao} onChange={(e) => setFilters((f) => ({ ...f, apropriacao: e.target.value }))}>
              <option value="">Apropriação (todas)</option>
              {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
          </div>
        </div>

        <div className="gf-table-wrap">
          <table className="gf-table gf-table-ledger">
            <thead>
              <tr>
                <th>Data</th><th>Tipo</th><th>Conta</th><th>Centro de custo</th><th>Fornecedor</th>
                <th>Apropriação</th><th>Descrição</th><th className="num">Valor</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr><td colSpan={9}><EmptyState icon={Receipt} title="Nenhum lançamento encontrado" body="Ajuste os filtros ou registre um novo lançamento acima." /></td></tr>
              )}
              {filtered.map((e) => (
                <tr key={e.id} className={e.tipo === "entrada" ? "row-pos" : "row-neg"}>
                  <td>{fmtDate(e.data)}</td>
                  <td><Badge tone={e.tipo === "entrada" ? "green" : "red"}>{e.tipo === "entrada" ? "Entrada" : "Saída"}</Badge></td>
                  <td>{accountName(e.contaId)}</td>
                  <td>{obraName(e.obraId)}</td>
                  <td>{e.fornecedor || "—"}</td>
                  <td>{e.apropriacao}</td>
                  <td className="gf-td-desc" title={e.descricao}>
                    {e.descricao || "—"}
                    {e.dividaId && <span className="gf-mini-tag" title="Vinculado a uma dívida"> · Dívida</span>}
                    {e.investimentoId && <span className="gf-mini-tag" title="Vinculado a um investimento"> · Invest.</span>}
                  </td>
                  <td>
                    <div className="gf-row-actions">
                      {confirmDeleteId === e.id ? (
                        <>
                          <button className="gf-icon-btn danger" title="Confirmar exclusão" onClick={() => remove(e.id)}><Check size={14} /></button>
                          <button className="gf-icon-btn" title="Cancelar" onClick={() => setConfirmDeleteId(null)}><X size={14} /></button>
                        </>
                      ) : (
                        <>
                          {podeEditar && <button className="gf-icon-btn" title="Editar" onClick={() => startEdit(e)}><Pencil size={14} /></button>}
                          {podeCriar && <button className="gf-icon-btn" title="Duplicar" onClick={() => duplicate(e)}><Copy size={14} /></button>}
                          {podeExcluir && <button className="gf-icon-btn" title="Excluir" onClick={() => setConfirmDeleteId(e.id)}><Trash2 size={14} /></button>}
                        </>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

/* ---------------------------------------------------------------
   Página: Saldos nas Contas Bancárias
--------------------------------------------------------------- */

function SaldosPage({ accounts, setAccounts, accountStats, entries, transfers, setTransfers, movimentosSocios, currentUser, logAudit }) {
  const [showNew, setShowNew] = useState(false);
  const [showTransfer, setShowTransfer] = useState(false);
  const [editingAccount, setEditingAccount] = useState(null);
  const podeCadastrar = can(currentUser, "contasbancarias_cadastrar");
  const podeSaldoInicial = can(currentUser, "contasbancarias_saldoinicial");
  const podeTransferir = can(currentUser, "contasbancarias_transferir");

  const saveAccount = (data) => {
    if (data.id) {
      const anterior = accounts.find((a) => a.id === data.id);
      // Reforço de segurança: mesmo que o formulário seja manipulado, o saldo inicial só muda se o usuário tiver permissão.
      const saldoInicialFinal = podeSaldoInicial ? data.saldoInicial : anterior?.saldoInicial;
      setAccounts((prev) => prev.map((a) => (a.id === data.id ? { ...data, saldoInicial: saldoInicialFinal } : a)));
      if (anterior && podeSaldoInicial && anterior.saldoInicial !== data.saldoInicial) {
        logAudit?.({
          tipoOperacao: "Alteração de saldo inicial", registroAlterado: data.nome,
          valorAnterior: BRL(anterior.saldoInicial), valorNovo: BRL(data.saldoInicial),
        });
      } else if (anterior && (anterior.nome !== data.nome || anterior.status !== data.status || anterior.agencia !== data.agencia || anterior.numero !== data.numero)) {
        logAudit?.({ tipoOperacao: "Alteração de conta bancária", registroAlterado: data.nome, valorAnterior: anterior.nome, valorNovo: data.nome });
      }
    } else {
      if (!podeCadastrar) return;
      setAccounts((prev) => [...prev, { ...data, id: uid("acc") }]);
      logAudit?.({ tipoOperacao: "Cadastro de conta bancária", registroAlterado: data.nome, valorAnterior: "—", valorNovo: BRL(data.saldoInicial) });
    }
    setShowNew(false); setEditingAccount(null);
  };

  const doTransfer = (t) => {
    if (!podeTransferir) return;
    setTransfers((prev) => [...prev, { ...t, id: uid("trf"), createdAt: Date.now() }]);
    setShowTransfer(false);
  };

  // evolução do saldo total ao longo do tempo
  const evolutionData = useMemo(() => {
    const baseline = accounts.reduce((s, a) => s + (Number(a.saldoInicial) || 0), 0);
    const movs = [];
    entries.forEach((e) => movs.push({ data: e.data, delta: e.tipo === "entrada" ? e.valor : -e.valor }));
    (movimentosSocios || []).forEach((m) => movs.push({ data: m.data, delta: m.tipo === "aporte" ? m.valor : -m.valor }));
    // transferências não afetam o total da empresa (saem de uma conta, entram em outra)
    const byDate = {};
    movs.forEach((m) => { byDate[m.data] = (byDate[m.data] || 0) + m.delta; });
    const dates = Object.keys(byDate).sort();
    let running = baseline;
    const points = [{ data: "Início", saldo: baseline }];
    dates.forEach((d) => { running += byDate[d]; points.push({ data: fmtDate(d), saldo: running }); });
    return points;
  }, [accounts, entries, movimentosSocios]);

  return (
    <div className="gf-stack">
      <div className="gf-toolbar">
        <h2>Contas bancárias</h2>
        <div className="gf-toolbar-actions">
          {podeTransferir && (
            <button className="gf-btn gf-btn-ghost" onClick={() => setShowTransfer(true)}>
              <ArrowLeftRight size={15} /> Transferir entre contas
            </button>
          )}
          {podeCadastrar && (
            <button className="gf-btn gf-btn-primary" onClick={() => setShowNew(true)}>
              <Plus size={15} /> Nova conta bancária
            </button>
          )}
        </div>
      </div>

      {accounts.length === 0 ? (
        <div className="gf-card"><EmptyState icon={Wallet} title="Nenhuma conta cadastrada" body="Cadastre a primeira conta bancária para começar a lançar movimentos." /></div>
      ) : (
        <div className="gf-grid-cards">
          {accounts.map((a) => {
            const s = accountStats[a.id] || { entradas: 0, saidas: 0, saldoAtual: a.saldoInicial };
            return (
              <div className={`gf-card gf-account-card ${a.status !== "Ativa" ? "is-inactive" : ""}`} key={a.id}>
                <div className="gf-card-head">
                  <div>
                    <div className="gf-eyebrow">{a.banco}</div>
                    <h3>{a.nome}</h3>
                  </div>
                  <div className="gf-row-actions">
                    <Badge tone={a.status === "Ativa" ? "green" : "gray"}>{a.status}</Badge>
                    {podeCadastrar && <button className="gf-icon-btn" title="Editar conta" onClick={() => setEditingAccount(a)}><Pencil size={14} /></button>}
                  </div>
                </div>
                <div className="gf-account-meta">Ag. {a.agencia} · Conta {a.numero}</div>
                <div className="gf-account-rows">
                  <div><span>Saldo inicial</span><b className="mono">{BRL(a.saldoInicial)}</b></div>
                  <div><span>Entradas</span><b className="mono pos">+ {BRL(s.entradas)}</b></div>
                  <div><span>Saídas</span><b className="mono neg">− {BRL(s.saidas)}</b></div>
                  {(s.transfIn || s.transfOut) ? (
                    <div><span>Transferências (líq.)</span><b className={`mono ${s.transfIn - s.transfOut >= 0 ? "pos" : "neg"}`}>{BRL(s.transfIn - s.transfOut)}</b></div>
                  ) : null}
                  {(s.socioIn || s.socioOut) ? (
                    <div><span>Sócios (líq.)</span><b className={`mono ${s.socioIn - s.socioOut >= 0 ? "pos" : "neg"}`}>{BRL(s.socioIn - s.socioOut)}</b></div>
                  ) : null}
                </div>
                <div className="gf-account-balance">
                  <span>Saldo atual</span>
                  <strong className={`mono ${s.saldoAtual >= 0 ? "pos" : "neg"}`}>{BRL(s.saldoAtual)}</strong>
                </div>
              </div>
            );
          })}
        </div>
      )}

      <section className="gf-card">
        <div className="gf-card-head"><h2>Evolução do saldo bancário</h2></div>
        <div className="gf-chart-box">
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={evolutionData} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
              <CartesianGrid stroke="var(--gf-border)" vertical={false} />
              <XAxis dataKey="data" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false}
                tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
              <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
              <Line type="monotone" dataKey="saldo" stroke="var(--gf-accent)" strokeWidth={2.2} dot={false} activeDot={{ r: 4 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </section>

      {(showNew || editingAccount) && (
        <AccountModal account={editingAccount} podeSaldoInicial={podeSaldoInicial} onClose={() => { setShowNew(false); setEditingAccount(null); }} onSave={saveAccount} />
      )}
      {showTransfer && (
        <TransferModal accounts={accounts.filter((a) => a.status === "Ativa")} onClose={() => setShowTransfer(false)} onSave={doTransfer} />
      )}
    </div>
  );
}

function AccountModal({ account, podeSaldoInicial, onClose, onSave }) {
  const [f, setF] = useState(account || {
    banco: "", nome: "", agencia: "", numero: "", saldoInicial: "", dataSaldoInicial: todayISO(), status: "Ativa",
  });
  const [tentouEditarSaldo, setTentouEditarSaldo] = useState(false);
  const upd = (p) => setF((s) => ({ ...s, ...p }));
  const submit = (e) => {
    e.preventDefault();
    if (!f.banco || !f.nome) return;
    onSave({ ...f, saldoInicial: Number(f.saldoInicial) || 0 });
  };
  return (
    <Modal title={account ? "Editar conta bancária" : "Nova conta bancária"} onClose={onClose}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <Field label="Banco"><input required value={f.banco} onChange={(e) => upd({ banco: e.target.value })} placeholder="Ex: Itaú" /></Field>
        <Field label="Nome / apelido da conta"><input required value={f.nome} onChange={(e) => upd({ nome: e.target.value })} placeholder="Ex: Itaú - Conta Principal" /></Field>
        <Field label="Agência"><input value={f.agencia} onChange={(e) => upd({ agencia: e.target.value })} /></Field>
        <Field label="Número da conta"><input value={f.numero} onChange={(e) => upd({ numero: e.target.value })} /></Field>
        <Field label="Saldo inicial (R$)">
          {podeSaldoInicial ? (
            <input type="number" step="0.01" value={f.saldoInicial} onChange={(e) => upd({ saldoInicial: e.target.value })} />
          ) : (
            <div className="gf-locked-field" onClick={() => setTentouEditarSaldo(true)}>
              <span className="mono">{BRL(f.saldoInicial)}</span>
              <Lock size={14} />
            </div>
          )}
        </Field>
        {tentouEditarSaldo && !podeSaldoInicial && (
          <div className="gf-warn gf-span-2">Acesso restrito. A alteração do saldo inicial só pode ser realizada por um usuário administrador.</div>
        )}
        <Field label="Data do saldo inicial"><input type="date" disabled={!podeSaldoInicial} value={f.dataSaldoInicial} onChange={(e) => upd({ dataSaldoInicial: e.target.value })} /></Field>
        <Field label="Status">
          <select value={f.status} onChange={(e) => upd({ status: e.target.value })}>
            <option>Ativa</option><option>Inativa</option>
          </select>
        </Field>
        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">Salvar conta</button>
        </div>
      </form>
    </Modal>
  );
}

function TransferModal({ accounts, onClose, onSave }) {
  const [f, setF] = useState({ origemId: accounts[0]?.id || "", destinoId: accounts[1]?.id || "", valor: "", data: todayISO(), descricao: "" });
  const upd = (p) => setF((s) => ({ ...s, ...p }));
  const submit = (e) => {
    e.preventDefault();
    if (!f.origemId || !f.destinoId || f.origemId === f.destinoId || !f.valor) return;
    onSave({ ...f, valor: Number(f.valor) });
  };
  return (
    <Modal title="Transferência entre contas" onClose={onClose}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <Field label="Conta de origem">
          <select required value={f.origemId} onChange={(e) => upd({ origemId: e.target.value })}>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
          </select>
        </Field>
        <Field label="Conta de destino">
          <select required value={f.destinoId} onChange={(e) => upd({ destinoId: e.target.value })}>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
          </select>
        </Field>
        <Field label="Valor (R$)"><input type="number" min="0.01" step="0.01" required value={f.valor} onChange={(e) => upd({ valor: e.target.value })} /></Field>
        <Field label="Data"><input type="date" required value={f.data} onChange={(e) => upd({ data: e.target.value })} /></Field>
        <Field label="Descrição" className="gf-span-2"><input value={f.descricao} onChange={(e) => upd({ descricao: e.target.value })} placeholder="Opcional" /></Field>
        {f.origemId === f.destinoId && <div className="gf-warn gf-span-2">Escolha contas diferentes para origem e destino.</div>}
        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">Confirmar transferência</button>
          <span className="gf-hint">Não afeta receitas/despesas nem o resultado das obras.</span>
        </div>
      </form>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Página: Painel de Controle - Obras
--------------------------------------------------------------- */

const PIE_COLORS = ["#2F8F5B", "#0E3B2E", "#7FB69A", "#B23A34", "#C98A2B", "#5D7566", "#3E6B57", "#9AA39B", "#1D5A45", "#84A98C", "#4A5259"];

function ObrasPage({ obras, setObras, obraStats, obraSemObra, entries, apropriacoes, setApropriacoes, contasPagar, orcamentoItens, setOrcamentoItens, currentUser }) {
  const [showNew, setShowNew] = useState(false);
  const [editingObra, setEditingObra] = useState(null);
  const [detailObraId, setDetailObraId] = useState(null);
  const podeGerenciar = can(currentUser, "centroscusto_gerenciar");

  const saveObra = (data) => {
    if (!podeGerenciar) return;
    if (data.id) setObras((prev) => prev.map((o) => (o.id === data.id ? data : o)));
    else setObras((prev) => [...prev, { ...data, id: uid("obra") }]);
    setShowNew(false); setEditingObra(null);
  };

  const comparativo = obras.map((o) => ({ nome: o.nome.length > 14 ? o.nome.slice(0, 13) + "…" : o.nome, resultado: obraStats[o.id]?.resultado || 0 }));
  const detailObra = obras.find((o) => o.id === detailObraId);

  return (
    <div className="gf-stack">
      <div className="gf-toolbar">
        <h2>Obras</h2>
        {podeGerenciar && <button className="gf-btn gf-btn-primary" onClick={() => setShowNew(true)}><Plus size={15} /> Nova obra</button>}
      </div>

      {obras.length === 0 ? (
        <div className="gf-card"><EmptyState icon={Building2} title="Nenhuma obra cadastrada" body="Cadastre uma obra para acompanhar receitas, despesas e resultado." /></div>
      ) : (
        <div className="gf-grid-cards">
          {obras.map((o) => {
            const s = obraStats[o.id] || { receitas: 0, despesas: 0, resultado: 0, margem: null, count: 0 };
            const orc = computeOrcamentoObra(o.id, orcamentoItens, contasPagar, entries);
            return (
              <button className="gf-card gf-obra-card" key={o.id} onClick={() => setDetailObraId(o.id)}>
                <div className="gf-card-head">
                  <div>
                    <div className="gf-eyebrow">{o.codigo}</div>
                    <h3>{o.nome}</h3>
                  </div>
                  <Badge tone={o.status === "Concluída" ? "gray" : o.status === "Cancelada" ? "red" : "green"}>{o.status}</Badge>
                </div>
                <div className="gf-account-rows">
                  <div><span>Receitas</span><b className="mono pos">{BRL(s.receitas)}</b></div>
                  <div><span>Despesas</span><b className="mono neg">{BRL(s.despesas)}</b></div>
                  <div><span>Margem</span><b className="mono">{s.margem === null ? "—" : `${s.margem.toFixed(1)}%`}</b></div>
                  <div><span>Lançamentos</span><b className="mono">{s.count}</b></div>
                </div>
                <div className="gf-account-balance">
                  <span>Resultado</span>
                  <strong className={`mono ${s.resultado >= 0 ? "pos" : "neg"}`}>{BRL(s.resultado)}</strong>
                </div>
                {orc.temOrcamento && (
                  <div className="gf-orcamento-chip">
                    <span className={`gf-dot gf-dot-${orc.totais.pct > 100 ? "red" : orc.totais.pct >= 80 ? "orange" : "green"}`} />
                    Orçamento {orc.totais.pct.toFixed(0)}% comprometido
                  </div>
                )}
                <span className="gf-card-cta">Ver detalhes <ChevronRight size={14} /></span>
              </button>
            );
          })}
          {(obraSemObra.count > 0) && (
            <div className="gf-card gf-obra-card gf-obra-card-static">
              <div className="gf-card-head">
                <div><div className="gf-eyebrow">—</div><h3>Administrativo / Custos internos</h3></div>
                <Badge tone="gray">Geral</Badge>
              </div>
              <div className="gf-account-rows">
                <div><span>Receitas</span><b className="mono pos">{BRL(obraSemObra.receitas)}</b></div>
                <div><span>Despesas</span><b className="mono neg">{BRL(obraSemObra.despesas)}</b></div>
                <div><span>Lançamentos</span><b className="mono">{obraSemObra.count}</b></div>
              </div>
              <div className="gf-account-balance">
                <span>Resultado</span>
                <strong className={`mono ${obraSemObra.resultado >= 0 ? "pos" : "neg"}`}>{BRL(obraSemObra.resultado)}</strong>
              </div>
            </div>
          )}
        </div>
      )}

      {obras.length > 1 && (
        <section className="gf-card">
          <div className="gf-card-head"><h2>Comparativo entre obras</h2></div>
          <div className="gf-chart-box">
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={comparativo} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
                <CartesianGrid stroke="var(--gf-border)" vertical={false} />
                <XAxis dataKey="nome" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
                <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                <Bar dataKey="resultado" radius={[4, 4, 0, 0]}>
                  {comparativo.map((c, i) => <Cell key={i} fill={c.resultado >= 0 ? "var(--gf-accent)" : "var(--gf-danger)"} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </section>
      )}

      {(showNew || editingObra) && (
        <ObraModal obra={editingObra} onClose={() => { setShowNew(false); setEditingObra(null); }} onSave={saveObra} />
      )}
      {detailObra && (
        <ObraDetail
          obra={detailObra} stats={obraStats[detailObra.id]} entries={entries.filter((e) => e.obraId === detailObra.id)}
          apropriacoes={apropriacoes} setApropriacoes={setApropriacoes} onClose={() => setDetailObraId(null)} podeEditar={podeGerenciar}
          onEdit={() => { setEditingObra(detailObra); setDetailObraId(null); }}
          contasPagar={contasPagar} orcamentoItens={orcamentoItens} setOrcamentoItens={setOrcamentoItens}
          allEntries={entries}
        />
      )}
    </div>
  );
}

function ObraModal({ obra, onClose, onSave }) {
  const [f, setF] = useState(obra || {
    nome: "", codigo: "", cliente: "", localizacao: "", dataInicio: todayISO(), dataFim: "", orcamento: "", status: "Em planejamento",
  });
  const upd = (p) => setF((s) => ({ ...s, ...p }));
  const submit = (e) => {
    e.preventDefault();
    if (!f.nome) return;
    onSave({ ...f, orcamento: Number(f.orcamento) || 0 });
  };
  return (
    <Modal title={obra ? "Editar obra" : "Nova obra"} onClose={onClose} width={600}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <Field label="Nome da obra" className="gf-span-2"><input required value={f.nome} onChange={(e) => upd({ nome: e.target.value })} /></Field>
        <Field label="Código da obra"><input value={f.codigo} onChange={(e) => upd({ codigo: e.target.value })} placeholder="Ex: OB-004" /></Field>
        <Field label="Cliente"><input value={f.cliente} onChange={(e) => upd({ cliente: e.target.value })} /></Field>
        <Field label="Localização"><input value={f.localizacao} onChange={(e) => upd({ localizacao: e.target.value })} /></Field>
        <Field label="Orçamento previsto (R$)"><input type="number" step="0.01" value={f.orcamento} onChange={(e) => upd({ orcamento: e.target.value })} /></Field>
        <Field label="Data de início"><input type="date" value={f.dataInicio} onChange={(e) => upd({ dataInicio: e.target.value })} /></Field>
        <Field label="Data prevista de término"><input type="date" value={f.dataFim} onChange={(e) => upd({ dataFim: e.target.value })} /></Field>
        <Field label="Status">
          <select value={f.status} onChange={(e) => upd({ status: e.target.value })}>
            {OBRA_STATUS.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">Salvar obra</button>
        </div>
      </form>
    </Modal>
  );
}

function ObraDetail({ obra, stats, entries, apropriacoes, setApropriacoes, onClose, onEdit, contasPagar, orcamentoItens, setOrcamentoItens, allEntries }) {
  const [tab, setTab] = useState("financeiro");
  const [apropFilter, setApropFilter] = useState("");
  const filtered = apropFilter ? entries.filter((e) => e.apropriacao === apropFilter) : entries;

  const porMes = useMemo(() => {
    const map = {};
    filtered.forEach((e) => {
      const k = monthKey(e.data);
      if (!map[k]) map[k] = { mes: k, receitas: 0, despesas: 0 };
      if (e.tipo === "entrada") map[k].receitas += e.valor; else map[k].despesas += e.valor;
    });
    return Object.values(map).sort((a, b) => a.mes.localeCompare(b.mes)).map((r) => ({ ...r, mesLabel: monthLabel(r.mes) }));
  }, [filtered]);

  const acumulado = useMemo(() => {
    let running = 0;
    return porMes.map((m) => { running += m.receitas - m.despesas; return { mesLabel: m.mesLabel, resultado: running }; });
  }, [porMes]);

  const porApropriacao = useMemo(() => {
    const map = {};
    filtered.filter((e) => e.tipo === "saida").forEach((e) => { map[e.apropriacao] = (map[e.apropriacao] || 0) + e.valor; });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  return (
    <Modal title={obra.nome} onClose={onClose} width={800}>
      <div className="gf-obra-detail">
        <div className="gf-detail-meta">
          <div><span>Cliente</span><b>{obra.cliente || "—"}</b></div>
          <div><span>Local</span><b>{obra.localizacao || "—"}</b></div>
          <div><span>Orçamento previsto</span><b className="mono">{BRL(obra.orcamento)}</b></div>
          <div><span>Período</span><b>{fmtDate(obra.dataInicio)} — {fmtDate(obra.dataFim)}</b></div>
          <button className="gf-btn gf-btn-ghost gf-detail-edit" onClick={onEdit}><Pencil size={13} /> Editar obra</button>
        </div>

        <div className="gf-segmented" style={{ marginTop: 4, marginBottom: 14 }}>
          <button className={tab === "financeiro" ? "is-active" : ""} onClick={() => setTab("financeiro")}>Visão financeira</button>
          <button className={tab === "orcamento" ? "is-active" : ""} onClick={() => setTab("orcamento")}>Controle orçamentário</button>
        </div>

        {tab === "financeiro" && (
          <>
            <div className="gf-account-rows gf-detail-stats">
              <div><span>Receitas</span><b className="mono pos">{BRL(stats?.receitas || 0)}</b></div>
              <div><span>Despesas</span><b className="mono neg">{BRL(stats?.despesas || 0)}</b></div>
              <div><span>Resultado</span><b className={`mono ${(stats?.resultado || 0) >= 0 ? "pos" : "neg"}`}>{BRL(stats?.resultado || 0)}</b></div>
              <div><span>Margem</span><b className="mono">{stats?.margem === null || stats?.margem === undefined ? "—" : `${stats.margem.toFixed(1)}%`}</b></div>
            </div>

            <div className="gf-field" style={{ maxWidth: 260, margin: "4px 0 14px" }}>
              <span className="gf-field-label">Filtrar por apropriação</span>
              <select value={apropFilter} onChange={(e) => setApropFilter(e.target.value)}>
                <option value="">Todas</option>
                {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
              </select>
            </div>

            <h4 className="gf-chart-title">Receitas x Despesas por mês</h4>
            <div className="gf-chart-box">
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={porMes} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                  <CartesianGrid stroke="var(--gf-border)" vertical={false} />
                  <XAxis dataKey="mesLabel" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={40} />
                  <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="receitas" name="Receitas" fill="var(--gf-accent)" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="despesas" name="Despesas" fill="var(--gf-danger)" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <h4 className="gf-chart-title">Resultado acumulado</h4>
            <div className="gf-chart-box">
              <ResponsiveContainer width="100%" height={200}>
                <LineChart data={acumulado} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                  <CartesianGrid stroke="var(--gf-border)" vertical={false} />
                  <XAxis dataKey="mesLabel" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={40} />
                  <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                  <Line type="monotone" dataKey="resultado" stroke="var(--gf-primary)" strokeWidth={2.2} dot={{ r: 3 }} />
                </LineChart>
              </ResponsiveContainer>
            </div>

            <h4 className="gf-chart-title">Despesas por apropriação</h4>
            <div className="gf-chart-box gf-chart-box-pie">
              {porApropriacao.length === 0 ? <EmptyState icon={Building2} title="Sem despesas registradas" body="Nenhuma saída lançada para esta obra ainda." /> : (
                <ResponsiveContainer width="100%" height={240}>
                  <PieChart>
                    <Pie data={porApropriacao} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={85} label={(d) => d.name}>
                      {porApropriacao.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                    </Pie>
                    <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                  </PieChart>
                </ResponsiveContainer>
              )}
            </div>
          </>
        )}

        {tab === "orcamento" && (
          <OrcamentoTab
            obra={obra} contasPagar={contasPagar} orcamentoItens={orcamentoItens} setOrcamentoItens={setOrcamentoItens}
            allEntries={allEntries} apropriacoes={apropriacoes} setApropriacoes={setApropriacoes}
          />
        )}
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Painel de Controle Orçamentário (dentro de ObraDetail)
--------------------------------------------------------------- */

function OrcamentoTab({ obra, contasPagar, orcamentoItens, setOrcamentoItens, allEntries, apropriacoes, setApropriacoes }) {
  const [showImport, setShowImport] = useState(false);
  const [apropDetail, setApropDetail] = useState(null);

  const orc = useMemo(() => computeOrcamentoObra(obra.id, orcamentoItens, contasPagar, allEntries), [obra.id, orcamentoItens, contasPagar, allEntries]);

  const graficoData = orc.linhas.map((l) => ({
    nome: l.apropriacao.length > 12 ? l.apropriacao.slice(0, 11) + "…" : l.apropriacao,
    Orçado: l.orcado, Comprometido: l.comprometido, Realizado: l.realizado,
  }));

  const confirmarImportacao = (itens) => {
    setOrcamentoItens((prev) => [...prev.filter((o) => o.obraId !== obra.id), ...itens]);
    const novasApropriacoes = [...new Set(itens.map((i) => i.apropriacao))].filter((a) => !apropriacoes.includes(a));
    if (novasApropriacoes.length) setApropriacoes((prev) => [...prev, ...novasApropriacoes]);
    setShowImport(false);
  };

  const alertas = orc.linhas.filter((l) => l.status === "vermelho");

  return (
    <div>
      <div className="gf-toolbar" style={{ marginBottom: 12 }}>
        <div />
        <button className="gf-btn gf-btn-ghost" onClick={() => setShowImport(true)}>
          <Layers size={14} /> Carregar orçamento (Excel)
        </button>
      </div>

      {!orc.temOrcamento ? (
        <EmptyState icon={Layers} title="Nenhum orçamento importado" body='Clique em "Carregar orçamento (Excel)" para importar a planilha de orçamento desta obra e começar o acompanhamento.' />
      ) : (
        <>
          {alertas.length > 0 && (
            <div className="gf-budget-alert">
              <AlertTriangle size={16} />
              <div>
                <strong>Orçamento estourado em {alertas.length} apropriaç{alertas.length > 1 ? "ões" : "ão"}</strong>
                <span>{alertas.map((a) => `${a.apropriacao} (${BRL(a.comprometido - a.orcado)} acima)`).join(" · ")}</span>
              </div>
            </div>
          )}

          <div className="gf-account-rows gf-detail-stats" style={{ flexWrap: "wrap" }}>
            <div><span>Orçamento total</span><b className="mono">{BRL(orc.totais.orcado)}</b></div>
            <div><span>Total comprometido</span><b className="mono">{BRL(orc.totais.comprometido)}</b></div>
            <div><span>Total realizado</span><b className="mono pos">{BRL(orc.totais.realizado)}</b></div>
            <div><span>Contas a pagar (pendente)</span><b className="mono">{BRL(orc.totais.pendente)}</b></div>
            <div><span>Saldo orçamentário</span><b className={`mono ${orc.totais.saldo >= 0 ? "pos" : "neg"}`}>{BRL(orc.totais.saldo)}</b></div>
            <div><span>% comprometido</span><b className="mono">{orc.totais.pct.toFixed(0)}%</b></div>
          </div>

          <h4 className="gf-chart-title">Orçado × Comprometido × Realizado</h4>
          <div className="gf-chart-box">
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={graficoData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                <CartesianGrid stroke="var(--gf-border)" vertical={false} />
                <XAxis dataKey="nome" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
                <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="Orçado" fill="var(--gf-gray-100)" stroke="var(--gf-border)" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Comprometido" fill="#C98A2B" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Realizado" fill="var(--gf-accent)" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>

          <h4 className="gf-chart-title">Controle orçamentário por apropriação</h4>
          <div className="gf-table-wrap">
            <table className="gf-table gf-table-orcamento">
              <thead>
                <tr><th>Apropriação</th><th className="num">Orçado</th><th className="num">Comprometido</th><th className="num">Realizado</th><th className="num">Saldo</th><th className="num">% Utilizado</th><th>Status</th></tr>
              </thead>
              <tbody>
                {orc.linhas.map((l) => (
                  <tr key={l.apropriacao} className="gf-row-clickable" onClick={() => setApropDetail(l)}>
                    <td><b>{l.apropriacao}</b></td>
                    <td className="num mono">{l.orcado > 0 ? BRL(l.orcado) : "—"}</td>
                    <td className="num mono">{BRL(l.comprometido)}</td>
                    <td className="num mono pos">{BRL(l.realizado)}</td>
                    <td className={`num mono ${l.saldo >= 0 ? "pos" : "neg"}`}>{l.orcado > 0 ? BRL(l.saldo) : "—"}</td>
                    <td className="num mono">{l.orcado > 0 ? `${l.pct.toFixed(0)}%` : "—"}</td>
                    <td><span className={`gf-dot gf-dot-${l.status === "vermelho" ? "red" : l.status === "amarelo" ? "orange" : "green"}`} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {showImport && (
        <ImportOrcamentoModal obra={obra} onClose={() => setShowImport(false)} onConfirm={confirmarImportacao} />
      )}
      {apropDetail && (
        <ApropriacaoOrcamentoModal
          linha={apropDetail} obra={obra} contasPagar={contasPagar} allEntries={allEntries}
          onClose={() => setApropDetail(null)}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------
   Importação de orçamento (Excel) e detalhe por apropriação
--------------------------------------------------------------- */

const CAMPOS_ORCAMENTO = [
  { key: "apropriacao", label: "Apropriação", obrigatorio: true, keywords: ["apropria", "categoria", "item"] },
  { key: "descricao", label: "Descrição", obrigatorio: false, keywords: ["descri", "detalhe"] },
  { key: "valorTotal", label: "Verba orçada / Valor total", obrigatorio: true, keywords: ["verba", "valor total", "orçad", "orcad", "total"] },
  { key: "unidade", label: "Unidade", obrigatorio: false, keywords: ["unid", "un."] },
  { key: "quantidade", label: "Quantidade", obrigatorio: false, keywords: ["quant", "qtd"] },
  { key: "valorUnitario", label: "Valor unitário", obrigatorio: false, keywords: ["unitário", "unitario", "vlr unit"] },
];

function guessColumn(headers, keywords) {
  const idx = headers.findIndex((h) => keywords.some((k) => String(h || "").toLowerCase().includes(k)));
  return idx >= 0 ? idx : "";
}

function parseNumeroPlanilha(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  const s = String(v).replace(/[^\d,.-]/g, "").replace(/\.(?=\d{3}(\D|$))/g, "").replace(",", ".");
  return Number(s) || 0;
}

function ImportOrcamentoModal({ obra, onClose, onConfirm }) {
  const [step, setStep] = useState("upload"); // upload | mapping
  const [rows, setRows] = useState([]); // array de arrays (inclui header na linha 0)
  const [nomeArquivo, setNomeArquivo] = useState("");
  const [mapping, setMapping] = useState({});
  const [erro, setErro] = useState("");

  const headers = rows[0] || [];
  const dataRows = rows.slice(1).filter((r) => r.some((c) => c !== undefined && c !== ""));

  const handleFile = async (file) => {
    setErro("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const asRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      if (!asRows.length) { setErro("Não foi possível ler nenhuma linha desta planilha."); return; }
      setRows(asRows);
      setNomeArquivo(file.name);
      const hdrs = asRows[0].map((h) => String(h || ""));
      const auto = {};
      CAMPOS_ORCAMENTO.forEach((c) => { auto[c.key] = guessColumn(hdrs, c.keywords); });
      setMapping(auto);
      setStep("mapping");
    } catch (e) {
      setErro("Não foi possível ler este arquivo. Verifique se é um .xlsx, .xls ou .csv válido.");
    }
  };

  const previewItens = useMemo(() => {
    if (mapping.apropriacao === "" || mapping.apropriacao === undefined) return [];
    return dataRows.map((r) => ({
      apropriacao: String(r[mapping.apropriacao] || "").trim(),
      descricao: mapping.descricao !== "" && mapping.descricao !== undefined ? String(r[mapping.descricao] || "") : "",
      unidade: mapping.unidade !== "" && mapping.unidade !== undefined ? String(r[mapping.unidade] || "") : "",
      quantidade: mapping.quantidade !== "" && mapping.quantidade !== undefined ? parseNumeroPlanilha(r[mapping.quantidade]) : "",
      valorUnitario: mapping.valorUnitario !== "" && mapping.valorUnitario !== undefined ? parseNumeroPlanilha(r[mapping.valorUnitario]) : "",
      valorTotal: mapping.valorTotal !== "" && mapping.valorTotal !== undefined ? parseNumeroPlanilha(r[mapping.valorTotal]) : 0,
    })).filter((it) => it.apropriacao);
  }, [dataRows, mapping]);

  const totalOrcado = previewItens.reduce((s, i) => s + i.valorTotal, 0);
  const mappingCompleto = CAMPOS_ORCAMENTO.filter((c) => c.obrigatorio).every((c) => mapping[c.key] !== "" && mapping[c.key] !== undefined);

  const confirmar = () => {
    const itens = previewItens.map((it) => ({
      id: uid("orc"), obraId: obra.id, apropriacao: it.apropriacao, descricao: it.descricao,
      unidade: it.unidade, quantidade: it.quantidade, valorUnitario: it.valorUnitario, valorTotal: it.valorTotal,
      importadoEm: Date.now(),
    }));
    onConfirm(itens);
  };

  return (
    <Modal title={`Carregar orçamento — ${obra.nome}`} onClose={onClose} width={720}>
      {step === "upload" && (
        <div className="gf-import-upload">
          <input
            type="file" accept=".xlsx,.xls,.csv"
            onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
            id="gf-orc-file"
            style={{ display: "none" }}
          />
          <label htmlFor="gf-orc-file" className="gf-import-dropzone">
            <Layers size={26} />
            <strong>Clique para selecionar a planilha de orçamento</strong>
            <span>.xlsx, .xls ou .csv — deve conter ao menos as colunas Apropriação e Verba orçada / Valor total</span>
          </label>
          {erro && <div className="gf-warn" style={{ marginTop: 10 }}>{erro}</div>}
        </div>
      )}

      {step === "mapping" && (
        <div>
          <p className="gf-field-hint" style={{ marginBottom: 10 }}>
            Arquivo <b>{nomeArquivo}</b> — {dataRows.length} linha(s) encontrada(s). Confira o mapeamento das colunas antes de importar.
          </p>
          <div className="gf-form-grid gf-modal-form" style={{ marginBottom: 16 }}>
            {CAMPOS_ORCAMENTO.map((c) => (
              <Field label={`${c.label}${c.obrigatorio ? " *" : ""}`} key={c.key}>
                <select value={mapping[c.key]} onChange={(e) => setMapping((m) => ({ ...m, [c.key]: e.target.value === "" ? "" : Number(e.target.value) }))}>
                  <option value="">— não mapear —</option>
                  {headers.map((h, i) => <option key={i} value={i}>{String(h || `Coluna ${i + 1}`)}</option>)}
                </select>
              </Field>
            ))}
          </div>

          {!mappingCompleto && <div className="gf-warn" style={{ marginBottom: 12 }}>Mapeie ao menos Apropriação e Verba orçada / Valor total para continuar.</div>}

          <h4 className="gf-chart-title" style={{ marginTop: 0 }}>Prévia da importação</h4>
          <div className="gf-table-wrap" style={{ maxHeight: 240, marginBottom: 12 }}>
            <table className="gf-table">
              <thead><tr><th>Apropriação</th><th>Descrição</th><th className="num">Qtd.</th><th className="num">Vlr. unitário</th><th className="num">Verba orçada</th></tr></thead>
              <tbody>
                {previewItens.slice(0, 12).map((it, i) => (
                  <tr key={i}>
                    <td>{it.apropriacao}</td><td className="gf-td-desc">{it.descricao || "—"}</td>
                    <td className="num mono">{it.quantidade || "—"}</td>
                    <td className="num mono">{it.valorUnitario ? BRL(it.valorUnitario) : "—"}</td>
                    <td className="num mono">{BRL(it.valorTotal)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="gf-summary-inline" style={{ marginBottom: 14 }}>
            <span>Linhas a importar <b>{previewItens.length}</b></span>
            <span>Verba orçada total <b className="mono">{BRL(totalOrcado)}</b></span>
          </div>
          <div className="gf-warn" style={{ marginBottom: 14 }}>
            Isso substituirá qualquer orçamento já importado anteriormente para esta obra.
          </div>

          <div className="gf-form-actions">
            <button className="gf-btn gf-btn-ghost" onClick={() => setStep("upload")}>Voltar</button>
            <button className="gf-btn gf-btn-primary" disabled={!mappingCompleto || previewItens.length === 0} onClick={confirmar}>
              Confirmar importação
            </button>
          </div>
        </div>
      )}
    </Modal>
  );
}

function ApropriacaoOrcamentoModal({ linha, obra, contasPagar, allEntries, onClose }) {
  const contasRelacionadas = useMemo(() => contasPagar.filter((c) =>
    !c.cancelada && (c.rateio || []).some((r) => r.centroCusto === obra.id) && c.apropriacao === linha.apropriacao
  ), [contasPagar, obra.id, linha.apropriacao]);

  const lancamentosRelacionados = useMemo(() => allEntries.filter((e) =>
    e.tipo === "saida" && e.obraId === obra.id && e.apropriacao === linha.apropriacao
  ).sort((a, b) => (a.data < b.data ? 1 : -1)), [allEntries, obra.id, linha.apropriacao]);

  const fornecedoresEnvolvidos = [...new Set(contasRelacionadas.map((c) => c.fornecedor).filter(Boolean))];

  return (
    <Modal title={`${linha.apropriacao} — Orçamento ${BRL(linha.orcado)}`} onClose={onClose} width={720}>
      <div className="gf-account-rows gf-detail-stats" style={{ flexWrap: "wrap" }}>
        <div><span>Orçamento original</span><b className="mono">{BRL(linha.orcado)}</b></div>
        <div><span>Comprometido</span><b className="mono">{BRL(linha.comprometido)}</b></div>
        <div><span>Realizado (pago)</span><b className="mono pos">{BRL(linha.realizado)}</b></div>
        <div><span>Ainda a pagar</span><b className="mono">{BRL(linha.pendente)}</b></div>
        <div><span>Saldo disponível</span><b className={`mono ${linha.saldo >= 0 ? "pos" : "neg"}`}>{BRL(linha.saldo)}</b></div>
        <div><span>% utilizado</span><b className="mono">{linha.pct.toFixed(1)}%</b></div>
      </div>

      {fornecedoresEnvolvidos.length > 0 && (
        <div style={{ margin: "12px 0" }}>
          <span className="gf-field-label" style={{ display: "block", marginBottom: 6 }}>Fornecedores</span>
          <div className="gf-chip-row">{fornecedoresEnvolvidos.map((f) => <span className="gf-chip" key={f}>{f}</span>)}</div>
        </div>
      )}

      <h4 className="gf-chart-title">Contas a pagar vinculadas</h4>
      <div className="gf-table-wrap" style={{ maxHeight: 200, marginBottom: 14 }}>
        <table className="gf-table">
          <thead><tr><th>Fornecedor</th><th>Vencimento</th><th className="num">Valor</th><th>Status</th></tr></thead>
          <tbody>
            {contasRelacionadas.length === 0 && <tr><td colSpan={4}><EmptyState icon={CalendarClock} title="Nenhuma conta vinculada" body="Nenhuma conta a pagar lançada para esta apropriação ainda." /></td></tr>}
            {contasRelacionadas.map((c) => {
              const r = c.rateio.find((x) => x.centroCusto === obra.id);
              const status = deriveStatusConta(c);
              return (
                <tr key={c.id}>
                  <td>{c.fornecedor || "—"}</td><td>{fmtDate(c.dataVencimento)}</td>
                  <td className="num mono">{BRL(r?.valor || 0)}</td>
                  <td><Badge tone={STATUS_TONE[status]}>{status}</Badge></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <h4 className="gf-chart-title">Lançamentos realizados</h4>
      <div className="gf-table-wrap" style={{ maxHeight: 200 }}>
        <table className="gf-table">
          <thead><tr><th>Data</th><th>Fornecedor</th><th>Descrição</th><th className="num">Valor</th></tr></thead>
          <tbody>
            {lancamentosRelacionados.length === 0 && <tr><td colSpan={4}><EmptyState icon={Receipt} title="Nenhum lançamento" body="Nenhuma saída lançada para esta apropriação ainda." /></td></tr>}
            {lancamentosRelacionados.map((e) => (
              <tr key={e.id}>
                <td>{fmtDate(e.data)}</td><td>{e.fornecedor || "—"}</td>
                <td className="gf-td-desc">{e.descricao || "—"}</td>
                <td className="num mono neg">{BRL(e.valor)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Página: Fornecedores
--------------------------------------------------------------- */

function FornecedoresPage({ supplierNames, supplierStats, entries, obras }) {
  const [selected, setSelected] = useState(null);
  const obraName = (id) => (id === SEM_OBRA ? "Administrativo / Custos internos" : obras.find((o) => o.id === id)?.nome || "—");

  const rows = supplierNames
    .map((name) => ({ name, ...(supplierStats[name] || { total: 0, count: 0, obras: new Set(), last: null }) }))
    .filter((r) => r.count > 0)
    .sort((a, b) => b.total - a.total);

  return (
    <div className="gf-stack">
      <div className="gf-toolbar"><h2>Fornecedores</h2></div>

      {rows.length === 0 ? (
        <div className="gf-card"><EmptyState icon={Users} title="Nenhum fornecedor com pagamentos" body="Fornecedores aparecem aqui automaticamente ao lançar saídas com o campo Fornecedor preenchido." /></div>
      ) : (
        <section className="gf-card">
          <div className="gf-table-wrap">
            <table className="gf-table">
              <thead>
                <tr><th>Fornecedor</th><th className="num">Total pago</th><th className="num">Pagamentos</th><th>Obras relacionadas</th><th>Último pagamento</th><th></th></tr>
              </thead>
              <tbody>
                {rows.map((r) => (
                  <tr key={r.name} className="gf-row-clickable" onClick={() => setSelected(r.name)}>
                    <td><b>{r.name}</b></td>
                    <td className="num mono neg">{BRL(r.total)}</td>
                    <td className="num mono">{r.count}</td>
                    <td>
                      <div className="gf-chip-row">
                        {[...r.obras].slice(0, 2).map((oid) => <span className="gf-chip" key={oid}>{obraName(oid)}</span>)}
                        {r.obras.size > 2 && <span className="gf-chip">+{r.obras.size - 2}</span>}
                        {r.obras.size === 0 && <span className="gf-chip gf-chip-muted">—</span>}
                      </div>
                    </td>
                    <td>{fmtDate(r.last)}</td>
                    <td><ChevronRight size={15} className="gf-muted" /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {selected && (
        <FornecedorDetail
          nome={selected} entries={entries.filter((e) => e.fornecedor === selected && e.tipo === "saida")}
          obras={obras} onClose={() => setSelected(null)}
        />
      )}
    </div>
  );
}

function FornecedorDetail({ nome, entries, obras, onClose }) {
  const [filters, setFilters] = useState({ de: "", ate: "", obraId: "", apropriacao: "" });
  const obraName = (id) => (id === SEM_OBRA ? "Administrativo / Custos internos" : obras.find((o) => o.id === id)?.nome || "—");

  const filtered = entries
    .filter((e) => (filters.de ? e.data >= filters.de : true))
    .filter((e) => (filters.ate ? e.data <= filters.ate : true))
    .filter((e) => (filters.obraId ? e.obraId === filters.obraId : true))
    .filter((e) => (filters.apropriacao ? e.apropriacao === filters.apropriacao : true))
    .sort((a, b) => (a.data < b.data ? 1 : -1));

  const apropriacoesDisponiveis = [...new Set(entries.map((e) => e.apropriacao))];

  const totalPeriodo = filtered.reduce((s, e) => s + e.valor, 0);
  const porObra = useMemo(() => {
    const map = {};
    filtered.forEach((e) => { const k = obraName(e.obraId); map[k] = (map[k] || 0) + e.valor; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [filtered]);
  const porApropriacao = useMemo(() => {
    const map = {};
    filtered.forEach((e) => { map[e.apropriacao] = (map[e.apropriacao] || 0) + e.valor; });
    return Object.entries(map).sort((a, b) => b[1] - a[1]);
  }, [filtered]);

  const exportCSV = () => {
    const header = ["Data", "Centro de custo", "Apropriação", "Descrição", "Valor"];
    const lines = filtered.map((e) => [fmtDate(e.data), obraName(e.obraId), e.apropriacao, (e.descricao || "").replace(/,/g, ";"), e.valor.toFixed(2).replace(".", ",")]);
    const csv = [header, ...lines].map((r) => r.join(";")).join("\n");
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `extrato-${nome.replace(/\s+/g, "_")}.csv`; a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <Modal title={nome} onClose={onClose} width={720}>
      <div className="gf-filters-grid gf-fornecedor-filters">
        <input type="date" title="De" value={filters.de} onChange={(e) => setFilters((f) => ({ ...f, de: e.target.value }))} />
        <input type="date" title="Até" value={filters.ate} onChange={(e) => setFilters((f) => ({ ...f, ate: e.target.value }))} />
        <select value={filters.obraId} onChange={(e) => setFilters((f) => ({ ...f, obraId: e.target.value }))}>
          <option value="">Centro de custo (todos)</option>
          {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
        </select>
        <select value={filters.apropriacao} onChange={(e) => setFilters((f) => ({ ...f, apropriacao: e.target.value }))}>
          <option value="">Apropriação (todas)</option>
          {apropriacoesDisponiveis.map((a) => <option key={a} value={a}>{a}</option>)}
        </select>
        <button className="gf-btn gf-btn-ghost" onClick={exportCSV}><Download size={14} /> Exportar CSV</button>
      </div>

      <div className="gf-table-wrap" style={{ maxHeight: 260 }}>
        <table className="gf-table">
          <thead><tr><th>Data</th><th>Centro de custo</th><th>Apropriação</th><th>Descrição</th><th className="num">Valor</th></tr></thead>
          <tbody>
            {filtered.length === 0 && <tr><td colSpan={5}><EmptyState icon={Receipt} title="Nenhum pagamento no período" body="Ajuste os filtros para ver mais resultados." /></td></tr>}
            {filtered.map((e) => (
              <tr key={e.id}>
                <td>{fmtDate(e.data)}</td><td>{obraName(e.obraId)}</td><td>{e.apropriacao}</td>
                <td className="gf-td-desc">{e.descricao || "—"}</td>
                <td className="num mono neg">{BRL(e.valor)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="gf-fornecedor-summary">
        <div className="gf-summary-block">
          <span>Total pago no período</span>
          <strong className="mono neg">{BRL(totalPeriodo)}</strong>
        </div>
        <div className="gf-summary-block">
          <span>Total por obra</span>
          {porObra.map(([k, v]) => <div className="gf-summary-line" key={k}><span>{k}</span><b className="mono">{BRL(v)}</b></div>)}
        </div>
        <div className="gf-summary-block">
          <span>Total por apropriação</span>
          {porApropriacao.map(([k, v]) => <div className="gf-summary-line" key={k}><span>{k}</span><b className="mono">{BRL(v)}</b></div>)}
        </div>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Página: Contas a Pagar
--------------------------------------------------------------- */

function ContasPagarPage({ contasPagar, setContasPagar, entries, setEntries, accounts, obras, apropriacoes, supplierNames, saldoAtualTotal, orcamentoItens, currentUser, logAudit }) {
  const [showNew, setShowNew] = useState(false);
  const [baixaFor, setBaixaFor] = useState(null);
  const [grupoId, setGrupoId] = useState(null);
  const [editingConta, setEditingConta] = useState(null);
  const [rateioView, setRateioView] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [confirmCancelId, setConfirmCancelId] = useState(null);
  const [horizonte, setHorizonte] = useState(30);
  const [filters, setFilters] = useState({
    status: "", centroCusto: "", fornecedor: "", apropriacao: "", tipo: "", de: "", ate: "", min: "", max: "", q: "",
  });

  const podeCriar = can(currentUser, "contaspagar_criar");
  const podeEditar = can(currentUser, "contaspagar_editar");
  const podeRatear = can(currentUser, "contaspagar_ratear");
  const podeBaixar = can(currentUser, "contaspagar_baixar");
  const podeExcluir = can(currentUser, "contaspagar_excluir");

  const centroCustoName = (id) => (id === SEM_OBRA ? "Administrativo / Custos internos" : obras.find((o) => o.id === id)?.nome || "—");
  const accountName = (id) => accounts.find((a) => a.id === id)?.nome || "—";

  const withStatus = useMemo(() => contasPagar.map((c) => ({ ...c, statusCalc: deriveStatusConta(c) })), [contasPagar]);

  /* --------- ações --------- */

  const saveNovaConta = (novasContas) => {
    if (!podeCriar) return;
    setContasPagar((prev) => [...novasContas, ...prev]);
    setShowNew(false);
    logAudit?.({
      tipoOperacao: "Cadastro de conta a pagar", registroAlterado: novasContas[0]?.descricao || novasContas[0]?.fornecedor || "Conta a pagar",
      valorAnterior: "—", valorNovo: BRL(novasContas.reduce((s, c) => s + c.valorPrevisto, 0)),
    });
  };

  const saveEdicao = (updated, scope) => {
    if (!podeEditar) return;
    // Se o valor previsto mudou, recalcula o rateio proporcionalmente mantendo os mesmos percentuais.
    const aplicarValor = (c) => {
      const novoRateio = c.rateio
        ? computeRateioValores(c.rateio.map((r) => ({ centroCusto: r.centroCusto, percentual: r.percentual })), updated.valorPrevisto)
        : c.rateio;
      return { ...c, valorPrevisto: updated.valorPrevisto, apropriacao: updated.apropriacao, descricao: updated.descricao, rateio: novoRateio };
    };
    const valorAnterior = editingConta.valorPrevisto;
    if (scope === "todas" && editingConta.groupId) {
      setContasPagar((prev) => prev.map((c) =>
        c.groupId === editingConta.groupId && c.seq >= editingConta.seq && !c.pagamento && !c.cancelada
          ? aplicarValor(c)
          : c
      ));
    } else {
      setContasPagar((prev) => prev.map((c) => (c.id === editingConta.id ? { ...aplicarValor(c), dataVencimento: updated.dataVencimento } : c)));
    }
    if (valorAnterior !== updated.valorPrevisto) {
      logAudit?.({
        tipoOperacao: "Alteração de valor de conta a pagar", registroAlterado: updated.descricao || editingConta.fornecedor || "Conta a pagar",
        valorAnterior: BRL(valorAnterior), valorNovo: BRL(updated.valorPrevisto),
      });
    }
    setEditingConta(null);
  };

  const cancelarConta = (id) => {
    if (!podeEditar) return;
    setContasPagar((prev) => prev.map((c) => (c.id === id ? { ...c, cancelada: true } : c)));
    setConfirmCancelId(null);
  };
  const reativarConta = (id) => { if (podeEditar) setContasPagar((prev) => prev.map((c) => (c.id === id ? { ...c, cancelada: false } : c))); };
  const excluirConta = (id) => {
    if (!podeExcluir) return;
    const conta = contasPagar.find((c) => c.id === id);
    setContasPagar((prev) => prev.filter((c) => c.id !== id));
    setConfirmDeleteId(null);
    if (conta) logAudit?.({ tipoOperacao: "Exclusão de conta a pagar", registroAlterado: conta.descricao || conta.fornecedor || "Conta a pagar", valorAnterior: BRL(conta.valorPrevisto), valorNovo: "—" });
  };

  // Baixa de conta: gera 1 lançamento por centro de custo do rateio, ratando também
  // multa/juros/acréscimos/descontos proporcionalmente, mas mantendo um único pagamento bancário.
  const baixarConta = (conta, pagamentoBase) => {
    if (!podeBaixar) return;
    const acrescimoLiquido = Math.round(
      (pagamentoBase.multa + pagamentoBase.juros + pagamentoBase.outrosAcrescimos - pagamentoBase.descontos) * 100
    ) / 100;
    const rateioPago = rateiaAcrescimo(conta.rateio, acrescimoLiquido);

    const novosLancamentos = rateioPago.map((r) => ({
      id: uid("lan"), tipo: "saida", contaId: pagamentoBase.contaId, obraId: r.centroCusto,
      data: pagamentoBase.data, valor: r.valorFinal, fornecedor: conta.fornecedor, apropriacao: conta.apropriacao,
      descricao: conta.rateio.length > 1
        ? `${conta.descricao || "Baixa de conta a pagar"} (rateio ${r.percentual.toFixed(0)}% — ${centroCustoName(r.centroCusto)})`
        : `${conta.descricao || "Baixa de conta a pagar"} (baixa de conta a pagar)`,
      createdAt: Date.now(), contaPagarId: conta.id,
    }));

    const valorFinalTotal = Math.round((pagamentoBase.valorOriginal + acrescimoLiquido) * 100) / 100;
    setEntries((prev) => [...novosLancamentos, ...prev]);
    setContasPagar((prev) => prev.map((c) => (c.id === conta.id ? {
      ...c,
      pagamento: {
        ...pagamentoBase, valorFinal: valorFinalTotal,
        rateioPago: rateioPago.map((r, i) => ({ ...r, lancamentoId: novosLancamentos[i].id })),
      },
    } : c)));
    logAudit?.({
      tipoOperacao: "Registro de pagamento", registroAlterado: conta.descricao || conta.fornecedor || "Conta a pagar",
      valorAnterior: BRL(pagamentoBase.valorOriginal), valorNovo: BRL(valorFinalTotal),
      centroCusto: conta.rateio.length === 1 ? conta.rateio[0].centroCusto : `Rateado (${conta.rateio.length})`,
    });
    setBaixaFor(null);
  };

  const estornarBaixa = (conta) => {
    if (!podeBaixar) return;
    if (!window.confirm("Isso vai excluir o(s) lançamento(s) gerado(s) e reabrir esta conta. Continuar?")) return;
    const idsParaRemover = conta.pagamento?.rateioPago?.map((r) => r.lancamentoId).filter(Boolean)
      || (conta.pagamento?.lancamentoId ? [conta.pagamento.lancamentoId] : []);
    if (idsParaRemover.length) setEntries((prev) => prev.filter((e) => !idsParaRemover.includes(e.id)));
    setContasPagar((prev) => prev.map((c) => (c.id === conta.id ? { ...c, pagamento: null } : c)));
    logAudit?.({ tipoOperacao: "Estorno de pagamento", registroAlterado: conta.descricao || conta.fornecedor || "Conta a pagar", valorAnterior: BRL(conta.pagamento?.valorFinal || 0), valorNovo: "—" });
  };

  /* --------- indicadores (6.12) --------- */

  const today = todayISO();
  const proximos7 = addDaysISO(today, 7);
  const proximos30 = addDaysISO(today, 30);
  const curMonthKey = monthKey(today);

  const indicadores = useMemo(() => {
    const abertos = withStatus.filter((c) => !c.pagamento && !c.cancelada);
    const totalVencido = abertos.filter((c) => c.statusCalc === "Vencida").reduce((s, c) => s + c.valorPrevisto, 0);
    const venceHoje = abertos.filter((c) => c.dataVencimento === today).reduce((s, c) => s + c.valorPrevisto, 0);
    const prox7 = abertos.filter((c) => c.dataVencimento >= today && c.dataVencimento <= proximos7).reduce((s, c) => s + c.valorPrevisto, 0);
    const prox30 = abertos.filter((c) => c.dataVencimento >= today && c.dataVencimento <= proximos30).reduce((s, c) => s + c.valorPrevisto, 0);
    const totalAberto = abertos.reduce((s, c) => s + c.valorPrevisto, 0);
    const totalPagoMes = withStatus.filter((c) => c.pagamento && monthKey(c.pagamento.data) === curMonthKey).reduce((s, c) => s + (c.pagamento.valorFinal || 0), 0);
    return { totalVencido, venceHoje, prox7, prox30, totalAberto, totalPagoMes };
  }, [withStatus, today, proximos7, proximos30, curMonthKey]);

  /* --------- vencimentos da semana (6.9) --------- */

  const semana = useMemo(() => {
    const { monday, sunday } = weekRange();
    const abertos = withStatus.filter((c) => !c.pagamento && !c.cancelada);
    const daSemana = abertos.filter((c) => c.dataVencimento >= monday && c.dataVencimento <= sunday).sort((a, b) => (a.dataVencimento < b.dataVencimento ? -1 : 1));
    return {
      monday, sunday, itens: daSemana,
      totalContas: daSemana.length,
      valorTotal: daSemana.reduce((s, c) => s + c.valorPrevisto, 0),
      valorHoje: abertos.filter((c) => c.dataVencimento === today).reduce((s, c) => s + c.valorPrevisto, 0),
      valorVencido: abertos.filter((c) => c.dataVencimento < today).reduce((s, c) => s + c.valorPrevisto, 0),
    };
  }, [withStatus, today]);

  /* --------- filtros e listagem --------- */

  const filtered = withStatus
    .filter((c) => (filters.status ? c.statusCalc === filters.status : true))
    .filter((c) => (filters.centroCusto ? (c.rateio || []).some((r) => r.centroCusto === filters.centroCusto) : true))
    .filter((c) => (filters.fornecedor ? c.fornecedor === filters.fornecedor : true))
    .filter((c) => (filters.apropriacao ? c.apropriacao === filters.apropriacao : true))
    .filter((c) => (filters.tipo ? c.tipo === filters.tipo : true))
    .filter((c) => (filters.de ? c.dataVencimento >= filters.de : true))
    .filter((c) => (filters.ate ? c.dataVencimento <= filters.ate : true))
    .filter((c) => (filters.min ? c.valorPrevisto >= Number(filters.min) : true))
    .filter((c) => (filters.max ? c.valorPrevisto <= Number(filters.max) : true))
    .filter((c) => {
      if (!filters.q) return true;
      const q = filters.q.toLowerCase();
      return (c.fornecedor || "").toLowerCase().includes(q) || (c.descricao || "").toLowerCase().includes(q) || (c.numeroDocumento || "").toLowerCase().includes(q);
    })
    .sort((a, b) => (a.dataVencimento < b.dataVencimento ? -1 : a.dataVencimento > b.dataVencimento ? 1 : 0));

  /* --------- projeção de fluxo de caixa (6.10) --------- */

  const HORIZONTES = [
    { v: 7, l: "7 dias" }, { v: 15, l: "15 dias" }, { v: 30, l: "30 dias" },
    { v: 60, l: "60 dias" }, { v: 90, l: "90 dias" }, { v: 365, l: "12 meses" },
  ];

  const projecao = useMemo(() => {
    const limite = addDaysISO(today, horizonte);
    const pendentes = withStatus.filter((c) => !c.pagamento && !c.cancelada && c.dataVencimento <= limite);
    const byDate = {};
    pendentes.forEach((c) => { byDate[c.dataVencimento] = (byDate[c.dataVencimento] || 0) + c.valorPrevisto; });
    const dates = Object.keys(byDate).sort();
    let running = saldoAtualTotal;
    const points = [{ data: "Hoje", saldo: running }];
    dates.forEach((d) => { running -= byDate[d]; points.push({ data: fmtDate(d), saldo: running }); });
    const totalPrevisto = pendentes.reduce((s, c) => s + c.valorPrevisto, 0);
    return { points, totalPrevisto, saldoProjetado: saldoAtualTotal - totalPrevisto };
  }, [withStatus, horizonte, saldoAtualTotal, today]);

  const grupo = grupoId ? withStatus.filter((c) => c.groupId === grupoId).sort((a, b) => a.seq - b.seq) : null;

  return (
    <div className="gf-stack">
      <div className="gf-toolbar">
        <h2>Contas a pagar</h2>
        {podeCriar && <button className="gf-btn gf-btn-primary" onClick={() => setShowNew(true)}><Plus size={15} /> Nova conta a pagar</button>}
      </div>

      <div className="gf-tiles gf-tiles-cp">
        <div className="gf-tile tone-red"><span className="gf-tile-label">Total vencido</span><span className="gf-tile-value neg">{BRL(indicadores.totalVencido)}</span></div>
        <div className="gf-tile tone-orange"><span className="gf-tile-label">Vence hoje</span><span className="gf-tile-value">{BRL(indicadores.venceHoje)}</span></div>
        <div className="gf-tile tone-yellow"><span className="gf-tile-label">Próximos 7 dias</span><span className="gf-tile-value">{BRL(indicadores.prox7)}</span></div>
        <div className="gf-tile tone-blue"><span className="gf-tile-label">Próximos 30 dias</span><span className="gf-tile-value">{BRL(indicadores.prox30)}</span></div>
        <div className="gf-tile"><span className="gf-tile-label">Total em aberto</span><span className="gf-tile-value">{BRL(indicadores.totalAberto)}</span></div>
        <div className="gf-tile tone-green"><span className="gf-tile-label">Total pago no mês</span><span className="gf-tile-value pos">{BRL(indicadores.totalPagoMes)}</span></div>
      </div>

      <section className="gf-card">
        <div className="gf-card-head">
          <h2>Vencimentos da semana</h2>
          <span className="gf-eyebrow">{fmtDate(semana.monday)} — {fmtDate(semana.sunday)}</span>
        </div>
        <div className="gf-summary-inline" style={{ marginBottom: 12 }}>
          <span>Contas a vencer <b>{semana.totalContas}</b></span>
          <span>Valor total a vencer <b>{BRL(semana.valorTotal)}</b></span>
          <span>Vencendo hoje <b>{BRL(semana.valorHoje)}</b></span>
          <span>Vencido (geral) <b className="neg">{BRL(semana.valorVencido)}</b></span>
        </div>
        {semana.itens.length === 0 ? (
          <EmptyState icon={CalendarClock} title="Nada vencendo esta semana" body="Nenhuma conta em aberto com vencimento entre segunda e domingo." />
        ) : (
          <div className="gf-table-wrap">
            <table className="gf-table">
              <thead><tr><th>Vencimento</th><th>Fornecedor</th><th>Centro de custo</th><th>Apropriação</th><th>Descrição</th><th className="num">Valor</th><th>Status</th></tr></thead>
              <tbody>
                {semana.itens.map((c) => (
                  <tr key={c.id}>
                    <td>{fmtDate(c.dataVencimento)}</td><td>{c.fornecedor || "—"}</td><td>{centroCustoName(c.centroCusto)}</td>
                    <td>{c.apropriacao}</td><td className="gf-td-desc">{c.descricao}</td>
                    <td className="num mono">{BRL(c.valorPrevisto)}</td>
                    <td><Badge tone={STATUS_TONE[c.statusCalc]}>{c.statusCalc}</Badge></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </section>

      <section className="gf-card">
        <div className="gf-card-head">
          <h2>Projeção de fluxo de caixa</h2>
          <select value={horizonte} onChange={(e) => setHorizonte(Number(e.target.value))} style={{ maxWidth: 160 }}>
            {HORIZONTES.map((h) => <option key={h.v} value={h.v}>Próximos {h.l}</option>)}
          </select>
        </div>
        <div className="gf-summary-inline" style={{ marginBottom: 10 }}>
          <span>Saldo atual <b>{BRL(saldoAtualTotal)}</b></span>
          <span>Contas a pagar no período <b className="neg">{BRL(projecao.totalPrevisto)}</b></span>
          <span>Saldo projetado <b className={projecao.saldoProjetado >= 0 ? "pos" : "neg"}>{BRL(projecao.saldoProjetado)}</b></span>
        </div>
        <div className="gf-chart-box">
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={projecao.points} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
              <CartesianGrid stroke="var(--gf-border)" vertical={false} />
              <XAxis dataKey="data" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
              <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
              <Line type="stepAfter" dataKey="saldo" stroke="var(--gf-primary)" strokeWidth={2.2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
        <p className="gf-field-hint">Considera o saldo atual em bancos menos as contas a pagar ainda não quitadas. Preparado para futuramente somar entradas previstas (Contas a Receber).</p>
      </section>

      <section className="gf-card">
        <div className="gf-card-head"><h2>Todas as contas</h2></div>
        <div className="gf-filters">
          <div className="gf-search">
            <Search size={15} />
            <input placeholder="Pesquisar por fornecedor, descrição ou número do documento…" value={filters.q} onChange={(e) => setFilters((f) => ({ ...f, q: e.target.value }))} />
          </div>
          <div className="gf-filters-grid gf-filters-grid-cp">
            <input type="date" title="Vencimento de" value={filters.de} onChange={(e) => setFilters((f) => ({ ...f, de: e.target.value }))} />
            <input type="date" title="Vencimento até" value={filters.ate} onChange={(e) => setFilters((f) => ({ ...f, ate: e.target.value }))} />
            <select value={filters.status} onChange={(e) => setFilters((f) => ({ ...f, status: e.target.value }))}>
              <option value="">Status (todos)</option>
              {["A vencer", "Vencendo hoje", "Vencida", "Paga", "Cancelada"].map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
            <select value={filters.tipo} onChange={(e) => setFilters((f) => ({ ...f, tipo: e.target.value }))}>
              <option value="">Tipo (todos)</option>
              <option value="unica">Única</option><option value="parcelada">Parcelada</option><option value="recorrente">Recorrente</option>
            </select>
            <select value={filters.centroCusto} onChange={(e) => setFilters((f) => ({ ...f, centroCusto: e.target.value }))}>
              <option value="">Centro de custo (todos)</option>
              <option value={SEM_OBRA}>Administrativo / Custos internos</option>
              {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
            </select>
            <select value={filters.fornecedor} onChange={(e) => setFilters((f) => ({ ...f, fornecedor: e.target.value }))}>
              <option value="">Fornecedor (todos)</option>
              {supplierNames.map((f) => <option key={f} value={f}>{f}</option>)}
            </select>
            <select value={filters.apropriacao} onChange={(e) => setFilters((f) => ({ ...f, apropriacao: e.target.value }))}>
              <option value="">Apropriação (todas)</option>
              {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
            <input type="number" placeholder="Valor mín." value={filters.min} onChange={(e) => setFilters((f) => ({ ...f, min: e.target.value }))} />
            <input type="number" placeholder="Valor máx." value={filters.max} onChange={(e) => setFilters((f) => ({ ...f, max: e.target.value }))} />
          </div>
        </div>

        <div className="gf-table-wrap">
          <table className="gf-table gf-table-ledger">
            <thead>
              <tr>
                <th>Vencimento</th><th>Status</th><th>Fornecedor</th><th>Centro de custo</th><th>Apropriação</th>
                <th>Parcela</th><th>Descrição</th><th className="num">Valor previsto</th><th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr><td colSpan={9}><EmptyState icon={CalendarClock} title="Nenhuma conta encontrada" body="Ajuste os filtros ou cadastre uma nova conta a pagar." /></td></tr>
              )}
              {filtered.map((c) => (
                <tr key={c.id} className={c.statusCalc === "Vencida" ? "row-neg" : c.statusCalc === "Paga" ? "row-pos" : ""}>
                  <td>{fmtDate(c.dataVencimento)}</td>
                  <td><Badge tone={STATUS_TONE[c.statusCalc]}>{c.statusCalc}</Badge></td>
                  <td>{c.fornecedor || "—"}</td>
                  <td>
                    {c.rateio && c.rateio.length > 1 ? (
                      <button className="gf-rateio-pill" onClick={() => setRateioView(c)} title="Ver rateio entre centros de custo">
                        <Layers size={12} /> Rateado ({c.rateio.length})
                      </button>
                    ) : (
                      centroCustoName(c.rateio?.[0]?.centroCusto ?? c.centroCusto)
                    )}
                  </td>
                  <td>{c.apropriacao}</td>
                  <td>{c.tipo === "unica" ? "—" : `${c.seq}/${c.total}`}</td>
                  <td className="gf-td-desc" title={c.descricao}>{c.descricao || "—"}</td>
                  <td className="num mono">{BRL(c.valorPrevisto)}</td>
                  <td>
                    <div className="gf-row-actions">
                      {confirmDeleteId === c.id ? (
                        <>
                          <button className="gf-icon-btn danger" title="Confirmar exclusão" onClick={() => excluirConta(c.id)}><Check size={14} /></button>
                          <button className="gf-icon-btn" title="Cancelar" onClick={() => setConfirmDeleteId(null)}><X size={14} /></button>
                        </>
                      ) : confirmCancelId === c.id ? (
                        <>
                          <button className="gf-icon-btn danger" title="Confirmar cancelamento" onClick={() => cancelarConta(c.id)}><Check size={14} /></button>
                          <button className="gf-icon-btn" title="Voltar" onClick={() => setConfirmCancelId(null)}><X size={14} /></button>
                        </>
                      ) : (
                        <>
                          {c.total > 1 && (
                            <button className="gf-icon-btn" title="Ver grupo" onClick={() => setGrupoId(c.groupId)}><Layers size={14} /></button>
                          )}
                          {!c.pagamento && !c.cancelada && podeBaixar && (
                            <button className="gf-icon-btn" title="Baixar conta" onClick={() => setBaixaFor(c)}><CircleDollarSign size={14} /></button>
                          )}
                          {c.pagamento && podeBaixar && (
                            <button className="gf-icon-btn" title="Estornar baixa" onClick={() => estornarBaixa(c)}><Undo2 size={14} /></button>
                          )}
                          {!c.pagamento && !c.cancelada && podeEditar && (
                            <button className="gf-icon-btn" title="Editar" onClick={() => setEditingConta(c)}><Pencil size={14} /></button>
                          )}
                          {!c.pagamento && !c.cancelada && podeEditar && (
                            <button className="gf-icon-btn" title="Cancelar conta" onClick={() => setConfirmCancelId(c.id)}><X size={14} /></button>
                          )}
                          {c.cancelada && podeEditar && (
                            <button className="gf-icon-btn" title="Reativar" onClick={() => reativarConta(c.id)}><Undo2 size={14} /></button>
                          )}
                          {podeExcluir && <button className="gf-icon-btn" title="Excluir" onClick={() => setConfirmDeleteId(c.id)}><Trash2 size={14} /></button>}
                        </>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {showNew && (
        <NovaContaPagarModal
          obras={obras} apropriacoes={apropriacoes} supplierNames={supplierNames} podeRatear={podeRatear}
          orcamentoItens={orcamentoItens} contasPagar={contasPagar} allEntries={entries}
          onClose={() => setShowNew(false)} onSave={saveNovaConta}
        />
      )}
      {baixaFor && (
        <BaixaModal conta={baixaFor} accounts={accounts} centroCustoName={centroCustoName} onClose={() => setBaixaFor(null)} onSave={(pag) => baixarConta(baixaFor, pag)} />
      )}
      {editingConta && (
        <EditContaModal conta={editingConta} apropriacoes={apropriacoes} onClose={() => setEditingConta(null)} onSave={saveEdicao} />
      )}
      {grupo && (
        <GrupoModal
          itens={grupo} centroCustoName={centroCustoName} onClose={() => setGrupoId(null)}
          onBaixar={(c) => { setGrupoId(null); setBaixaFor(c); }}
          onEditar={(c) => { setGrupoId(null); setEditingConta(c); }}
        />
      )}
      {rateioView && (
        <RateioViewModal conta={rateioView} centroCustoName={centroCustoName} onClose={() => setRateioView(null)} />
      )}
    </div>
  );
}

function NovaContaPagarModal({ obras, apropriacoes, supplierNames, podeRatear = true, orcamentoItens = [], contasPagar = [], allEntries = [], onClose, onSave }) {
  const [tipo, setTipo] = useState("unica");
  const [f, setF] = useState({
    centroCusto: SEM_OBRA, fornecedor: "", apropriacao: apropriacoes[0] || "", descricao: "",
    numeroDocumento: "", observacao: "",
    dataCompetencia: todayISO(), dataVencimento: todayISO(), valorPrevisto: "",
    valorTotal: "", numeroParcelas: 2, periodicidadeParcela: "mensal", intervaloDias: 30,
    valorRecorrente: "", periodicidadeRecorrente: "mensal", modoFim: "quantidade", quantidadeOcorrencias: 12, dataFinal: "", diaPreferencial: "",
  });
  const upd = (p) => setF((s) => ({ ...s, ...p }));

  const [rateioAtivo, setRateioAtivo] = useState(false);
  const [rateioRows, setRateioRows] = useState([
    { id: uid("rr"), centroCusto: SEM_OBRA, valor: "" },
    { id: uid("rr"), centroCusto: SEM_OBRA, valor: "" },
  ]);
  const updRateioRow = (id, patch) => setRateioRows((prev) => prev.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const addRateioRow = () => setRateioRows((prev) => [...prev, { id: uid("rr"), centroCusto: SEM_OBRA, valor: "" }]);
  const removeRateioRow = (id) => setRateioRows((prev) => (prev.length > 2 ? prev.filter((r) => r.id !== id) : prev));

  const valorReferencia = tipo === "unica" ? Number(f.valorPrevisto) || 0
    : tipo === "parcelada" ? Number(f.valorTotal) || 0
    : Number(f.valorRecorrente) || 0;

  const somaRateio = rateioRows.reduce((s, r) => s + (Number(r.valor) || 0), 0);
  const diferencaRateio = Math.round((valorReferencia - somaRateio) * 100) / 100;
  const rateioFecha = Math.abs(diferencaRateio) < 0.01;
  const rateioCompleto = rateioRows.length >= 2 && rateioRows.every((r) => r.centroCusto && Number(r.valor) > 0);
  const rateioValido = !rateioAtivo || (rateioFecha && rateioCompleto && valorReferencia > 0);

  const base = () => ({
    fornecedor: f.fornecedor.trim(), apropriacao: f.apropriacao,
    descricao: f.descricao.trim(), numeroDocumento: f.numeroDocumento.trim(), observacao: f.observacao.trim(),
  });

  const rateioParaValor = (valorDaConta) => {
    if (!rateioAtivo) return [{ centroCusto: f.centroCusto, percentual: 100, valor: valorDaConta }];
    const percentuais = rateioRows.map((r) => ({ centroCusto: r.centroCusto, percentual: (Number(r.valor) / valorReferencia) * 100 }));
    return computeRateioValores(percentuais, valorDaConta);
  };

  const montaContaComRateio = (campos, valorDaConta) => {
    const rateio = rateioParaValor(valorDaConta);
    return { ...campos, valorPrevisto: valorDaConta, rateio, centroCusto: rateio.length === 1 ? rateio[0].centroCusto : null };
  };

  const totalOcorrenciasRecorrente = () => {
    if (tipo !== "recorrente") return 1;
    if (f.modoFim === "quantidade") return Math.max(1, Number(f.quantidadeOcorrencias) || 1);
    let total = 0; let venc = f.dataVencimento;
    while (venc <= f.dataFinal && total < 240) { total += 1; venc = addPeriodo(venc, f.periodicidadeRecorrente); }
    return Math.max(1, total);
  };

  const centroCustoName = (id) => (id === SEM_OBRA ? "Administrativo / Custos internos" : obras.find((o) => o.id === id)?.nome || "—");

  const alertasOrcamento = useMemo(() => {
    if (valorReferencia <= 0 || !f.apropriacao) return [];
    const valorTotalAgora = tipo === "recorrente" ? valorReferencia * totalOcorrenciasRecorrente() : valorReferencia;
    const linhas = rateioAtivo
      ? (rateioFecha && rateioCompleto ? computeRateioValores(rateioRows.map((r) => ({ centroCusto: r.centroCusto, percentual: (Number(r.valor) / valorReferencia) * 100 })), valorTotalAgora) : [])
      : [{ centroCusto: f.centroCusto, valor: valorTotalAgora }];
    return linhas
      .map((l) => ({ centroCusto: l.centroCusto, valor: l.valor, alerta: checarAlertaOrcamento(l.centroCusto, f.apropriacao, l.valor, orcamentoItens, contasPagar, allEntries) }))
      .filter((l) => l.alerta);
  }, [valorReferencia, tipo, f.apropriacao, f.centroCusto, f.modoFim, f.quantidadeOcorrencias, f.dataFinal, f.dataVencimento, f.periodicidadeRecorrente, rateioAtivo, rateioRows, rateioFecha, rateioCompleto, orcamentoItens, contasPagar, allEntries]);

  const submit = (e) => {
    e.preventDefault();
    if (!rateioValido) return;

    if (tipo === "unica") {
      if (!f.valorPrevisto || Number(f.valorPrevisto) <= 0 || !f.dataVencimento) return;
      onSave([{
        id: uid("cp"), groupId: uid("grp"), seq: 1, total: 1, tipo: "unica", ...base(),
        ...montaContaComRateio({ dataCompetencia: f.dataCompetencia, dataVencimento: f.dataVencimento }, Number(f.valorPrevisto)),
        pagamento: null, cancelada: false, createdAt: Date.now(),
      }]);
    } else if (tipo === "parcelada") {
      const total = Number(f.numeroParcelas) || 0;
      const valorTotal = Number(f.valorTotal) || 0;
      if (total < 2 || valorTotal <= 0 || !f.dataVencimento) return;
      const parcela = Math.round((valorTotal / total) * 100) / 100;
      const grpId = uid("grp");
      const list = [];
      let venc = f.dataVencimento;
      for (let i = 1; i <= total; i++) {
        const valor = i === total ? Math.round((valorTotal - parcela * (total - 1)) * 100) / 100 : parcela;
        list.push({
          id: uid("cp"), groupId: grpId, seq: i, total, tipo: "parcelada", ...base(),
          ...montaContaComRateio({ dataCompetencia: venc, dataVencimento: venc }, valor),
          pagamento: null, cancelada: false, createdAt: Date.now(),
        });
        venc = addPeriodo(venc, f.periodicidadeParcela, f.intervaloDias);
      }
      onSave(list);
    } else {
      const valor = Number(f.valorRecorrente) || 0;
      if (valor <= 0 || !f.dataVencimento) return;
      let total;
      if (f.modoFim === "quantidade") total = Math.max(1, Number(f.quantidadeOcorrencias) || 1);
      else {
        total = 0;
        let venc = f.dataVencimento;
        while (venc <= f.dataFinal && total < 240) { total += 1; venc = addPeriodo(venc, f.periodicidadeRecorrente); }
        total = Math.max(1, total);
      }
      const grpId = uid("grp");
      const list = [];
      let venc = f.dataVencimento;
      for (let i = 1; i <= total; i++) {
        list.push({
          id: uid("cp"), groupId: grpId, seq: i, total, tipo: "recorrente", ...base(),
          ...montaContaComRateio({ dataCompetencia: venc, dataVencimento: venc }, valor),
          pagamento: null, cancelada: false, createdAt: Date.now(),
        });
        venc = addPeriodo(venc, f.periodicidadeRecorrente);
      }
      onSave(list);
    }
  };

  return (
    <Modal title="Nova conta a pagar" onClose={onClose} width={680}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <div className="gf-segmented gf-span-2 gf-segmented-3">
          <button type="button" className={tipo === "unica" ? "is-active" : ""} onClick={() => setTipo("unica")}>Única</button>
          <button type="button" className={tipo === "parcelada" ? "is-active" : ""} onClick={() => setTipo("parcelada")}>Parcelada</button>
          <button type="button" className={tipo === "recorrente" ? "is-active" : ""} onClick={() => setTipo("recorrente")}>Recorrente</button>
        </div>

        <Field label="Fornecedor">
          <input list="gf-cp-fornecedores" value={f.fornecedor} onChange={(e) => upd({ fornecedor: e.target.value })} placeholder="Digite ou selecione…" />
          <datalist id="gf-cp-fornecedores">{supplierNames.map((s) => <option key={s} value={s} />)}</datalist>
        </Field>
        <Field label="Apropriação">
          <select value={f.apropriacao} onChange={(e) => upd({ apropriacao: e.target.value })}>
            {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
        </Field>
        <Field label="Número / documento" hint="Opcional"><input value={f.numeroDocumento} onChange={(e) => upd({ numeroDocumento: e.target.value })} /></Field>
        <Field label="Descrição"><input value={f.descricao} onChange={(e) => upd({ descricao: e.target.value })} /></Field>
        <Field label="Observação" className="gf-span-2" hint="Opcional"><input value={f.observacao} onChange={(e) => upd({ observacao: e.target.value })} /></Field>

        {tipo === "unica" && (
          <>
            <Field label="Data de competência"><input type="date" value={f.dataCompetencia} onChange={(e) => upd({ dataCompetencia: e.target.value })} /></Field>
            <Field label="Data de vencimento"><input type="date" required value={f.dataVencimento} onChange={(e) => upd({ dataVencimento: e.target.value })} /></Field>
            <Field label="Valor previsto (R$)" className="gf-span-2"><input type="number" min="0.01" step="0.01" required value={f.valorPrevisto} onChange={(e) => upd({ valorPrevisto: e.target.value })} /></Field>
          </>
        )}

        {tipo === "parcelada" && (
          <>
            <Field label="Valor total (R$)"><input type="number" min="0.01" step="0.01" required value={f.valorTotal} onChange={(e) => upd({ valorTotal: e.target.value })} /></Field>
            <Field label="Número de parcelas"><input type="number" min="2" step="1" required value={f.numeroParcelas} onChange={(e) => upd({ numeroParcelas: e.target.value })} /></Field>
            <Field label="Data do 1º vencimento"><input type="date" required value={f.dataVencimento} onChange={(e) => upd({ dataVencimento: e.target.value })} /></Field>
            <Field label="Periodicidade">
              <select value={f.periodicidadeParcela} onChange={(e) => upd({ periodicidadeParcela: e.target.value })}>
                {PERIODICIDADES_PARCELA.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
              </select>
            </Field>
            {f.periodicidadeParcela === "personalizada" && (
              <Field label="Intervalo entre parcelas (dias)" className="gf-span-2">
                <input type="number" min="1" value={f.intervaloDias} onChange={(e) => upd({ intervaloDias: e.target.value })} />
              </Field>
            )}
            {Number(f.valorTotal) > 0 && Number(f.numeroParcelas) > 1 && (
              <div className="gf-hint gf-span-2">{f.numeroParcelas}x de aproximadamente {BRL(Number(f.valorTotal) / Number(f.numeroParcelas))}</div>
            )}
          </>
        )}

        {tipo === "recorrente" && (
          <>
            <Field label="Valor por ocorrência (R$)"><input type="number" min="0.01" step="0.01" required value={f.valorRecorrente} onChange={(e) => upd({ valorRecorrente: e.target.value })} /></Field>
            <Field label="Data do 1º vencimento"><input type="date" required value={f.dataVencimento} onChange={(e) => upd({ dataVencimento: e.target.value })} /></Field>
            <Field label="Periodicidade">
              <select value={f.periodicidadeRecorrente} onChange={(e) => upd({ periodicidadeRecorrente: e.target.value })}>
                {PERIODICIDADES_RECORRENTE.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
              </select>
            </Field>
            <Field label="Dia preferencial de vencimento" hint="Opcional"><input type="number" min="1" max="31" value={f.diaPreferencial} onChange={(e) => upd({ diaPreferencial: e.target.value })} /></Field>
            <div className="gf-segmented gf-span-2">
              <button type="button" className={f.modoFim === "quantidade" ? "is-active" : ""} onClick={() => upd({ modoFim: "quantidade" })}>Por quantidade</button>
              <button type="button" className={f.modoFim === "data" ? "is-active" : ""} onClick={() => upd({ modoFim: "data" })}>Por data final</button>
            </div>
            {f.modoFim === "quantidade" ? (
              <Field label="Quantidade de ocorrências" className="gf-span-2"><input type="number" min="1" value={f.quantidadeOcorrencias} onChange={(e) => upd({ quantidadeOcorrencias: e.target.value })} /></Field>
            ) : (
              <Field label="Data final da recorrência" className="gf-span-2"><input type="date" value={f.dataFinal} onChange={(e) => upd({ dataFinal: e.target.value })} /></Field>
            )}
          </>
        )}

        <div className="gf-span-2" style={{ borderTop: "1px solid var(--gf-border)", paddingTop: 14, marginTop: 4 }}>
          {podeRatear && (
            <div className="gf-rateio-toggle">
              <label className="gf-checkbox">
                <input type="checkbox" checked={rateioAtivo} onChange={(e) => setRateioAtivo(e.target.checked)} />
                <span>Ratear esta conta entre múltiplos centros de custo</span>
              </label>
            </div>
          )}

          {!rateioAtivo ? (
            <Field label="Centro de custo">
              <select value={f.centroCusto} onChange={(e) => upd({ centroCusto: e.target.value })}>
                <option value={SEM_OBRA}>Administrativo / Custos internos</option>
                {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
              </select>
            </Field>
          ) : (
            <div className="gf-rateio-box">
              {valorReferencia <= 0 && <p className="gf-hint">Informe o valor da conta acima para poder ratear.</p>}
              {rateioRows.map((r, i) => {
                const pct = valorReferencia > 0 ? ((Number(r.valor) || 0) / valorReferencia) * 100 : 0;
                return (
                  <div className="gf-rateio-row" key={r.id}>
                    <select value={r.centroCusto} onChange={(e) => updRateioRow(r.id, { centroCusto: e.target.value })}>
                      <option value={SEM_OBRA}>Administrativo / Custos internos</option>
                      {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
                    </select>
                    <input type="number" min="0" step="0.01" placeholder="Valor (R$)" value={r.valor} onChange={(e) => updRateioRow(r.id, { valor: e.target.value })} />
                    <span className="gf-rateio-pct mono">{pct.toFixed(1)}%</span>
                    {rateioRows.length > 2 && (
                      <button type="button" className="gf-icon-btn danger" onClick={() => removeRateioRow(r.id)}><Trash2 size={13} /></button>
                    )}
                  </div>
                );
              })}
              <button type="button" className="gf-btn gf-btn-ghost" style={{ marginTop: 4 }} onClick={addRateioRow}>
                <Plus size={14} /> Adicionar centro de custo
              </button>
              <div className={`gf-rateio-total ${rateioFecha ? "is-ok" : "is-warn"}`}>
                <span>Total rateado: {BRL(somaRateio)} de {BRL(valorReferencia)}</span>
                {rateioFecha
                  ? <b className="pos">Rateio fecha 100% ✓</b>
                  : <b className="neg">{diferencaRateio > 0 ? `Falta ratear ${BRL(diferencaRateio)}` : `Excedeu em ${BRL(-diferencaRateio)}`}</b>}
              </div>
            </div>
          )}
        </div>

        {alertasOrcamento.length > 0 && (
          <div className="gf-span-2">
            {alertasOrcamento.map((l, i) => (
              <div className={`gf-budget-alert ${l.alerta.ultrapassa ? "" : "is-atencao"}`} key={i}>
                <AlertTriangle size={16} />
                <div>
                  <strong>
                    {l.alerta.ultrapassa
                      ? `Atenção: este lançamento fará "${f.apropriacao}" (${centroCustoName(l.centroCusto)}) ultrapassar o orçamento em ${BRL(l.alerta.estouro)}.`
                      : `Atenção: "${f.apropriacao}" (${centroCustoName(l.centroCusto)}) ficará em ${l.alerta.pctNovo.toFixed(0)}% do orçamento após este lançamento.`}
                  </strong>
                  <span>Orçado: {BRL(l.alerta.orcado)} · Comprometido antes: {BRL(l.alerta.comprometidoAtual)} · Novo comprometido: {BRL(l.alerta.novoComprometido)}</span>
                </div>
              </div>
            ))}
            <p className="gf-field-hint" style={{ marginTop: 6 }}>O lançamento não será bloqueado — você pode prosseguir mesmo assim; o estouro fica registrado no Controle Orçamentário da obra.</p>
          </div>
        )}

        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary" disabled={!rateioValido}>Cadastrar conta a pagar</button>
          <span className="gf-hint">Não gera saída financeira até que seja baixada como paga.</span>
        </div>
      </form>
    </Modal>
  );
}

function BaixaModal({ conta, accounts, centroCustoName, onClose, onSave }) {
  const [f, setF] = useState({
    data: todayISO(), contaId: accounts[0]?.id || "", multa: "0", juros: "0", outrosAcrescimos: "0", descontos: "0",
  });
  const upd = (p) => setF((s) => ({ ...s, ...p }));
  const valorOriginal = conta.valorPrevisto;
  const acrescimoLiquido = (Number(f.multa) || 0) + (Number(f.juros) || 0) + (Number(f.outrosAcrescimos) || 0) - (Number(f.descontos) || 0);
  const valorFinal = valorOriginal + acrescimoLiquido;
  const rateioPreview = conta.rateio && conta.rateio.length > 1 ? rateiaAcrescimo(conta.rateio, Math.round(acrescimoLiquido * 100) / 100) : null;

  const submit = (e) => {
    e.preventDefault();
    if (!f.contaId || !f.data) return;
    onSave({
      data: f.data, contaId: f.contaId, valorOriginal,
      multa: Number(f.multa) || 0, juros: Number(f.juros) || 0,
      outrosAcrescimos: Number(f.outrosAcrescimos) || 0, descontos: Number(f.descontos) || 0,
      valorFinal: Math.max(0, Math.round(valorFinal * 100) / 100),
    });
  };

  return (
    <Modal title="Baixar conta" onClose={onClose}>
      <div className="gf-detail-meta" style={{ gridTemplateColumns: "repeat(2,1fr)" }}>
        <div><span>Fornecedor</span><b>{conta.fornecedor || "—"}</b></div>
        <div><span>Descrição</span><b>{conta.descricao || "—"}</b></div>
        <div><span>Vencimento</span><b>{fmtDate(conta.dataVencimento)}</b></div>
        <div><span>Valor previsto</span><b className="mono">{BRL(conta.valorPrevisto)}</b></div>
      </div>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form" style={{ marginTop: 14 }}>
        <Field label="Data efetiva do pagamento"><input type="date" required value={f.data} onChange={(e) => upd({ data: e.target.value })} /></Field>
        <Field label="Conta bancária utilizada">
          <select required value={f.contaId} onChange={(e) => upd({ contaId: e.target.value })}>
            <option value="" disabled>Selecione…</option>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
          </select>
        </Field>
        <Field label="Multa (R$)"><input type="number" step="0.01" value={f.multa} onChange={(e) => upd({ multa: e.target.value })} /></Field>
        <Field label="Juros (R$)"><input type="number" step="0.01" value={f.juros} onChange={(e) => upd({ juros: e.target.value })} /></Field>
        <Field label="Outros acréscimos (R$)"><input type="number" step="0.01" value={f.outrosAcrescimos} onChange={(e) => upd({ outrosAcrescimos: e.target.value })} /></Field>
        <Field label="Descontos (R$)"><input type="number" step="0.01" value={f.descontos} onChange={(e) => upd({ descontos: e.target.value })} /></Field>
        <div className="gf-account-balance gf-span-2">
          <span>Valor final pago</span>
          <strong className="mono">{BRL(valorFinal)}</strong>
        </div>

        {rateioPreview && (
          <div className="gf-span-2">
            <span className="gf-field-label" style={{ display: "block", marginBottom: 6 }}>Distribuição por centro de custo (rateio proporcional)</span>
            <div className="gf-table-wrap">
              <table className="gf-table">
                <thead><tr><th>Centro de custo</th><th className="num">%</th><th className="num">Valor original</th><th className="num">Acréscimo/desc.</th><th className="num">Valor final</th></tr></thead>
                <tbody>
                  {rateioPreview.map((r) => (
                    <tr key={r.centroCusto}>
                      <td>{centroCustoName(r.centroCusto)}</td>
                      <td className="num mono">{r.percentual.toFixed(1)}%</td>
                      <td className="num mono">{BRL(r.valor)}</td>
                      <td className={`num mono ${r.acrescimo >= 0 ? "" : "pos"}`}>{r.acrescimo >= 0 ? "+" : ""}{BRL(r.acrescimo)}</td>
                      <td className="num mono">{BRL(r.valorFinal)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="gf-field-hint">O boleto continua sendo um único pagamento bancário; o sistema apenas gera um lançamento por centro de custo, proporcional ao rateio original.</p>
          </div>
        )}

        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">Confirmar pagamento</button>
          <span className="gf-hint">Cria automaticamente o(s) lançamento(s) de saída e atualiza o saldo da conta.</span>
        </div>
      </form>
    </Modal>
  );
}

function EditContaModal({ conta, apropriacoes, onClose, onSave }) {
  const [f, setF] = useState({
    valorPrevisto: String(conta.valorPrevisto), dataVencimento: conta.dataVencimento,
    apropriacao: conta.apropriacao, descricao: conta.descricao, scope: "somente",
  });
  const upd = (p) => setF((s) => ({ ...s, ...p }));
  const submit = (e) => {
    e.preventDefault();
    if (!f.valorPrevisto || Number(f.valorPrevisto) <= 0) return;
    onSave({ valorPrevisto: Number(f.valorPrevisto), dataVencimento: f.dataVencimento, apropriacao: f.apropriacao, descricao: f.descricao }, f.scope);
  };
  return (
    <Modal title="Editar conta a pagar" onClose={onClose}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <Field label="Valor previsto (R$)"><input type="number" min="0.01" step="0.01" required value={f.valorPrevisto} onChange={(e) => upd({ valorPrevisto: e.target.value })} /></Field>
        <Field label="Data de vencimento" hint={conta.total > 1 ? "Aplica-se apenas a esta ocorrência" : undefined}>
          <input type="date" required value={f.dataVencimento} onChange={(e) => upd({ dataVencimento: e.target.value })} />
        </Field>
        <Field label="Apropriação">
          <select value={f.apropriacao} onChange={(e) => upd({ apropriacao: e.target.value })}>
            {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
        </Field>
        <Field label="Descrição"><input value={f.descricao} onChange={(e) => upd({ descricao: e.target.value })} /></Field>
        {conta.total > 1 && (
          <div className="gf-span-2">
            <span className="gf-field-label" style={{ display: "block", marginBottom: 6 }}>Aplicar alteração de valor a</span>
            <div className="gf-segmented">
              <button type="button" className={f.scope === "somente" ? "is-active" : ""} onClick={() => upd({ scope: "somente" })}>Somente esta ocorrência</button>
              <button type="button" className={f.scope === "todas" ? "is-active" : ""} onClick={() => upd({ scope: "todas" })}>Esta e as próximas</button>
            </div>
          </div>
        )}
        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">Salvar alterações</button>
          {conta.rateio && conta.rateio.length > 1 && (
            <span className="gf-hint">O rateio entre os {conta.rateio.length} centros de custo será recalculado proporcionalmente ao novo valor.</span>
          )}
        </div>
      </form>
    </Modal>
  );
}

function GrupoModal({ itens, centroCustoName, onClose, onBaixar, onEditar }) {
  const total = itens.length;
  const valorTotal = itens.reduce((s, c) => s + c.valorPrevisto, 0);
  const pagos = itens.filter((c) => c.pagamento);
  const valorPago = pagos.reduce((s, c) => s + (c.pagamento?.valorFinal || 0), 0);
  const emAberto = itens.filter((c) => !c.pagamento && !c.cancelada);
  const valorEmAberto = emAberto.reduce((s, c) => s + c.valorPrevisto, 0);

  return (
    <Modal title={`${itens[0]?.descricao || "Grupo"} — ${total} ocorrências`} onClose={onClose} width={680}>
      <div className="gf-detail-meta" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        <div><span>Valor total contratado</span><b className="mono">{BRL(valorTotal)}</b></div>
        <div><span>Valor já pago</span><b className="mono pos">{BRL(valorPago)}</b></div>
        <div><span>Valor em aberto</span><b className="mono neg">{BRL(valorEmAberto)}</b></div>
        <div><span>Pagas / em aberto</span><b className="mono">{pagos.length} / {emAberto.length}</b></div>
      </div>
      <div className="gf-table-wrap" style={{ marginTop: 14, maxHeight: 320 }}>
        <table className="gf-table">
          <thead><tr><th>Parcela</th><th>Vencimento</th><th>Centro de custo</th><th className="num">Valor</th><th>Status</th><th></th></tr></thead>
          <tbody>
            {itens.map((c) => {
              const status = deriveStatusConta(c);
              return (
                <tr key={c.id}>
                  <td>{c.seq}/{c.total}</td>
                  <td>{fmtDate(c.dataVencimento)}</td>
                  <td>{c.rateio && c.rateio.length > 1 ? `Rateado (${c.rateio.length})` : centroCustoName(c.rateio?.[0]?.centroCusto ?? c.centroCusto)}</td>
                  <td className="num mono">{BRL(c.valorPrevisto)}</td>
                  <td><Badge tone={STATUS_TONE[status]}>{status}</Badge></td>
                  <td>
                    <div className="gf-row-actions">
                      {!c.pagamento && !c.cancelada && (
                        <>
                          <button className="gf-icon-btn" title="Baixar" onClick={() => onBaixar(c)}><CircleDollarSign size={14} /></button>
                          <button className="gf-icon-btn" title="Editar" onClick={() => onEditar(c)}><Pencil size={14} /></button>
                        </>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

function RateioViewModal({ conta, centroCustoName, onClose }) {
  const pago = conta.pagamento?.rateioPago;
  return (
    <Modal title="Rateio por centro de custo" onClose={onClose} width={640}>
      <div className="gf-detail-meta" style={{ gridTemplateColumns: "repeat(2,1fr)" }}>
        <div><span>Fornecedor</span><b>{conta.fornecedor || "—"}</b></div>
        <div><span>Descrição</span><b>{conta.descricao || "—"}</b></div>
        <div><span>Vencimento</span><b>{fmtDate(conta.dataVencimento)}</b></div>
        <div><span>Valor total do boleto</span><b className="mono">{BRL(conta.valorPrevisto)}</b></div>
      </div>
      <p className="gf-field-hint" style={{ margin: "10px 0" }}>
        Este é um único título/boleto, apenas apropriado entre {conta.rateio.length} centros de custo. Veja abaixo o valor original de cada um e, se já pago, quanto foi efetivamente desembolsado (incluindo a parte proporcional de multas/juros/descontos).
      </p>
      <div className="gf-table-wrap">
        <table className="gf-table">
          <thead>
            <tr>
              <th>Centro de custo</th><th className="num">%</th><th className="num">Valor original</th>
              {pago && <th className="num">Acréscimo/desc.</th>}
              {pago && <th className="num">Valor final pago</th>}
            </tr>
          </thead>
          <tbody>
            {conta.rateio.map((r) => {
              const p = pago?.find((x) => x.centroCusto === r.centroCusto);
              return (
                <tr key={r.centroCusto}>
                  <td>{centroCustoName(r.centroCusto)}</td>
                  <td className="num mono">{r.percentual.toFixed(1)}%</td>
                  <td className="num mono">{BRL(r.valor)}</td>
                  {pago && <td className={`num mono ${p && p.acrescimo < 0 ? "pos" : ""}`}>{p ? `${p.acrescimo >= 0 ? "+" : ""}${BRL(p.acrescimo)}` : "—"}</td>}
                  {pago && <td className="num mono">{p ? BRL(p.valorFinal) : "—"}</td>}
                </tr>
              );
            })}
            <tr>
              <td><b>Total</b></td>
              <td className="num mono"><b>100%</b></td>
              <td className="num mono"><b>{BRL(conta.valorPrevisto)}</b></td>
              {pago && <td className="num mono"><b>{BRL((conta.pagamento.multa || 0) + (conta.pagamento.juros || 0) + (conta.pagamento.outrosAcrescimos || 0) - (conta.pagamento.descontos || 0))}</b></td>}
              {pago && <td className="num mono"><b>{BRL(conta.pagamento.valorFinal)}</b></td>}
            </tr>
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Página: Custos Internos
--------------------------------------------------------------- */

function CustosInternosPage({ entries, apropriacoes, accounts }) {
  const [filters, setFilters] = useState({ apropriacao: "", de: "", ate: "" });
  const accountName = (id) => accounts.find((a) => a.id === id)?.nome || "—";

  const internos = useMemo(
    () => entries.filter((e) => e.obraId === SEM_OBRA && e.tipo === "saida"),
    [entries]
  );

  const filtered = internos
    .filter((e) => (filters.apropriacao ? e.apropriacao === filters.apropriacao : true))
    .filter((e) => (filters.de ? e.data >= filters.de : true))
    .filter((e) => (filters.ate ? e.data <= filters.ate : true))
    .sort((a, b) => (a.data < b.data ? 1 : -1));

  const now = new Date();
  const curKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const prevD = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevKey = `${prevD.getFullYear()}-${String(prevD.getMonth() + 1).padStart(2, "0")}`;

  const totalGeral = filtered.reduce((s, e) => s + e.valor, 0);
  const totalMesAtual = internos.filter((e) => monthKey(e.data) === curKey).reduce((s, e) => s + e.valor, 0);
  const totalMesAnterior = internos.filter((e) => monthKey(e.data) === prevKey).reduce((s, e) => s + e.valor, 0);
  const variacao = totalMesAnterior ? ((totalMesAtual - totalMesAnterior) / totalMesAnterior) * 100 : null;

  const porMes = useMemo(() => {
    const map = {};
    filtered.forEach((e) => {
      const k = monthKey(e.data);
      if (!map[k]) map[k] = { mes: k, total: 0 };
      map[k].total += e.valor;
    });
    return Object.values(map).sort((a, b) => a.mes.localeCompare(b.mes)).map((r) => ({ ...r, mesLabel: monthLabel(r.mes) }));
  }, [filtered]);

  const porApropriacao = useMemo(() => {
    const map = {};
    filtered.forEach((e) => { map[e.apropriacao] = (map[e.apropriacao] || 0) + e.valor; });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [filtered]);

  return (
    <div className="gf-stack">
      <div className="gf-toolbar"><h2>Custos internos</h2></div>

      <div className="gf-tiles" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        <div className="gf-tile">
          <span className="gf-tile-label">Total no período filtrado</span>
          <span className="gf-tile-value neg">{BRL(totalGeral)}</span>
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Gasto no mês atual</span>
          <span className="gf-tile-value neg">{BRL(totalMesAtual)}</span>
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Gasto no mês anterior</span>
          <span className="gf-tile-value">{BRL(totalMesAnterior)}</span>
        </div>
        <div className="gf-tile">
          <span className="gf-tile-label">Variação vs mês anterior</span>
          <span className={`gf-tile-value ${variacao === null ? "" : variacao <= 0 ? "pos" : "neg"}`}>
            {variacao === null ? "—" : `${variacao > 0 ? "+" : ""}${variacao.toFixed(0)}%`}
          </span>
        </div>
      </div>

      <section className="gf-card">
        <div className="gf-card-head"><h2>Filtros</h2></div>
        <div className="gf-filters-grid" style={{ gridTemplateColumns: "repeat(3,1fr)" }}>
          <select value={filters.apropriacao} onChange={(e) => setFilters((f) => ({ ...f, apropriacao: e.target.value }))}>
            <option value="">Apropriação (todas)</option>
            {apropriacoes.map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
          <input type="date" title="De" value={filters.de} onChange={(e) => setFilters((f) => ({ ...f, de: e.target.value }))} />
          <input type="date" title="Até" value={filters.ate} onChange={(e) => setFilters((f) => ({ ...f, ate: e.target.value }))} />
        </div>
      </section>

      {internos.length === 0 ? (
        <div className="gf-card">
          <EmptyState icon={BarChart3} title="Nenhum custo interno lançado" body='Registre saídas na página Lançamentos com o centro de custo "Administrativo / Custos internos" para vê-las aqui.' />
        </div>
      ) : (
        <>
          <section className="gf-card">
            <div className="gf-card-head"><h2>Custos internos por mês</h2></div>
            <div className="gf-chart-box">
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={porMes} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
                  <CartesianGrid stroke="var(--gf-border)" vertical={false} />
                  <XAxis dataKey="mesLabel" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
                  <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                  <Bar dataKey="total" name="Custos internos" fill="var(--gf-danger)" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section className="gf-card">
            <div className="gf-card-head"><h2>Distribuição por apropriação</h2></div>
            <div className="gf-chart-box gf-chart-box-pie">
              <ResponsiveContainer width="100%" height={260}>
                <PieChart>
                  <Pie data={porApropriacao} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} label={(d) => d.name}>
                    {porApropriacao.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          </section>

          <section className="gf-card">
            <div className="gf-card-head"><h2>Lançamentos de custos internos</h2></div>
            <div className="gf-table-wrap">
              <table className="gf-table gf-table-ledger">
                <thead><tr><th>Data</th><th>Conta</th><th>Apropriação</th><th>Fornecedor</th><th>Descrição</th><th className="num">Valor</th></tr></thead>
                <tbody>
                  {filtered.map((e) => (
                    <tr key={e.id} className="row-neg">
                      <td>{fmtDate(e.data)}</td>
                      <td>{accountName(e.contaId)}</td>
                      <td>{e.apropriacao}</td>
                      <td>{e.fornecedor || "—"}</td>
                      <td className="gf-td-desc" title={e.descricao}>{e.descricao || "—"}</td>
                      <td className="num mono neg">{BRL(e.valor)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}

      <p className="gf-field-hint">
        Pró-labore dos sócios é controlado separadamente na página Sócios e não entra nestes totais.
      </p>
    </div>
  );
}

/* ---------------------------------------------------------------
   Página: Usuários e Permissões
--------------------------------------------------------------- */

function UsuariosPage({ usuarios, setUsuarios, currentUser, auditLog, logAudit }) {
  const [tab, setTab] = useState("usuarios");
  const [editingUser, setEditingUser] = useState(null);
  const [showNew, setShowNew] = useState(false);
  const [filtros, setFiltros] = useState({ usuario: "", tipo: "", de: "", ate: "" });

  const salvarUsuario = (data) => {
    if (data.id) {
      const anterior = usuarios.find((u) => u.id === data.id);
      setUsuarios((prev) => prev.map((u) => (u.id === data.id ? data : u)));
      if (anterior) {
        const mudouPerms = JSON.stringify(anterior.perms) !== JSON.stringify(data.perms);
        const mudouPapel = anterior.papel !== data.papel;
        const mudouAtivo = anterior.ativo !== data.ativo;
        if (mudouPerms || mudouPapel || mudouAtivo) {
          logAudit?.({
            tipoOperacao: "Alteração de permissões de usuário", registroAlterado: data.nome,
            valorAnterior: `${PAPEL_LABEL[anterior.papel] || anterior.papel}${anterior.ativo ? "" : " (inativo)"}`,
            valorNovo: `${PAPEL_LABEL[data.papel] || data.papel}${data.ativo ? "" : " (inativo)"}`,
          });
        }
      }
    } else {
      setUsuarios((prev) => [...prev, { ...data, id: uid("user") }]);
      logAudit?.({ tipoOperacao: "Cadastro de usuário", registroAlterado: data.nome, valorAnterior: "—", valorNovo: PAPEL_LABEL[data.papel] || data.papel });
    }
    setEditingUser(null); setShowNew(false);
  };

  const auditFiltrado = auditLog
    .filter((a) => (filtros.usuario ? a.usuario === filtros.usuario : true))
    .filter((a) => (filtros.tipo ? a.tipoOperacao === filtros.tipo : true));

  const usuariosNomes = [...new Set(auditLog.map((a) => a.usuario))];
  const tiposOperacao = [...new Set(auditLog.map((a) => a.tipoOperacao))];

  return (
    <div className="gf-stack">
      <div className="gf-toolbar">
        <h2>Usuários e permissões</h2>
        {tab === "usuarios" && <button className="gf-btn gf-btn-primary" onClick={() => setShowNew(true)}><Plus size={15} /> Novo usuário</button>}
      </div>

      <div className="gf-segmented" style={{ marginBottom: 4 }}>
        <button className={tab === "usuarios" ? "is-active" : ""} onClick={() => setTab("usuarios")}>Usuários</button>
        <button className={tab === "auditoria" ? "is-active" : ""} onClick={() => setTab("auditoria")}>Log de auditoria</button>
      </div>

      {tab === "usuarios" && (
        <div className="gf-grid-cards">
          {usuarios.map((u) => (
            <div className={`gf-card gf-account-card ${!u.ativo ? "is-inactive" : ""}`} key={u.id}>
              <div className="gf-card-head">
                <div>
                  <div className="gf-eyebrow">{PAPEL_LABEL[u.papel] || u.papel}</div>
                  <h3>{u.nome}</h3>
                </div>
                <div className="gf-row-actions">
                  <Badge tone={u.ativo ? "green" : "gray"}>{u.ativo ? "Ativo" : "Inativo"}</Badge>
                  <button className="gf-icon-btn" title="Editar permissões" onClick={() => setEditingUser(u)}><Pencil size={14} /></button>
                </div>
              </div>
              <p className="gf-field-hint">{ALL_PERM_KEYS.filter((k) => u.perms?.[k]).length} de {ALL_PERM_KEYS.length} permissões ativas</p>
            </div>
          ))}
        </div>
      )}

      {tab === "auditoria" && (
        <section className="gf-card">
          <div className="gf-card-head"><h2>Log de auditoria</h2></div>
          <div className="gf-filters-grid" style={{ gridTemplateColumns: "repeat(2,1fr)", marginBottom: 14 }}>
            <select value={filtros.usuario} onChange={(e) => setFiltros((f) => ({ ...f, usuario: e.target.value }))}>
              <option value="">Usuário (todos)</option>
              {usuariosNomes.map((n) => <option key={n} value={n}>{n}</option>)}
            </select>
            <select value={filtros.tipo} onChange={(e) => setFiltros((f) => ({ ...f, tipo: e.target.value }))}>
              <option value="">Tipo de operação (todos)</option>
              {tiposOperacao.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
          <div className="gf-table-wrap">
            <table className="gf-table">
              <thead><tr><th>Data/hora</th><th>Usuário</th><th>Operação</th><th>Registro</th><th>Anterior</th><th>Novo</th></tr></thead>
              <tbody>
                {auditFiltrado.length === 0 && (
                  <tr><td colSpan={6}><EmptyState icon={ShieldCheck} title="Nenhum registro de auditoria" body="As operações sensíveis (saldo inicial, exclusões, pagamentos, permissões) aparecerão aqui." /></td></tr>
                )}
                {auditFiltrado.map((a) => (
                  <tr key={a.id}>
                    <td className="mono">{a.data}</td>
                    <td>{a.usuario}</td>
                    <td>{a.tipoOperacao}</td>
                    <td className="gf-td-desc">{a.registroAlterado}</td>
                    <td className="mono">{a.valorAnterior}</td>
                    <td className="mono">{a.valorNovo}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {(editingUser || showNew) && (
        <UsuarioModal usuario={editingUser} onClose={() => { setEditingUser(null); setShowNew(false); }} onSave={salvarUsuario} />
      )}
    </div>
  );
}

function UsuarioModal({ usuario, onClose, onSave }) {
  const [f, setF] = useState(usuario ? { ...usuario, perms: { ...usuario.perms } } : {
    nome: "", papel: "customizavel", ativo: true, perms: permsAllFalse(),
  });
  const upd = (p) => setF((s) => ({ ...s, ...p }));
  const togglePerm = (key) => setF((s) => ({ ...s, perms: { ...s.perms, [key]: !s.perms[key] } }));
  const aplicarTemplate = (papel) => {
    const perms = papel === "admin" ? permsAllTrue() : papel === "operacional" ? { ...ASSISTENTE_PERMS } : permsAllFalse();
    upd({ papel, perms });
  };
  const submit = (e) => {
    e.preventDefault();
    if (!f.nome.trim()) return;
    onSave({ ...f, nome: f.nome.trim() });
  };

  return (
    <Modal title={usuario ? "Editar usuário" : "Novo usuário"} onClose={onClose} width={640}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <Field label="Nome"><input required value={f.nome} onChange={(e) => upd({ nome: e.target.value })} /></Field>
        <Field label="Status">
          <select value={f.ativo ? "ativo" : "inativo"} onChange={(e) => upd({ ativo: e.target.value === "ativo" })}>
            <option value="ativo">Ativo</option><option value="inativo">Inativo</option>
          </select>
        </Field>
        <div className="gf-span-2">
          <span className="gf-field-label" style={{ display: "block", marginBottom: 6 }}>Perfil (aplica um modelo de permissões — você ainda pode ajustar individualmente abaixo)</span>
          <div className="gf-segmented gf-segmented-3">
            <button type="button" className={f.papel === "admin" ? "is-active" : ""} onClick={() => aplicarTemplate("admin")}>Administrador</button>
            <button type="button" className={f.papel === "operacional" ? "is-active" : ""} onClick={() => aplicarTemplate("operacional")}>Operacional</button>
            <button type="button" className={f.papel === "customizavel" ? "is-active" : ""} onClick={() => aplicarTemplate("customizavel")}>Personalizado</button>
          </div>
        </div>

        <div className="gf-span-2">
          {PERMISSION_GROUPS.map((g) => (
            <div key={g.group} style={{ marginBottom: 12 }}>
              <span className="gf-field-label" style={{ display: "block", marginBottom: 6 }}>{g.group}</span>
              <div className="gf-perm-grid">
                {g.keys.map((k) => (
                  <label className="gf-checkbox gf-perm-item" key={k.key}>
                    <input type="checkbox" checked={!!f.perms[k.key]} onChange={() => togglePerm(k.key)} />
                    <span>{k.label}</span>
                  </label>
                ))}
              </div>
            </div>
          ))}
        </div>

        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">Salvar usuário</button>
        </div>
      </form>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Página: Dívidas e Investimentos
--------------------------------------------------------------- */

function DividasInvestimentosPage({ dividas, investimentos, entries, accounts, obras }) {
  const [filters, setFilters] = useState({ de: "", ate: "", instituicao: "", tipoDivida: "", tipoInvestimento: "", status: "", obraId: "" });
  const [dividaDetail, setDividaDetail] = useState(null);
  const [investimentoDetail, setInvestimentoDetail] = useState(null);

  const obraName = (id) => (id === SEM_OBRA ? "Administrativo / Custos internos" : obras.find((o) => o.id === id)?.nome || "—");
  const instituicoes = [...new Set([...dividas.map((d) => d.instituicao), ...investimentos.map((i) => i.instituicao)])].filter(Boolean);
  const statusOptions = [...new Set([...dividas.map((d) => d.status), ...investimentos.map((i) => i.status)])].filter(Boolean);

  const dentroPeriodo = (data) => (!filters.de || data >= filters.de) && (!filters.ate || data <= filters.ate);

  const filteredDividas = dividas
    .filter((d) => (filters.instituicao ? d.instituicao === filters.instituicao : true))
    .filter((d) => (filters.tipoDivida ? d.tipoOperacao === filters.tipoDivida : true))
    .filter((d) => (filters.status ? d.status === filters.status : true))
    .filter((d) => (filters.obraId ? d.obraId === filters.obraId : true))
    .filter((d) => dentroPeriodo(d.dataContratacao));

  const filteredInvestimentos = investimentos
    .filter((i) => (filters.instituicao ? i.instituicao === filters.instituicao : true))
    .filter((i) => (filters.tipoInvestimento ? i.tipoInvestimento === filters.tipoInvestimento : true))
    .filter((i) => (filters.status ? i.status === filters.status : true))
    .filter((i) => (filters.obraId ? i.obraId === filters.obraId : true))
    .filter((i) => dentroPeriodo(i.dataAplicacao));

  const dividasComStats = filteredDividas.map((d) => ({ ...d, stats: computeDividaStats(d, entries) }));
  const investimentosComStats = filteredInvestimentos.map((i) => ({ ...i, stats: computeInvestimentoStats(i, entries) }));

  const curMonthKey = monthKey(todayISO());
  const anoAtual = new Date().getFullYear();
  const idsDividas = new Set(filteredDividas.map((d) => d.id));
  const idsInvestimentos = new Set(filteredInvestimentos.map((i) => i.id));

  const totais = useMemo(() => {
    const dividaTotalAtual = dividasComStats.filter((d) => d.status !== "Quitada").reduce((s, d) => s + d.stats.saldoDevedor, 0);
    const totalJurosPagos = dividasComStats.reduce((s, d) => s + d.stats.jurosPagos, 0);
    const jurosPagosNoMes = entries.filter((e) => e.dividaSubtipo === "pagamento" && idsDividas.has(e.dividaId) && monthKey(e.data) === curMonthKey).reduce((s, e) => s + (Number(e.juros) || 0), 0);
    const parcelasAPagar = dividasComStats.filter((d) => d.status === "Ativa").reduce((s, d) => s + d.stats.parcelasRestantes, 0);

    const totalInvestido = investimentosComStats.filter((i) => i.status === "Ativo").reduce((s, i) => s + i.stats.principalLiquido, 0);
    const valorAtualInvestimentos = investimentosComStats.reduce((s, i) => s + i.stats.valorAtual, 0);
    const rendimentosAcumulados = investimentosComStats.reduce((s, i) => s + i.stats.rendimentos, 0);
    const rendimentosNoMes = entries.filter((e) => e.investimentoSubtipo === "rendimento" && idsInvestimentos.has(e.investimentoId) && monthKey(e.data) === curMonthKey).reduce((s, e) => s + (Number(e.valor) || 0), 0);
    const rentabilidadeAcumulada = totalInvestido > 0 ? (rendimentosAcumulados / totalInvestido) * 100 : 0;

    const jurosPagosAno = entries.filter((e) => e.dividaSubtipo === "pagamento" && idsDividas.has(e.dividaId) && e.data?.slice(0, 4) === String(anoAtual)).reduce((s, e) => s + (Number(e.juros) || 0), 0);
    const rendimentosAno = entries.filter((e) => e.investimentoSubtipo === "rendimento" && idsInvestimentos.has(e.investimentoId) && e.data?.slice(0, 4) === String(anoAtual)).reduce((s, e) => s + (Number(e.valor) || 0), 0);

    return {
      dividaTotalAtual, totalJurosPagos, jurosPagosNoMes, parcelasAPagar, totalAPagarFuturo: dividaTotalAtual,
      totalInvestido, valorAtualInvestimentos, rendimentosAcumulados, rendimentosNoMes, rentabilidadeAcumulada,
      dividaLiquida: dividaTotalAtual - valorAtualInvestimentos, patrimonioLiquido: valorAtualInvestimentos - dividaTotalAtual,
      jurosPagosAno, rendimentosAno, custoFinanceiroLiquido: jurosPagosAno - rendimentosAno,
    };
  }, [dividasComStats, investimentosComStats, entries, curMonthKey, anoAtual]);

  // Evolução mensal (últimos meses com movimentação)
  const evolucao = useMemo(() => {
    const todosMeses = new Set();
    entries.forEach((e) => { if (e.dividaId || e.investimentoId) todosMeses.add(monthKey(e.data)); });
    const meses = [...todosMeses].sort();
    if (meses.length === 0) return [];
    let dividaAcum = 0, investAcum = 0, rendAcum = 0;
    const contratacoesPorMes = {}, amortPorMes = {}, jurosPorMes = {}, aplicPorMes = {}, rendPorMes = {};
    entries.forEach((e) => {
      const mk = monthKey(e.data);
      if (e.dividaSubtipo === "contratacao") contratacoesPorMes[mk] = (contratacoesPorMes[mk] || 0) + e.valor;
      if (e.dividaSubtipo === "pagamento") { amortPorMes[mk] = (amortPorMes[mk] || 0) + (e.amortizacao || 0); jurosPorMes[mk] = (jurosPorMes[mk] || 0) + (e.juros || 0); }
      if (e.investimentoSubtipo === "aplicacao" || e.investimentoSubtipo === "aporte") aplicPorMes[mk] = (aplicPorMes[mk] || 0) + e.valor;
      if (e.investimentoSubtipo === "resgate") aplicPorMes[mk] = (aplicPorMes[mk] || 0) - e.valor;
      if (e.investimentoSubtipo === "rendimento") rendPorMes[mk] = (rendPorMes[mk] || 0) + e.valor;
    });
    return meses.map((mk) => {
      dividaAcum += (contratacoesPorMes[mk] || 0) - (amortPorMes[mk] || 0);
      investAcum += (aplicPorMes[mk] || 0);
      rendAcum += (rendPorMes[mk] || 0);
      return {
        mesLabel: monthLabel(mk), dividaTotal: Math.max(0, dividaAcum), jurosPagos: jurosPorMes[mk] || 0,
        investimentos: investAcum + rendAcum, rendimentos: rendPorMes[mk] || 0, dividaLiquida: Math.max(0, dividaAcum) - (investAcum + rendAcum),
      };
    });
  }, [entries]);

  return (
    <div className="gf-stack">
      <div className="gf-toolbar"><h2>Dívidas e investimentos</h2></div>

      <section className="gf-card">
        <div className="gf-card-head"><h2>Filtros</h2></div>
        <div className="gf-filters-grid" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
          <input type="date" title="De" value={filters.de} onChange={(e) => setFilters((f) => ({ ...f, de: e.target.value }))} />
          <input type="date" title="Até" value={filters.ate} onChange={(e) => setFilters((f) => ({ ...f, ate: e.target.value }))} />
          <select value={filters.instituicao} onChange={(e) => setFilters((f) => ({ ...f, instituicao: e.target.value }))}>
            <option value="">Instituição (todas)</option>
            {instituicoes.map((i) => <option key={i} value={i}>{i}</option>)}
          </select>
          <select value={filters.status} onChange={(e) => setFilters((f) => ({ ...f, status: e.target.value }))}>
            <option value="">Status (todos)</option>
            {statusOptions.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <select value={filters.tipoDivida} onChange={(e) => setFilters((f) => ({ ...f, tipoDivida: e.target.value }))}>
            <option value="">Tipo de dívida (todos)</option>
            {TIPOS_DIVIDA.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          <select value={filters.tipoInvestimento} onChange={(e) => setFilters((f) => ({ ...f, tipoInvestimento: e.target.value }))}>
            <option value="">Tipo de investimento (todos)</option>
            {TIPOS_INVESTIMENTO.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          <select value={filters.obraId} onChange={(e) => setFilters((f) => ({ ...f, obraId: e.target.value }))}>
            <option value="">Obra/centro de custo (todos)</option>
            <option value={SEM_OBRA}>Administrativo / Custos internos</option>
            {obras.map((o) => <option key={o.id} value={o.id}>{o.nome}</option>)}
          </select>
        </div>
      </section>

      <div className="gf-dvi-columns">
        <section className="gf-card">
          <div className="gf-card-head"><h2>Dívidas</h2></div>
          <div className="gf-account-rows">
            <div><span>Dívida total atual</span><b className="mono neg">{BRL(totais.dividaTotalAtual)}</b></div>
            <div><span>Total de juros pagos</span><b className="mono">{BRL(totais.totalJurosPagos)}</b></div>
            <div><span>Juros pagos no mês</span><b className="mono">{BRL(totais.jurosPagosNoMes)}</b></div>
            <div><span>Parcelas a pagar</span><b className="mono">{totais.parcelasAPagar}</b></div>
            <div><span>Total a pagar futuro (saldo devedor)</span><b className="mono">{BRL(totais.totalAPagarFuturo)}</b></div>
          </div>
        </section>
        <section className="gf-card">
          <div className="gf-card-head"><h2>Investimentos</h2></div>
          <div className="gf-account-rows">
            <div><span>Total investido</span><b className="mono">{BRL(totais.totalInvestido)}</b></div>
            <div><span>Valor atual dos investimentos</span><b className="mono pos">{BRL(totais.valorAtualInvestimentos)}</b></div>
            <div><span>Rendimentos acumulados</span><b className="mono pos">{BRL(totais.rendimentosAcumulados)}</b></div>
            <div><span>Rendimentos no mês</span><b className="mono pos">{BRL(totais.rendimentosNoMes)}</b></div>
            <div><span>Rentabilidade acumulada</span><b className="mono">{totais.rentabilidadeAcumulada.toFixed(1)}%</b></div>
          </div>
        </section>
      </div>

      <section className="gf-card">
        <div className="gf-card-head"><h2>Visão consolidada</h2></div>
        <div className="gf-account-rows" style={{ flexWrap: "wrap" }}>
          <div><span>Dívidas</span><b className="mono neg">{BRL(totais.dividaTotalAtual)}</b></div>
          <div><span>Investimentos (valor atual)</span><b className="mono pos">{BRL(totais.valorAtualInvestimentos)}</b></div>
          <div><span>Patrimônio financeiro líquido</span><b className={`mono ${totais.patrimonioLiquido >= 0 ? "pos" : "neg"}`}>{BRL(totais.patrimonioLiquido)}</b></div>
          <div><span>Dívida líquida</span><b className={`mono ${totais.dividaLiquida <= 0 ? "pos" : "neg"}`}>{BRL(totais.dividaLiquida)}</b></div>
          <div><span>Juros pagos no ano</span><b className="mono">{BRL(totais.jurosPagosAno)}</b></div>
          <div><span>Rendimentos no ano</span><b className="mono pos">{BRL(totais.rendimentosAno)}</b></div>
          <div><span>Custo financeiro líquido (ano)</span><b className={`mono ${totais.custoFinanceiroLiquido >= 0 ? "neg" : "pos"}`}>{BRL(totais.custoFinanceiroLiquido)}</b></div>
        </div>
      </section>

      {evolucao.length > 0 && (
        <section className="gf-card">
          <div className="gf-card-head"><h2>Evolução mensal</h2></div>
          <div className="gf-chart-box">
            <ResponsiveContainer width="100%" height={260}>
              <LineChart data={evolucao} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
                <CartesianGrid stroke="var(--gf-border)" vertical={false} />
                <XAxis dataKey="mesLabel" tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
                <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Line type="monotone" dataKey="dividaTotal" name="Dívida total" stroke="var(--gf-danger)" strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="investimentos" name="Investimentos" stroke="var(--gf-accent)" strokeWidth={2} dot={false} />
                <Line type="monotone" dataKey="dividaLiquida" name="Dívida líquida" stroke="var(--gf-primary)" strokeWidth={2} strokeDasharray="4 3" dot={false} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </section>
      )}

      <section className="gf-card">
        <div className="gf-card-head"><h2>Dívidas cadastradas</h2></div>
        {dividasComStats.length === 0 ? (
          <EmptyState icon={TrendingDown} title="Nenhuma dívida encontrada" body='Classifique um lançamento de entrada como "Dívida / Crédito" na página Lançamentos para registrar uma contratação.' />
        ) : (
          <div className="gf-grid-cards">
            {dividasComStats.map((d) => (
              <button className="gf-card gf-obra-card" key={d.id} onClick={() => setDividaDetail(d)}>
                <div className="gf-card-head">
                  <div><div className="gf-eyebrow">{d.tipoOperacao}</div><h3>{d.instituicao}</h3></div>
                  <Badge tone={d.status === "Quitada" ? "gray" : "red"}>{d.status}</Badge>
                </div>
                <div className="gf-account-rows">
                  <div><span>Valor contratado</span><b className="mono">{BRL(d.valorContratado)}</b></div>
                  <div><span>Saldo devedor</span><b className="mono neg">{BRL(d.stats.saldoDevedor)}</b></div>
                  <div><span>Parcelas pagas/restantes</span><b className="mono">{d.stats.parcelasPagas}/{d.qtdParcelas}</b></div>
                  <div><span>Juros pagos</span><b className="mono">{BRL(d.stats.jurosPagos)}</b></div>
                </div>
                <span className="gf-card-cta">Ver detalhes <ChevronRight size={14} /></span>
              </button>
            ))}
          </div>
        )}
      </section>

      <section className="gf-card">
        <div className="gf-card-head"><h2>Investimentos cadastrados</h2></div>
        {investimentosComStats.length === 0 ? (
          <EmptyState icon={TrendingUp} title="Nenhum investimento encontrado" body='Classifique um lançamento de saída como "Investimento" na página Lançamentos para registrar uma aplicação.' />
        ) : (
          <div className="gf-grid-cards">
            {investimentosComStats.map((i) => (
              <button className="gf-card gf-obra-card" key={i.id} onClick={() => setInvestimentoDetail(i)}>
                <div className="gf-card-head">
                  <div><div className="gf-eyebrow">{i.tipoInvestimento}</div><h3>{i.instituicao}</h3></div>
                  <Badge tone={i.status === "Ativo" ? "green" : "gray"}>{i.status}</Badge>
                </div>
                <div className="gf-account-rows">
                  <div><span>Valor investido líquido</span><b className="mono">{BRL(i.stats.principalLiquido)}</b></div>
                  <div><span>Valor atual</span><b className="mono pos">{BRL(i.stats.valorAtual)}</b></div>
                  <div><span>Rendimentos</span><b className="mono pos">{BRL(i.stats.rendimentos)}</b></div>
                  <div><span>Rentabilidade</span><b className="mono">{i.stats.rentabilidade.toFixed(1)}%</b></div>
                </div>
                <span className="gf-card-cta">Ver detalhes <ChevronRight size={14} /></span>
              </button>
            ))}
          </div>
        )}
      </section>

      {dividaDetail && <DividaDetailModal divida={dividaDetail} obraName={obraName} accounts={accounts} onClose={() => setDividaDetail(null)} />}
      {investimentoDetail && <InvestimentoDetailModal investimento={investimentoDetail} obraName={obraName} accounts={accounts} onClose={() => setInvestimentoDetail(null)} />}
    </div>
  );
}

function DividaDetailModal({ divida, obraName, accounts, onClose }) {
  const s = divida.stats;
  const accountName = (id) => accounts.find((a) => a.id === id)?.nome || "—";
  return (
    <Modal title={`${divida.instituicao} — ${divida.tipoOperacao}`} onClose={onClose} width={720}>
      <div className="gf-detail-meta" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        <div><span>Data de contratação</span><b>{fmtDate(divida.dataContratacao)}</b></div>
        <div><span>Valor contratado</span><b className="mono">{BRL(divida.valorContratado)}</b></div>
        <div><span>Taxa de juros</span><b>{divida.taxaJuros || "—"}</b></div>
        <div><span>Sistema de amortização</span><b>{divida.sistemaAmortizacao || "—"}</b></div>
        <div><span>Centro de custo</span><b>{obraName(divida.obraId)}</b></div>
        <div><span>Próximo vencimento</span><b>{divida.proximoVencimento ? fmtDate(divida.proximoVencimento) : "—"}</b></div>
        <div><span>Valor da próxima parcela</span><b className="mono">{divida.valorParcelaEstimado ? BRL(divida.valorParcelaEstimado) : "—"}</b></div>
        <div><span>Status</span><b>{divida.status}</b></div>
      </div>

      <div className="gf-account-rows gf-detail-stats" style={{ flexWrap: "wrap" }}>
        <div><span>Saldo devedor</span><b className="mono neg">{BRL(s.saldoDevedor)}</b></div>
        <div><span>Parcelas pagas</span><b className="mono">{s.parcelasPagas}</b></div>
        <div><span>Parcelas restantes</span><b className="mono">{s.parcelasRestantes}</b></div>
        <div><span>Total amortizado</span><b className="mono">{BRL(s.totalAmortizado)}</b></div>
        <div><span>Juros pagos</span><b className="mono">{BRL(s.jurosPagos)}</b></div>
        <div><span>Encargos pagos</span><b className="mono">{BRL(s.encargosPagos)}</b></div>
        <div><span>Total desembolsado</span><b className="mono">{BRL(s.totalDesembolsado)}</b></div>
      </div>

      <h4 className="gf-chart-title">Histórico de parcelas pagas</h4>
      <div className="gf-table-wrap" style={{ maxHeight: 260 }}>
        <table className="gf-table">
          <thead><tr><th>Data</th><th>Conta</th><th className="num">Amortização</th><th className="num">Juros</th><th className="num">Encargos</th><th className="num">Valor total</th></tr></thead>
          <tbody>
            {s.pagamentos.length === 0 && <tr><td colSpan={6}><EmptyState icon={Receipt} title="Nenhum pagamento registrado" body="Nenhuma parcela paga ainda para esta dívida." /></td></tr>}
            {s.pagamentos.sort((a, b) => (a.data < b.data ? 1 : -1)).map((p) => (
              <tr key={p.id}>
                <td>{fmtDate(p.data)}</td><td>{accountName(p.contaId)}</td>
                <td className="num mono">{BRL(p.amortizacao || 0)}</td>
                <td className="num mono">{BRL(p.juros || 0)}</td>
                <td className="num mono">{BRL(p.encargos || 0)}</td>
                <td className="num mono">{BRL(p.valor)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

function InvestimentoDetailModal({ investimento, obraName, accounts, onClose }) {
  const s = investimento.stats;
  const accountName = (id) => accounts.find((a) => a.id === id)?.nome || "—";
  const tone = (sub) => (sub === "rendimento" ? "green" : sub === "resgate" ? "orange" : "blue");
  const label = (sub) => ({ aplicacao: "Aplicação", aporte: "Aporte", resgate: "Resgate", rendimento: "Rendimento" }[sub] || sub);
  return (
    <Modal title={`${investimento.instituicao} — ${investimento.tipoInvestimento}`} onClose={onClose} width={720}>
      <div className="gf-detail-meta" style={{ gridTemplateColumns: "repeat(4,1fr)" }}>
        <div><span>Data da aplicação</span><b>{fmtDate(investimento.dataAplicacao)}</b></div>
        <div><span>Valor aplicado (inicial)</span><b className="mono">{BRL(investimento.valorAplicado)}</b></div>
        <div><span>Vencimento</span><b>{investimento.dataVencimento ? fmtDate(investimento.dataVencimento) : "—"}</b></div>
        <div><span>Liquidez</span><b>{investimento.liquidez || "—"}</b></div>
        <div><span>Centro de custo</span><b>{obraName(investimento.obraId)}</b></div>
        <div><span>Status</span><b>{investimento.status}</b></div>
      </div>

      <div className="gf-account-rows gf-detail-stats" style={{ flexWrap: "wrap" }}>
        <div><span>Aportes adicionais</span><b className="mono">{BRL(s.aportes)}</b></div>
        <div><span>Resgates</span><b className="mono neg">{BRL(s.resgates)}</b></div>
        <div><span>Valor investido líquido</span><b className="mono">{BRL(s.principalLiquido)}</b></div>
        <div><span>Rendimentos</span><b className="mono pos">{BRL(s.rendimentos)}</b></div>
        <div><span>Valor atual</span><b className="mono pos">{BRL(s.valorAtual)}</b></div>
        <div><span>Rentabilidade</span><b className="mono">{s.rentabilidade.toFixed(1)}%</b></div>
      </div>

      <h4 className="gf-chart-title">Movimentações</h4>
      <div className="gf-table-wrap" style={{ maxHeight: 260 }}>
        <table className="gf-table">
          <thead><tr><th>Data</th><th>Tipo</th><th>Conta</th><th className="num">Valor</th></tr></thead>
          <tbody>
            {s.relacionados.length === 0 && <tr><td colSpan={4}><EmptyState icon={TrendingUp} title="Nenhuma movimentação" body="Nenhum aporte, resgate ou rendimento registrado ainda." /></td></tr>}
            {[...s.relacionados].sort((a, b) => (a.data < b.data ? 1 : -1)).map((e) => (
              <tr key={e.id}>
                <td>{fmtDate(e.data)}</td>
                <td><Badge tone={tone(e.investimentoSubtipo)}>{label(e.investimentoSubtipo)}</Badge></td>
                <td>{accountName(e.contaId)}</td>
                <td className="num mono">{BRL(e.valor)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   Página: Aportes e Retiradas dos Sócios
--------------------------------------------------------------- */

const PRO_LABORE_PADRAO = 10000;

function SociosPage({ socios, setSocios, socioStats, movimentos, setMovimentos, accounts }) {
  const [showNew, setShowNew] = useState(false);
  const [editingMov, setEditingMov] = useState(null);
  const [editingParticipacao, setEditingParticipacao] = useState(false);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [filters, setFilters] = useState({ socioId: "", tipo: "", de: "", ate: "" });

  const socioName = (id) => socios.find((s) => s.id === id)?.nome || "—";
  const accountName = (id) => accounts.find((a) => a.id === id)?.nome || "—";

  const saveMov = (data) => {
    if (data.id) setMovimentos((prev) => prev.map((m) => (m.id === data.id ? data : m)));
    else setMovimentos((prev) => [{ ...data, id: uid("soc"), createdAt: Date.now() }, ...prev]);
    setShowNew(false); setEditingMov(null);
  };

  const remove = (id) => { setMovimentos((prev) => prev.filter((m) => m.id !== id)); setConfirmDeleteId(null); };

  const filtered = movimentos
    .filter((m) => (filters.socioId ? m.socioId === filters.socioId : true))
    .filter((m) => (filters.tipo ? m.tipo === filters.tipo : true))
    .filter((m) => (filters.de ? m.data >= filters.de : true))
    .filter((m) => (filters.ate ? m.data <= filters.ate : true))
    .sort((a, b) => (a.data < b.data ? 1 : a.data > b.data ? -1 : b.createdAt - a.createdAt));

  const totalAportes = socios.reduce((s, p) => s + (socioStats[p.id]?.aportes || 0), 0);
  const totalRetiradas = socios.reduce((s, p) => s + (socioStats[p.id]?.retiradas || 0), 0);
  const totalProLabore = socios.reduce((s, p) => s + (socioStats[p.id]?.proLabore || 0), 0);

  const comparativo = socios.map((p) => ({
    nome: p.nome,
    aportes: socioStats[p.id]?.aportes || 0,
    retiradas: socioStats[p.id]?.retiradas || 0,
  }));

  // Equilíbrio: as retiradas (excluindo pró-labore e aportes) devem seguir a % de participação de cada sócio.
  const balanco = useMemo(() => {
    const rows = socios.map((p) => {
      const real = socioStats[p.id]?.retiradas || 0;
      const ideal = totalRetiradas * ((Number(p.participacao) || 0) / 100);
      return { id: p.id, nome: p.nome, participacao: p.participacao, real, ideal, diff: real - ideal };
    });
    return rows.sort((a, b) => b.diff - a.diff);
  }, [socios, socioStats, totalRetiradas]);

  const EPS = 0.5;
  const devedor = balanco[0];
  const credor = balanco[balanco.length - 1];
  const desequilibrio = balanco.length >= 2 && devedor && credor && (devedor.diff - credor.diff) > EPS * 2;

  return (
    <div className="gf-stack">
      <div className="gf-toolbar">
        <h2>Sócios</h2>
        <div className="gf-toolbar-actions">
          <button className="gf-btn gf-btn-ghost" onClick={() => setEditingParticipacao(true)}>
            <Scale size={15} /> Participação societária
          </button>
          <button className="gf-btn gf-btn-primary" onClick={() => setShowNew(true)}>
            <Plus size={15} /> Novo lançamento
          </button>
        </div>
      </div>

      <section className={`gf-card gf-balance-card ${desequilibrio ? "is-unbalanced" : "is-balanced"}`}>
        <div className="gf-card-head"><h2>Situação entre os sócios</h2></div>
        {totalRetiradas <= 0 ? (
          <p className="gf-hint">Ainda não há retiradas registradas para calcular o equilíbrio entre os sócios.</p>
        ) : desequilibrio && socios.length === 2 ? (
          <div className="gf-balance-banner">
            <strong>{devedor.nome} está à frente nas retiradas.</strong>
            <span>
              Considerando a participação de {devedor.participacao}% / {credor.participacao}%, {devedor.nome} deve{" "}
              <b className="mono neg">{BRL(devedor.diff)}</b> a {credor.nome} para equilibrar as retiradas conforme a sociedade.
            </span>
          </div>
        ) : (
          <div className="gf-balance-banner is-ok">
            <strong>Retiradas equilibradas.</strong>
            <span>As retiradas de cada sócio estão de acordo com a participação societária.</span>
          </div>
        )}
        <div className="gf-table-wrap" style={{ marginTop: 12 }}>
          <table className="gf-table">
            <thead><tr><th>Sócio</th><th className="num">Participação</th><th className="num">Retirado (real)</th><th className="num">Retirada ideal</th><th className="num">Diferença</th></tr></thead>
            <tbody>
              {balanco.map((r) => (
                <tr key={r.id}>
                  <td><b>{r.nome}</b></td>
                  <td className="num mono">{r.participacao}%</td>
                  <td className="num mono">{BRL(r.real)}</td>
                  <td className="num mono">{BRL(r.ideal)}</td>
                  <td className={`num mono ${r.diff > EPS ? "neg" : r.diff < -EPS ? "pos" : ""}`}>
                    {r.diff > EPS ? `deve ${BRL(r.diff)}` : r.diff < -EPS ? `a receber ${BRL(-r.diff)}` : "equilibrado"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="gf-field-hint" style={{ marginTop: 10 }}>
          Cálculo baseado apenas nas retiradas (não considera aportes nem pró-labore, que não entram na divisão societária).
        </p>
      </section>

      <div className="gf-grid-cards">
        {socios.map((p) => {
          const s = socioStats[p.id] || { aportes: 0, retiradas: 0, proLabore: 0, saldo: 0 };
          return (
            <div className="gf-card gf-account-card" key={p.id}>
              <div className="gf-card-head">
                <div>
                  <div className="gf-eyebrow">Sócio</div>
                  <h3>{p.nome}</h3>
                </div>
                <Badge tone="green">{p.participacao}% da empresa</Badge>
              </div>
              <div className="gf-account-rows">
                <div><span>Total aportado</span><b className="mono pos">+ {BRL(s.aportes)}</b></div>
                <div><span>Total retirado</span><b className="mono neg">− {BRL(s.retiradas)}</b></div>
                <div><span>Pró-labore recebido</span><b className="mono">{BRL(s.proLabore)}</b></div>
              </div>
              <div className="gf-account-balance">
                <span>Saldo líquido (aportes − retiradas)</span>
                <strong className={`mono ${s.saldo >= 0 ? "pos" : "neg"}`}>{BRL(s.saldo)}</strong>
              </div>
            </div>
          );
        })}
      </div>

      <section className="gf-card">
        <div className="gf-card-head">
          <h2>Aportes x Retiradas por sócio</h2>
          <div className="gf-summary-inline">
            <span>Total aportado <b className="pos">{BRL(totalAportes)}</b></span>
            <span>Total retirado <b className="neg">{BRL(totalRetiradas)}</b></span>
            <span>Pró-labore pago <b>{BRL(totalProLabore)}</b></span>
          </div>
        </div>
        <div className="gf-chart-box">
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={comparativo} margin={{ top: 10, right: 20, left: 0, bottom: 0 }}>
              <CartesianGrid stroke="var(--gf-border)" vertical={false} />
              <XAxis dataKey="nome" tick={{ fontSize: 12, fill: "var(--gf-muted)" }} axisLine={{ stroke: "var(--gf-border)" }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "var(--gf-muted)" }} axisLine={false} tickLine={false} tickFormatter={(v) => `${(v / 1000).toFixed(0)}k`} width={44} />
              <Tooltip formatter={(v) => BRL(v)} contentStyle={{ fontFamily: "var(--gf-font-mono)", fontSize: 12, borderRadius: 8, border: "1px solid var(--gf-border)" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="aportes" name="Aportes" fill="var(--gf-accent)" radius={[4, 4, 0, 0]} />
              <Bar dataKey="retiradas" name="Retiradas" fill="var(--gf-danger)" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </section>

      <section className="gf-card">
        <div className="gf-card-head"><h2>Movimentações</h2></div>
        <div className="gf-filters-grid" style={{ marginBottom: 14 }}>
          <select value={filters.socioId} onChange={(e) => setFilters((f) => ({ ...f, socioId: e.target.value }))}>
            <option value="">Sócio (todos)</option>
            {socios.map((s) => <option key={s.id} value={s.id}>{s.nome}</option>)}
          </select>
          <select value={filters.tipo} onChange={(e) => setFilters((f) => ({ ...f, tipo: e.target.value }))}>
            <option value="">Tipo (todos)</option>
            <option value="aporte">Aporte</option><option value="retirada">Retirada</option><option value="pro-labore">Pró-labore</option>
          </select>
          <input type="date" title="De" value={filters.de} onChange={(e) => setFilters((f) => ({ ...f, de: e.target.value }))} />
          <input type="date" title="Até" value={filters.ate} onChange={(e) => setFilters((f) => ({ ...f, ate: e.target.value }))} />
        </div>
        <div className="gf-table-wrap">
          <table className="gf-table gf-table-ledger">
            <thead>
              <tr><th>Data</th><th>Tipo</th><th>Sócio</th><th>Conta</th><th>Descrição</th><th className="num">Valor</th><th></th></tr>
            </thead>
            <tbody>
              {filtered.length === 0 && (
                <tr><td colSpan={7}><EmptyState icon={Scale} title="Nenhuma movimentação encontrada" body="Registre um aporte, retirada ou pró-labore acima." /></td></tr>
              )}
              {filtered.map((m) => (
                <tr key={m.id} className={m.tipo === "aporte" ? "row-pos" : m.tipo === "pro-labore" ? "" : "row-neg"}>
                  <td>{fmtDate(m.data)}</td>
                  <td>
                    <Badge tone={m.tipo === "aporte" ? "green" : m.tipo === "pro-labore" ? "gray" : "red"}>
                      {m.tipo === "aporte" ? "Aporte" : m.tipo === "pro-labore" ? "Pró-labore" : "Retirada"}
                    </Badge>
                  </td>
                  <td>{socioName(m.socioId)}</td>
                  <td>{accountName(m.contaId)}</td>
                  <td className="gf-td-desc" title={m.descricao}>{m.descricao || "—"}</td>
                  <td className={`num mono ${m.tipo === "aporte" ? "pos" : m.tipo === "pro-labore" ? "" : "neg"}`}>
                    {m.tipo === "aporte" ? "+" : "−"} {BRL(m.valor)}
                  </td>
                  <td>
                    <div className="gf-row-actions">
                      {confirmDeleteId === m.id ? (
                        <>
                          <button className="gf-icon-btn danger" title="Confirmar exclusão" onClick={() => remove(m.id)}><Check size={14} /></button>
                          <button className="gf-icon-btn" title="Cancelar" onClick={() => setConfirmDeleteId(null)}><X size={14} /></button>
                        </>
                      ) : (
                        <>
                          <button className="gf-icon-btn" title="Editar" onClick={() => setEditingMov(m)}><Pencil size={14} /></button>
                          <button className="gf-icon-btn" title="Excluir" onClick={() => setConfirmDeleteId(m.id)}><Trash2 size={14} /></button>
                        </>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {(showNew || editingMov) && (
        <SocioMovModal
          mov={editingMov} socios={socios} accounts={accounts}
          onClose={() => { setShowNew(false); setEditingMov(null); }} onSave={saveMov}
        />
      )}
      {editingParticipacao && (
        <ParticipacaoModal socios={socios} onClose={() => setEditingParticipacao(false)} onSave={(s) => { setSocios(s); setEditingParticipacao(false); }} />
      )}
    </div>
  );
}

function SocioMovModal({ mov, socios, accounts, onClose, onSave }) {
  const [f, setF] = useState(mov || {
    tipo: "aporte", socioId: socios[0]?.id || "", contaId: accounts[0]?.id || "", data: todayISO(), valor: "", descricao: "",
  });
  const upd = (p) => setF((s) => ({ ...s, ...p }));

  const chooseTipo = (tipo) => {
    if (tipo === "pro-labore" && (!f.valor || f.tipo !== "pro-labore")) {
      upd({ tipo, valor: String(PRO_LABORE_PADRAO), descricao: f.descricao || "Pró-labore mensal" });
    } else {
      upd({ tipo });
    }
  };

  const submit = (e) => {
    e.preventDefault();
    if (!f.socioId || !f.contaId || !f.valor || Number(f.valor) <= 0) return;
    onSave({ ...f, valor: Number(f.valor) });
  };

  return (
    <Modal title={mov ? "Editar movimentação" : "Novo lançamento de sócio"} onClose={onClose}>
      <form onSubmit={submit} className="gf-form-grid gf-modal-form">
        <div className="gf-segmented gf-span-2 gf-segmented-3">
          <button type="button" className={f.tipo === "aporte" ? "is-active pos" : ""} onClick={() => chooseTipo("aporte")}>
            <ArrowUpRight size={14} /> Aporte
          </button>
          <button type="button" className={f.tipo === "retirada" ? "is-active neg" : ""} onClick={() => chooseTipo("retirada")}>
            <ArrowDownRight size={14} /> Retirada
          </button>
          <button type="button" className={f.tipo === "pro-labore" ? "is-active" : ""} onClick={() => chooseTipo("pro-labore")}>
            <Scale size={14} /> Pró-labore
          </button>
        </div>
        <Field label="Sócio">
          <select required value={f.socioId} onChange={(e) => upd({ socioId: e.target.value })}>
            {socios.map((s) => <option key={s.id} value={s.id}>{s.nome} ({s.participacao}%)</option>)}
          </select>
        </Field>
        <Field label="Conta bancária">
          <select required value={f.contaId} onChange={(e) => upd({ contaId: e.target.value })}>
            <option value="" disabled>Selecione…</option>
            {accounts.map((a) => <option key={a.id} value={a.id}>{a.nome}</option>)}
          </select>
        </Field>
        <Field label="Data"><input type="date" required value={f.data} onChange={(e) => upd({ data: e.target.value })} /></Field>
        <Field label="Valor (R$)" hint={f.tipo === "pro-labore" ? `Sugestão: ${BRL(PRO_LABORE_PADRAO)} (salário fixo mensal)` : undefined}>
          <input type="number" min="0.01" step="0.01" required value={f.valor} onChange={(e) => upd({ valor: e.target.value })} />
        </Field>
        <Field label="Descrição" className="gf-span-2">
          <input value={f.descricao} onChange={(e) => upd({ descricao: e.target.value })} placeholder={f.tipo === "pro-labore" ? "Ex: Pró-labore mensal" : "Ex: Aporte inicial de capital"} />
        </Field>
        <div className="gf-form-actions gf-span-2">
          <button type="submit" className="gf-btn gf-btn-primary">{mov ? "Salvar alterações" : "Registrar movimentação"}</button>
          <span className="gf-hint">
            {f.tipo === "pro-labore"
              ? "Pró-labore não entra na divisão societária (retiradas), apenas reduz o saldo da conta."
              : "Não entra no resultado das obras nem nas entradas/saídas da empresa, mas atualiza o saldo da conta."}
          </span>
        </div>
      </form>
    </Modal>
  );
}

function ParticipacaoModal({ socios, onClose, onSave }) {
  const [list, setList] = useState(socios.map((s) => ({ ...s })));
  const total = list.reduce((s, p) => s + (Number(p.participacao) || 0), 0);
  const upd = (id, patch) => setList((prev) => prev.map((p) => (p.id === id ? { ...p, ...patch } : p)));
  const addSocio = () => setList((prev) => [...prev, { id: uid("socio"), nome: "", participacao: 0 }]);
  const removeSocio = (id) => setList((prev) => prev.filter((p) => p.id !== id));
  const submit = (e) => {
    e.preventDefault();
    if (total !== 100) return;
    if (list.some((p) => !p.nome.trim())) return;
    onSave(list.map((p) => ({ ...p, participacao: Number(p.participacao) || 0 })));
  };
  return (
    <Modal title="Participação societária" onClose={onClose}>
      <form onSubmit={submit} className="gf-modal-form" style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {list.map((p) => (
          <div key={p.id} style={{ display: "flex", gap: 8, alignItems: "flex-end" }}>
            <Field label="Nome do sócio"><input required value={p.nome} onChange={(e) => upd(p.id, { nome: e.target.value })} /></Field>
            <Field label="Participação (%)"><input type="number" min="0" max="100" step="0.1" required value={p.participacao} onChange={(e) => upd(p.id, { participacao: e.target.value })} /></Field>
            {list.length > 1 && <button type="button" className="gf-icon-btn danger" onClick={() => removeSocio(p.id)}><Trash2 size={14} /></button>}
          </div>
        ))}
        <button type="button" className="gf-btn gf-btn-ghost" style={{ alignSelf: "flex-start" }} onClick={addSocio}><Plus size={14} /> Adicionar sócio</button>
        <div className={`gf-hint ${total !== 100 ? "gf-warn" : ""}`}>
          Soma atual: {total.toFixed(1)}% {total !== 100 && "— a soma precisa ser exatamente 100% para salvar."}
        </div>
        <div className="gf-form-actions">
          <button type="submit" className="gf-btn gf-btn-primary" disabled={total !== 100}>Salvar participação</button>
        </div>
      </form>
    </Modal>
  );
}

/* ---------------------------------------------------------------
   CSS
--------------------------------------------------------------- */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');

:root{
  --gf-bg:#F3F5F1;
  --gf-surface:#FFFFFF;
  --gf-ink:#1B2420;
  --gf-muted:#6B7570;
  --gf-primary:#0E3B2E;
  --gf-primary-light:#1D5A45;
  --gf-accent:#2F8F5B;
  --gf-accent-soft:#E4F2E9;
  --gf-danger:#B23A34;
  --gf-danger-soft:#F7E7E5;
  --gf-border:#DEE3DB;
  --gf-gray-100:#F0F1EE;
  --gf-font-display:'Space Grotesk', sans-serif;
  --gf-font-body:'Inter', sans-serif;
  --gf-font-mono:'IBM Plex Mono', monospace;
}
.gf-app{
  display:flex; min-height:100%; background:var(--gf-bg); color:var(--gf-ink);
  font-family:var(--gf-font-body); font-size:14px; line-height:1.45;
}
.gf-app *{ box-sizing:border-box; }
.gf-loading{ align-items:center; justify-content:center; padding:60px; }

/* Sidebar */
.gf-sidebar{
  width:216px; flex:0 0 216px; background:var(--gf-primary); color:#EFF4F0;
  display:flex; flex-direction:column; padding:20px 14px; gap:22px; min-height:100vh;
}
.gf-brand{ display:flex; flex-direction:column; align-items:flex-start; gap:6px; padding:2px 6px 4px; }
.gf-brand-logo{ width:100%; max-width:180px; height:auto; display:block; }
.gf-brand-mark{
  width:34px; height:34px; border-radius:8px; background:var(--gf-accent); color:#08201A;
  display:flex; align-items:center; justify-content:center; font-family:var(--gf-font-display);
  font-weight:700; font-size:13px; letter-spacing:.5px;
}
.gf-brand-name{ font-family:var(--gf-font-display); font-weight:700; font-size:15px; letter-spacing:1.5px; }
.gf-brand-sub{ font-size:10.5px; color:#9FC2AC; text-transform:uppercase; letter-spacing:1.5px; padding-left:2px; }
.gf-nav{ display:flex; flex-direction:column; gap:2px; }
.gf-nav-item{
  display:flex; align-items:center; gap:10px; padding:9px 12px; border-radius:8px; border:none;
  background:transparent; color:#CBDCD1; font-family:var(--gf-font-body); font-size:13.5px; font-weight:500;
  cursor:pointer; text-align:left;
}
.gf-nav-item:hover{ background:rgba(255,255,255,.06); color:#fff; }
.gf-nav-item.is-active{ background:var(--gf-accent); color:#08201A; font-weight:600; }
.gf-sidebar-foot{ margin-top:auto; font-size:11px; color:#7FA592; line-height:1.5; padding:0 6px; }

/* Main / TopBar */
.gf-main{ flex:1; min-width:0; display:flex; flex-direction:column; }
.gf-topbar{ background:var(--gf-surface); border-bottom:1px solid var(--gf-border); padding:20px 28px 18px; }
.gf-topbar-row{ display:flex; align-items:flex-end; justify-content:space-between; gap:16px; margin-bottom:16px; }
.gf-eyebrow{ font-size:11px; text-transform:uppercase; letter-spacing:1.2px; color:var(--gf-muted); font-weight:600; margin-bottom:2px; }
.gf-topbar h1{ font-family:var(--gf-font-display); font-size:22px; font-weight:700; margin:0; color:var(--gf-primary); }
.gf-topbar-actions{ display:flex; align-items:center; gap:12px; }
.gf-saved{ font-size:11.5px; color:var(--gf-muted); }
.gf-menu-wrap{ position:relative; }
.gf-menu{ position:absolute; right:0; top:calc(100% + 6px); background:#fff; border:1px solid var(--gf-border); border-radius:10px; box-shadow:0 8px 24px rgba(14,59,46,.12); display:flex; flex-direction:column; min-width:220px; z-index:20; overflow:hidden; }
.gf-menu button{ padding:10px 14px; text-align:left; background:none; border:none; font-size:13px; cursor:pointer; font-family:var(--gf-font-body); color:var(--gf-ink); }
.gf-menu button:hover{ background:var(--gf-gray-100); }

.gf-tiles{ display:grid; grid-template-columns:repeat(6,1fr); gap:10px; }
.gf-tile{ background:var(--gf-gray-100); border:1px solid var(--gf-border); border-radius:10px; padding:12px 14px; display:flex; flex-direction:column; gap:4px; min-width:0; }
.gf-tile-primary{ background:var(--gf-primary); border-color:var(--gf-primary); }
.gf-tile-primary .gf-tile-label{ color:#9FC2AC; }
.gf-tile-primary .gf-tile-value{ color:#fff; }
.gf-tile-label{ font-size:11px; color:var(--gf-muted); font-weight:500; }
.gf-tile-value{ font-family:var(--gf-font-mono); font-weight:600; font-size:16px; }
.gf-delta{ display:inline-flex; align-items:center; gap:2px; font-size:10.5px; color:var(--gf-muted); }
.gf-delta.up{ color:var(--gf-accent); }
.gf-delta.down{ color:var(--gf-danger); }

.gf-page{ padding:22px 28px 60px; }
.gf-stack{ display:flex; flex-direction:column; gap:18px; max-width:1240px; }

/* Cards */
.gf-card{ background:var(--gf-surface); border:1px solid var(--gf-border); border-radius:12px; padding:18px 20px; }
.gf-card-head{ display:flex; align-items:center; justify-content:space-between; gap:10px; margin-bottom:14px; flex-wrap:wrap; }
.gf-card-head h2{ font-family:var(--gf-font-display); font-size:16px; font-weight:600; color:var(--gf-primary); margin:0; }
.gf-card-head h3{ font-family:var(--gf-font-display); font-size:14.5px; font-weight:600; margin:0; }
.gf-toolbar{ display:flex; align-items:center; justify-content:space-between; }
.gf-toolbar h2{ font-family:var(--gf-font-display); font-size:19px; color:var(--gf-primary); margin:0; }
.gf-toolbar-actions{ display:flex; gap:10px; }

/* Buttons / inputs */
.gf-btn{ display:inline-flex; align-items:center; gap:7px; padding:9px 14px; border-radius:8px; font-size:13px; font-weight:600; border:1px solid transparent; cursor:pointer; font-family:var(--gf-font-body); white-space:nowrap; }
.gf-btn-primary{ background:var(--gf-accent); color:#fff; }
.gf-btn-primary:hover{ background:#277A4C; }
.gf-btn-ghost{ background:#fff; color:var(--gf-primary); border-color:var(--gf-border); }
.gf-btn-ghost:hover{ background:var(--gf-gray-100); }
.gf-icon-btn{ display:inline-flex; align-items:center; justify-content:center; width:28px; height:28px; border-radius:7px; border:1px solid var(--gf-border); background:#fff; color:var(--gf-muted); cursor:pointer; }
.gf-icon-btn:hover{ background:var(--gf-gray-100); color:var(--gf-ink); }
.gf-icon-btn.danger{ color:var(--gf-danger); border-color:var(--gf-danger); }

label.gf-field{ display:flex; flex-direction:column; gap:5px; }
.gf-field-label{ font-size:11.5px; font-weight:600; color:var(--gf-muted); text-transform:uppercase; letter-spacing:.4px; }
.gf-field-hint{ font-size:11px; color:var(--gf-muted); }
input, select{
  font-family:var(--gf-font-body); font-size:13.5px; padding:9px 10px; border-radius:8px;
  border:1px solid var(--gf-border); background:#fff; color:var(--gf-ink); width:100%;
}
input:focus, select:focus{ outline:2px solid var(--gf-accent); outline-offset:1px; border-color:var(--gf-accent); }
.gf-span-2{ grid-column:span 2; }
.gf-inline-select{ display:flex; gap:6px; align-items:center; }
.gf-inline-select select, .gf-inline-select input{ flex:1; }

/* Entry form */
.gf-entry-card{ border-top:3px solid var(--gf-accent); }
.gf-segmented{ display:inline-flex; border:1px solid var(--gf-border); border-radius:8px; overflow:hidden; margin-bottom:14px; }
.gf-segmented button{ display:flex; align-items:center; gap:6px; padding:8px 16px; border:none; background:#fff; font-size:13px; font-weight:600; cursor:pointer; color:var(--gf-muted); }
.gf-segmented button.is-active.pos{ background:var(--gf-accent-soft); color:var(--gf-accent); }
.gf-segmented button.is-active.neg{ background:var(--gf-danger-soft); color:var(--gf-danger); }
.gf-segmented-3 button.is-active{ background:var(--gf-gray-100); color:var(--gf-primary); }
.gf-segmented-3 button{ flex:1; justify-content:center; }

/* Balanço entre sócios */
.gf-balance-card{ border-top:3px solid var(--gf-border); }
.gf-balance-card.is-unbalanced{ border-top-color:var(--gf-danger); }
.gf-balance-card.is-balanced{ border-top-color:var(--gf-accent); }
.gf-balance-banner{ display:flex; flex-direction:column; gap:4px; background:var(--gf-danger-soft); border-radius:10px; padding:12px 14px; }
.gf-balance-banner strong{ color:var(--gf-danger); font-family:var(--gf-font-display); font-size:14px; }
.gf-balance-banner span{ font-size:13px; color:var(--gf-ink); }
.gf-balance-banner.is-ok{ background:var(--gf-accent-soft); }
.gf-balance-banner.is-ok strong{ color:var(--gf-accent); }
.gf-badge-blue{ background:#E4ECF2; color:#2A5A82; }
.gf-badge-orange{ background:#FBE8DA; color:#B0530E; }
.gf-badge-yellow{ background:#FBF1D2; color:#8A6D06; }

.gf-tile.tone-red{ background:#F7E7E5; border-color:#F0D2CE; }
.gf-tile.tone-orange{ background:#FBE8DA; border-color:#F3D6BC; }
.gf-tile.tone-yellow{ background:#FBF1D2; border-color:#F0E3AC; }
.gf-tile.tone-blue{ background:#E4ECF2; border-color:#CBDBE8; }
.gf-tile.tone-green{ background:var(--gf-accent-soft); border-color:#CCE7D8; }
.gf-tiles-cp{ grid-template-columns:repeat(6,1fr); }
.gf-filters-grid-cp{ grid-template-columns:repeat(4,1fr); }

.gf-rateio-toggle{ margin-bottom:10px; }
.gf-checkbox{ display:inline-flex; align-items:center; gap:8px; font-size:13px; font-weight:600; color:var(--gf-ink); cursor:pointer; }
.gf-checkbox input{ width:auto; accent-color:var(--gf-accent); }
.gf-rateio-box{ background:var(--gf-gray-100); border:1px solid var(--gf-border); border-radius:10px; padding:12px; display:flex; flex-direction:column; gap:8px; }
.gf-rateio-row{ display:grid; grid-template-columns:1.6fr 1fr auto auto; gap:8px; align-items:center; }
.gf-rateio-pct{ font-size:12.5px; color:var(--gf-muted); min-width:48px; text-align:right; }
.gf-rateio-total{ display:flex; justify-content:space-between; align-items:center; font-size:12.5px; margin-top:4px; padding-top:8px; border-top:1px dashed var(--gf-border); }
.gf-rateio-total.is-ok b{ color:var(--gf-accent); }
.gf-rateio-total.is-warn b{ color:var(--gf-danger); }
.gf-rateio-pill{
  display:inline-flex; align-items:center; gap:5px; padding:4px 9px; border-radius:100px; border:1px solid var(--gf-border);
  background:#fff; font-size:11.5px; font-weight:600; color:var(--gf-primary); cursor:pointer;
}
.gf-rateio-pill:hover{ background:var(--gf-gray-100); }

.gf-perm-grid{ display:grid; grid-template-columns:repeat(2,1fr); gap:6px 14px; }
.gf-perm-item{ font-weight:500 !important; font-size:12.5px !important; }
.gf-locked-field{
  display:flex; align-items:center; justify-content:space-between; padding:9px 10px; border-radius:8px;
  border:1px solid var(--gf-border); background:var(--gf-gray-100); color:var(--gf-muted); cursor:not-allowed;
}
.gf-sidebar-user{ display:flex; align-items:center; gap:9px; padding-top:12px; border-top:1px solid rgba(255,255,255,.12); margin-top:4px; }
.gf-sidebar-avatar{ width:30px; height:30px; border-radius:100px; background:var(--gf-accent); color:#08201A; display:flex; align-items:center; justify-content:center; font-weight:700; font-size:13px; flex:0 0 auto; }
.gf-sidebar-user-name{ font-size:12.5px; color:#EFF4F0; font-weight:600; }
.gf-sidebar-user-role{ font-size:10.5px; color:#9FC2AC; }
.gf-topbar-user{ display:flex; align-items:center; gap:8px; padding-left:10px; border-left:1px solid var(--gf-border); }
.gf-topbar-user-name{ font-size:12.5px; font-weight:600; color:var(--gf-primary); }

.gf-login{ flex:1; display:flex; align-items:center; justify-content:center; min-height:100vh; background:var(--gf-bg); padding:20px; }
.gf-login-card{ background:#fff; border:1px solid var(--gf-border); border-radius:16px; padding:28px; width:100%; max-width:400px; box-shadow:0 20px 50px rgba(14,59,46,.08); }
.gf-login-users{ display:flex; flex-direction:column; gap:8px; }
.gf-login-user{
  display:flex; align-items:center; gap:10px; padding:10px 12px; border-radius:10px; border:1px solid var(--gf-border);
  background:#fff; cursor:pointer; text-align:left; width:100%;
}
.gf-login-user:hover{ background:var(--gf-gray-100); }
.gf-login-user.is-selected{ border-color:var(--gf-accent); background:var(--gf-accent-soft); }
.gf-login-user-name{ font-size:13.5px; font-weight:600; color:var(--gf-ink); }
.gf-login-user-role{ font-size:11px; color:var(--gf-muted); }
.gf-login-logo-tile{ background:var(--gf-primary); border-radius:12px; padding:16px 20px; display:inline-flex; margin-bottom:10px; }
.gf-login-logo-img{ width:170px; height:auto; display:block; }

/* Controle orçamentário */
.gf-orcamento-chip{ display:flex; align-items:center; gap:6px; font-size:11.5px; color:var(--gf-muted); margin-top:8px; font-weight:600; }
.gf-dot{ width:9px; height:9px; border-radius:100px; display:inline-block; flex:0 0 auto; }
.gf-dot-green{ background:var(--gf-accent); }
.gf-dot-orange{ background:#C98A2B; }
.gf-dot-red{ background:var(--gf-danger); }
.gf-budget-alert{
  display:flex; gap:10px; align-items:flex-start; background:var(--gf-danger-soft); border:1px solid #EFC6C1;
  border-radius:10px; padding:12px 14px; margin-bottom:14px; color:var(--gf-danger);
}
.gf-budget-alert.is-atencao{ background:#FBF1D2; border-color:#F0E3AC; color:#8A6D06; }
.gf-budget-alert svg{ flex:0 0 auto; margin-top:1px; }
.gf-budget-alert strong{ display:block; font-size:13px; font-weight:600; }
.gf-budget-alert span{ display:block; font-size:12px; margin-top:2px; color:var(--gf-ink); }
.gf-table-orcamento td, .gf-table-orcamento th{ white-space:nowrap; }

.gf-import-upload{ padding:4px 0; }
.gf-import-dropzone{
  display:flex; flex-direction:column; align-items:center; justify-content:center; gap:8px; text-align:center;
  border:2px dashed var(--gf-border); border-radius:12px; padding:40px 20px; cursor:pointer; color:var(--gf-muted);
}
.gf-import-dropzone:hover{ background:var(--gf-gray-100); border-color:var(--gf-accent); }
.gf-import-dropzone strong{ color:var(--gf-ink); font-size:14px; }
.gf-import-dropzone span{ font-size:12px; max-width:340px; }

.gf-classificacao-box{ background:var(--gf-gray-100); border:1px solid var(--gf-border); border-radius:10px; padding:14px; margin-top:2px; }
.gf-mini-tag{ font-size:11px; color:var(--gf-muted); font-weight:600; }
.gf-dvi-columns{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }
@media (max-width: 900px){ .gf-dvi-columns{ grid-template-columns:1fr; } }
.gf-form-grid{ display:grid; grid-template-columns:repeat(4,1fr); gap:14px; }
.gf-modal-form{ grid-template-columns:repeat(2,1fr); }
.gf-form-actions{ display:flex; align-items:center; gap:14px; margin-top:16px; }
.gf-hint{ font-size:12px; color:var(--gf-muted); }
.gf-warn{ font-size:12px; color:var(--gf-danger); }

/* Summary inline */
.gf-summary-inline{ display:flex; gap:18px; font-size:13px; color:var(--gf-muted); flex-wrap:wrap; }
.gf-summary-inline b{ font-family:var(--gf-font-mono); margin-left:4px; }

/* Filters */
.gf-filters{ display:flex; flex-direction:column; gap:10px; margin-bottom:14px; }
.gf-search{ display:flex; align-items:center; gap:8px; border:1px solid var(--gf-border); border-radius:8px; padding:0 10px; background:var(--gf-gray-100); }
.gf-search input{ border:none; background:transparent; padding:9px 0; }
.gf-search input:focus{ outline:none; }
.gf-filters-grid{ display:grid; grid-template-columns:repeat(6,1fr); gap:8px; }
.gf-fornecedor-filters{ grid-template-columns:repeat(5,1fr); margin-bottom:14px; }

/* Table */
.gf-table-wrap{ overflow:auto; }
.gf-table{ width:100%; border-collapse:collapse; font-size:13px; }
.gf-table th{ text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:.4px; color:var(--gf-muted); padding:8px 10px; border-bottom:1px solid var(--gf-border); white-space:nowrap; }
.gf-table td{ padding:9px 10px; border-bottom:1px solid var(--gf-border); vertical-align:middle; }
.gf-table td.num, .gf-table th.num{ text-align:right; }
.gf-td-desc{ max-width:220px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.gf-table-ledger tr.row-pos{ border-left:3px solid var(--gf-accent); }
.gf-table-ledger tr.row-neg{ border-left:3px solid var(--gf-danger); }
.gf-row-actions{ display:flex; gap:5px; justify-content:flex-end; }
.gf-row-clickable{ cursor:pointer; }
.gf-row-clickable:hover{ background:var(--gf-gray-100); }
.mono{ font-family:var(--gf-font-mono); }
.pos{ color:var(--gf-accent); }
.neg{ color:var(--gf-danger); }

/* Badges / chips */
.gf-badge{ display:inline-block; padding:3px 9px; border-radius:100px; font-size:11px; font-weight:600; }
.gf-badge-green{ background:var(--gf-accent-soft); color:var(--gf-accent); }
.gf-badge-red{ background:var(--gf-danger-soft); color:var(--gf-danger); }
.gf-badge-gray{ background:var(--gf-gray-100); color:var(--gf-muted); }
.gf-chip-row{ display:flex; gap:5px; flex-wrap:wrap; }
.gf-chip{ background:var(--gf-gray-100); border-radius:6px; padding:3px 8px; font-size:11px; color:var(--gf-ink); }
.gf-chip-muted{ color:var(--gf-muted); }

/* Grid cards (accounts/obras) */
.gf-grid-cards{ display:grid; grid-template-columns:repeat(auto-fill, minmax(270px,1fr)); gap:14px; }
.gf-account-card.is-inactive{ opacity:.55; }
.gf-account-meta{ font-size:12px; color:var(--gf-muted); margin-bottom:10px; }
.gf-account-rows{ display:flex; flex-direction:column; gap:6px; padding:10px 0; border-top:1px solid var(--gf-border); border-bottom:1px solid var(--gf-border); }
.gf-account-rows div{ display:flex; justify-content:space-between; font-size:12.5px; }
.gf-account-rows span{ color:var(--gf-muted); }
.gf-account-balance{ display:flex; justify-content:space-between; align-items:center; padding-top:10px; }
.gf-account-balance span{ font-size:12px; color:var(--gf-muted); font-weight:600; }
.gf-account-balance strong{ font-size:17px; }

.gf-obra-card{ text-align:left; cursor:pointer; border:1px solid var(--gf-border); font-family:inherit; position:relative; }
.gf-obra-card-static{ cursor:default; }
.gf-card-cta{ display:inline-flex; align-items:center; gap:2px; font-size:12px; color:var(--gf-accent); font-weight:600; margin-top:10px; }

.gf-chart-box{ padding-top:4px; }
.gf-chart-title{ font-family:var(--gf-font-display); font-size:13px; font-weight:600; color:var(--gf-primary); margin:18px 0 6px; }

/* Modal */
.gf-modal-backdrop{ position:fixed; inset:0; background:rgba(14,20,17,.45); display:flex; align-items:center; justify-content:center; z-index:50; padding:20px; }
.gf-modal{ background:#fff; border-radius:14px; width:100%; max-height:88vh; overflow:auto; box-shadow:0 24px 60px rgba(0,0,0,.25); }
.gf-modal-head{ display:flex; align-items:center; justify-content:space-between; padding:16px 20px; border-bottom:1px solid var(--gf-border); position:sticky; top:0; background:#fff; z-index:2; }
.gf-modal-head h3{ font-family:var(--gf-font-display); font-size:16px; margin:0; color:var(--gf-primary); }
.gf-modal-body{ padding:18px 20px 22px; }

.gf-obra-detail{ display:flex; flex-direction:column; }
.gf-detail-meta{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; padding-bottom:14px; border-bottom:1px solid var(--gf-border); margin-bottom:12px; align-items:end; }
.gf-detail-meta div span{ display:block; font-size:11px; color:var(--gf-muted); }
.gf-detail-edit{ grid-column:span 4; justify-self:start; }
.gf-detail-stats{ border:none; padding:0 0 4px; }

.gf-fornecedor-summary{ display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin-top:16px; padding-top:14px; border-top:1px solid var(--gf-border); }
.gf-summary-block{ display:flex; flex-direction:column; gap:6px; }
.gf-summary-block > span{ font-size:11px; text-transform:uppercase; letter-spacing:.4px; color:var(--gf-muted); font-weight:600; }
.gf-summary-line{ display:flex; justify-content:space-between; font-size:12.5px; }
.gf-summary-block strong{ font-size:17px; }

.gf-empty{ display:flex; flex-direction:column; align-items:center; justify-content:center; gap:6px; padding:36px 10px; color:var(--gf-muted); text-align:center; }
.gf-empty strong{ color:var(--gf-ink); font-size:13.5px; }
.gf-empty p{ margin:0; font-size:12.5px; max-width:320px; }

@media (max-width: 900px){
  .gf-tiles{ grid-template-columns:repeat(3,1fr); }
  .gf-form-grid{ grid-template-columns:repeat(2,1fr); }
  .gf-filters-grid{ grid-template-columns:repeat(3,1fr); }
  .gf-detail-meta{ grid-template-columns:repeat(2,1fr); }
  .gf-detail-edit{ grid-column:span 2; }
  .gf-fornecedor-summary{ grid-template-columns:1fr; }
}
@media (max-width: 640px){
  .gf-app{ flex-direction:column; }
  .gf-sidebar{ width:100%; flex-direction:row; align-items:center; min-height:auto; padding:12px 14px; overflow-x:auto; }
  .gf-nav{ flex-direction:row; }
  .gf-sidebar-foot{ display:none; }
  .gf-tiles{ grid-template-columns:repeat(2,1fr); }
  .gf-form-grid{ grid-template-columns:1fr; }
  .gf-modal-form{ grid-template-columns:1fr; }
  .gf-span-2{ grid-column:span 1; }
  .gf-filters-grid{ grid-template-columns:repeat(2,1fr); }
  .gf-page{ padding:16px; }
  .gf-topbar{ padding:16px; }
}
`;
