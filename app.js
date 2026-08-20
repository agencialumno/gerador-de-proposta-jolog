// Caminho do modelo .pptx dentro do próprio repositório (GitHub Pages serve isso como arquivo estático)
const TEMPLATE_URL = "template/modelo_proposta.pptx";

// Catálogo de funções pré-cadastradas (ajuste os valores/lista conforme sua tabela real)
const CATALOGO_SERVICOS = [
  { funcao: "PORTEIROS DIURNOS", valorUnitario: 4894.14 },
  { funcao: "PORTEIROS NOTURNOS", valorUnitario: 5684.72 },
  { funcao: "AUX DE SERVIÇOS GERAIS 6X1 S/INSAL", valorUnitario: 5088.29 },
  { funcao: "ASSISTENTE ADMINI 6X1", valorUnitario: 5659.01 },
  { funcao: "VIGIA", valorUnitario: 0 },
  { funcao: "JARDINAGEM", valorUnitario: 0 },
  { funcao: "MANUTENÇÃO PREDIAL", valorUnitario: 0 },
];

const corpoServicos = document.getElementById("corpoServicos");
const totalQtdEl = document.getElementById("totalQtd");
const totalValorEl = document.getElementById("totalValor");

function formatarBRLExibicao(valor) {
  return valor.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function criarLinhaServico() {
  const tr = document.createElement("tr");

  const tdFuncao = document.createElement("td");
  const select = document.createElement("select");
  CATALOGO_SERVICOS.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s.funcao;
    opt.textContent = s.funcao;
    opt.dataset.valor = s.valorUnitario;
    select.appendChild(opt);
  });
  const optCustom = document.createElement("option");
  optCustom.value = "__custom__";
  optCustom.textContent = "Personalizado...";
  select.appendChild(optCustom);
  tdFuncao.appendChild(select);

  const inputFuncaoCustom = document.createElement("input");
  inputFuncaoCustom.type = "text";
  inputFuncaoCustom.placeholder = "Nome da função";
  inputFuncaoCustom.style.display = "none";
  tdFuncao.appendChild(inputFuncaoCustom);

  const tdQtd = document.createElement("td");
  const inputQtd = document.createElement("input");
  inputQtd.type = "number";
  inputQtd.min = "1";
  inputQtd.value = "1";
  tdQtd.appendChild(inputQtd);

  const tdValor = document.createElement("td");
  const inputValor = document.createElement("input");
  inputValor.type = "number";
  inputValor.step = "0.01";
  inputValor.min = "0";
  inputValor.value = CATALOGO_SERVICOS[0].valorUnitario;
  tdValor.appendChild(inputValor);

  const tdRemover = document.createElement("td");
  const btnRemover = document.createElement("button");
  btnRemover.type = "button";
  btnRemover.textContent = "Remover";
  btnRemover.className = "btn-remover";
  btnRemover.addEventListener("click", () => {
    tr.remove();
    atualizarTotal();
  });
  tdRemover.appendChild(btnRemover);

  tr.append(tdFuncao, tdQtd, tdValor, tdRemover);

  select.addEventListener("change", () => {
    if (select.value === "__custom__") {
      inputFuncaoCustom.style.display = "inline-block";
      inputValor.value = 0;
    } else {
      inputFuncaoCustom.style.display = "none";
      const opt = select.selectedOptions[0];
      inputValor.value = opt.dataset.valor;
    }
    atualizarTotal();
  });
  inputQtd.addEventListener("input", atualizarTotal);
  inputValor.addEventListener("input", atualizarTotal);

  corpoServicos.appendChild(tr);
  atualizarTotal();
}

function atualizarTotal() {
  let totalQtd = 0;
  let totalValor = 0;

  corpoServicos.querySelectorAll("tr").forEach((tr) => {
    const inputs = tr.querySelectorAll('input[type="number"]');
    const quantidade = parseFloat(inputs[0]?.value) || 0;
    const valorUnit = parseFloat(inputs[1]?.value) || 0;
    totalQtd += quantidade;
    totalValor += quantidade * valorUnit;
  });

  totalQtdEl.textContent = totalQtd;
  totalValorEl.textContent = formatarBRLExibicao(totalValor);
}

function coletarServicos() {
  const servicos = [];
  corpoServicos.querySelectorAll("tr").forEach((tr) => {
    const select = tr.querySelector("select");
    const inputCustom = tr.querySelector('input[type="text"]');
    const inputs = tr.querySelectorAll('input[type="number"]');
    const quantidade = parseFloat(inputs[0]?.value) || 0;
    const valorUnit = parseFloat(inputs[1]?.value) || 0;
    const funcao = select.value === "__custom__" ? inputCustom.value : select.value;

    servicos.push({
      funcao,
      quantidade,
      valor: quantidade * valorUnit, // valor TOTAL da linha
    });
  });
  return servicos;
}

document.getElementById("btnAddServico").addEventListener("click", () => criarLinhaServico());

document.getElementById("btnGerar").addEventListener("click", async () => {
  const statusEl = document.getElementById("status");
  const cliente = document.getElementById("clienteNome").value || "Condominio";
  const servicos = coletarServicos();

  if (servicos.length === 0) {
    statusEl.textContent = "Adicione pelo menos um serviço.";
    return;
  }

  statusEl.textContent = "Gerando proposta...";

  try {
    const resp = await fetch(TEMPLATE_URL);
    if (!resp.ok) throw new Error("Não foi possível carregar o modelo .pptx");
    const templateBuffer = await resp.arrayBuffer();

    // gerarProposta vem de js/pptxEditor.js (carregado antes deste script no index.html)
    const blob = await gerarProposta(templateBuffer, servicos, JSZip);

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Proposta_Jolog_${cliente}.pptx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    statusEl.textContent = "Proposta gerada com sucesso!";
  } catch (e) {
    console.error(e);
    statusEl.textContent = "Erro: " + e.message;
  }
});

// Linha inicial
criarLinhaServico();
