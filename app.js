// Caminho do modelo .pptx dentro do próprio repositório (GitHub Pages serve isso como arquivo estático)
const TEMPLATE_URL = "template/modelo_proposta.pptx";

const corpoServicos = document.getElementById("corpoServicos");
const totalQtdEl = document.getElementById("totalQtd");
const totalValorEl = document.getElementById("totalValor");

function formatarBRLExibicao(valor) {
  return valor.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function renderEstadoVazio() {
  corpoServicos.innerHTML = `
    <tr class="linha-vazia">
      <td colspan="4" class="empty-state">
        Nenhum serviço adicionado ainda. Clique em <strong>+ Adicionar serviço</strong> para começar.
      </td>
    </tr>`;
}

function criarLinhaServico() {
  const linhaVazia = corpoServicos.querySelector(".linha-vazia");
  if (linhaVazia) linhaVazia.remove();

  const tr = document.createElement("tr");

  const tdFuncao = document.createElement("td");
  const inputFuncao = document.createElement("input");
  inputFuncao.type = "text";
  inputFuncao.placeholder = "Ex: Porteiro Diurno";
  tdFuncao.appendChild(inputFuncao);

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
  inputValor.value = "0";
  tdValor.appendChild(inputValor);

  const tdRemover = document.createElement("td");
  const btnRemover = document.createElement("button");
  btnRemover.type = "button";
  btnRemover.textContent = "Remover";
  btnRemover.className = "btn-remover";
  btnRemover.addEventListener("click", () => {
    tr.remove();
    if (!corpoServicos.querySelector("tr")) renderEstadoVazio();
    atualizarTotal();
  });
  tdRemover.appendChild(btnRemover);

  tr.append(tdFuncao, tdQtd, tdValor, tdRemover);

  inputQtd.addEventListener("input", atualizarTotal);
  inputValor.addEventListener("input", atualizarTotal);

  corpoServicos.appendChild(tr);
  atualizarTotal();
}

function atualizarTotal() {
  let totalQtd = 0;
  let totalValor = 0;

  corpoServicos.querySelectorAll("tr").forEach((tr) => {
    if (tr.classList.contains("linha-vazia")) return;
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
    if (tr.classList.contains("linha-vazia")) return;
    const inputFuncao = tr.querySelector('input[type="text"]');
    const inputs = tr.querySelectorAll('input[type="number"]');
    const quantidade = parseFloat(inputs[0]?.value) || 0;
    const valorUnit = parseFloat(inputs[1]?.value) || 0;

    servicos.push({
      funcao: inputFuncao.value,
      quantidade,
      valor: quantidade * valorUnit, // valor TOTAL da linha
    });
  });
  return servicos;
}

document.getElementById("btnAddServico").addEventListener("click", () => criarLinhaServico());

// Guarda o modelo já baixado nessa sessão, pra não baixar de novo a cada proposta gerada
let templateEmCache = null;

document.getElementById("btnGerar").addEventListener("click", async () => {
  const statusEl = document.getElementById("status");
  const cliente = document.getElementById("clienteNome").value || "Condominio";
  const servicos = coletarServicos();

  if (servicos.length === 0) {
    statusEl.textContent = "Adicione pelo menos um serviço.";
    return;
  }

  try {
    let templateBuffer;

    if (templateEmCache) {
      templateBuffer = templateEmCache;
      statusEl.textContent = "Montando a proposta...";
    } else {
      statusEl.textContent = "Carregando modelo... 0%";

      const resp = await fetch(TEMPLATE_URL);
      if (!resp.ok) throw new Error("Não foi possível carregar o modelo .pptx");

      const tamanhoTotal = parseInt(resp.headers.get("Content-Length") || "0", 10);
      const reader = resp.body.getReader();
      const pedacos = [];
      let recebido = 0;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        pedacos.push(value);
        recebido += value.length;
        if (tamanhoTotal) {
          const pct = Math.round((recebido / tamanhoTotal) * 100);
          statusEl.textContent = `Carregando modelo... ${pct}%`;
        } else {
          statusEl.textContent = `Carregando modelo... ${(recebido / 1024 / 1024).toFixed(1)}MB`;
        }
      }

      templateBuffer = await new Blob(pedacos).arrayBuffer();
      templateEmCache = templateBuffer;
      statusEl.textContent = "Montando a proposta...";
    }

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

// Começa sem nenhum serviço, mostrando a instrução de estado vazio
renderEstadoVazio();