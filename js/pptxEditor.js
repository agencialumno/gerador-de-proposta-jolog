/**
 * Edita a tabela de preços de um .pptx inteiramente via manipulação de string
 * no XML do slide (sem precisar de servidor / Python).
 *
 * Como funciona:
 * 1. O .pptx é um .zip contendo XMLs. A tabela de preços fica em
 *    ppt/slides/slideN.xml, dentro de uma tag <a:tbl>.
 * 2. Cada linha da tabela é um bloco <a:tr>...</a:tr> com exatamente
 *    3 blocos <a:t>...</a:t> (Função, Quantidade, Valores).
 * 3. Usamos a primeira linha (cabeçalho) como está, a segunda linha como
 *    "molde" de estilo pra clonar pra cada serviço, e a última como molde
 *    da linha de Total.
 */

const SLIDE_FILE = "ppt/slides/slide12.xml"; // slide "IV - Do Preço"

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function formatarBRL(valor) {
  const num = Number(valor);
  const s = num.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return `R$ ${s}`;
}

/**
 * Substitui, em ordem, o conteúdo dos N blocos <a:t> de uma linha (a:tr) por
 * novos textos. Preserva toda a formatação (rPr, tcPr, bordas, etc.), só troca
 * o texto visível.
 */
function substituirTextosDaLinha(rowXml, novosTextos) {
  let i = 0;
  return rowXml.replace(/<a:t>(.*?)<\/a:t>/g, () => {
    const texto = novosTextos[i] ?? "";
    i += 1;
    return `<a:t>${escapeXml(texto)}</a:t>`;
  });
}

function extrairAltura(rowXml) {
  const m = rowXml.match(/<a:tr[^>]*\sh="(\d+)"/);
  return m ? parseInt(m[1], 10) : 300000;
}

function ajustarAltura(rowXml, novaAltura) {
  return rowXml.replace(/(<a:tr[^>]*\sh=")(\d+)(")/, `$1${novaAltura}$3`);
}

/**
 * @param {ArrayBuffer} templateArrayBuffer - conteúdo binário do modelo .pptx
 * @param {Array<{funcao: string, quantidade: number, valor: number}>} servicos
 * @param {JSZip} JSZipLib - referência à lib JSZip (window.JSZip no navegador)
 * @returns {Promise<Blob|Buffer>} pptx pronto
 */
async function gerarProposta(templateArrayBuffer, servicos, JSZipLib) {
  if (!servicos || servicos.length === 0) {
    throw new Error("Lista de serviços vazia.");
  }

  const zip = await JSZipLib.loadAsync(templateArrayBuffer);
  const slideFile = zip.file(SLIDE_FILE);
  if (!slideFile) {
    throw new Error(`Slide de preço não encontrado em ${SLIDE_FILE}`);
  }

  const xml = await slideFile.async("string");

  // Extrai todas as linhas <a:tr>...</a:tr> da tabela
  const rowRegex = /<a:tr[^>]*>[\s\S]*?<\/a:tr>/g;
  const rows = xml.match(rowRegex);
  if (!rows || rows.length < 3) {
    throw new Error("Tabela de preço não encontrada ou com estrutura inesperada.");
  }

  const headerRow = rows[0];
  const templateServiceRow = rows[1];
  const templateTotalRow = rows[rows.length - 1];

  // Área original ocupada pela tabela (header + linhas de serviço + total),
  // usada pra recalcular a altura de cada linha e não sobrepor o texto abaixo
  // quando o número de serviços é diferente do template original.
  const alturaHeaderOriginal = extrairAltura(headerRow);
  const alturaLinhaOriginal = extrairAltura(templateServiceRow);
  const qtdLinhasOriginais = rows.length - 2; // exclui header e total
  const alturaAreaTotal =
    alturaHeaderOriginal + alturaLinhaOriginal * qtdLinhasOriginais + extrairAltura(templateTotalRow);

  const novaAlturaLinha = Math.max(
    280000, // ~0,3cm mínimo, pra não ficar ilegível
    Math.floor((alturaAreaTotal - alturaHeaderOriginal * 2) / (servicos.length + 1))
  );

  // Monta as novas linhas de serviço a partir do molde de estilo
  let total = 0;
  let totalQtd = 0;
  const novasLinhasServico = servicos.map((s) => {
    const qtd = Number(s.quantidade) || 0;
    const valor = Number(s.valor) || 0;
    total += valor;
    totalQtd += qtd;
    const linha = substituirTextosDaLinha(templateServiceRow, [
      String(s.funcao).toUpperCase(),
      String(qtd),
      formatarBRL(valor),
    ]);
    return ajustarAltura(linha, novaAlturaLinha);
  });

  let novaLinhaTotal = substituirTextosDaLinha(templateTotalRow, [
    "Total",
    String(totalQtd),
    formatarBRL(total),
  ]);
  novaLinhaTotal = ajustarAltura(novaLinhaTotal, novaAlturaLinha);

  const novasLinhas = [headerRow, ...novasLinhasServico, novaLinhaTotal].join("");

  // Substitui o bloco de linhas original (do início do primeiro <a:tr> ao fim do último </a:tr>)
  const primeiroIdx = xml.indexOf(rows[0]);
  const ultimoIdx = xml.lastIndexOf(rows[rows.length - 1]) + rows[rows.length - 1].length;

  const novoXml = xml.slice(0, primeiroIdx) + novasLinhas + xml.slice(ultimoIdx);

  zip.file(SLIDE_FILE, novoXml);

  return zip.generateAsync({
    type: typeof window === "undefined" ? "nodebuffer" : "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

if (typeof module !== "undefined") {
  module.exports = { gerarProposta, formatarBRL };
}
