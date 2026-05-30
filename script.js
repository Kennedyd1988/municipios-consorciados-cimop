
let dadosMunicipios = [];
let dadosFiltrados = [];
let paginaAtual = 1;
const itensPorPagina = 10;

async function carregarJSON(caminho){
  try{
    const resposta = await fetch(caminho);
    if(!resposta.ok) throw new Error(`Erro ao carregar ${caminho}`);
    return await resposta.json();
  }catch(erro){
    console.error(erro);
    return [];
  }
}

function preencherData(){
  const el = document.getElementById("dataHoje");
  if(el){
    const hoje = new Date();
    el.innerHTML = hoje.toLocaleDateString("pt-BR") + " às " + hoje.toLocaleTimeString("pt-BR", {
      hour:"2-digit",
      minute:"2-digit"
    });
  }
}

function filtrarTabela(){
  const termo = document.getElementById("campoBusca").value.toLowerCase();
  dadosFiltrados = dadosMunicipios.filter(item =>
    JSON.stringify(item).toLowerCase().includes(termo)
  );
  paginaAtual = 1;
  renderizarTabela();
}

function fatiaPagina(){
  const inicio = (paginaAtual - 1) * itensPorPagina;
  return dadosFiltrados.slice(inicio, inicio + itensPorPagina);
}

function atualizarPaginacao(){
  const info = document.getElementById("infoPagina");
  const anterior = document.getElementById("btnAnterior");
  const proxima = document.getElementById("btnProxima");
  const totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / itensPorPagina));

  info.innerText = `Página ${paginaAtual} de ${totalPaginas} — ${dadosFiltrados.length} município(s)`;
  anterior.disabled = paginaAtual <= 1;
  proxima.disabled = paginaAtual >= totalPaginas;
}

function paginaAnterior(){
  if(paginaAtual > 1){
    paginaAtual--;
    renderizarTabela();
  }
}

function proximaPagina(){
  const totalPaginas = Math.max(1, Math.ceil(dadosFiltrados.length / itensPorPagina));
  if(paginaAtual < totalPaginas){
    paginaAtual++;
    renderizarTabela();
  }
}

function botoesPdf(pdf){
  if(!pdf){
    return "PDF não localizado";
  }

  return `
    <a class="botao" href="${pdf}" target="_blank">Visualizar Lei</a>
    <a class="botao-secundario" href="${pdf}" download>Baixar PDF</a>
  `;
}

function renderizarTabela(){
  const tabela = document.getElementById("tabelaMunicipios");

  if(!dadosFiltrados.length){
    tabela.innerHTML = `<tr><td colspan="6">Nenhum município localizado.</td></tr>`;
    atualizarPaginacao();
    return;
  }

  tabela.innerHTML = fatiaPagina().map((item, index) => `
    <tr>
      <td>${((paginaAtual - 1) * itensPorPagina) + index + 1}</td>
      <td>${item.municipio}</td>
      <td>${item.lei}</td>
      <td>${item.dataSancao || "-"}</td>
      <td>${item.dataPublicacao || "-"}</td>
      <td>${botoesPdf(item.pdf)}</td>
    </tr>
  `).join("");

  atualizarPaginacao();
}

async function iniciarPagina(){
  preencherData();
  dadosMunicipios = await carregarJSON("dados/municipios.json");
  dadosFiltrados = dadosMunicipios;

  const total = document.getElementById("totalMunicipios");
  if(total){
    total.innerText = dadosMunicipios.length;
  }

  renderizarTabela();
}

function dadosExportacao(){
  return dadosFiltrados.map(item => ({
    "Município": item.municipio,
    "Lei Ratificadora": item.lei,
    "Data da Sanção": item.dataSancao,
    "Data da Publicação": item.dataPublicacao,
    "PDF": item.pdf
  }));
}

function baixarArquivo(conteudo,nomeArquivo,tipo){
  const blob = new Blob([conteudo], {type: tipo});
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = nomeArquivo;
  link.click();
}

function nomeBase(){
  return document.title.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"");
}

function exportarCSV(){
  const dados = dadosExportacao();
  if(!dados.length) return alert("Nenhum dado disponível.");
  const colunas = Object.keys(dados[0]);
  const linhas = [colunas.join(";"), ...dados.map(item =>
    colunas.map(coluna => `"${String(item[coluna] ?? "").replace(/"/g,'""')}"`).join(";")
  )];
  baixarArquivo("\uFEFF" + linhas.join("\n"), nomeBase() + ".csv", "text/csv;charset=utf-8;");
}

function exportarJSON(){
  baixarArquivo(JSON.stringify(dadosExportacao(), null, 2), nomeBase() + ".json", "application/json;charset=utf-8;");
}

function exportarXML(){
  let xml = `<?xml version="1.0" encoding="UTF-8"?>\n<municipios>\n`;
  dadosExportacao().forEach(item => {
    xml += "  <municipio>\n";
    Object.entries(item).forEach(([chave, valor]) => {
      const tag = chave.normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-zA-Z0-9]/g,"_").toLowerCase();
      xml += `    <${tag}>${String(valor || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</${tag}>\n`;
    });
    xml += "  </municipio>\n";
  });
  xml += "</municipios>";
  baixarArquivo(xml, nomeBase() + ".xml", "application/xml;charset=utf-8;");
}

function exportarXLSX(){
  const dados = dadosExportacao();
  if(!dados.length) return alert("Nenhum dado disponível.");
  const planilha = XLSX.utils.json_to_sheet(dados);
  const pasta = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(pasta, planilha, "Municípios");
  XLSX.writeFile(pasta, nomeBase() + ".xlsx");
}

function exportarPDF(){
  const dados = dadosExportacao();
  const {jsPDF} = window.jspdf;
  const doc = new jsPDF("landscape");

  doc.setFontSize(14);
  doc.text(document.title, 14, 15);

  if(!dados.length){
    doc.text("Nenhum dado disponível.", 14, 30);
    doc.save(nomeBase() + ".pdf");
    return;
  }

  const colunas = Object.keys(dados[0]);
  const linhas = dados.map(item => colunas.map(coluna => item[coluna]));

  doc.autoTable({
    head:[colunas],
    body:linhas,
    startY:25,
    styles:{fontSize:7,cellPadding:2},
    headStyles:{fillColor:[7,55,99]}
  });

  doc.save(nomeBase() + ".pdf");
}

iniciarPagina();
