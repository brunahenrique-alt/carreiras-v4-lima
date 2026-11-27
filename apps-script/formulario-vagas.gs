/* ------------------------------------------------------------------
   Formulário de vagas (menu Vagas > Nova vaga / Editar vaga selecionada)
   Os campos seguem o modelo de vaga usado pelo time de People.
   ------------------------------------------------------------------ */

// Textos que já vêm preenchidos ao abrir "Nova vaga". Tudo pode ser alterado no formulário antes de salvar.
const PADRAO_VAGA = {
  'Modelo': 'Presencial',
  'Local': 'Rio de Janeiro',
  'Endereço': 'Centro do Rio de Janeiro, próximo à Candelária',
  'Horário': 'Segunda a sexta-feira',
  'Ambiente de trabalho': [
    'Somos guiados por atitudes claras e transparentes em nossa rotina, que nos permitem fazer o máximo para nossos clientes. Trabalhamos de forma colaborativa, confiando nas pessoas que estão ao nosso lado.',
    'Aqui ninguém é empregado de ninguém. O ambiente é o mais horizontal possível, proporcionando autonomia para que você tome as melhores decisões dentro das suas responsabilidades. Não espere pelos outros aquilo que você pode fazer por si mesmo.'
  ].join('\n'),
  'O que a posição oferece': [
    'Staage: plataforma com mais de 60 cursos.',
    'Qulture Rocks: plataforma de desenvolvimento, feedbacks e desenvolvimento do escopo.',
    'Growth Leaders Academy: trilhas de formação e desenvolvimento de lideranças.',
    'Reservatório de Dopamina: curadoria de conteúdos para o seu desenvolvimento contínuo.',
    'OnHappy: até 50% de desconto em pacotes de viagem.',
    'Parcerias Universitárias: bolsas e descontos em graduação e pós em universidades parceiras, com opções para você escolher.',
    'English Pass: aulas de inglês personalizadas, com IA e um dependente gratuito.'
  ].join('\n'),
  'Sobre a empresa': [
    'Há 13 anos, lideramos o caminho no Marketing Digital, sempre em busca de inovação, nos tornando a maior assessoria de marketing digital do país. Nosso foco é na geração de vendas e na transformação da realidade de milhares de empreendedores, dentro e fora do Brasil. Existimos para vender o seu produto, vender mais vezes, vender para mais pessoas e vender pelo maior valor.',
    'Estruturados em um sistema de franquias, temos mais de 250 escritórios no Brasil, mais de 6.000 clientes ativos e reunimos mais de 3.500 profissionais e mais de 300 sócios, entre Matriz e Franqueados.',
    'Nossa jornada continua, queremos elevar a V4 Company a novos patamares, e sabemos que isso só acontece com profissionais excepcionais ao nosso lado.'
  ].join('\n'),
  'Status': 'Aberta'
};

// Mesmo endereço que o site usa para montar o link de cada vaga
function slugVaga(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase()
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
}

function abrirPainel(linha) {
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_VAGAS);
  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  let dados = {};
  if (linha) {
    const valores = aba.getRange(linha, 1, 1, cabecalho.length).getDisplayValues()[0];
    cabecalho.forEach((h, i) => dados[h] = valores[i]);
    // Vaga antiga sem "Sobre a empresa": mostra o texto padrão (é o mesmo que o site já exibe)
    if (!dados['Sobre a empresa']) dados['Sobre a empresa'] = PADRAO_VAGA['Sobre a empresa'];
  } else {
    dados = Object.assign({}, PADRAO_VAGA);
  }
  const t = HtmlService.createTemplate(PAINEL_HTML);
  t.dados = JSON.stringify(dados).replace(/</g, '\\u003c');
  t.linha = linha || 0;
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle(linha ? 'Editar vaga' : 'Nova vaga'));
}

function salvarVaga(dados, linha) {
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_VAGAS);
  let cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];

  // Cria na aba Vagas as colunas que ainda não existem (ex.: Remuneração, Horário)
  const novas = Object.keys(dados).filter(h => cabecalho.indexOf(h) === -1);
  if (novas.length) {
    const ini = cabecalho.length + 1, fim = cabecalho.length + novas.length;
    if (aba.getMaxColumns() < fim) aba.insertColumnsAfter(aba.getMaxColumns(), fim - aba.getMaxColumns());
    aba.getRange(1, ini, 1, novas.length).setValues([novas]);
    aba.getRange(1, 1).copyFormatToRange(aba, ini, fim, 1, 1);
    cabecalho = cabecalho.concat(novas);
  }

  // Evita duas vagas abertas com o mesmo nome (as duas teriam o mesmo link no site)
  const id = slugVaga(dados['Vaga']);
  const iVaga = cabecalho.indexOf('Vaga'), iStatus = cabecalho.indexOf('Status');
  if (aba.getLastRow() > 1 && String(dados['Status']).toLowerCase() !== 'fechada') {
    const linhas = aba.getRange(2, 1, aba.getLastRow() - 1, cabecalho.length).getDisplayValues();
    const repetida = linhas.some((l, i) => (i + 2) !== linha && slugVaga(l[iVaga]) === id &&
      String(l[iStatus]).toLowerCase() !== 'fechada');
    if (repetida) throw new Error('Já existe uma vaga aberta com esse nome. Mude o nome ou feche a outra vaga.');
  }

  const limpa = v => { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; };
  if (linha) {
    const atual = aba.getRange(linha, 1, 1, cabecalho.length).getValues()[0];
    const nova = cabecalho.map((h, i) => (h in dados) ? limpa(dados[h]) : atual[i]);
    aba.getRange(linha, 1, 1, cabecalho.length).setValues([nova]);
  } else {
    aba.appendRow(cabecalho.map(h => (h in dados) ? limpa(dados[h]) : ''));
  }
  return id;
}

const PAINEL_HTML = `<!DOCTYPE html>
<html><head><base target="_top">
<style>
  body{font-family:Arial,sans-serif;font-size:13px;color:#111;margin:0;padding:16px}
  h3{font-size:12px;text-transform:uppercase;letter-spacing:.4px;color:#E3151F;margin:24px 0 0;padding-top:16px;border-top:1px solid #e6e6e6}
  h3.primeiro{border-top:0;padding-top:0;margin-top:8px}
  label{display:block;font-weight:bold;margin:12px 0 5px}
  small{font-weight:normal;color:#666}
  input[type=text],select,textarea{width:100%;box-sizing:border-box;padding:8px;border:1px solid #ccc;border-radius:4px;font:inherit}
  input:focus,select:focus,textarea:focus{outline:none;border-color:#E3151F}
  textarea{min-height:90px;resize:vertical}
  textarea.alto{min-height:180px}
  .chk{display:flex;gap:8px;align-items:center;font-weight:normal;margin-top:16px}
  button{margin-top:20px;width:100%;padding:12px;background:#E3151F;color:#fff;border:0;border-radius:4px;font-weight:bold;font-size:14px;cursor:pointer}
  button:disabled{opacity:.6;cursor:default}
  #msg{margin-top:12px;line-height:1.4}
  #msg a{color:#E3151F;font-weight:bold}
  .ok{color:#137333}.erro{color:#c5221f}
  .dica{color:#666;margin:0 0 4px;line-height:1.4}
</style></head>
<body>
<p class="dica">Tudo o que você preencher aqui aparece na página da vaga, na mesma ordem deste formulário. Campo em branco não aparece no site. Para quebrar linha, use Enter.</p>
<form id="f">
  <h3 class="primeiro">Vaga</h3>
  <label for="c1">Nome da vaga</label><input type="text" id="c1" name="Vaga" placeholder="Ex.: BDR Comercial">
  <label for="c2">Nível <small>(opcional, só aparece no site se escolher)</small></label><select id="c2" name="Nível"><option value="">Não informar</option><option>Estágio</option><option>Júnior</option><option>Pleno</option><option>Sênior</option></select>
  <label for="c3">Modelo de trabalho</label><select id="c3" name="Modelo"><option>Presencial</option><option>Híbrido</option><option>Remoto</option></select>
  <label for="c4">Cidade <small>(aparece na etiqueta e no filtro do site)</small></label><input type="text" id="c4" name="Local">
  <label for="c5">Endereço <small>(aparece em Informações da vaga)</small></label><input type="text" id="c5" name="Endereço" placeholder="Ex.: Centro do Rio de Janeiro, próximo à Candelária">

  <h3>Descrição da vaga</h3>
  <label for="c6">Sobre a vaga <small>(texto de abertura; cada linha vira um parágrafo)</small></label><textarea class="alto" id="c6" name="Descrição"></textarea>
  <label for="c19">Sobre o ambiente de trabalho <small>(cada linha vira um parágrafo)</small></label><textarea class="alto" id="c19" name="Ambiente de trabalho"></textarea>
  <label for="c20">Missão <small>(aparece como "Sua missão como [nome da vaga]")</small></label><textarea id="c20" name="Missão"></textarea>
  <label for="c7">Principais responsabilidades <small>(uma por linha)</small></label><textarea id="c7" name="Responsabilidades"></textarea>
  <label for="c8">Requisitos <small>(um por linha)</small></label><textarea id="c8" name="Requisitos"></textarea>
  <label for="c9">Diferenciais <small>(opcional, um por linha)</small></label><textarea id="c9" name="Diferenciais"></textarea>

  <h3>Informações da vaga</h3>
  <label for="c10">Formato de contratação <small>(ex.: CLT, PJ, Estágio)</small></label><input type="text" id="c10" name="Contratação">
  <label for="c11">Remuneração</label><input type="text" id="c11" name="Remuneração" placeholder="Ex.: R$ 4.000 a R$ 7.000 entre fixo e variável">
  <label for="c12">Horário</label><input type="text" id="c12" name="Horário">
  <label for="c13">Benefícios <small>(opcional, resumo em uma linha)</small></label><input type="text" id="c13" name="Benefícios">
  <label for="c14">Outras informações <small>(uma por linha, no formato Rótulo: texto)</small></label><textarea id="c14" name="Outras informações" placeholder="Requisito: notebook próprio"></textarea>

  <h3>O que oferecemos</h3>
  <label for="c15">O que oferecemos <small>(um por linha, no formato Nome: descrição)</small></label><textarea id="c15" name="O que a posição oferece"></textarea>
  <label for="c16">Sobre a empresa <small>(cada linha vira um parágrafo)</small></label><textarea class="alto" id="c16" name="Sobre a empresa"></textarea>

  <h3>Inscrição e publicação</h3>
  <label for="c17">Perguntas da inscrição <small>(uma por linha, aparecem na 2ª etapa do formulário)</small></label><textarea id="c17" name="Perguntas da inscrição"></textarea>
  <label for="c18">Status</label><select id="c18" name="Status"><option>Aberta</option><option>Fechada</option></select>
  <label class="chk"><input type="checkbox" name="Destaque"> Marcar como prioridade</label>
  <button type="submit" id="b">Salvar vaga</button>
  <div id="msg"></div>
</form>
<script>
  var dados = <?!= dados ?>, linha = <?!= linha ?>;
  var SITE = 'https://carreiras.v4lima.com/vaga.html?v=';
  var f = document.getElementById('f');
  var campos = ['Vaga','Nível','Modelo','Local','Endereço','Descrição','Ambiente de trabalho','Missão','Responsabilidades','Requisitos','Diferenciais',
    'Contratação','Remuneração','Horário','Benefícios','Outras informações','O que a posição oferece','Sobre a empresa',
    'Perguntas da inscrição','Status'];
  function preencher(d) {
    Object.keys(d).forEach(function (k) {
      var el = f.elements[k]; if (!el) return;
      if (el.type === 'checkbox') el.checked = String(d[k]).toLowerCase() === 'sim';
      else el.value = d[k];
    });
  }
  preencher(dados);
  if (linha) document.getElementById('b').textContent = 'Salvar alterações';
  f.addEventListener('submit', function (e) {
    e.preventDefault();
    var o = {};
    campos.forEach(function (k) { o[k] = f.elements[k].value.trim(); });
    o['Destaque'] = f.elements['Destaque'].checked ? 'Sim' : 'Não';
    if (!o['Vaga']) { msg('Preencha o nome da vaga.', 'erro'); f.elements['Vaga'].focus(); return; }
    var b = document.getElementById('b'), txt = b.textContent;
    b.disabled = true; b.textContent = 'Salvando...';
    google.script.run
      .withSuccessHandler(function (id) {
        b.disabled = false; b.textContent = txt;
        var aberta = o['Status'] !== 'Fechada';
        msg((linha ? 'Vaga atualizada. ' : 'Vaga criada. ') + (aberta ? 'Ela já aparece no site.' : 'Ela está fechada e não aparece no site.'), 'ok',
          aberta && id ? ' <a href="' + SITE + id + '" target="_blank">Ver a vaga no site</a>' : '');
        if (!linha) { f.reset(); preencher(dados); }
      })
      .withFailureHandler(function (err) {
        b.disabled = false; b.textContent = txt;
        msg('Não foi possível salvar: ' + err.message, 'erro');
      })
      .salvarVaga(o, linha);
  });
  function msg(t, c, extra) {
    var m = document.getElementById('msg'); m.textContent = t; m.className = c;
    if (extra) m.insertAdjacentHTML('beforeend', extra);
  }
</script>
</body></html>`;
