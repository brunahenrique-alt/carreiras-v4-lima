/* ------------------------------------------------------------------
   Janela salarial: cada inscrição recebe a Situação "Em análise" ou
   "Desclassificado: pretensão fora da faixa", comparando a pretensão
   do candidato com as colunas "Pretensão mínima" e "Pretensão máxima"
   da vaga na aba Vagas. Vaga sem faixa preenchida: sempre "Em análise".
   O candidato não vê nada disso no site.
   ------------------------------------------------------------------ */
function valorReais(v) {
  let s = String(v == null ? '' : v).replace(/[^\d.,]/g, '');
  if (!s) return NaN;
  if (s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  return parseFloat(s);
}

function reais(n) {
  return 'R$ ' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function situacaoPretensao(d) {
  d.faixa = '';
  const vaga = String(d.vagas || '').trim();
  const aba = SpreadsheetApp.getActive().getSheetByName(ABA_VAGAS);
  const [cab, ...linhas] = aba.getDataRange().getDisplayValues();
  const iV = cab.indexOf('Vaga'), iMin = cab.indexOf('Pretensão mínima'), iMax = cab.indexOf('Pretensão máxima');
  const l = linhas.find(x => String(x[iV]).trim() === vaga);
  if (!l) return 'Em análise';
  const min = iMin >= 0 ? valorReais(l[iMin]) : NaN;
  const max = iMax >= 0 ? valorReais(l[iMax]) : NaN;
  if (isNaN(min) && isNaN(max)) return 'Em análise';
  d.faixa = !isNaN(min) && !isNaN(max) ? reais(min) + ' a ' + reais(max)
          : !isNaN(min) ? reais(min) + ' ou mais' : 'até ' + reais(max);
  const p = valorReais(d.pretensao);
  if (isNaN(p)) return 'Em análise';
  if ((!isNaN(min) && p < min) || (!isNaN(max) && p > max)) return 'Desclassificado: pretensão fora da faixa';
  return 'Em análise';
}

