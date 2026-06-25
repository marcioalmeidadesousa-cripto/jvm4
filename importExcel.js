const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('BaseNegociacaoJVM4.xlsx');

// Dados sheet
const wsDados = wb.Sheets['Dados'];
const rawDados = XLSX.utils.sheet_to_json(wsDados, { header: 1 });

// Servico sheet: Produto -> classifNova
const wsServico = wb.Sheets['Servico'];
const rawServico = XLSX.utils.sheet_to_json(wsServico, { header: 1 });
const classifNovaMap = {};
for (let i = 1; i < rawServico.length; i++) {
  const [produto, , classifNova] = rawServico[i];
  if (produto) classifNovaMap[produto.trim()] = classifNova || '';
}

// Mapa de override: garante classifNova esperados pelo app (chaves em precosCat/NOVA_ORDER)
// Necessário quando o nome no Servico sheet difere do nome esperado pelo app
const CLASSIF_NOVA_OVERRIDE = {
  'PET AGUA MINERAL SW 500 ML'        : '500ml SW',
  'PET ÁGUA MINERAL S/GAS 500 PREMIUM': '500ml premium sem',
  'PET ÁGUA MINERAL C/GAS 500 PREMIUM': '500ml premium com',
};

// Mapa de fallback para produtos não encontrados no Servico sheet (ex: OTTO)
// Inferido pelo nome do produto
function inferClassifNova(produto) {
  const p = produto.trim().toUpperCase();
  if (p.includes('10 LT') || p.includes('10LT') || p.includes('10 LTS')) return '10L';
  if (p.includes('COPO'))      return 'Copo';
  if (p.includes('5 LT') || p.includes('5LT') || p.includes('5000') || p.includes('5L') || p.includes('5 LITROS')) return '5L';
  if (p.includes('1500') && p.includes('SEM')) return '1500ml sem';
  if (p.includes('1500') && p.includes('COM')) return '1500ml com';
  if (p.includes('1500'))      return '1500ml sem';
  if (p.includes('500') && (p.includes('SEM') || p.includes('S/GÁS') || p.includes('S/GAS'))) return '500ml sem';
  if (p.includes('500') && (p.includes('COM') || p.includes('C/GÁS') || p.includes('C/GAS'))) return '500ml com';
  return '';
}

// Month names in Portuguese
const MONTHS = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

function excelDateToStr(serial) {
  const adjusted = serial >= 60 ? serial - 1 : serial;
  const date = new Date(Date.UTC(1900, 0, 0) + adjusted * 86400000);
  const m = MONTHS[date.getUTCMonth()];
  const y = date.getUTCFullYear();
  return `${m}/${y}`;
}

const result = [];
for (let i = 1; i < rawDados.length; i++) {
  const [dataSerial, produto, volume, classifReduzida] = rawDados[i];
  if (!produto || volume === undefined) continue;

  const data = typeof dataSerial === 'number' ? excelDateToStr(dataSerial) : String(dataSerial);
  const classif = String(classifReduzida || '').trim();
  const produtoTrimmed = String(produto).trim();

  // Determina classifNova: override > servico sheet > inferido
  let classifNova = CLASSIF_NOVA_OVERRIDE[produtoTrimmed]
    || classifNovaMap[produtoTrimmed]
    || inferClassifNova(produtoTrimmed);

  result.push({ data, produto: produtoTrimmed, volume: Number(volume) || 0, classif, classifNova });
}

const js = `const dadosRaw = ${JSON.stringify(result)};`;
fs.writeFileSync('dadosRaw.js', js, 'utf8');
console.log(`Gerado dadosRaw.js com ${result.length} registros.`);

// Sumário
const summary = {};
result.forEach(r => {
  if (!summary[r.classif]) summary[r.classif] = new Set();
  summary[r.classif].add(r.produto);
});
console.log('\nClassificacoes (classif):');
Object.entries(summary).forEach(([cat, prods]) => {
  console.log(`  ${cat}: ${prods.size} produto(s)`);
  [...prods].forEach(p => console.log(`    - ${p}`));
});

// Check classifNova coverage
const noNova = result.filter(r => !r.classifNova);
if (noNova.length) {
  const nomes = [...new Set(noNova.map(r => r.produto))];
  console.log('\nProdutos SEM classifNova:', nomes);
}
