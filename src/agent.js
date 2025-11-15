Office.onReady(() => { /* ready */ });

let generator = null;

async function runAgent() {
  try {
    await Excel.run(async (ctx) => {
      const range = ctx.workbook.getSelectedRange();
      range.load(["values", "address", "rowCount", "columnCount"]);
      await ctx.sync();

      const data = range.values;
      const address = range.address.split('!')[1];
      const [startCell] = address.split(':');
      const cols = data[0]?.length || 0;

      const userPrompt = prompt("What would you like to compute?")?.trim();
      if (!userPrompt) return;

      const intent = await inferIntent(userPrompt, cols);
      const result = intent.action === 'correlation' ? computeCorrelation(data, intent) : null;

      await writeResult(ctx, range, result, intent, startCell);
      await ctx.sync();
    });
  } catch (err) {
    console.error(err);
    alert("Error: " + (err.message || err));
  }
}

// --- Micro-LM (transformers.js + ONNX) ---
async function loadLM() {
  if (!generator) {
    const { pipeline } = await import('https://cdn.jsdelivr.net/npm/@xenova/transformers@2.17.2');
    generator = await pipeline('text-generation', 'onnx-community/Qwen2.5-0.5B', {
      quantized: true,
      device: 'wasm',
      progress_callback: (p) => console.log('LM load:', p)
    });
  }
}

const PROMPT = `
You are an Excel formula assistant. Return ONLY valid JSON:
{ "action": "correlation|formula", "colA": int, "colB": int, "output": "A1" }

User: {{PROMPT}}
Data has {{COLS}} columns. Use 0-based column index.
JSON:`;

async function inferIntent(prompt, cols) {
  if (!generator) await loadLM();
  const filled = PROMPT
    .replace('{{PROMPT}}', prompt)
    .replace('{{COLS}}', cols);

  const output = await generator(filled, {
    max_new_tokens: 80,
    temperature: 0.0,
    do_sample: false
  });

  const jsonStr = output[0].generated_text.split('JSON:')[1]?.trim() || '{}';
  try {
    const parsed = JSON.parse(jsonStr);
    if (parsed.action === 'correlation' && parsed.colA != null && parsed.colB != null) {
      return parsed;
    }
  } catch {}
  return { action: 'formula', formula: prompt };
}

// --- Stats: Correlation + MoE ---
function correlation(x, y) {
  const n = x.length;
  const sum = (a) => a.reduce((s, v) => s + v, 0);
  const sumXY = x.map((v, i) => v * y[i]).reduce((s, v) => s + v, 0);
  const sumX2 = sum(x.map(v => v * v));
  const sumY2 = sum(y.map(v => v * v));
  const num = n * sumXY - sum(x) * sum(y);
  const den = Math.sqrt((n * sumX2 - sum(x) ** 2) * (n * sumY2 - sum(y) ** 2));
  return num / den;
}

function moe(r, n) {
  const z = 0.5 * Math.log((1 + r) / (1 - r));
  const se = 1 / Math.sqrt(n - 3);
  const ci = 1.96 * se;
  return { lower: Math.tanh(z - ci), upper: Math.tanh(z + ci) };
}

function computeCorrelation(data, intent) {
  const x = data.map(r => r[intent.colA]).filter(v => typeof v === 'number');
  const y = data.map(r => r[intent.colB]).filter(v => typeof v === 'number');
  if (x.length < 3) return null;
  const r = correlation(x, y);
  const { lower, upper } = moe(r, x.length);
  const colA = numToCol(intent.colA);
  const colB = numToCol(intent.colB);
  const row1 = 2, rowN = data.length + 1;
  return {
    r, lower, upper,
    formula: `=CORREL(${colA}${row1}:${colA}${rowN},${colB}${row1}:${colB}${rowN})`
  };
}

function numToCol(n) {
  let s = '';
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// --- Write Result ---
async function writeResult(ctx, range, result, intent, startCell) {
  const sheet = range.worksheet;
  const outputRef = intent.output || 'B1';
  const [startCol, startRow] = startCell.match(/[A-Z]+|[0-9]+/g);
  const offsetCol = colToNum(outputRef.match(/[A-Z]+/)?.[0] || 'A') - colToNum(startCol);
  const offsetRow = (parseInt(outputRef.match(/[0-9]+/)?.[0] || 1) - parseInt(startRow));

  const target = range.getOffsetRange(offsetRow, offsetCol);
  target.load('address');
  await ctx.sync();

  if (result?.formula) {
    target.formulas = [[result.formula]];
    const rCell = target.getOffsetRange(0, 1);
    const moeCell = target.getOffsetRange(0, 2);
    rCell.values = [[result.r]];
    moeCell.formulas = [[`=CONFIDENCE.T(0.05,1,${result.r})`]]; // placeholder
  } else if (intent.action === 'formula') {
    target.formulas = [[ '=' + intent.formula ]];
  }
}

function colToNum(col) {
  return col.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}