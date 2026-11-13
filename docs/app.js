// The demo of the SheetRS package for WebAssembly: every tool of the package on a workbook that
// stays in the browser. The package is `./pkg` (`scripts/package-wasm.sh --demo` puts it there).
import init, {
    Book,
    initPanicHook,
    lintRules,
    evaluateFormula,
} from './pkg/sheetrs_wasm.js';

let wasmReady = false;
let currentTool = 'lint';
let currentFileData = null;
let currentFileName = '';
let rulesMetadata = [];
// The operations of the edit tool: a list the page shows, applied to a `Book` by `applyPlan`.
let editPlan = [];
// The module the compile tool made, instantiated to run it in the page.
let compiled = null;
// The model the calculate tool read: its inputs and outputs.
let model = null;

const $ = (id) => document.getElementById(id);
const fileZone = $('file-zone');
const fileInput = $('file-input');
const output = $('output');
const tableOutput = $('table-output');
const downloads = $('downloads');
const statusText = $('status');
const toolBtns = document.querySelectorAll('.tool-btn');
const playBtn = $('play-btn');
const fileInfo = $('file-info');
const rulesAccordion = $('rules-accordion');

const TOOLS = ['lint', 'stats', 'query', 'edit', 'calculate', 'evaluate', 'compile'];
const TITLES = {
    lint: 'Lint', stats: 'Statistics', query: 'Query', edit: 'Edit',
    calculate: 'Calculate', evaluate: 'Evaluate', compile: 'Compile',
};
// What each tool needs: the evaluate tool reads no workbook.
const NEEDS_FILE = { lint: true, stats: true, query: true, edit: true, calculate: true, evaluate: false, compile: true };

async function run() {
    try {
        await init();
        initPanicHook();
        wasmReady = true;
        updateStatus('WASM Loaded');
        rulesMetadata = lintRules();
        buildRulesUI();
        refreshPlay();
    } catch (e) {
        console.error('Initialization error:', e);
        updateStatus('Initialization Error');
    }
}

// ---- Tool selection ----------------------------------------------------------------------------
toolBtns.forEach((btn) => {
    btn.addEventListener('click', () => selectTool(btn.dataset.tool));
});

function selectTool(tool) {
    currentTool = tool;
    toolBtns.forEach((b) => b.classList.toggle('active', b.dataset.tool === tool));
    for (const t of TOOLS) {
        $(`${t}-setup`).style.display = t === tool ? 'block' : 'none';
    }
    $('output-title').textContent = TITLES[tool];
    refreshPlay();
}

function refreshPlay() {
    playBtn.disabled = !wasmReady || (NEEDS_FILE[currentTool] && !currentFileData);
}

// ---- The rules of the linter --------------------------------------------------------------------
function buildRulesUI() {
    if (!rulesMetadata || !Array.isArray(rulesMetadata)) return;
    rulesAccordion.innerHTML = '';
    const categories = {};
    rulesMetadata.forEach((rule) => {
        (categories[rule.category] ??= []).push(rule);
    });
    // The categories in the order of the numeric prefix of their first rule (1xx, 2xx, ...).
    const order = Object.keys(categories).sort((a, b) => {
        const number = (cat) => parseInt(categories[cat][0].id.match(/\d+/)?.[0] || '0', 10);
        return number(a) - number(b);
    });
    order.forEach((cat) => {
        const catDiv = document.createElement('div');
        catDiv.className = 'accordion-category';
        const header = document.createElement('div');
        header.className = 'category-header';
        header.innerHTML = `<span>${cat}</span> <span>▾</span>`;
        const content = document.createElement('div');
        content.className = 'category-content';
        categories[cat].forEach((rule) => {
            const item = document.createElement('div');
            item.className = 'rule-item';
            item.innerHTML = `
                <input type="checkbox" id="rule-${rule.id}" data-id="${rule.id}" ${rule.is_default ? 'checked' : ''}>
                <label for="rule-${rule.id}" class="rule-name" title="${rule.id}: ${rule.name}">${rule.name}</label>
            `;
            content.appendChild(item);
        });
        header.addEventListener('click', () => {
            content.classList.toggle('active');
            header.querySelector('span:last-child').textContent = content.classList.contains('active') ? '▴' : '▾';
        });
        catDiv.appendChild(header);
        catDiv.appendChild(content);
        rulesAccordion.appendChild(catDiv);
    });
}

function buildConfigToml() {
    const enabled = [];
    const disabled = [];
    rulesAccordion.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
        (cb.checked ? enabled : disabled).push(cb.dataset.id);
    });
    return `[global]\nenabled_rules = ${JSON.stringify(enabled)}\ndisabled_rules = ${JSON.stringify(disabled)}`;
}

// ---- The file -----------------------------------------------------------------------------------
fileZone.addEventListener('click', () => fileInput.click());
fileZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    fileZone.classList.add('dragover');
});
fileZone.addEventListener('dragleave', () => fileZone.classList.remove('dragover'));
fileZone.addEventListener('drop', (e) => {
    e.preventDefault();
    fileZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
});

function handleFile(file) {
    currentFileName = file.name;
    fileInfo.textContent = `File: ${file.name} (${formatSize(file.size)})`;
    clearResults();
    model = null;
    compiled = null;
    $('calculate-inputs').innerHTML = '';
    $('compile-run').style.display = 'none';
    updateStatus('Loading file...');
    const reader = new FileReader();
    reader.onload = (e) => {
        currentFileData = new Uint8Array(e.target.result);
        refreshPlay();
        updateStatus('✓ File loaded - Ready to run');
    };
    reader.onerror = (err) => {
        console.error('FileReader error:', err);
        updateStatus('File Read Error');
    };
    reader.readAsArrayBuffer(file);
}

const extension = () => currentFileName.split('.').pop().toLowerCase();
const stem = () => currentFileName.replace(/\.[^.]+$/, '') || 'workbook';

// ---- Results ------------------------------------------------------------------------------------
function clearResults() {
    output.textContent = 'Welcome to SheetRS WASM Demo.\n1. Select a tool and configure it.\n2. Upload a spreadsheet.\n3. Press Play!';
    tableOutput.style.display = 'none';
    tableOutput.innerHTML = '';
    downloads.innerHTML = '';
}

function showText(text) {
    tableOutput.style.display = 'none';
    tableOutput.innerHTML = '';
    output.textContent = text;
}

// A table, and the text under it (a summary, the time).
function showTable(columns, rows, note = '') {
    const table = document.createElement('table');
    const head = table.createTHead().insertRow();
    columns.forEach((c) => {
        const th = document.createElement('th');
        th.textContent = c;
        head.appendChild(th);
    });
    const body = table.createTBody();
    rows.forEach((row) => {
        const tr = body.insertRow();
        columns.forEach((c) => {
            const td = tr.insertCell();
            const v = row[c];
            td.textContent = v === undefined || v === null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v);
        });
    });
    tableOutput.innerHTML = '';
    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';
    wrap.appendChild(table);
    tableOutput.appendChild(wrap);
    tableOutput.style.display = 'block';
    output.textContent = note;
}

function offerDownload(name, data, type = 'application/octet-stream', extraClass = '') {
    const blob = new Blob([data], { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.className = `action-btn download-link ${extraClass}`.trim();
    a.textContent = `⬇ ${name} (${formatSize(blob.size)})`;
    downloads.appendChild(a);
}

const pretty = (value) => JSON.stringify(value, null, 2);

// On a narrow screen the results are under the tool: bring them into view.
function showResults() {
    if (window.matchMedia('(max-width: 800px)').matches) {
        $('output-header').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

// ---- Running a tool -----------------------------------------------------------------------------
playBtn.addEventListener('click', processFile);

function processFile() {
    if (!wasmReady || (NEEDS_FILE[currentTool] && !currentFileData)) return;
    updateStatus('⏳ Processing...');
    playBtn.disabled = true;
    downloads.innerHTML = '';
    output.textContent = `⏳ Processing ${currentFileName || ''}...\n\n`;
    tableOutput.style.display = 'none';
    // Let the page paint before the work: a tool runs on this thread.
    setTimeout(async () => {
        const start = performance.now();
        try {
            await TOOL_RUNNERS[currentTool]();
            const ms = (performance.now() - start).toFixed(2);
            output.textContent += `\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n⏱️  Finished in ${ms}ms\n`;
            updateStatus('✓ Success');
            showResults();
        } catch (e) {
            tableOutput.style.display = 'none';
            output.textContent += `\n❌ Error: ${e}\n`;
            updateStatus('❌ Error occurred');
            console.warn('Processing error:', e);
        } finally {
            refreshPlay();
        }
    }, 10);
}

const TOOL_RUNNERS = {
    lint() {
        const config = buildConfigToml();
        const format = $('lint-format').value;
        if (format === 'sarif') {
            const sarif = new Book(currentFileData).lint(config).toSarif(currentFileName);
            const log = JSON.parse(sarif);
            const results = log.runs[0].results;
            showText(`SARIF ${log.version}: ${results.length} results, ${log.runs[0].tool.driver.rules.length} rules\n\n${pretty(log).slice(0, 200000)}`);
            offerDownload(`${stem()}.sarif`, sarif, 'application/sarif+json');
        } else if (format === 'json') {
            const json = new Book(currentFileData).lint(config).toJsonText(currentFileName);
            const doc = JSON.parse(json);
            showText(`JSON: ${doc.summary ? JSON.stringify(doc.summary) : ''}\n\n${pretty(doc).slice(0, 200000)}`);
            offerDownload(`${stem()}.json`, json, 'application/json');
        } else if (format === 'table') {
            const records = new Book(currentFileData).lint(config).records;
            showTable(['rule_id', 'severity', 'message', 'sheet', 'cell'], records, `${records.length} violations`);
        } else {
            showText(new Book(currentFileData).lint(config).toText());
        }
    },

    stats() {
        const stats = new Book(currentFileData).stats();
        const json = stats.toJSON();
        if ($('stats-json').checked) {
            showText(pretty(json));
        } else {
            const names = Object.keys(stats.definedNames);
            showText(
                `📊 Workbook Statistics\n${'━'.repeat(40)}\n\n` +
                `📄 Total Sheets: ${json.total_sheets}\n\nSheet Names:\n` +
                json.sheets.map((sheet, i) => `  ${i + 1}. ${sheet.name}\n`).join('') +
                `\n🏷️  Named Ranges: ${names.length}\n` +
                (names.length ? `\nNamed Range List:\n${names.map((name, i) => `  ${i + 1}. ${name}\n`).join('')}` : '') +
                `\n🔧 Contains Macros: ${stats.hasMacros ? 'Yes' : 'No'}\n`,
            );
        }
    },

    query() {
        const kind = $('query-scope-kind').value;
        const scopeText = $('query-scope').value.trim();
        const query = {};
        if (kind === 'sheet') query.scope = { sheet: scopeText };
        else if (kind === 'range') query.scope = { range: scopeText };
        else query.scope = 'workbook';
        const properties = [...$('query-properties').querySelectorAll('input:checked')].map((c) => c.dataset.property);
        if (properties.length > 0) query.properties = properties;
        const text = $('query-find').value;
        if (text !== '') {
            query.find = {
                text,
                mode: $('query-mode').value,
                search_in: $('query-search-in').value,
                match_case: $('query-match-case').checked,
                entire_cells: $('query-entire-cells').checked,
            };
            delete query.scope;
            if (kind !== 'workbook' && scopeText !== '') query.scope = kind === 'sheet' ? { sheet: scopeText } : { range: scopeText };
        }
        const { cells } = new Book(currentFileData).query(query);
        const columns = ['sheet', 'cell'];
        for (const p of ['value', 'formula', 'number_format', 'display', 'comment']) {
            if (cells.some((c) => c[p] !== undefined && c[p] !== null)) columns.push(p);
        }
        showTable(columns, cells.slice(0, 2000), `${cells.length} cells${cells.length > 2000 ? ' (the first 2000 are shown)' : ''}`);
        offerDownload(`${stem()}.cells.json`, pretty({ cells }), 'application/json');
    },

    edit() {
        if (editPlan.length === 0) throw new Error('the plan is empty: add an operation first');
        const plan = pretty(editPlan);
        const book = new Book(currentFileData);
        applyPlan(book, editPlan);
        const explained = book.explain();
        const workbook = book.save();
        const { replaced, skipped } = book.report;
        showText(`The plan\n${plan}\n\nThe passes over the file\n${explained}\n\n` +
            `${workbook.length} bytes written, ${replaced} cells replaced, ${skipped.length} skipped` +
            (skipped.length ? `\n${pretty(skipped)}` : ''));
        offerDownload(`${stem()}.edited.${extension()}`, workbook);
    },

    async calculate() {
        const mode = $('calculate-mode').value;
        const request = {};
        const inputs = {};
        $('calculate-inputs').querySelectorAll('input[data-input]').forEach((field) => {
            if (field.value.trim() === '') return;
            try {
                inputs[field.dataset.input] = JSON.parse(field.value);
            } catch {
                inputs[field.dataset.input] = field.value;
            }
        });
        if (Object.keys(inputs).length > 0) request.inputs = inputs;
        const ranges = $('calculate-ranges').value.split('\n').map((s) => s.trim()).filter(Boolean);
        if (ranges.length > 0) request.ranges = ranges;
        const book = new Book(currentFileData);
        book.calculate(request);
        const answer = book.valueRanges().toJSON();
        if (mode === 'embed_template') {
            const workbook = book.save();
            showText(`${pretty(answer)}\n\nThe workbook with the results: ${workbook.length} bytes`);
            offerDownload(`${stem()}.calculated.${extension()}`, workbook);
        } else {
            showText(pretty(answer));
        }
        const notes = book.report?.diagnostics ?? [];
        if (notes.length) output.textContent += `\n\nNotes\n${notes.map((d) => d.message).join('\n')}`;
    },

    evaluate() {
        const locale = $('evaluate-locale').value.trim();
        const value = evaluateFormula($('evaluate-formula').value, locale === '' ? undefined : locale, undefined);
        showText(pretty(value));
    },

    async compile() {
        const mode = $('compile-mode').value;
        const options = {};
        if (mode !== 'values') options[mode] = true;
        const declarations = $('compile-declarations').value.trim();
        if (declarations !== '') options.declarations = JSON.parse(declarations);
        const built = new Book(currentFileData).compile(options);
        compiled = await instantiate(built.module);
        const meta = compiled.meta();
        showText(
            `${stem()}.wasm: ${formatSize(built.module.length)}\n` +
            `${built.cells} cells, ${built.formulas} formulas (${built.kept} kept, ${built.folded} folded at build time), ${built.cycles} cycles\n` +
            `${built.diagnostics.length ? `\nNotes\n${built.diagnostics.map((d) => d.message).join('\n')}\n` : ''}` +
            `\nWhat the module says about itself (sheetrs_meta)\n${pretty(meta)}`,
        );
        offerDownload(`${stem()}.wasm`, built.module, 'application/wasm');
        $('compile-run').style.display = 'block';
        $('compile-json-label').style.display = mode === 'values' ? 'none' : 'block';
        $('compile-request').value = '{}';
    },
};

// ---- A module made by the compiler ---------------------------------------------------------------
// The module has no imports: the exports are described in the README of the package. A page must
// instantiate it asynchronously (a browser refuses to compile a big module on the main thread).
async function instantiate(bytes) {
    const { instance } = await WebAssembly.instantiate(bytes, {});
    const x = instance.exports;
    x.sheetrs_set_clock(Date.now() / 1000);
    const encoder = new TextEncoder();
    const decoder = new TextDecoder();
    const result = () => decoder.decode(new Uint8Array(x.memory.buffer, x.sheetrs_result_ptr(), x.sheetrs_result_len()));
    const call = (fn, text, ...rest) => {
        const request = encoder.encode(text);
        const ptr = x.sheetrs_alloc(request.length);
        new Uint8Array(x.memory.buffer, ptr, request.length).set(request);
        const status = fn(ptr, request.length, ...rest);
        x.sheetrs_free(ptr, request.length);
        return status;
    };
    return {
        writer: x.sheetrs_has_writer() === 1,
        meta() {
            if (x.sheetrs_meta() !== 0) throw new Error(result());
            return JSON.parse(result());
        },
        calculate(request) {
            if (call(x.sheetrs_calculate, request) !== 0) throw new Error(result());
            return result();
        },
        write(request, withJson) {
            if (call(x.sheetrs_write, request, withJson ? 1 : 0) !== 0) throw new Error(result());
            const workbook = new Uint8Array(x.memory.buffer, x.sheetrs_workbook_ptr(), x.sheetrs_workbook_len()).slice();
            return { workbook, json: result() };
        },
    };
}

$('compile-run-btn').addEventListener('click', async () => {
    if (!compiled) return;
    // The module stays on offer; the files of an earlier run go.
    downloads.querySelectorAll('.answer-link').forEach((link) => link.remove());
    const request = $('compile-request').value.trim() || '{}';
    try {
        const start = performance.now();
        let text;
        if (compiled.writer) {
            const withJson = $('compile-json').checked;
            const { workbook, json } = compiled.write(request, withJson);
            text = `${workbook.length} bytes of workbook` + (withJson ? `\n\n${json}` : '');
            offerDownload(`${stem()}.answer.${extension()}`, workbook, 'application/octet-stream', 'answer-link');
        } else {
            text = compiled.calculate(request);
        }
        showText(`The module answered (${(performance.now() - start).toFixed(2)}ms):\n\n${text}`);
        updateStatus('✓ The module ran');
    } catch (e) {
        showText(`❌ The module said: ${e.message ?? e}`);
        updateStatus('❌ Error occurred');
    }
});

// ---- The inputs of a model (the calculate tool) --------------------------------------------------
$('calculate-read').addEventListener('click', async () => {
    if (!wasmReady || !currentFileData) return;
    updateStatus('⏳ Reading the model...');
    try {
        const built = new Book(currentFileData).compile();
        model = (await instantiate(built.module)).meta();
        const inputs = model.meta.filter((m) => m.meta_type === 'input');
        const outputs = model.meta.filter((m) => m.meta_type === 'output');
        const box = $('calculate-inputs');
        box.innerHTML = '';
        const title = document.createElement('p');
        title.className = 'hint';
        title.textContent = `${inputs.length} inputs, ${outputs.length} outputs (empty fields keep the value of the workbook)`;
        box.appendChild(title);
        inputs.slice(0, 40).forEach((input) => {
            const label = document.createElement('label');
            label.className = 'field';
            label.textContent = `${input.key} (${input.value})`;
            const field = document.createElement('input');
            field.type = 'text';
            field.dataset.input = input.key;
            label.appendChild(field);
            box.appendChild(label);
        });
        if (inputs.length > 40) {
            const more = document.createElement('p');
            more.className = 'hint';
            more.textContent = `${inputs.length - 40} more inputs are not shown`;
            box.appendChild(more);
        }
        updateStatus('✓ Model read');
    } catch (e) {
        updateStatus('❌ Error occurred');
        showText(`❌ ${e}`);
    }
});

// ---- The plan of the edit tool ---------------------------------------------------------------------
const EDIT_HELP = {
    delete_sheets: ['Sheet names, one per line', 'Tmp'],
    delete_names: ['Defined names, one per line', 'old_range'],
    delete_rows: ['Rows, one per line', 'Data!3:10'],
    delete_columns: ['Columns, one per line', 'Data!B:D'],
    formula_to_value: ['A sheet, columns or rows, one per line', 'Report!B:D'],
    replace: ['Find: one text and its replacement in the field below, or one pair per line as “find → with”', 'text to find'],
};

$('edit-op').addEventListener('change', () => {
    const [label, placeholder] = EDIT_HELP[$('edit-op').value];
    $('edit-arg-label').firstChild.textContent = label;
    $('edit-arg').placeholder = placeholder;
    $('edit-replace').style.display = $('edit-op').value === 'replace' ? 'block' : 'none';
});

$('edit-add').addEventListener('click', () => {
    const op = $('edit-op').value;
    const lines = $('edit-arg').value.split('\n').map((s) => s.trim()).filter(Boolean);
    if (lines.length === 0) {
        updateStatus('❌ The operation needs its argument');
        return;
    }
    if (op === 'replace') {
        // A sweep: one pair per line, `find → with`, each an operation in order; a single line is the
        // text to find and the replacement is the field below.
        const options = {
            mode: $('edit-regex').checked ? 'regex' : 'literal',
            match_case: $('edit-match-case').checked,
            entire_cells: $('edit-entire-cells').checked,
        };
        const pairs = lines.map((line) => {
            const at = line.indexOf(' → ');
            return at >= 0 ? [line.slice(0, at), line.slice(at + 3)] : null;
        });
        if (lines.length > 1 && pairs.some((pair) => pair === null)) {
            updateStatus('❌ With several lines, every line is “find → with”');
            return;
        }
        if (pairs.every((pair) => pair !== null)) {
            editPlan.push({ replace: { pairs, ...options } });
        } else {
            editPlan.push({ replace: { text: lines[0], with: $('edit-with').value, ...options } });
        }
    } else {
        editPlan.push({ [op]: lines });
    }
    renderPlan();
    $('edit-arg').value = '';
});

$('edit-clear-plan').addEventListener('click', () => {
    editPlan = [];
    renderPlan();
});

// The operations of the page, as calls on a `Book`: they accumulate, and `save` runs them.
const CAMEL = { delete_sheets: 'deleteSheets', delete_names: 'deleteNames', delete_rows: 'deleteRows', delete_columns: 'deleteColumns', formula_to_value: 'formulaToValue' };
function applyPlan(book, plan) {
    for (const operation of plan) {
        const [name] = Object.keys(operation);
        if (name === 'replace') {
            const { pairs, text, with: replacement, mode, match_case: matchCase, entire_cells: entireCells } = operation.replace;
            book.replace(pairs ?? [[text, replacement]], { mode, matchCase, entireCells });
        } else {
            book[CAMEL[name]](operation[name]);
        }
    }
}

function renderPlan() {
    const list = $('edit-plan');
    list.innerHTML = '';
    editPlan.forEach((op, i) => {
        const item = document.createElement('li');
        const [name] = Object.keys(op);
        const body = op[name];
        item.textContent = `${name.replaceAll('_', ' ')}: ${Array.isArray(body) ? body.join(', ') : body.pairs ? body.pairs.map(([a, b]) => `${a} → ${b}`).join(', ') : `${body.text} → ${body.with}`} `;
        const remove = document.createElement('button');
        remove.className = 'action-btn';
        remove.textContent = '✕';
        remove.title = 'Remove from the plan';
        remove.addEventListener('click', () => {
            editPlan.splice(i, 1);
            renderPlan();
        });
        item.appendChild(remove);
        list.appendChild(item);
    });
}

// ---- Utility actions -----------------------------------------------------------------------------
$('copy-btn').addEventListener('click', () => {
    navigator.clipboard.writeText(output.textContent || tableOutput.innerText);
    updateStatus('Copied to clipboard');
});

$('download-btn').addEventListener('click', () => {
    const blob = new Blob([output.textContent || tableOutput.innerText], { type: 'text/plain' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${currentFileName || 'sheetrs'}_${currentTool}.txt`;
    a.click();
    URL.revokeObjectURL(a.href);
});

$('clear-btn').addEventListener('click', () => {
    output.textContent = '';
    tableOutput.style.display = 'none';
    tableOutput.innerHTML = '';
    downloads.innerHTML = '';
    updateStatus('Cleared');
});

function updateStatus(msg) {
    statusText.textContent = msg;
}

function formatSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${sizes[i]}`;
}

selectTool('lint');
run();
