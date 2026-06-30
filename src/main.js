import './styles.css';

const apiBase = 'http://127.0.0.1:5174';

const state = {
  active: 'quotation',
  meta: { themes: [] },
  status: '',
  result: null,
  preview: null,
  previewDirty: false,
  previewLoading: false,
  previewRevision: 0,
  previewRequestId: 0,
  previewStatus: '',
  previewTimer: null,
  previewZoom: {
    quotation: 1,
    supply: 0.85,
    contract: 0.85,
  },
  forms: {
    quotation: {
      themeId: 'example_invoice',
      stampType: 'contract',
      customerName: '',
      date: '2026-06-29',
      rows: [
        {},
      ],
    },
    supply: {
      themeId: 'example_invoice',
      stampType: 'contract',
      date: '2026-06-29',
      remark: '含增值税普通发票',
      rows: [
        {},
      ],
    },
    contract: {
      themeId: 'example_invoice',
      stampType: 'contract',
      buyer: '',
      buyerInfo: '',
      date: '2026-06-29',
      remark: '含增值税普通发票',
      packaging: '塑料袋包装',
      items: [
        {},
      ],
    },
  },
};

const labels = {
  quotation: '报价单',
  supply: '供货清单',
  contract: '普票合同',
};

const app = document.querySelector('#app');
await loadMeta();
Object.keys(state.forms).forEach(syncCalculatedFields);
render();

async function loadMeta() {
  try {
    const res = await fetch(`${apiBase}/api/templates`);
    state.meta = await res.json();
  } catch (error) {
    state.status = `无法连接本地生成服务：${error.message}`;
  }
}

function render() {
  const active = state.active;
  app.innerHTML = `
    <aside class="sidebar">
      <div class="brand">
        <strong>Office 文档生成器</strong>
        <span>Word / Excel 模板自动生成 PDF</span>
      </div>
      <nav>
        ${Object.entries(labels).map(([key, label]) => `
          <button class="nav-item ${key === active ? 'active' : ''}" data-nav="${key}">
            ${label}<small>${subtitle(key)}</small>
          </button>
        `).join('')}
      </nav>
    </aside>
    <main class="workspace">
      <header class="topbar">
        <div>
          <h1>${labels[active]}</h1>
          <p>${subtitle(active)}</p>
        </div>
        <button class="primary" id="generateBtn">生成 Word/Excel + PDF</button>
      </header>
      <section class="content office-content">
        <form class="panel" id="formPanel">${formHtml(active)}</form>
        <section class="result-panel" id="previewPanel">${rightPanelHtml()}</section>
      </section>
    </main>
  `;

  document.querySelectorAll('[data-nav]').forEach((button) => {
    button.addEventListener('click', () => {
      state.active = button.dataset.nav;
      state.result = null;
      state.preview = null;
      state.previewDirty = false;
      state.previewLoading = false;
      state.previewRequestId += 1;
      state.previewStatus = '';
      state.status = '';
      render();
    });
  });
  const form = document.querySelector('#formPanel');
  form.addEventListener('input', handleInput);
  form.addEventListener('change', handleInput);
  form.addEventListener('click', handleInput);
  document.querySelector('#previewPanel').addEventListener('click', handlePreviewClick);
  document.querySelector('#previewPanel').addEventListener('wheel', handlePreviewWheel, { passive: false });
  document.querySelector('#generateBtn').addEventListener('click', generate);
}

function subtitle(key) {
  if (key === 'quotation') return 'Excel 报价表模板';
  if (key === 'supply') return 'Excel 供货清单模板';
  return 'Word 合同模板';
}

function formHtml(type) {
  const data = state.forms[type];
  return `
    <div class="two-col">
      ${themeSelect(type, data.themeId)}
      ${stampTypeSelect(type, data.stampType)}
    </div>
    ${type === 'quotation' ? quotationFields(data) : ''}
    ${type === 'supply' ? supplyFields(data) : ''}
    ${type === 'contract' ? contractFields(data) : ''}
  `;
}

function themeSelect(type, value) {
  const themes = state.meta.themes || [];
  const filtered = themes.filter((theme) => {
    if (type === 'supply') return theme.hasSupply;
    if (type === 'contract') return theme.hasContract;
    return theme.hasQuotation !== false;
  });
  return `
    <label class="field">
      <span>供方名称</span>
      <select data-path="${type}.themeId">
        ${filtered.map((theme) => `<option value="${theme.id}" ${theme.id === value ? 'selected' : ''}>${themeLabel(type, theme)}</option>`).join('')}
      </select>
    </label>
  `;
}

function stampTypeSelect(type, value = 'official') {
  return `
    <label class="field">
      <span>印章选择</span>
      <select data-path="${type}.stampType">
        <option value="official" ${value !== 'contract' ? 'selected' : ''}>公章</option>
        <option value="contract" ${value === 'contract' ? 'selected' : ''}>合同章</option>
      </select>
    </label>
  `;
}

function themeLabel(type, theme) {
  return type === 'quotation' ? (theme.quotationLabel || theme.label) : theme.label;
}

function quotationFields(data) {
  return `
    <div class="two-col">
      ${field('客户名称', 'quotation.customerName', data.customerName)}
      ${field('日期', 'quotation.date', data.date)}
    </div>
    ${rowsEditor('quotation', data.rows, ['name', 'quantity', 'unitPrice', 'amount'])}
  `;
}

function supplyFields(data) {
  return `
    <div class="two-col">
      ${field('供货时间', 'supply.date', data.date)}
      ${field('备注', 'supply.remark', data.remark)}
    </div>
    ${rowsEditor('supply', data.rows, ['name', 'spec', 'quantity', 'unitPrice', 'amount'])}
  `;
}

function contractFields(data) {
  return `
    <div class="two-col">
      ${field('需方', 'contract.buyer', data.buyer)}
      ${field('签订时间', 'contract.date', data.date)}
      ${field('备注', 'contract.remark', data.remark)}
      ${field('包装方式', 'contract.packaging', data.packaging)}
    </div>
    ${textareaField('需方下方内容', 'contract.buyerInfo', data.buyerInfo)}
    ${rowsEditor('contract', data.items, ['name', 'quantity', 'unitPrice', 'amount'], 'items')}
  `;
}

function rowsEditor(type, rows, keys, collection = 'rows') {
  const title = collection === 'items' ? '明细' : '货品明细';
  const labelMap = { name: '名称', spec: '规格', quantity: '数量', unitPrice: '单价', amount: '金额', remark: '备注' };
  return `
    <div class="section-title">${title}</div>
    <div class="rows">
      ${rows.map((row, index) => `
        <div class="row-card">
          <div class="row-head">第 ${index + 1} 行</div>
          <div class="row-grid">
            ${keys.map((key) => field(labelMap[key], `${type}.${collection}.${index}.${key}`, row[key] ?? '')).join('')}
          </div>
        </div>
      `).join('')}
    </div>
    <div class="actions">
      <button type="button" class="secondary" data-add-row="${type}" data-collection="${collection}">增加明细行</button>
      <button type="button" class="secondary" data-remove-row="${type}" data-collection="${collection}">删除最后一行</button>
    </div>
  `;
}

function field(label, path, value, options = {}) {
  const readonly = options.readonly ? ' readonly aria-readonly="true"' : '';
  const type = options.type || (path.endsWith('.date') ? 'date' : 'text');
  return `<label class="field"><span>${label}</span><input type="${type}" data-path="${path}" value="${escapeAttr(value)}"${readonly} /></label>`;
}

function textareaField(label, path, value) {
  return `<label class="field field-wide"><span>${label}</span><textarea data-path="${path}" rows="5">${escapeHtml(value)}</textarea></label>`;
}

function rightPanelHtml() {
  return `
    <div id="instantPreviewSlot">${documentPreview()}</div>
    <div id="statusSlot">${statusHtml()}</div>
    <div id="resultSlot">${resultHtml()}</div>
  `;
}

function officePreviewHtml() {
  const title = `${labels[state.active]}最终 PDF 核对`;
  const status = previewStatusText();
  const statusClass = previewStatusClass();
  const buttonText = state.preview?.pdfUrl ? '刷新最终 PDF' : '生成最终 PDF';
  const action = `
    <div class="preview-actions">
      <span class="preview-badge ${statusClass}" data-preview-status>${escapeHtml(status)}</span>
      <button type="button" class="secondary compact" data-refresh-preview ${state.previewLoading ? 'disabled' : ''}>${buttonText}</button>
    </div>
  `;
  if (state.preview?.pdfUrl) {
    return `
      <div class="preview-card office-pdf-preview">
        <div class="preview-toolbar">
          <strong>${title}</strong>
          ${action}
        </div>
        <div class="preview-theme" data-preview-theme>${escapeHtml(currentTheme()?.label ?? '')}</div>
        <div class="pdf-preview-wrap">
          <iframe class="pdf-preview-frame" src="${apiBase}${state.preview.pdfUrl}#toolbar=0&navpanes=0&view=FitH"></iframe>
          ${state.previewLoading ? '<div class="preview-overlay" data-preview-overlay>正在刷新最终 PDF，当前仍显示上一次结果</div>' : ''}
        </div>
      </div>
    `;
  }
  return `
    <div class="preview-card office-pdf-preview">
      <div class="preview-toolbar">
        <strong>${title}</strong>
        ${action}
      </div>
      <div class="preview-theme" data-preview-theme>${escapeHtml(currentTheme()?.label ?? '')}</div>
      <div class="preview-loading">${escapeHtml(state.previewStatus || '点击“生成最终 PDF”核对导出版式。')}</div>
    </div>
  `;
}

function documentPreview() {
  if (state.active === 'quotation') return quotationPreview();
  if (state.active === 'supply') return supplyPreview();
  return contractPreview();
}

function instantPreviewToolbar(theme) {
  const zoomControls = `
        <button type="button" class="secondary icon-button" data-preview-zoom="out" title="缩小">-</button>
        <span class="zoom-value">${Math.round(state.previewZoom[state.active] * 100)}%</span>
        <button type="button" class="secondary icon-button" data-preview-zoom="in" title="放大">+</button>
        <button type="button" class="secondary compact" data-preview-zoom="reset">重置</button>
  `;
  return `
    <div class="preview-toolbar">
      <strong>即时预览</strong>
      <div class="preview-actions">
        <span>${escapeHtml(theme?.label ?? '')}</span>
        ${zoomControls}
      </div>
    </div>
  `;
}

function quotationPreview() {
  const data = state.forms.quotation;
  const theme = currentTheme();
  const rows = visibleRows(data.rows);
  const total = sumAmounts(data.rows);
  const totalChinese = rmbUpper(total);
  const displayDate = documentDateText(data.date);
  return `
    <div class="preview-card landscape-preview instant-preview">
      ${instantPreviewToolbar(theme)}
      <div class="instant-preview-scroll">
      <div class="paper paper-landscape quotation-paper" style="--preview-zoom: ${state.previewZoom[state.active]}">
        <table class="quote-preview-table">
          <thead>
            <tr><th colspan="6" class="quote-title">报价表</th></tr>
            <tr><th colspan="6" class="quote-customer">客户名称：${escapeHtml(data.customerName)}</th></tr>
            <tr><th>序号</th><th>日期</th><th>产品名称</th><th>单价（RMB/元）</th><th>数量（件）</th><th>金额（RMB/元）</th></tr>
          </thead>
          <tbody>
            ${rows.map((row, index) => quoteRow(row, index, displayDate)).join('')}
            <tr><td colspan="5" class="quote-total-label">共计</td><td class="quote-money">${moneyText(total)}</td></tr>
            <tr><td colspan="6" class="quote-upper">大写金额：${escapeHtml(totalChinese)}</td></tr>
            <tr><td colspan="6" class="quote-supplier"><span>供货方（盖章）：${escapeHtml(theme?.company ?? '')}</span>${stampImage(theme)}</td></tr>
          </tbody>
        </table>
      </div>
      </div>
    </div>
  `;
}

function quoteRow(row, index, date) {
  return `
    <tr>
      <td>${index + 1}</td>
      <td>${escapeHtml(row.date ?? date ?? '')}</td>
      <td>${escapeHtml(row.name ?? '')}</td>
      <td>${escapeHtml(row.unitPrice ?? '')}</td>
      <td>${escapeHtml(row.quantity ?? '')}</td>
      <td class="quote-money">${moneyText(row.amount)}</td>
    </tr>
  `;
}

function supplyPreview() {
  const data = state.forms.supply;
  const theme = currentTheme();
  const rows = visibleRows(data.rows);
  const total = sumAmounts(data.rows);
  return `
    <div class="preview-card portrait-preview instant-preview">
      ${instantPreviewToolbar(theme)}
      <div class="instant-preview-scroll">
      <div class="paper paper-portrait" style="--preview-zoom: ${state.previewZoom[state.active]}">
        <h2>供货结算清单</h2>
        <div class="paper-line">供货时间：${escapeHtml(documentDateText(data.date))}</div>
        <table class="paper-table">
          <colgroup>
            <col class="supply-col-index" />
            <col class="supply-col-name" />
            <col class="supply-col-spec" />
            <col class="supply-col-qty" />
            <col class="supply-col-price" />
            <col class="supply-col-money" />
            <col class="supply-col-remark" />
            <col class="supply-col-remark" />
          </colgroup>
          <thead>
            <tr><th>序号</th><th>货物名称</th><th>规格</th><th>数量</th><th>单价</th><th>金额</th><th colspan="2">备注</th></tr>
          </thead>
          <tbody>
            ${rows.map((row, index) => `
              <tr>
                <td>${index + 1}</td>
                <td>${escapeHtml(row.name ?? '')}</td>
                <td>${escapeHtml(row.spec ?? '')}</td>
                <td>${escapeHtml(row.quantity ?? '')}</td>
                <td>${escapeHtml(row.unitPrice ?? '')}</td>
                <td>${moneyText(row.amount)}</td>
                ${index === 0 ? `<td colspan="2" rowspan="${rows.length + 1}" class="supply-remark">${escapeHtml(data.remark ?? '')}</td>` : ''}
              </tr>
            `).join('')}
            <tr>
              <td colspan="2" class="supply-total">合计实收：${moneyText(total)}</td>
              <td colspan="4" class="supply-upper">大写：${escapeHtml(rmbUpper(total))}</td>
            </tr>
          </tbody>
        </table>
        <pre class="supply-info">供货单位：
${escapeHtml(theme?.supplierInfo ?? '')}</pre>
        <div class="paper-stamp-row">
          <span>供货单位（盖章）：</span>
          ${stampImage(theme)}
        </div>
      </div>
      </div>
    </div>
  `;
}

function contractPreview() {
  const data = state.forms.contract;
  const theme = currentTheme();
  const rows = visibleRows(data.items);
  const total = sumAmounts(data.items);
  return `
    <div class="preview-card portrait-preview instant-preview">
      ${instantPreviewToolbar(theme)}
      <div class="instant-preview-scroll">
      <div class="paper paper-portrait contract-paper" style="--preview-zoom: ${state.previewZoom[state.active]}">
        <h2>合同</h2>
        <div class="contract-meta">
          <div>
            <div>需方：${escapeHtml(data.buyer)}</div>
            <div>供方：${escapeHtml(theme?.company ?? '')}</div>
          </div>
          <div>
            <div>签订时间：${escapeHtml(documentDateText(data.date))}</div>
          </div>
        </div>
        <div class="paper-line">一、品名、规格型号、数量、单价、金额</div>
        <table class="paper-table">
          <thead>
            <tr><th>产品名称</th><th>件数</th><th>单价（元）</th><th>金额（元）</th><th>备注</th></tr>
          </thead>
          <tbody>
            ${rows.map((row, index) => `
              <tr>
                <td>${escapeHtml(row.name ?? '')}</td>
                <td>${escapeHtml(row.quantity ?? '')}</td>
                <td>${escapeHtml(row.unitPrice ?? '')}</td>
                <td>${moneyText(row.amount)}</td>
                ${index === 0 ? `<td rowspan="${rows.length}" class="contract-remark">${escapeHtml(data.remark ?? '')}</td>` : ''}
              </tr>
            `).join('')}
            <tr><td colspan="5">总计：人民币（小写）${moneyText(total)} 人民币（大写）：${escapeHtml(rmbUpper(total))}</td></tr>
          </tbody>
        </table>
        <div class="contract-terms">${contractClausesHtml(data, theme)}</div>
        <div class="contract-sign">
          <div>
            <strong>需方</strong>
            <pre>${escapeHtml(data.buyerInfo)}</pre>
          </div>
          <div class="contract-supplier-sign">
            <strong>供方</strong>
            <pre>${escapeHtml(theme?.supplierInfo ?? '')}</pre>
            ${stampImage(theme)}
          </div>
        </div>
      </div>
      </div>
    </div>
  `;
}

function contractClausesHtml(data, theme) {
  return contractClausesText(data, theme)
    .split('\n')
    .map((line) => `<p>${escapeHtml(line)}</p>`)
    .join('');
}

function contractClausesText(data, theme) {
  const payment = contractPaymentMethod(theme);
  const packaging = contractPackagingText(data.packaging);
  return [
    '二、结算方式与发货时间：',
    `    ${payment}，定制产品3-5天左右出货。`,
    '三、包装、运输方式：',
    `    ${packaging}，圆通快递包邮。`,
    '四、违约责任：',
    '双方应严格遵守本合同的约定，如出现问题，根据问题所属承担责任。解决合同纠纷的',
    '方式：协商调解不成时，依法向双方所在地人民法院诉讼。',
    '五、其它事项：',
    '1.定制产品不支持退货，保证质量。',
    '2.如货物出现质量问题，需方在收到货七天内通知供方，供方负责调换并承担相关运输费用。',
    '3.供方保证产品符合相关国家和行业标准，因供方产品给需方造成的损失，供方承担赔偿责任及包括律师费在内的维权成本。',
    '六、本合同邮寄具有法律效力，需方收到邮寄签字盖章回传后生效。',
  ].join('\n');
}

function contractPaymentMethod(theme) {
  const offlineIds = new Set(['example_invoice', 'example_vat']);
  const offlineLabels = new Set(['Example Invoice', 'Example VAT']);
  return offlineIds.has(theme?.id) || offlineLabels.has(theme?.label) ? '对公线下付款' : '淘宝平台下单';
}

function contractPackagingText(value) {
  const text = String(value ?? '').trim();
  return text || '塑料袋包装';
}

function statusHtml() {
  return state.status ? `<div class="status">${escapeHtml(state.status)}</div>` : '';
}

function resultHtml() {
  if (!state.result) {
    return `<div class="empty">生成后会在这里显示 Office 文件和 PDF 链接。</div>`;
  }
  return `
    <div class="info-card">
      <h2>生成完成</h2>
      <a class="download" href="${apiBase}${state.result.officeUrl}" target="_blank">打开 Office 文件</a>
      <a class="download" href="${apiBase}${state.result.pdfUrl}" target="_blank">打开 PDF 文件</a>
      <p>印章：${state.result.stampUsed ? '已插入' : '未插入'}</p>
    </div>
  `;
}

function handleInput(event) {
  const add = event.target.closest('[data-add-row]');
  const remove = event.target.closest('[data-remove-row]');
  if (event.type === 'click' && !add && !remove) return;
  if (add) {
    const type = add.dataset.addRow;
    const collection = add.dataset.collection;
    state.forms[type][collection].push({});
    syncCalculatedFields(type);
    state.result = null;
    state.status = '';
    markPreviewDirty();
    updateFormPanel();
    updateInstantPreview();
    updatePreviewChrome();
    updateStatusAndResult();
    return;
  }
  if (remove) {
    const type = remove.dataset.removeRow;
    const collection = remove.dataset.collection;
    if (state.forms[type][collection].length > 1) state.forms[type][collection].pop();
    syncCalculatedFields(type);
    state.result = null;
    state.status = '';
    markPreviewDirty();
    updateFormPanel();
    updateInstantPreview();
    updatePreviewChrome();
    updateStatusAndResult();
    return;
  }
  const path = event.target.dataset.path;
  if (!path) return;
  const originalParts = path.split('.');
  const parts = [...originalParts];
  const type = parts[0];
  let target = state.forms;
  while (parts.length > 1) target = target[parts.shift()];
  target[parts[0]] = event.target.value;
  const change = pathChangeInfo(originalParts);
  markManualAmount(type, change, target);
  syncCalculatedFields(type, change);
  refreshCalculatedInputs(type, change);
  state.result = null;
  state.status = '';
  markPreviewDirty();
  updateInstantPreview();
  updatePreviewChrome();
  updateStatusAndResult();
}

function handlePreviewClick(event) {
  const zoom = event.target.closest('[data-preview-zoom]');
  if (zoom) {
    setPreviewZoom(zoom.dataset.previewZoom);
    return;
  }
  const refresh = event.target.closest('[data-refresh-preview]');
  if (!refresh || state.previewLoading) return;
  refreshPreview();
}

function handlePreviewWheel(event) {
  if (!event.ctrlKey || !event.target.closest('.instant-preview')) return;
  event.preventDefault();
  setPreviewZoom(event.deltaY < 0 ? 'in' : 'out');
}

function markPreviewDirty() {
  clearTimeout(state.previewTimer);
  state.previewDirty = true;
  state.previewRevision += 1;
  state.previewStatus = '';
}

function schedulePreview(delay = 0) {
  clearTimeout(state.previewTimer);
  state.previewTimer = setTimeout(refreshPreview, delay);
}

async function refreshPreview() {
  const type = state.active;
  const requestRevision = state.previewRevision;
  const requestId = state.previewRequestId + 1;
  state.previewRequestId = requestId;
  const payload = JSON.stringify(payloadFor(type));
  state.previewLoading = true;
  state.previewDirty = false;
  state.previewStatus = '正在生成最终 PDF...';
  if (state.preview?.pdfUrl) {
    updatePreviewChrome();
  } else {
    updateRightPanel();
  }
  try {
    const res = await fetch(`${apiBase}/api/preview/${type}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: payload,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '预览生成失败');
    if (type !== state.active || requestId !== state.previewRequestId) return;
    state.preview = data;
    state.previewLoading = false;
    state.previewDirty = state.previewRevision !== requestRevision;
    state.previewStatus = '';
  } catch (error) {
    if (type !== state.active || requestId !== state.previewRequestId) return;
    state.previewLoading = false;
    state.previewDirty = true;
    state.previewStatus = error.message;
  }
  updateRightPanel();
}

function updateRightPanel() {
  const panel = document.querySelector('#previewPanel');
  if (panel) panel.innerHTML = rightPanelHtml();
}

function updateFormPanel() {
  const form = document.querySelector('#formPanel');
  if (form) form.innerHTML = formHtml(state.active);
}

function updateInstantPreview() {
  const slot = document.querySelector('#instantPreviewSlot');
  if (slot) slot.innerHTML = documentPreview();
}

function setPreviewZoom(action) {
  const current = state.previewZoom[state.active] || 1;
  const next = action === 'reset'
    ? 1
    : current + (action === 'in' ? 0.1 : -0.1);
  state.previewZoom[state.active] = Math.min(1.8, Math.max(0.45, Number(next.toFixed(2))));
  updateInstantPreview();
}

function updateStatusAndResult() {
  const status = document.querySelector('#statusSlot');
  const result = document.querySelector('#resultSlot');
  if (status) status.innerHTML = statusHtml();
  if (result) result.innerHTML = resultHtml();
}

function updatePreviewChrome() {
  const status = document.querySelector('[data-preview-status]');
  const theme = document.querySelector('[data-preview-theme]');
  const button = document.querySelector('[data-refresh-preview]');
  const wrap = document.querySelector('.pdf-preview-wrap');
  if (status) {
    status.textContent = previewStatusText();
    status.className = `preview-badge ${previewStatusClass()}`;
  }
  if (theme) theme.textContent = currentTheme()?.label ?? '';
  if (button) button.disabled = state.previewLoading;
  if (wrap) {
    const overlay = wrap.querySelector('[data-preview-overlay]');
    if (state.previewLoading && !overlay) {
      wrap.insertAdjacentHTML('beforeend', '<div class="preview-overlay" data-preview-overlay>正在刷新最终 PDF，当前仍显示上一次结果</div>');
    }
    if (!state.previewLoading && overlay) overlay.remove();
  }
}

async function generate() {
  const type = state.active;
  state.status = '正在调用 Office/WPS 生成文件...';
  state.result = null;
  updateStatusAndResult();
  try {
    const res = await fetch(`${apiBase}/api/generate/${type}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payloadFor(type)),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '生成失败');
    state.result = data;
    state.preview = data;
    state.previewDirty = false;
    state.previewLoading = false;
    state.previewStatus = '';
    state.status = '生成完成。';
  } catch (error) {
    state.status = error.message;
  }
  updateRightPanel();
}

function currentTheme() {
  const form = state.forms[state.active];
  return (state.meta.themes || []).find((item) => item.id === form.themeId);
}

function payloadFor(type) {
  const payload = JSON.parse(JSON.stringify(state.forms[type]));
  if (payload.date) payload.date = documentDateText(payload.date);
  return payload;
}

function documentDateText(value) {
  const text = String(value ?? '').trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) return text;
  return `${Number(match[1])}年${Number(match[2])}月${Number(match[3])}日`;
}

function previewStatusText() {
  if (state.previewLoading) return state.preview?.pdfUrl ? '正在刷新最终 PDF' : '正在生成最终 PDF';
  if (state.previewStatus) return state.previewStatus;
  if (state.previewDirty) return '表单已修改，最终 PDF 待核对';
  if (state.preview?.pdfUrl) return '最终 PDF 已生成';
  return '未生成最终 PDF';
}

function previewStatusClass() {
  if (state.previewLoading) return 'loading';
  if (state.previewStatus) return 'error';
  if (state.previewDirty) return 'dirty';
  if (state.preview?.pdfUrl) return 'synced';
  return 'idle';
}

function stampImage(theme) {
  const stamp = selectedStampUrl(theme);
  return stamp ? `<img class="doc-stamp" src="${apiBase}${stamp}" alt="印章预览" onerror="this.hidden=true" />` : '';
}

function selectedStampUrl(theme) {
  if (!theme) return '';
  const form = state.forms[state.active];
  return form?.stampType === 'contract' ? (theme.contractStamp || '') : (theme.stamp || '');
}

function padRows(rows, count) {
  const padded = rows.slice(0, count);
  while (padded.length < count) padded.push({});
  return padded;
}

function visibleRows(rows) {
  return rows.length ? rows : [{}];
}

function sumAmounts(rows) {
  return rows.reduce((sum, row) => sum + numberValue(row.amount), 0);
}

function pathChangeInfo(parts) {
  return {
    collection: parts[1],
    index: Number(parts[2]),
    key: parts[3],
  };
}

function markManualAmount(type, change, row) {
  const collection = type === 'contract' ? 'items' : 'rows';
  if (change.collection !== collection || change.key !== 'amount') return;
  row.amountManual = !isBlank(row.amount);
}

function syncCalculatedFields(type, change = null) {
  const form = state.forms[type];
  if (!form) return;
  const collection = type === 'contract' ? 'items' : 'rows';
  const rows = form[collection] || [];
  rows.forEach((row, index) => {
    syncRowAmount(row, shouldRecalculateRowAmount(collection, index, change));
  });
  const total = sumAmounts(rows);
  if (type === 'contract') {
    form.totalAmount = total ? formatPlainNumber(total) : '';
    form.amountChinese = total ? rmbUpper(total).replace(/整$/, '') : '';
  }
}

function shouldRecalculateRowAmount(collection, index, change) {
  if (!change || change.collection !== collection || change.index !== index) return false;
  return change.key === 'quantity' || change.key === 'unitPrice';
}

function syncRowAmount(row, recalculate) {
  if (row.amountManual && recalculate && !row.amount) row.amountManual = false;
  if (!recalculate || row.amountManual) return;
  const hasQuantity = !isBlank(row.quantity);
  const hasUnitPrice = !isBlank(row.unitPrice);
  if (hasQuantity && hasUnitPrice) {
    row.amount = formatPlainNumber(numberValue(row.quantity) * numberValue(row.unitPrice));
  }
}

function refreshCalculatedInputs(type, change = null) {
  const form = state.forms[type];
  if (!form) return;
  const collection = type === 'contract' ? 'items' : 'rows';
  (form[collection] || []).forEach((row, index) => {
    if (shouldRecalculateRowAmount(collection, index, change) && !row.amountManual) {
      setInputValue(`${type}.${collection}.${index}.amount`, row.amount ?? '');
    }
  });
}

function setInputValue(path, value) {
  const input = document.querySelector(`[data-path="${cssEscape(path)}"]`);
  if (input) input.value = value;
}

function numberValue(value) {
  const number = Number(String(value ?? '').replace(/[^\d.-]/g, ''));
  return Number.isFinite(number) ? number : 0;
}

function isBlank(value) {
  return String(value ?? '').trim() === '';
}

function formatMoney(value) {
  return value ? value.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '';
}

function formatPlainNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return Number.isInteger(number) ? String(number) : String(Number(number.toFixed(2)));
}

function moneyText(value) {
  const number = numberValue(value);
  return number ? `￥${formatMoney(number)}` : '';
}

function rmbUpper(value) {
  const amount = Math.round((Number(value) || 0) * 100);
  if (!amount) return '零元整';
  const fraction = ['角', '分'];
  const digit = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
  const unit = [['元', '万', '亿'], ['', '拾', '佰', '仟']];
  let integer = Math.floor(amount / 100);
  let result = '';
  for (let i = 0; i < fraction.length; i++) {
    const number = Math.floor(amount / Math.pow(10, 1 - i)) % 10;
    result += number ? digit[number] + fraction[i] : '';
  }
  result = result || '整';
  for (let i = 0; integer > 0 && i < unit[0].length; i++) {
    let part = '';
    for (let j = 0; j < unit[1].length && integer > 0; j++) {
      part = digit[integer % 10] + unit[1][j] + part;
      integer = Math.floor(integer / 10);
    }
    part = part.replace(/(零.)*零$/, '').replace(/^$/, '零');
    result = part + unit[0][i] + result;
  }
  return result
    .replace(/零(拾|佰|仟)/g, '零')
    .replace(/零+/g, '零')
    .replace(/零(万|亿|元)/g, '$1')
    .replace(/亿万/g, '亿')
    .replace(/^元/, '零元')
    .replace(/零角零分$/, '整')
    .replace(/零分$/, '')
    .replace(/零角/, '零');
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/"/g, '&quot;');
}

function cssEscape(value) {
  if (window.CSS?.escape) return window.CSS.escape(value);
  return String(value).replace(/["\\]/g, '\\$&');
}
