(() => {
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

  const state = {
    token: null,
    preflight: null,
    reports: [],
    activeReport: 0,
    zipBase64: null,
    zipFilename: null,
  };

  const excelInput = $('#excelFile');
  const preflightBtn = $('#preflightBtn');
  const generateBtn = $('#generateBtn');
  const preflightPanel = $('#preflightPanel');
  const previewHost = $('#previewHost');
  const reportTabs = $('#reportTabs');
  const reportActions = $('#reportActions');
  const downloadAllBtn = $('#downloadAllBtn');

  function setStatus(text, type = 'idle') {
    $('#statusText').textContent = text;
    $('#statusDot').className = `status-dot ${type}`;
  }
  function setStage(name, done) {
    const item = $(`[data-stage="${name}"]`);
    if (item) item.classList.toggle('done', done);
  }
  function selectedScenarios() {
    return $$('.scenario-enabled:checked').map(x => x.dataset.scenario);
  }
  function inputsReady() {
    if (!excelInput.files[0]) return false;
    const scenarios = selectedScenarios();
    if (!scenarios.length) return false;
    return scenarios.every(s => $(`#model${s}`).files[0]);
  }

  function clearGeneratedReports() {
    state.reports = [];
    state.activeReport = 0;
    state.zipBase64 = null;
    state.zipFilename = null;
    reportTabs.classList.add('hidden');
    reportActions.classList.add('hidden');
    previewHost.classList.add('hidden');
    downloadAllBtn.classList.add('hidden');
    setStage('generation', false);
  }

  function invalidateGeneration() {
    if (!state.reports.length && !state.zipBase64) return;
    clearGeneratedReports();
    $('#generateHint').textContent = state.token
      ? 'Report settings changed. Generate again to apply them.'
      : 'Complete a successful preflight first.';
    setStatus(state.token ? 'Report settings changed' : 'Ready for inputs', 'idle');
  }

  function invalidatePreflight() {
    state.token = null;
    state.preflight = null;
    clearGeneratedReports();
    generateBtn.disabled = true;
    $('#generateHint').textContent = 'Complete a successful preflight first.';
    $('#preflightHint').textContent = inputsReady() ? 'Inputs changed. Run preflight to validate them.' : 'Select the workbook and a model for every enabled scenario.';
    preflightPanel.classList.add('hidden');
    setStage('preflight', false);
    setStage('files', inputsReady());
    setStatus(inputsReady() ? 'Inputs ready for preflight' : 'Ready for inputs', 'idle');
  }

  function bindDropZone(label, input) {
    ['dragenter', 'dragover'].forEach(evt => label.addEventListener(evt, e => { e.preventDefault(); label.classList.add('dragging'); }));
    ['dragleave', 'drop'].forEach(evt => label.addEventListener(evt, e => { e.preventDefault(); label.classList.remove('dragging'); }));
    label.addEventListener('drop', e => {
      if (!e.dataTransfer.files.length) return;
      const dt = new DataTransfer();
      dt.items.add(e.dataTransfer.files[0]);
      input.files = dt.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
  }
  bindDropZone($('#excelDrop'), excelInput);
  $$('.model-file').forEach(input => bindDropZone(input.closest('.model-picker'), input));

  excelInput.addEventListener('change', () => {
    $('#excelName').textContent = excelInput.files[0]?.name || 'Choose the current monthly workbook';
    invalidatePreflight();
  });

  $$('.scenario-enabled').forEach(toggle => {
    toggle.addEventListener('change', () => {
      const s = toggle.dataset.scenario;
      const card = $(`.scenario-card[data-scenario="${s}"]`);
      const input = $(`#model${s}`);
      card.classList.toggle('enabled', toggle.checked);
      if (!toggle.checked) {
        input.value = '';
        $(`#model${s}Name`).textContent = `Choose ${card.dataset.label} model`;
      }
      invalidatePreflight();
    });
  });

  $$('.model-file').forEach(input => {
    input.addEventListener('change', () => {
      const s = input.dataset.scenario;
      const card = $(`.scenario-card[data-scenario="${s}"]`);
      const toggle = $(`.scenario-enabled[data-scenario="${s}"]`);
      if (input.files[0] && !toggle.checked) {
        toggle.checked = true;
        card.classList.add('enabled');
      }
      $(`#model${s}Name`).textContent = input.files[0]?.name || `Choose ${card.dataset.label} model`;
      invalidatePreflight();
    });
  });

  function renderMessages(errors = [], warnings = []) {
    let html = '';
    for (const item of errors) html += `<div class="message error">${escapeHtml(item)}</div>`;
    for (const item of warnings) html += `<div class="message warning">${escapeHtml(item)}</div>`;
    return html;
  }

  function escapeHtml(value) {
    const div = document.createElement('div');
    div.textContent = value ?? '';
    return div.innerHTML;
  }

  function renderPreflight(data) {
    const rows = data.scenarios.map(row => `
      <tr>
        <td><strong>${escapeHtml(row.scenario_label)}</strong></td>
        <td>${escapeHtml(row.model_filename)}</td>
        <td>${escapeHtml(row.model_run_start)}<br><span class="muted">to ${escapeHtml(row.model_run_finish)}</span></td>
        <td>${row.report_years.join(', ')}</td>
        <td><span class="pass">${row.required_slots_found}/${row.required_slots_total} found</span></td>
        <td>${row.conservation_summary_included ? 'Yes' : 'No'}</td>
      </tr>`).join('');
    const scenarioWarnings = data.scenarios.flatMap(r => r.warnings || []);
    preflightPanel.innerHTML = `
      <div class="message success">Preflight passed. These parsed inputs are cached for report generation.</div>
      <div class="preflight-common">
        <span class="meta-chip">Workbook: ${escapeHtml(data.workbook_filename)}</span>
        <span class="meta-chip">Rules: ${escapeHtml(data.policy_version)}</span>
        <span class="meta-chip">Report years: ${data.report_years.join('–')}</span>
      </div>
      <table class="preflight-table">
        <thead><tr><th>Scenario</th><th>Model</th><th>Model run</th><th>Report years</th><th>RiverWare slots</th><th>Conservation summary</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
      ${renderMessages([], [...(data.warnings || []), ...scenarioWarnings])}`;
    preflightPanel.classList.remove('hidden');
  }

  preflightBtn.addEventListener('click', async () => {
    if (!inputsReady()) {
      preflightPanel.innerHTML = renderMessages(['Select the current workbook and a model file for every enabled scenario.']);
      preflightPanel.classList.remove('hidden');
      return;
    }
    const form = new FormData();
    form.append('excel_file', excelInput.files[0]);
    for (const s of selectedScenarios()) form.append(`model_${s}`, $(`#model${s}`).files[0]);

    preflightBtn.disabled = true;
    generateBtn.disabled = true;
    setStatus('Parsing models and validating workbook…', 'busy');
    $('#preflightHint').textContent = 'Reading RiverWare slots and scenario-specific workbook data…';
    try {
      const response = await fetch('/api/preflight', { method: 'POST', body: form });
      const data = await response.json();
      if (!data.ok) {
        state.token = null;
        preflightPanel.innerHTML = renderMessages(data.errors?.length ? data.errors : [data.error], data.warnings || []);
        preflightPanel.classList.remove('hidden');
        setStatus('Preflight needs attention', 'error');
        $('#preflightHint').textContent = 'Correct the listed issues and run preflight again.';
        return;
      }
      state.token = data.preflight.token;
      state.preflight = data.preflight;
      renderPreflight(data.preflight);
      generateBtn.disabled = false;
      $('#generateHint').textContent = 'Inputs are validated and ready.';
      $('#preflightHint').textContent = 'Preflight passed.';
      setStage('files', true);
      setStage('preflight', true);
      setStatus('Preflight passed', 'ok');
    } catch (error) {
      preflightPanel.innerHTML = renderMessages([`Could not complete preflight: ${error.message}`]);
      preflightPanel.classList.remove('hidden');
      setStatus('Preflight failed', 'error');
    } finally {
      preflightBtn.disabled = false;
    }
  });

  function base64ToBlob(base64, mime) {
    const binary = atob(base64);
    const chunk = 1024 * 1024;
    const parts = [];
    for (let offset = 0; offset < binary.length; offset += chunk) {
      const slice = binary.slice(offset, offset + chunk);
      const bytes = new Uint8Array(slice.length);
      for (let i = 0; i < slice.length; i++) bytes[i] = slice.charCodeAt(i);
      parts.push(bytes);
    }
    return new Blob(parts, { type: mime });
  }

  function downloadBase64(base64, filename, mime) {
    const blob = base64ToBlob(base64, mime);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  function renderReport(index) {
    state.activeReport = index;
    $$('.report-tab').forEach((b, i) => b.classList.toggle('active', i === index));
    const report = state.reports[index];
    previewHost.innerHTML = report.preview_html;
    previewHost.classList.remove('hidden');
    reportActions.innerHTML = `<button class="btn secondary" id="downloadReportBtn">Download ${escapeHtml(report.scenario_label)} Word report</button>`;
    reportActions.classList.remove('hidden');
    $('#downloadReportBtn').addEventListener('click', () => downloadBase64(
      report.docx_base64,
      report.filename,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    ));
  }

  function renderReports() {
    reportTabs.innerHTML = state.reports.map((r, i) => `<button class="report-tab ${i === 0 ? 'active' : ''}" data-index="${i}">${escapeHtml(r.scenario_label)}</button>`).join('');
    reportTabs.classList.remove('hidden');
    $$('.report-tab', reportTabs).forEach(btn => btn.addEventListener('click', () => renderReport(Number(btn.dataset.index))));
    if (state.zipBase64) downloadAllBtn.classList.remove('hidden'); else downloadAllBtn.classList.add('hidden');
    renderReport(0);
  }

  generateBtn.addEventListener('click', async () => {
    if (!state.token) return;
    const monYear = $('#studyMonth').value;
    if (!monYear) {
      $('#generateHint').textContent = 'Select the 24-Month Study month before generating.';
      $('#studyMonth').focus();
      return;
    }
    generateBtn.disabled = true;
    setStatus('Building canonical reports and Word documents…', 'busy');
    $('#generateHint').textContent = 'Rendering preview and Word output from the same report data…';
    try {
      const response = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ token: state.token, mon_year: monYear })
      });
      const data = await response.json();
      if (!data.ok) {
        $('#generateHint').textContent = data.error;
        if (response.status === 409) invalidatePreflight();
        setStatus('Generation failed', 'error');
        return;
      }
      state.reports = data.reports;
      state.zipBase64 = data.zip_base64 || null;
      state.zipFilename = data.zip_filename || null;
      renderReports();
      $('#generateHint').textContent = `${data.reports.length} report${data.reports.length === 1 ? '' : 's'} generated successfully.`;
      setStage('generation', true);
      setStatus('Reports generated', 'ok');
    } catch (error) {
      $('#generateHint').textContent = `Could not generate reports: ${error.message}`;
      setStatus('Generation failed', 'error');
    } finally {
      generateBtn.disabled = !state.token;
    }
  });

  downloadAllBtn.addEventListener('click', () => {
    if (state.zipBase64) downloadBase64(state.zipBase64, state.zipFilename, 'application/zip');
  });

  previewHost.addEventListener('click', event => {
    const button = event.target.closest('.source-link');
    if (!button || !state.reports.length) return;
    const sourceId = button.dataset.sourceId;
    const refs = state.reports[state.activeReport].source_map[sourceId] || [];
    showSources(refs);
  });

  function showSources(refs) {
    const content = $('#sourceContent');
    if (!refs.length) {
      content.innerHTML = '<p class="muted">No source details are available for this item.</p>';
    } else {
      content.innerHTML = refs.map(ref => `
        <article class="source-card">
          <span class="source-kind">${escapeHtml(ref.kind)}</span>
          <dl>
            <dt>Source</dt><dd>${escapeHtml(ref.source_name || '')}</dd>
            <dt>Location</dt><dd>${escapeHtml(ref.locator || '')}</dd>
            ${ref.year ? `<dt>Year</dt><dd>${escapeHtml(ref.year)}</dd>` : ''}
            ${ref.raw_value !== null && ref.raw_value !== undefined ? `<dt>Value</dt><dd>${escapeHtml(ref.raw_value)} ${escapeHtml(ref.units || '')}</dd>` : ''}
            ${ref.note ? `<dt>Detail</dt><dd>${escapeHtml(ref.note)}</dd>` : ''}
          </dl>
        </article>`).join('');
    }
    const drawer = $('#sourceDrawer');
    drawer.classList.add('open'); drawer.setAttribute('aria-hidden', 'false');
  }
  function closeSourceDrawer() {
    const drawer = $('#sourceDrawer');
    drawer.classList.remove('open');
    drawer.setAttribute('aria-hidden', 'true');
  }

  function closeRulesDrawer() {
    const drawer = $('#rulesDrawer');
    drawer.classList.remove('open');
    drawer.setAttribute('aria-hidden', 'true');
  }

  $$('[data-close-drawer]').forEach(el => el.addEventListener('click', closeSourceDrawer));
  $$('[data-close-rules]').forEach(el => el.addEventListener('click', closeRulesDrawer));

  function effectiveRules(items = []) {
    return [...items]
      .sort((a, b) => Number(a.effective_from_year || 0) - Number(b.effective_from_year || 0))
      .map(item => `<li><strong>From ${Number(item.effective_from_year || 0) || 'all years'}:</strong> ${escapeHtml(item.label || item.text || '')}${item.decompose !== undefined ? ` <span class="rule-detail">(detail decomposition: ${item.decompose ? 'yes' : 'no'})</span>` : ''}</li>`)
      .join('');
  }

  function renderRules(data) {
    const p = data.policy || {};
    const report = p.report || {};
    const rules = report.rules || {};
    const scenarios = Object.entries(p.scenarios || {}).map(([key, cfg]) => `
      <tr><td><strong>${escapeHtml(cfg.display_name || key)}</strong></td><td>${escapeHtml(key)}</td><td>${escapeHtml(cfg.excel?.label_column || '')} / ${escapeHtml(cfg.excel?.value_column || '')} / ${escapeHtml(cfg.excel?.notes_column || '')}</td><td>${(report.conservation_summary_scenarios || []).includes(key) ? 'Yes' : 'No'}</td><td>${(report.powell_release_subtitle_scenarios || []).includes(key) ? 'Yes' : 'No'}</td></tr>`).join('');
    const caRoll = rules.california_system_conservation_rollup || {};
    const azOverride = rules.az_water_use_reduction_override || {};
    const slots = p.riverware?.required_slots || [];
    $('#rulesContent').innerHTML = `
      <div class="rule-summary">
        <span class="meta-chip">Version ${escapeHtml(data.policy_version)}</span>
        <p>${escapeHtml(data.description || '')}</p>
        <p><strong>Editable configuration:</strong> <code>${escapeHtml(data.config_file)}</code>. Changes are loaded when the application is restarted.</p>
      </div>
      <section class="rule-section"><h3>Scenario rules</h3>
        <table class="rules-table"><thead><tr><th>Display name</th><th>Key</th><th>Excel label / value / notes</th><th>Conservation table</th><th>Powell subtitle</th></tr></thead><tbody>${scenarios}</tbody></table>
        <p class="rule-detail">Default enabled scenarios: ${(report.default_enabled_scenarios || []).map(escapeHtml).join(', ')}.</p>
      </section>
      <section class="rule-section"><h3>Report content</h3>
        <p><strong>Powell subtitle:</strong> ${escapeHtml(report.powell_release_subtitle || '')}</p>
        <p><strong>Notes and Disclaimers:</strong> ${escapeHtml(report.notes_conservation_disclaimer || '')}</p>
        <p><strong>California conservation roll-up:</strong> A California water-left-in-Mead value is added into Total California System Conservation when its workbook note contains: ${(caRoll.year_sheet_note_labels || []).map(v => `<code>${escapeHtml(v)}</code>`).join(', ') || 'none'}.</p>
      </section>
      <section class="rule-section"><h3>Effective-year wording and overrides</h3>
        <h4>Total-use disclaimer</h4><ul>${effectiveRules(report.disclaimers || [])}</ul>
        <h4>Mexico shortage/reduction wording</h4><ul>${effectiveRules(rules.mexico_shortage_wording || [])}</ul>
        <h4>Nevada shortage/reduction wording</h4><ul>${effectiveRules(rules.nevada_shortage_wording || [])}</ul>
        <h4>Nevada conservation wording</h4><ul>${effectiveRules(rules.nevada_system_conservation_wording || [])}</ul>
        <p><strong>Arizona reduction display override:</strong> ${azOverride.display_kaf !== undefined ? `${escapeHtml(azOverride.display_kaf)} kaf from ${escapeHtml(azOverride.effective_from_year)} onward` : 'None'}.</p>
      </section>
      <section class="rule-section"><h3>Source validation</h3>
        <p><strong>Excel:</strong> System Conservation sheet <code>${escapeHtml(p.excel?.system_conservation_sheet || '')}</code>; summary sheet <code>${escapeHtml(p.excel?.summary_sheet || '')}</code>.</p>
        <p><strong>RiverWare:</strong> ${slots.length} required annual slots; ${escapeHtml(p.riverware?.target_year_count || '')} target years; accepted volume units: ${(p.riverware?.accepted_volume_units || []).map(escapeHtml).join(', ')}.</p>
        <details><summary>Show required RiverWare slots</summary><ul class="slot-list">${slots.map(slot => `<li><code>${escapeHtml(slot)}</code></li>`).join('')}</ul></details>
      </section>
      <details class="raw-rules"><summary>Show full active configuration (JSON)</summary><pre>${escapeHtml(JSON.stringify(p, null, 2))}</pre></details>`;
  }

  $('#rulesButton')?.addEventListener('click', async () => {
    const drawer = $('#rulesDrawer');
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden', 'false');
    $('#rulesContent').innerHTML = '<p class="muted">Loading active rules…</p>';
    try {
      const response = await fetch('/api/rules');
      const data = await response.json();
      if (!data.ok) throw new Error(data.error || 'Could not load rules.');
      renderRules(data);
    } catch (error) {
      $('#rulesContent').innerHTML = `<div class="message error">Could not load active rules: ${escapeHtml(error.message)}</div>`;
    }
  });

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') { closeSourceDrawer(); closeRulesDrawer(); }
  });

  // Report settings do not affect source preflight, but changing them invalidates an already generated preview/download.
  $('#studyMonth').addEventListener('input', invalidateGeneration);

  // Default the study month to the current local month; users can change it.
  const now = new Date();
  $('#studyMonth').value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  invalidatePreflight();
})();
