const QUESTION_ROLES = [
  { key: "tooCheap", label: "Liian halpa" },
  { key: "bargain", label: "Edullinen" },
  { key: "expensive", label: "Kallis" },
  { key: "tooExpensive", label: "Liian kallis" },
];

const state = {
  variables: [],
  rows: [],
  chartMain: null,
  metrics: null,
  clusters: null,
  groupBy: "none",
  filterGroup: "all",
  filterRange: null,
  filterCategories: new Set(),
  optimizer: {
    running: false,
    cancel: false,
    selectedVarIds: new Set(),
    results: [],
  },
};

const psmLabelPlugin = {
  id: "psmLabelPlugin",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      if (!dataset.psmLabels) return;
      if (!chart.isDatasetVisible(datasetIndex)) return;
      const meta = chart.getDatasetMeta(datasetIndex);
      meta.data.forEach((point, index) => {
        const raw = dataset.data[index];
        if (!raw || !raw.label) return;
        const { x, y } = point.getProps(["x", "y"], true);
        ctx.save();
        ctx.font = "600 12px Segoe UI";
        ctx.fillStyle = raw.color || "#111";
        ctx.fillText(raw.label, x + 8, y - 8);
        ctx.restore();
      });
    });
  },
};

const ui = {
  tabButtons: document.querySelectorAll(".tab"),
  tabPanels: document.querySelectorAll(".tab-panel"),
  variableTableBody: document.querySelector("#variableTable tbody"),
  dataTableHead: document.querySelector("#dataTable thead"),
  dataTableBody: document.querySelector("#dataTable tbody"),
  addRowBtn: document.getElementById("addRowBtn"),
  addVarBtn: document.getElementById("addVarBtn"),
  fileLoader: document.getElementById("fileLoader"),
  jsonLoader: document.getElementById("jsonLoader"),
  saveJsonBtn: document.getElementById("saveJsonBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportPptBtn: document.getElementById("exportPptBtn"),
  groupSelect: document.getElementById("groupSelect"),
  filterControl: document.getElementById("filterControl"),
  filterLabel: document.getElementById("filterLabel"),
  clusterSelect: document.getElementById("clusterSelect"),
  clusterK: document.getElementById("clusterK"),
  runClusterBtn: document.getElementById("runClusterBtn"),
  runClusterTabulateBtn: document.getElementById("runClusterTabulateBtn"),
  clusterName: document.getElementById("clusterName"),
  clusterMeanLabels: document.getElementById("clusterMeanLabels"),
  clusterMeanLabelsRow: document.getElementById("clusterMeanLabelsRow"),
  clusterSelectMessage: document.getElementById("clusterSelectMessage"),
  optimizerVars: document.getElementById("optimizerVars"),
  optimizerTopN: document.getElementById("optimizerTopN"),
  optimizerStartBtn: document.getElementById("optimizerStartBtn"),
  optimizerStopBtn: document.getElementById("optimizerStopBtn"),
  optimizerResults: document.getElementById("optimizerResults"),
  optimizerStatus: document.getElementById("optimizerStatus"),
  filterCard: document.getElementById("filterCard"),
  filterSlotAnalysis: document.getElementById("filterSlotAnalysis"),
  filterSlotViz: document.getElementById("filterSlotViz"),
  metricsForecastCard: document.getElementById("metricsForecastCard"),
  pivotCard: document.getElementById("pivotCard"),
  vizChartCard: document.getElementById("vizChartCard"),
  pivotRow: document.getElementById("pivotRow"),
  pivotCol: document.getElementById("pivotCol"),
  pivotValue: document.getElementById("pivotValue"),
  pivotHead: document.querySelector("#pivotTable thead"),
  pivotBody: document.querySelector("#pivotTable tbody"),
  statusText: document.getElementById("statusText"),
  forecastPrice: document.getElementById("forecastPrice"),
  forecastTooCheap: document.getElementById("forecastTooCheap"),
  forecastBargain: document.getElementById("forecastBargain"),
  forecastExpensive: document.getElementById("forecastExpensive"),
  forecastTooExpensive: document.getElementById("forecastTooExpensive"),
  oppValue: document.getElementById("oppValue"),
  pmcValue: document.getElementById("pmcValue"),
  pmeValue: document.getElementById("pmeValue"),
  ippValue: document.getElementById("ippValue"),
  rangeValue: document.getElementById("rangeValue"),
  showTooCheap: null,
  showBargain: null,
  showExpensive: null,
  showTooExpensive: null,
  showPoints: null,
  colorByGroup: null,
  chartTheme: null,
  guideBadge: document.querySelector(".tab-badge"),
};

function init() {
  initVariables();
  initRows(14);
  bindEvents();
  renderAll();
  buildChart();
  setStatus("Valmis.");
  scheduleRecalc();
}

function initVariables() {
  state.variables = [
    { id: crypto.randomUUID(), name: "Liian halpa", type: "continuous", role: "tooCheap" },
    { id: crypto.randomUUID(), name: "Edullinen", type: "continuous", role: "bargain" },
    { id: crypto.randomUUID(), name: "Kallis", type: "continuous", role: "expensive" },
    { id: crypto.randomUUID(), name: "Liian kallis", type: "continuous", role: "tooExpensive" },
    { id: crypto.randomUUID(), name: "Ikä", type: "continuous", role: "background" },
    { id: crypto.randomUUID(), name: "Sukupuoli", type: "categorical", role: "background" },
  ];
}

function initRows(count) {
  state.rows = Array.from({ length: count }, () => createEmptyRow());
  seedDemoData();
}

function createEmptyRow() {
  const row = { id: crypto.randomUUID(), values: {}, _cluster: null };
  state.variables.forEach((v) => {
    row.values[v.id] = "";
  });
  return row;
}

function seedDemoData() {
  const questionVars = getQuestionVariables();
  const ageVar = state.variables.find((v) => v.name === "Ikä");
  const genderVar = state.variables.find((v) => v.name === "Sukupuoli");
  if (!ageVar || !genderVar) return;

  const demoData = {
    tooCheap: [2.15, 4.15, 2.64, 4.53, 4.76, 1.18, 3.11, 4.57, 3.21, 2.83, 4.83, 2.81, 3.71, 3.29, 1.41, 4.6, 1.98, 1.17, 2.31, 4.82],
    bargain: [5.56, 4.77, 4.56, 5.98, 4.62, 4.83, 4.18, 4.38, 3.16, 2.59, 5.85, 5.61, 4.76, 5.18, 2.1, 3.91, 5.03, 2.87, 3.27, 2.93],
    expensive: [3.57, 4.66, 4.65, 4.48, 3.61, 3.56, 3.93, 4.86, 4.06, 6.43, 3.18, 4.77, 6.2, 3.49, 5.24, 3.83, 3.51, 6.01, 6.58, 4.5],
    tooExpensive: [6.66, 4.38, 5.54, 5.1, 7.26, 5.79, 7.24, 7.25, 7.18, 5.76, 7.02, 6.52, 6.84, 4.0, 5.9, 4.88, 5.52, 6.45, 5.41, 4.44],
  };

  const ages = [19, 22, 25, 27, 29, 31, 33, 35, 37, 39, 41, 44, 46, 48, 51, 53, 55, 58, 60, 62];
  const genders = [
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
    "Nainen",
    "Mies",
  ];

  state.rows = demoData.tooCheap.map((value, idx) => {
    const row = createEmptyRow();
    row.values[questionVars.tooCheap.id] = value;
    row.values[questionVars.bargain.id] = demoData.bargain[idx];
    row.values[questionVars.expensive.id] = demoData.expensive[idx];
    row.values[questionVars.tooExpensive.id] = demoData.tooExpensive[idx];
    row.values[ageVar.id] = ages[idx];
    row.values[genderVar.id] = genders[idx];
    return row;
  });
}

function bindEvents() {
  ui.tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });

  ui.addRowBtn.addEventListener("click", () => {
    state.rows.push(createEmptyRow());
    renderDataTable();
    scheduleRecalc();
  });

  ui.addVarBtn.addEventListener("click", () => {
    addVariable();
    scheduleRecalc();
  });

  ui.fileLoader.addEventListener("change", handleImport);
  ui.jsonLoader.addEventListener("change", handleJsonLoad);
  ui.saveJsonBtn.addEventListener("click", handleJsonSave);
  ui.exportExcelBtn.addEventListener("click", exportExcel);
  ui.exportPptBtn.addEventListener("click", exportPpt);

  ui.groupSelect.addEventListener("change", () => {
    state.groupBy = ui.groupSelect.value;
    updateGroupFilter();
    scheduleRecalc();
  });

  ui.runClusterBtn.addEventListener("click", () => runKMeans());
  ui.runClusterTabulateBtn.addEventListener("click", () => runKMeansAndTabulate());

  ui.clusterName.addEventListener("input", () => {
    updateClusterName();
    scheduleRecalc();
  });
  if (ui.clusterK) {
    ui.clusterK.addEventListener("input", () => updateClusterActionButtons());
  }

  [ui.pivotRow, ui.pivotCol, ui.pivotValue].forEach((el) => {
    el.addEventListener("change", () => updatePivot());
  });

  ui.forecastPrice.addEventListener("input", () => updateForecast());

  if (ui.optimizerStartBtn) {
    ui.optimizerStartBtn.addEventListener("click", () => startOptimization());
  }
  if (ui.optimizerStopBtn) {
    ui.optimizerStopBtn.addEventListener("click", () => stopOptimization());
  }

  if (ui.guideBadge) {
    const stopGuideBadge = () => ui.guideBadge.remove();
    document.addEventListener("click", stopGuideBadge, { once: true });
  }

  // Visual settings removed; Chart.js legend handles visibility.
}

function switchTab(tabName) {
  ui.tabButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.tab === tabName));
  ui.tabPanels.forEach((panel) => panel.classList.toggle("active", panel.id === `tab-${tabName}`));
  if (ui.filterCard && ui.filterSlotAnalysis && ui.filterSlotViz) {
    if (tabName === "viz") {
      ui.filterSlotViz.appendChild(ui.filterCard);
    } else {
      ui.filterSlotAnalysis.appendChild(ui.filterCard);
    }
  }
}

function addVariable() {
  const newVar = {
    id: crypto.randomUUID(),
    name: `Muuttuja ${state.variables.length + 1}`,
    type: "continuous",
    role: "background",
  };
  state.variables.push(newVar);
  state.rows.forEach((row) => {
    row.values[newVar.id] = "";
  });
  renderAll();
}

function renderAll() {
  renderVariableTable();
  renderDataTable();
  updateGroupControls();
  renderOptimizerControls();
}

function renderVariableTable() {
  ui.variableTableBody.innerHTML = "";
  state.variables.forEach((variable, index) => {
    const tr = document.createElement("tr");

    const nameCell = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.value = variable.name;
    nameInput.addEventListener("input", (e) => {
      variable.name = e.target.value;
      renderDataTable();
      updateGroupControls();
      scheduleRecalc();
    });
    nameCell.appendChild(nameInput);

    const typeCell = document.createElement("td");
    const typeSelect = document.createElement("select");
    [
      { value: "continuous", label: "jatkuva" },
      { value: "categorical", label: "kategorinen" },
      { value: "other", label: "muu" },
    ].forEach((type) => {
      const option = document.createElement("option");
      option.value = type.value;
      option.textContent = type.label;
      if (variable.type === type.value) option.selected = true;
      typeSelect.appendChild(option);
    });
    typeSelect.addEventListener("change", (e) => {
      variable.type = e.target.value;
      renderDataTable();
      updateGroupControls();
      scheduleRecalc();
    });
    typeCell.appendChild(typeSelect);

    const roleCell = document.createElement("td");
    const roleSelect = document.createElement("select");
    const roleOptions = [...QUESTION_ROLES.map((r) => r.key), "background"];
    roleOptions.forEach((role) => {
      const option = document.createElement("option");
      option.value = role;
      const roleLabel = QUESTION_ROLES.find((r) => r.key === role)?.label || "Taustamuuttuja";
      option.textContent = roleLabel;
      if (variable.role === role) option.selected = true;
      roleSelect.appendChild(option);
    });
    roleSelect.addEventListener("change", (e) => {
      const newRole = e.target.value;
      if (newRole !== "background") {
        const existing = state.variables.find((v) => v.role === newRole);
        if (existing) existing.role = "background";
      }
      variable.role = newRole;
      renderVariableTable();
      updateGroupControls();
      scheduleRecalc();
    });
    roleCell.appendChild(roleSelect);

    const deleteCell = document.createElement("td");
    const deleteBtn = document.createElement("button");
    deleteBtn.className = "btn";
    deleteBtn.textContent = "Poista";
    deleteBtn.disabled = QUESTION_ROLES.some((r) => r.key === variable.role);
    deleteBtn.addEventListener("click", () => removeVariable(variable.id));
    deleteCell.appendChild(deleteBtn);

    tr.append(nameCell, typeCell, roleCell, deleteCell);
    ui.variableTableBody.appendChild(tr);
  });
}

function renderDataTable() {
  ui.dataTableHead.innerHTML = "";
  ui.dataTableBody.innerHTML = "";

  const headRow = document.createElement("tr");
  state.variables.forEach((variable) => {
    const th = document.createElement("th");
    const nameInput = document.createElement("input");
    nameInput.value = variable.name;
    nameInput.addEventListener("input", (e) => {
      variable.name = e.target.value;
      renderVariableTable();
      updateGroupControls();
      scheduleRecalc();
    });
    th.appendChild(nameInput);
    headRow.appendChild(th);
  });
  ui.dataTableHead.appendChild(headRow);

  state.rows.forEach((row, rowIndex) => {
    const tr = document.createElement("tr");
    state.variables.forEach((variable) => {
      const td = document.createElement("td");
      const input = document.createElement("input");
      input.value = row.values[variable.id] ?? "";
      input.addEventListener("input", (e) => {
        row.values[variable.id] = e.target.value;
        validateInput(e.target, variable);
        scheduleRecalc();
      });
      validateInput(input, variable);
      td.appendChild(input);
      tr.appendChild(td);
    });
    ui.dataTableBody.appendChild(tr);
  });
}

function validateInput(input, variable) {
  if (variable.type === "continuous") {
    const value = input.value.trim();
    const isValid = value !== "" && !Number.isNaN(Number(value));
    input.classList.toggle("invalid", !isValid);
  } else {
    input.classList.remove("invalid");
  }
}

function removeVariable(id) {
  state.variables = state.variables.filter((v) => v.id !== id);
  state.rows.forEach((row) => delete row.values[id]);
  renderAll();
  scheduleRecalc();
}

function updateGroupControls() {
  const backgroundVars = state.variables.filter(
    (v) => v.role === "background" && v.type !== "other"
  );
  ui.groupSelect.innerHTML = "";
  const noneOpt = document.createElement("option");
  noneOpt.value = "none";
  noneOpt.textContent = "Ei ryhmittelyä";
  ui.groupSelect.appendChild(noneOpt);

  backgroundVars.forEach((v) => {
    const option = document.createElement("option");
    option.value = v.id;
    option.textContent = v.name;
    ui.groupSelect.appendChild(option);
  });
  if (!backgroundVars.map((v) => v.id).includes(state.groupBy)) {
    state.groupBy = "none";
  }
  ui.groupSelect.value = state.groupBy;

  ui.clusterSelect.innerHTML = "";
  backgroundVars
    .filter((v) => v.type === "continuous")
    .forEach((v) => {
      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = v.id;
      checkbox.addEventListener("change", () => {
        setClusterSelectMessage("");
        updateClusterActionButtons();
      });
      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(v.name));
      ui.clusterSelect.appendChild(label);
    });

  updateClusterActionButtons();
  updateGroupFilter();
  updatePivotControls();
}

function updateClusterActionButtons() {
  const selected = getClusterSelectedIds();
  const showTabulate = selected.length === 1;
  ui.runClusterTabulateBtn.classList.toggle("hidden", !showTabulate);
  if (ui.clusterMeanLabelsRow) {
    ui.clusterMeanLabelsRow.classList.toggle("hidden", !showTabulate);
  }
  if (!showTabulate && ui.clusterMeanLabels) {
    ui.clusterMeanLabels.checked = false;
  }
}

function renderOptimizerControls() {
  if (!ui.optimizerVars) return;
  const backgroundVars = getBackgroundVars();
  ui.optimizerVars.innerHTML = "";
  if (!state.optimizer.selectedVarIds.size && backgroundVars.length) {
    state.optimizer.selectedVarIds = new Set(backgroundVars.map((v) => v.id));
  }

  backgroundVars.forEach((variable) => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = variable.id;
    checkbox.checked = state.optimizer.selectedVarIds.has(variable.id);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.optimizer.selectedVarIds.add(variable.id);
      else state.optimizer.selectedVarIds.delete(variable.id);
      updateOptimizerButtons();
    });
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(variable.name));
    ui.optimizerVars.appendChild(label);
  });

  updateOptimizerButtons();
  renderOptimizerResults(state.optimizer.results || []);
}

function updateOptimizerButtons() {
  if (!ui.optimizerStartBtn || !ui.optimizerStopBtn) return;
  const hasSelection = state.optimizer.selectedVarIds.size > 0;
  ui.optimizerStartBtn.disabled = state.optimizer.running || !hasSelection;
  ui.optimizerStopBtn.disabled = !state.optimizer.running;
}

function startOptimization() {
  if (state.optimizer.running) return;
  const questionVars = getQuestionVariables();
  if (Object.values(questionVars).some((v) => !v)) {
    setOptimizerStatus("Määritä kaikki neljä kysymysmuuttujaa ennen hakua.");
    return;
  }
  state.optimizer.cancel = false;
  state.optimizer.running = true;
  state.optimizer.results = [];
  renderOptimizerResults([]);
  updateOptimizerButtons();
  runOptimization();
}

function stopOptimization() {
  state.optimizer.cancel = true;
  setOptimizerStatus("Haku keskeytetään...");
}

async function runOptimization() {
  const topN = Math.max(1, Number(ui.optimizerTopN?.value || 10));
  const selectedVars = getBackgroundVars().filter((v) =>
    state.optimizer.selectedVarIds.has(v.id)
  );
  if (!selectedVars.length) {
    setOptimizerStatus("Valitse vähintään yksi taustamuuttuja.");
    state.optimizer.running = false;
    updateOptimizerButtons();
    return;
  }

  const predicatesByVar = selectedVars
    .map((variable) => buildVariablePredicates(variable))
    .filter((list) => list.length > 0);

  if (!predicatesByVar.length) {
    setOptimizerStatus("Valituista muuttujista ei löytynyt kelvollisia segmenttejä.");
    state.optimizer.running = false;
    updateOptimizerButtons();
    return;
  }

  const totalComb = predicatesByVar.reduce((acc, list) => acc * list.length, 1);
  const indices = new Array(predicatesByVar.length).fill(0);
  const rows = state.rows;
  const questionVars = getQuestionVariables();
  let checked = 0;
  const results = [];

  setOptimizerStatus(`Haku käynnissä... 0 / ${totalComb}`);

  while (true) {
    if (state.optimizer.cancel) break;
    const predicates = predicatesByVar.map((list, idx) => list[indices[idx]]);
    const segmentRows = rows.filter((row) => predicates.every((p) => p.test(row)));
    const responses = getResponsesForRows(segmentRows, questionVars);
    if (responses) {
      const metrics = computeMetrics(responses);
      const opp = metrics.opp;
      if (!Number.isNaN(opp)) {
        const size = responses.tooCheap.length;
        const score = size * opp;
        const description = predicates.map((p) => p.label).join("; ");
        results.push({ description, opp, size, score });
        results.sort((a, b) => b.score - a.score);
        if (results.length > topN) results.length = topN;
      }
    }

    checked += 1;
    if (checked % 200 === 0) {
      setOptimizerStatus(`Haku käynnissä... ${checked} / ${totalComb}`);
      renderOptimizerResults(results);
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    let pos = indices.length - 1;
    while (pos >= 0) {
      indices[pos] += 1;
      if (indices[pos] < predicatesByVar[pos].length) break;
      indices[pos] = 0;
      pos -= 1;
    }
    if (pos < 0) break;
  }

  state.optimizer.running = false;
  state.optimizer.results = results;
  renderOptimizerResults(results);
  updateOptimizerButtons();
  if (state.optimizer.cancel) {
    setOptimizerStatus("Haku keskeytetty.");
  } else {
    setOptimizerStatus(`Haku valmis. Tarkastettu ${checked} segmenttiä.`);
  }
}

function buildVariablePredicates(variable) {
  if (variable.type === "categorical") {
    const values = getUniqueValues(variable.id);
    return values.map((value) => ({
      label: `${variable.name} = ${value}`,
      test: (row) => String(row.values[variable.id]) === value,
    }));
  }
  if (variable.type === "continuous") {
    const values = getNumericValues(variable.id);
    const breakpoints = getBreakpoints(values, 10);
    const predicates = [];
    for (let i = 0; i < breakpoints.length - 1; i += 1) {
      for (let j = i + 1; j < breakpoints.length; j += 1) {
        const min = breakpoints[i];
        const max = breakpoints[j];
        predicates.push({
          label: `${formatNumber(min)} ≤ ${variable.name} ≤ ${formatNumber(max)}`,
          test: (row) => {
            const value = Number(row.values[variable.id]);
            if (Number.isNaN(value)) return false;
            return value >= min && value <= max;
          },
        });
      }
    }
    return predicates;
  }
  return [];
}

function getNumericValues(varId) {
  return state.rows
    .map((row) => Number(row.values[varId]))
    .filter((value) => !Number.isNaN(value));
}

function getUniqueValues(varId) {
  const set = new Set();
  state.rows.forEach((row) => {
    const value = row.values[varId];
    if (value !== "" && value != null) set.add(String(value));
  });
  return Array.from(set);
}

function getBreakpoints(values, maxPoints) {
  if (!values.length) return [];
  const sorted = Array.from(new Set(values)).sort((a, b) => a - b);
  if (sorted.length <= maxPoints) return sorted;
  const points = [];
  for (let i = 0; i < maxPoints; i += 1) {
    const idx = Math.round((i / (maxPoints - 1)) * (sorted.length - 1));
    points.push(sorted[idx]);
  }
  return Array.from(new Set(points));
}

function getBackgroundVars() {
  return state.variables.filter((v) => v.role === "background" && v.type !== "other");
}

function renderOptimizerResults(results) {
  if (!ui.optimizerResults) return;
  ui.optimizerResults.innerHTML = "";
  if (!results.length) {
    ui.optimizerResults.textContent = "Ei tuloksia.";
    return;
  }
  results.forEach((item, idx) => {
    const row = document.createElement("div");
    row.className = "result-row";
    row.textContent = `${idx + 1}. ${item.description} — OPP ${formatCurrency(item.opp)} (n=${item.size}); OPP * n = ${formatCurrency(item.score)}`;
    ui.optimizerResults.appendChild(row);
  });
}

function setOptimizerStatus(message) {
  if (!ui.optimizerStatus) return;
  ui.optimizerStatus.textContent = message;
}

function updatePivotControls() {
  const backgroundVars = state.variables.filter(
    (v) => v.role === "background" && v.type !== "other"
  );
  const numericVars = state.variables.filter((v) => v.type === "continuous");

  ui.pivotRow.innerHTML = "";
  ui.pivotCol.innerHTML = "";
  ui.pivotValue.innerHTML = "";

  const emptyRow = document.createElement("option");
  emptyRow.value = "";
  emptyRow.textContent = "(ei riviä)";
  ui.pivotRow.appendChild(emptyRow);

  const emptyCol = document.createElement("option");
  emptyCol.value = "";
  emptyCol.textContent = "(ei saraketta)";
  ui.pivotCol.appendChild(emptyCol);

  backgroundVars.forEach((v) => {
    const rowOpt = document.createElement("option");
    rowOpt.value = v.id;
    rowOpt.textContent = v.name;
    ui.pivotRow.appendChild(rowOpt);

    const colOpt = document.createElement("option");
    colOpt.value = v.id;
    colOpt.textContent = v.name;
    ui.pivotCol.appendChild(colOpt);
  });

  numericVars.forEach((v) => {
    const valOpt = document.createElement("option");
    valOpt.value = v.id;
    valOpt.textContent = v.name;
    ui.pivotValue.appendChild(valOpt);
  });

  if (!ui.pivotRow.value && backgroundVars[0]) ui.pivotRow.value = backgroundVars[0].id;
  if (!ui.pivotCol.value && backgroundVars[1]) ui.pivotCol.value = backgroundVars[1].id;
  if (!ui.pivotCol.value && backgroundVars[0]) ui.pivotCol.value = backgroundVars[0].id;
  if (!ui.pivotValue.value && numericVars[0]) ui.pivotValue.value = numericVars[0].id;

  updatePivot();
}

function updateGroupFilter() {
  if (!ui.filterControl || !ui.filterLabel) return;
  ui.filterControl.innerHTML = "";
  const groupVar = state.variables.find((v) => v.id === state.groupBy);
  if (!groupVar) {
    ui.filterLabel.textContent = "Suodata ryhmä";
    const select = document.createElement("select");
    const opt = document.createElement("option");
    opt.value = "all";
    opt.textContent = "Kaikki";
    select.appendChild(opt);
    select.disabled = true;
    ui.filterControl.appendChild(select);
    state.filterGroup = "all";
    state.filterRange = null;
    state.filterCategories = new Set();
    return;
  }

  if (groupVar.type === "continuous") {
    ui.filterLabel.textContent = "Suodata väli";
    const values = state.rows
      .map((row) => Number(row.values[groupVar.id]))
      .filter((v) => !Number.isNaN(v));
    const minVal = values.length ? Math.min(...values) : 0;
    const maxVal = values.length ? Math.max(...values) : 0;

    if (!state.filterRange || state.filterRange.varId !== groupVar.id) {
      state.filterRange = { varId: groupVar.id, min: minVal, max: maxVal };
    }

    const group = document.createElement("div");
    group.className = "range-group";

    const stack = document.createElement("div");
    stack.className = "range-stack";

    const minInput = document.createElement("input");
    minInput.type = "range";
    minInput.min = minVal;
    minInput.max = maxVal;
    minInput.step = "0.01";
    minInput.value = state.filterRange.min;
    minInput.className = "range-handle min-handle";

    const maxInput = document.createElement("input");
    maxInput.type = "range";
    maxInput.min = minVal;
    maxInput.max = maxVal;
    maxInput.step = "0.01";
    maxInput.value = state.filterRange.max;
    maxInput.className = "range-handle max-handle";

    const valuesRow = document.createElement("div");
    valuesRow.className = "range-values";
    const minValue = document.createElement("span");
    minValue.textContent = minInput.value;
    const maxValue = document.createElement("span");
    maxValue.textContent = maxInput.value;
    valuesRow.append(minValue, maxValue);

    minInput.addEventListener("input", () => {
      const nextMin = Number(minInput.value);
      const nextMax = Math.max(nextMin, Number(maxInput.value));
      maxInput.value = nextMax;
      state.filterRange.min = nextMin;
      state.filterRange.max = nextMax;
      minValue.textContent = nextMin;
      maxValue.textContent = nextMax;
      scheduleRecalc();
    });

    maxInput.addEventListener("input", () => {
      const nextMax = Number(maxInput.value);
      const nextMin = Math.min(nextMax, Number(minInput.value));
      minInput.value = nextMin;
      state.filterRange.min = nextMin;
      state.filterRange.max = nextMax;
      minValue.textContent = nextMin;
      maxValue.textContent = nextMax;
      scheduleRecalc();
    });

    stack.append(minInput, maxInput);
    group.append(stack, valuesRow);
    ui.filterControl.appendChild(group);
    return;
  }

  ui.filterLabel.textContent = "Suodata ryhmä";
  const groups = getGroups(state.groupBy);
  const boxList = document.createElement("div");
  boxList.className = "checkbox-list";
  if (!state.filterCategories || state.filterCategories.varId !== groupVar.id) {
    state.filterCategories = new Set(groups.map((g) => g.key));
    state.filterCategories.varId = groupVar.id;
  }

  groups.forEach((group) => {
    const label = document.createElement("label");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = group.key;
    checkbox.checked = state.filterCategories.has(group.key);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.filterCategories.add(group.key);
      else state.filterCategories.delete(group.key);
      scheduleRecalc();
    });
    label.appendChild(checkbox);
    label.appendChild(document.createTextNode(group.label));
    boxList.appendChild(label);
  });

  ui.filterControl.appendChild(boxList);
}

function getGroups(groupById) {
  if (groupById === "none") return [];
  const variable = state.variables.find((v) => v.id === groupById);
  if (!variable) return [];
  const values = new Set();
  state.rows.forEach((row) => {
    const value = row.values[variable.id];
    if (value !== "") values.add(value);
  });
  return Array.from(values).map((value, index) => ({
    key: String(value),
    label: `${variable.name}: ${value}`,
    color: getColor(index),
  }));
}

function getRowsByGroup(groupById, groupKey) {
  if (!groupById || groupById === "none") return state.rows;
  const groupVar = state.variables.find((v) => v.id === groupById);
  if (!groupVar || groupVar.type === "other") return state.rows;

  if (groupVar.type === "continuous" && state.filterRange?.varId === groupById) {
    const { min, max } = state.filterRange;
    return state.rows.filter((row) => {
      const value = Number(row.values[groupById]);
      if (Number.isNaN(value)) return false;
      return value >= min && value <= max;
    });
  }

  if (groupVar.type === "categorical" && state.filterCategories?.varId === groupById) {
    const selected = state.filterCategories;
    if (!selected.size) return [];
    return state.rows.filter((row) => selected.has(String(row.values[groupById])));
  }

  return state.rows;
}

function getQuestionVariables() {
  const map = {};
  QUESTION_ROLES.forEach((q) => {
    map[q.key] = state.variables.find((v) => v.role === q.key);
  });
  return map;
}

function calculatePSM() {
  const questionVars = getQuestionVariables();
  if (Object.values(questionVars).some((v) => !v)) {
    setStatus("Määritä kaikki neljä kysymysmuuttujaa Muuttujanäkymässä.");
    return null;
  }

  const rows = getRowsByGroup(state.groupBy, state.filterGroup);
  const responses = getResponsesForRows(rows, questionVars);
  if (!responses) {
    setStatus("Ei kelvollisia rivejä. Täytä numeeriset arvot kaikkiin neljään kysymykseen.");
    return null;
  }

  const metrics = computeMetrics(responses);
  const validCount = responses.tooCheap.length;
  setStatus(`Laskenta valmis. ${validCount} kelvollista riviä.`);
  return { metrics, responses };
}

function getResponsesForRows(rows, questionVars) {
  const filtered = rows.filter((row) => {
    return Object.values(questionVars).every((v) => isNumeric(row.values[v.id]));
  });
  if (!filtered.length) return null;
  const responses = {};
  Object.entries(questionVars).forEach(([key, variable]) => {
    responses[key] = filtered.map((row) => Number(row.values[variable.id]));
  });
  return responses;
}

function computeMetrics(responses) {
  const allValues = Object.values(responses).flat();
  const min = Math.min(...allValues);
  const max = Math.max(...allValues);
  const xValues = buildXValues(min, max, 120);

  const curves = buildCurves(responses, xValues);
  const opp = findRoot(xValues, (x) => curves.tooCheap(x) - curves.tooExpensive(x));
  const pmc = findRoot(xValues, (x) => curves.tooCheap(x) - curves.expensive(x));
  const pme = findRoot(xValues, (x) => curves.bargain(x) - curves.tooExpensive(x));
  const ipp = findRoot(xValues, (x) => curves.bargain(x) - curves.expensive(x));

  return {
    opp,
    pmc,
    pme,
    ipp,
    min,
    max,
  };
}

function buildCurves(responses, xValues) {
  const sorted = {
    tooCheap: [...responses.tooCheap].sort((a, b) => a - b),
    bargain: [...responses.bargain].sort((a, b) => a - b),
    expensive: [...responses.expensive].sort((a, b) => a - b),
    tooExpensive: [...responses.tooExpensive].sort((a, b) => a - b),
  };

  const tooCheap = (x) => percentAt(sorted.tooCheap, x, "decreasing");
  const bargain = (x) => percentAt(sorted.bargain, x, "decreasing");
  const expensive = (x) => percentAt(sorted.expensive, x, "increasing");
  const tooExpensive = (x) => percentAt(sorted.tooExpensive, x, "increasing");

  const F = (x) => tooCheap(x) + bargain(x);
  const G = (x) => expensive(x) + tooExpensive(x);

  return { xValues, tooCheap, bargain, expensive, tooExpensive, F, G };
}

function percentAt(sortedValues, x, mode) {
  const n = sortedValues.length;
  if (n === 0) return 0;
  const leCount = upperBound(sortedValues, x);
  const geCount = n - lowerBound(sortedValues, x);
  if (mode === "increasing") return leCount / n;
  return geCount / n;
}

function buildXValues(min, max, steps) {
  if (min === max) return [min];
  const delta = (max - min) / (steps - 1);
  return Array.from({ length: steps }, (_, i) => min + delta * i);
}

function findRoot(xValues, fn) {
  let best = { x: xValues[0], value: Math.abs(fn(xValues[0])) };
  for (let i = 1; i < xValues.length; i++) {
    const val = Math.abs(fn(xValues[i]));
    if (val < best.value) best = { x: xValues[i], value: val };
  }

  for (let i = 1; i < xValues.length; i++) {
    const a = xValues[i - 1];
    const b = xValues[i];
    const fa = fn(a);
    const fb = fn(b);
    if (fa === 0) return a;
    if (fb === 0) return b;
    if (fa * fb < 0) {
      return bisection(fn, a, b, 40);
    }
  }
  return best.x;
}

function bisection(fn, a, b, iterations) {
  let left = a;
  let right = b;
  for (let i = 0; i < iterations; i++) {
    const mid = (left + right) / 2;
    const fmid = fn(mid);
    if (Math.abs(fmid) < 1e-6) return mid;
    if (fn(left) * fmid < 0) {
      right = mid;
    } else {
      left = mid;
    }
  }
  return (left + right) / 2;
}

function lowerBound(arr, value) {
  let left = 0;
  let right = arr.length;
  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    if (arr[mid] < value) left = mid + 1;
    else right = mid;
  }
  return left;
}

function upperBound(arr, value) {
  let left = 0;
  let right = arr.length;
  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    if (arr[mid] <= value) left = mid + 1;
    else right = mid;
  }
  return left;
}

function updateMetrics(metrics) {
  ui.oppValue.textContent = formatCurrency(metrics.opp);
  ui.pmcValue.textContent = formatCurrency(metrics.pmc);
  ui.pmeValue.textContent = formatCurrency(metrics.pme);
  ui.ippValue.textContent = formatCurrency(metrics.ipp);
  ui.rangeValue.textContent = `${formatCurrency(metrics.pmc)} – ${formatCurrency(metrics.pme)}`;
}

function clearMetrics() {
  ui.oppValue.textContent = "-";
  ui.pmcValue.textContent = "-";
  ui.pmeValue.textContent = "-";
  ui.ippValue.textContent = "-";
  ui.rangeValue.textContent = "-";
}

let recalcTimer = null;
function scheduleRecalc() {
  if (recalcTimer) window.clearTimeout(recalcTimer);
  recalcTimer = window.setTimeout(() => {
    recalcTimer = null;
    runRecalc();
  }, 150);
}

function runRecalc() {
  const result = calculatePSM();
  if (result) {
    state.metrics = result.metrics;
    updateMetrics(result.metrics);
  } else {
    state.metrics = null;
    clearMetrics();
  }
  updateChart();
  updateForecast();
  updatePivot();
}

function updateForecast() {
  const price = Number(ui.forecastPrice.value);
  if (Number.isNaN(price)) {
    ui.forecastTooCheap.textContent = "-";
    ui.forecastBargain.textContent = "-";
    ui.forecastExpensive.textContent = "-";
    ui.forecastTooExpensive.textContent = "-";
    return;
  }

  const questionVars = getQuestionVariables();
  if (Object.values(questionVars).some((v) => !v)) return;
  const rows = getRowsByGroup(state.groupBy, state.filterGroup);
  const responses = getResponsesForRows(rows, questionVars);
  if (!responses) return;

  const curves = buildCurves(responses, [price]);
  ui.forecastTooCheap.textContent = formatPercent(curves.tooCheap(price));
  ui.forecastBargain.textContent = formatPercent(curves.bargain(price));
  ui.forecastExpensive.textContent = formatPercent(curves.expensive(price));
  ui.forecastTooExpensive.textContent = formatPercent(curves.tooExpensive(price));
}

function updatePivot() {
  if (!ui.pivotRow || !ui.pivotCol || !ui.pivotValue) return;
  const rowId = ui.pivotRow.value;
  const colId = ui.pivotCol.value;
  const valueId = ui.pivotValue.value;
  if (!valueId) return;

  if (!rowId && !colId) {
    ui.pivotHead.innerHTML = "";
    ui.pivotBody.innerHTML = "";
    return;
  }

  const rows = state.rows;
  const rowValues = new Set();
  const colValues = new Set();
  const cells = new Map();

  rows.forEach((row) => {
    const r = rowId ? row.values[rowId] : "";
    const c = colId ? row.values[colId] : "";
    const v = Number(row.values[valueId]);
    if ((rowId && r === "") || (colId && c === "") || Number.isNaN(v)) return;
    const rowKey = rowId ? String(r) : "Kaikki";
    const colKey = colId ? String(c) : "Kaikki";
    if (rowId) rowValues.add(rowKey);
    if (colId) colValues.add(colKey);
    const key = `${rowKey}||${colKey}`;
    if (!cells.has(key)) cells.set(key, []);
    cells.get(key).push(v);
  });

  const rowList = rowId ? Array.from(rowValues) : ["Kaikki"];
  const colList = colId ? Array.from(colValues) : ["Kaikki"];

  ui.pivotHead.innerHTML = "";
  ui.pivotBody.innerHTML = "";

  const headRow = document.createElement("tr");
  const corner = document.createElement("th");
  corner.textContent = "";
  headRow.appendChild(corner);
  colList.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  });
  ui.pivotHead.appendChild(headRow);

  rowList.forEach((rowVal) => {
    const tr = document.createElement("tr");
    const th = document.createElement("th");
    th.textContent = rowVal;
    tr.appendChild(th);
    colList.forEach((colVal) => {
      const td = document.createElement("td");
      const values = cells.get(`${rowVal}||${colVal}`) || [];
      if (!values.length) {
        td.textContent = "-";
      } else {
        const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
        td.textContent = `${mean.toFixed(1)} (n=${values.length})`;
      }
      tr.appendChild(td);
    });
    ui.pivotBody.appendChild(tr);
  });
}

function formatPercent(value) {
  if (value == null || Number.isNaN(value)) return "-";
  return `${(value * 100).toFixed(1)}%`;
}

function formatNumber(value) {
  if (value == null || Number.isNaN(value)) return "-";
  return new Intl.NumberFormat("fi-FI", {
    maximumFractionDigits: 2,
  }).format(value);
}

function formatCurrency(value) {
  if (value == null || Number.isNaN(value)) return "-";
  return new Intl.NumberFormat("fi-FI", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 2,
  }).format(value);
}

function buildChart() {
  const mainCtx = document.getElementById("psmChart");
  if (!mainCtx) return;
  Chart.register(psmLabelPlugin);
  state.chartMain = new Chart(mainCtx, {
    type: "line",
    data: { datasets: [] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "nearest", intersect: false },
      scales: {
        x: {
          type: "linear",
          title: { display: true, text: "Hinta" },
        },
        y: {
          title: { display: true, text: "Kertymä %" },
          min: 0,
          max: 1,
          ticks: {
            callback: (value) => `${Math.round(value * 100)}%`,
          },
        },
      },
      plugins: {
        legend: { position: "top" },
        tooltip: {
          callbacks: {
            label: (context) => {
              const label = context.dataset.label || "";
              const y = (context.parsed.y * 100).toFixed(1);
              return `${label}: ${y}% @ ${context.parsed.x.toFixed(2)}`;
            },
          },
        },
      },
    },
  });
  updateChart();
}

function updateChart() {
  if (!state.chartMain) return;
  const questionVars = getQuestionVariables();
  const valid = Object.values(questionVars).every((v) => v);
  if (!valid) return;

  const rows = getRowsByGroup(state.groupBy, state.filterGroup);
  const responses = getResponsesForRows(rows, questionVars);
  if (!responses) return;

  const allValues = Object.values(responses).flat();
  const min = Math.min(...allValues);
  const max = Math.max(...allValues);
  const xValues = buildXValues(min, max, 120);
  const curves = buildCurves(responses, xValues);

  const datasets = [];
  datasets.push(buildLineDataset("Liian halpa", xValues, curves.tooCheap, "#2f9e44"));
  datasets.push(buildLineDataset("Edullinen", xValues, curves.bargain, "#12b886"));
  datasets.push(buildLineDataset("Kallis", xValues, curves.expensive, "#f59f00"));
  datasets.push(buildLineDataset("Liian kallis", xValues, curves.tooExpensive, "#e8590c"));

  if (state.metrics) {
    const points = [
      {
        label: `OPP (${formatCurrency(state.metrics.opp)})`,
        x: state.metrics.opp,
        y: curves.tooCheap(state.metrics.opp),
        color: "#364fc7",
      },
      {
        label: `PMC (${formatCurrency(state.metrics.pmc)})`,
        x: state.metrics.pmc,
        y: curves.tooCheap(state.metrics.pmc),
        color: "#15aabf",
      },
      {
        label: `PME (${formatCurrency(state.metrics.pme)})`,
        x: state.metrics.pme,
        y: curves.bargain(state.metrics.pme),
        color: "#f03e3e",
      },
      {
        label: `IPP (${formatCurrency(state.metrics.ipp)})`,
        x: state.metrics.ipp,
        y: curves.bargain(state.metrics.ipp),
        color: "#845ef7",
      },
    ].filter((point) => !Number.isNaN(point.x) && !Number.isNaN(point.y));

    if (points.length) {
      datasets.push({
        label: "PSM-pisteet",
        data: points.map((point) => ({
          x: point.x,
          y: point.y,
          label: point.label,
          color: point.color,
        })),
        type: "scatter",
        showLine: false,
        pointRadius: 6,
        pointHoverRadius: 7,
        pointBackgroundColor: (ctx) => ctx.raw?.color || "#111",
        pointBorderColor: "#fff",
        pointBorderWidth: 1,
        psmLabels: true,
      });
    }
  }

  state.chartMain.data.datasets = datasets;
  state.chartMain.update();
}

function buildLineDataset(label, xValues, fn, color, width = 2, dash = []) {
  return {
    label,
    data: xValues.map((x) => ({ x, y: fn(x) })),
    borderColor: color,
    borderWidth: width,
    borderDash: dash,
    tension: 0.2,
    pointRadius: 0,
  };
}

function getColor(index) {
  const palette = ["#4c6ef5", "#12b886", "#f76707", "#ae3ec9", "#1098ad", "#e03131"];
  return palette[index % palette.length];
}

function isNumeric(value) {
  return value !== "" && value != null && !Number.isNaN(Number(value));
}

function setStatus(message) {
  ui.statusText.textContent = message;
}

function handleImport(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    let workbook;
    if (file.name.toLowerCase().endsWith(".csv")) {
      workbook = XLSX.read(e.target.result, { type: "string" });
    } else {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array" });
    }
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
    loadFromTable(rows);
  };
  if (file.name.toLowerCase().endsWith(".csv")) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function loadFromTable(table) {
  if (!table.length) return;
  const headers = table[0].map((h) => String(h).trim());
  const body = table.slice(1);
  state.variables = headers.map((name, colIndex) => {
    const roleMatch = QUESTION_ROLES.find((q) => q.label.toLowerCase() === name.toLowerCase());
    const columnValues = body.map((row) => row[colIndex]).filter((v) => v !== "");
    const isNumericColumn = columnValues.every((v) => v !== "" && !Number.isNaN(Number(v)));
    return {
      id: crypto.randomUUID(),
      name: name || "Muuttuja",
      type: roleMatch ? "continuous" : isNumericColumn ? "continuous" : "categorical",
      role: roleMatch ? roleMatch.key : "background",
    };
  });

  state.rows = body.map((row) => {
    const newRow = { id: crypto.randomUUID(), values: {}, _cluster: null };
    state.variables.forEach((variable, index) => {
      newRow.values[variable.id] = row[index] ?? "";
    });
    return newRow;
  });

  renderAll();
  scheduleRecalc();
  setStatus("Tuonti valmis.");
}

function handleJsonSave() {
  const payload = {
    variables: state.variables,
    rows: state.rows,
    metrics: state.metrics,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "psm-project.json";
  link.click();
  URL.revokeObjectURL(url);
  setStatus("Projektin JSON tallennettu.");
}

function handleJsonLoad(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    const payload = JSON.parse(e.target.result);
    state.variables = payload.variables || [];
    state.rows = (payload.rows || []).map((row) => {
      const values = row.values || {};
      state.variables.forEach((variable) => {
        if (!(variable.id in values)) values[variable.id] = "";
      });
      return { ...row, values };
    });
    state.metrics = payload.metrics || null;
    renderAll();
    updateMetrics(state.metrics || {});
    scheduleRecalc();
    setStatus("Projektin JSON ladattu.");
  };
  reader.readAsText(file);
}

function exportExcel() {
  const wb = XLSX.utils.book_new();
  const getVarById = (id) => state.variables.find((v) => v.id === id);
  const getVarName = (id) => getVarById(id)?.name || "";

  const variableSheet = XLSX.utils.json_to_sheet(
    state.variables.map((v) => ({
      name: v.name,
      type: v.type,
      role: v.role,
    }))
  );
  XLSX.utils.book_append_sheet(wb, variableSheet, "Variables");

  const header = state.variables.map((v) => v.name);
  const dataRows = state.rows.map((row) => state.variables.map((v) => row.values[v.id]));
  const dataSheet = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
  XLSX.utils.book_append_sheet(wb, dataSheet, "Data");

  const filterRows = [];
  const groupVar = state.groupBy !== "none" ? getVarById(state.groupBy) : null;
  filterRows.push(["Ryhmittele muuttujalla", groupVar ? groupVar.name : "Ei ryhmittelyä"]);
  if (!groupVar) {
    filterRows.push(["Suodatus", "Kaikki"]);
  } else if (groupVar.type === "continuous" && state.filterRange?.varId === groupVar.id) {
    filterRows.push(["Suodatus", "Väli"]);
    filterRows.push(["Min", state.filterRange.min]);
    filterRows.push(["Max", state.filterRange.max]);
  } else if (groupVar.type === "categorical" && state.filterCategories?.varId === groupVar.id) {
    const selected = Array.from(state.filterCategories).join(", ");
    filterRows.push(["Suodatus", "Luokat"]);
    filterRows.push(["Valitut", selected || "-"]);
  } else {
    filterRows.push(["Suodatus", "Kaikki"]);
  }
  const filterSheet = XLSX.utils.aoa_to_sheet(filterRows);
  XLSX.utils.book_append_sheet(wb, filterSheet, "Filter");

  if (state.metrics) {
    const metricsSheet = XLSX.utils.json_to_sheet([
      {
        opp: state.metrics.opp,
        pmc: state.metrics.pmc,
        pme: state.metrics.pme,
        ipp: state.metrics.ipp,
        rangeLow: state.metrics.pmc,
        rangeHigh: state.metrics.pme,
      },
    ]);
    XLSX.utils.book_append_sheet(wb, metricsSheet, "Metrics");
  }

  const forecastPrice = Number(ui.forecastPrice?.value);
  const questionVars = getQuestionVariables();
  if (!Number.isNaN(forecastPrice) && Object.values(questionVars).every((v) => v)) {
    const rows = getRowsByGroup(state.groupBy, state.filterGroup);
    const responses = getResponsesForRows(rows, questionVars);
    if (responses) {
      const curves = buildCurves(responses, [forecastPrice]);
      const forecastSheet = XLSX.utils.json_to_sheet([
        {
          price: forecastPrice,
          tooCheap: curves.tooCheap(forecastPrice),
          bargain: curves.bargain(forecastPrice),
          expensive: curves.expensive(forecastPrice),
          tooExpensive: curves.tooExpensive(forecastPrice),
        },
      ]);
      XLSX.utils.book_append_sheet(wb, forecastSheet, "Forecast");
    }
  }

  if (state.optimizer?.results?.length) {
    const selectedVars = Array.from(state.optimizer.selectedVarIds || [])
      .map((id) => getVarName(id))
      .filter(Boolean)
      .join(", ");
    const optimizerRows = [
      ["Valitut taustamuuttujat", selectedVars || "-"],
      [],
      ["Segmentti", "OPP", "n", "Score"],
      ...state.optimizer.results.map((item) => [item.description, item.opp, item.size, item.score]),
    ];
    const optimizerSheet = XLSX.utils.aoa_to_sheet(optimizerRows);
    XLSX.utils.book_append_sheet(wb, optimizerSheet, "Optimizer");
  }

  XLSX.writeFile(wb, "psm-export.xlsx");
  setStatus("Excel viety.");
}

function exportPpt() {
  if (!state.chartMain) return;
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  const img = state.chartMain.toBase64Image();
  slide.addImage({ data: img, x: 0.5, y: 0.8, w: 9, h: 4.8 });

  const metricsText = state.metrics
    ? `OPP: ${formatCurrency(state.metrics.opp)}\nPMC: ${formatCurrency(state.metrics.pmc)}\nPME: ${formatCurrency(state.metrics.pme)}\nIPP: ${formatCurrency(state.metrics.ipp)}`
    : "Metrics not calculated.";
  slide.addText(metricsText, { x: 0.5, y: 5.8, w: 9, h: 1.2, fontSize: 14 });

  pptx.writeFile({ fileName: "psm-chart.pptx" });
  setStatus("PowerPoint viety.");
}

function runKMeans() {
  const selected = getClusterSelectedIds();
  if (!selected.length) {
    setClusterSelectMessage("Valitse vähintään yksi muuttuja.");
    setStatus("Valitse numeeriset taustamuuttujat klusterointiin.");
    return false;
  }
  const k = Number(ui.clusterK.value);
  if (!k || k < 2) {
    setStatus("Klusterien määrä (k) vähintään 2.");
    return false;
  }

  const points = state.rows
    .map((row) => {
      const values = selected.map((id) => Number(row.values[id]));
      return { row, values };
    })
    .filter((item) => item.values.every((v) => !Number.isNaN(v)));

  if (points.length < k) {
    setStatus("Ei tarpeeksi rivejä valitulle k-arvolle.");
    return false;
  }

  const centroids = points.slice(0, k).map((p) => [...p.values]);
  for (let iter = 0; iter < 15; iter++) {
    points.forEach((p) => {
      p.cluster = nearestCentroid(p.values, centroids);
    });
    const sums = Array.from({ length: k }, () => new Array(selected.length).fill(0));
    const counts = new Array(k).fill(0);
    points.forEach((p) => {
      counts[p.cluster] += 1;
      p.values.forEach((v, i) => {
        sums[p.cluster][i] += v;
      });
    });
    centroids.forEach((c, idx) => {
      if (!counts[idx]) return;
      c.forEach((_, i) => {
        c[i] = sums[idx][i] / counts[idx];
      });
    });
  }

  points.forEach((p) => {
    p.row._cluster = p.cluster;
  });

  state.groupBy = "cluster";
  const useMeanLabels = ui.clusterMeanLabels?.checked && selected.length === 1;
  const meanLabels = useMeanLabels ? centroids.map((c) => formatNumber(c[0])) : null;
  updateClusterGrouping(k, meanLabels);
  renderAll();
  scheduleRecalc();
  setStatus(`K-means valmis. ${k} klusteria.`);
  return true;
}

function runKMeansAndTabulate() {
  const selected = getClusterSelectedIds();
  if (selected.length !== 1) {
    setClusterSelectMessage("Valitse täsmälleen yksi muuttuja.");
    return;
  }
  const ok = runKMeans();
  if (!ok) return;

  ui.pivotRow.value = "cluster";
  ui.pivotCol.value = "";
  ui.pivotValue.value = selected[0];
  updatePivot();
  switchTab("analysis");
}

function setClusterSelectMessage(message) {
  if (!ui.clusterSelectMessage) return;
  ui.clusterSelectMessage.textContent = message;
}

function getClusterSelectedIds() {
  const checkboxes = ui.clusterSelect.querySelectorAll("input[type=checkbox]");
  return Array.from(checkboxes)
    .filter((cb) => cb.checked)
    .map((cb) => cb.value);
}

function updateClusterGrouping(k, meanLabels = null) {
  const clusterName = (ui.clusterName?.value || "Klusteri").trim() || "Klusteri";
  if (!state.variables.find((v) => v.id === "cluster")) {
    state.variables.push({
      id: "cluster",
      name: clusterName,
      type: "categorical",
      role: "background",
    });
  } else {
    const clusterVar = state.variables.find((v) => v.id === "cluster");
    if (clusterVar) {
      clusterVar.name = clusterName;
      clusterVar.type = "categorical";
    }
  }
  state.rows.forEach((row) => {
    if (row._cluster == null || row._cluster === "") {
      row.values["cluster"] = "";
      return;
    }
    const label = meanLabels ? meanLabels[row._cluster] : `Klusteri ${row._cluster + 1}`;
    row.values["cluster"] = label;
  });
}

function updateClusterName() {
  const clusterVar = state.variables.find((v) => v.id === "cluster");
  if (!clusterVar) return;
  const name = (ui.clusterName.value || "Klusteri").trim() || "Klusteri";
  clusterVar.name = name;
  renderAll();
}

function nearestCentroid(point, centroids) {
  let best = 0;
  let bestDist = Infinity;
  centroids.forEach((centroid, idx) => {
    const dist = Math.sqrt(
      centroid.reduce((sum, value, i) => sum + (value - point[i]) ** 2, 0)
    );
    if (dist < bestDist) {
      bestDist = dist;
      best = idx;
    }
  });
  return best;
}

init();
