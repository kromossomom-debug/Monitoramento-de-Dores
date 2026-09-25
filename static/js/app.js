// State
let allRecords = [];
let filteredRecords = [];
let appParameters = {};
let appStatus = {};
let charts = {};
let activeQuickFilter = '';
let currentRecordInDetail = null;

// Initialize when DOM ready
document.addEventListener("DOMContentLoaded", async () => {
    lucide.createIcons();
    setDefaultFormDates();
    await loadParameters();
    await loadStatus();
    await refreshAllData();

    // Close modals on ESC
    document.addEventListener("keydown", (e) => {
        if (e.key === "Escape") {
            closeNewModal();
            closeDetailsModal();
            closeEditModal();
        }
    });
});

// Formatters
function formatCurrency(val) {
    if (val === null || val === undefined || isNaN(val)) return "R$ 0,00";
    return new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" }).format(val);
}

function formatDateBR(val) {
    if (!val) return "-";
    if (typeof val === "string" && val.includes("T")) {
        val = val.split("T")[0];
    }
    if (typeof val === "string" && val.length === 10 && val.includes("-")) {
        const [y, m, d] = val.split("-");
        return `${d}/${m}/${y}`;
    }
    return val;
}

function setDefaultFormDates() {
    const today = new Date().toISOString().split("T")[0];
    const dateInputs = [
        "cad_data_ocorrencia", "cad_data_registro",
        "modal_data_ocorrencia", "modal_data_registro"
    ];
    dateInputs.forEach(id => {
        const el = document.getElementById(id);
        if (el && !el.value) el.value = today;
    });
}

// Toast Notifications
function showToast(message, type = "info") {
    const container = document.getElementById("toast-container");
    if (!container) return;

    const toast = document.createElement("div");
    toast.className = `toast px-4 py-3 rounded-xl shadow-lg text-xs font-semibold flex items-center gap-2 text-white max-w-sm border ${
        type === "success" ? "bg-emerald-600 border-emerald-700" :
        type === "error" ? "bg-red-600 border-red-700" :
        type === "warning" ? "bg-amber-600 border-amber-700" :
        "bg-blue-600 border-blue-700"
    }`;

    const iconName = type === "success" ? "check-circle" : type === "error" ? "alert-circle" : "info";
    toast.innerHTML = `<i data-lucide="${iconName}" class="w-4 h-4"></i> <span>${message}</span>`;
    container.appendChild(toast);
    lucide.createIcons();

    setTimeout(() => {
        toast.style.opacity = "0";
        toast.style.transform = "translateY(20px)";
        toast.style.transition = "all 0.3s ease-out";
        setTimeout(() => toast.remove(), 300);
    }, 4500);
}

// Tab Switching
function switchTab(tabName) {
    const tabs = ["dashboard", "consulta", "cadastro", "status"];
    tabs.forEach(t => {
        const btn = document.getElementById(`tab-btn-${t}`);
        const content = document.getElementById(`tab-content-${t}`);
        if (t === tabName) {
            btn?.classList.add("active");
            content?.classList.remove("hidden");
        } else {
            btn?.classList.remove("active");
            content?.classList.add("hidden");
        }
    });

    if (tabName === "dashboard") {
        setTimeout(renderCharts, 100);
    }
    lucide.createIcons();
}

// Parameters Loading
async function loadParameters() {
    try {
        const res = await fetch("/api/parameters");
        const json = await res.json();
        if (json.success) {
            appParameters = json.data;
            populateDropdowns();
        }
    } catch (e) {
        console.error("Error loading parameters:", e);
    }
}

function populateDropdowns() {
    const p = appParameters;
    if (!p) return;

    // Filter dropdowns
    populateSelect("filter-status", p.status || [], "Status: Todos");
    populateSelect("filter-prioridade", p.prioridade || [], "Prioridade: Todas");
    populateSelect("filter-filial", p.filial || [], "Filial: Todas");

    // Modal & Form dropdowns
    ["cad", "modal"].forEach(prefix => {
        populateSelect(`${prefix}_filial`, p.filial || []);
        populateSelect(`${prefix}_setor_responsavel`, p.setor || []);
        populateSelect(`${prefix}_setor_impactado`, [""].concat(p.setor || []), "-- Selecione ou Deixe em Branco --");
        populateSelect(`${prefix}_categoria_dor`, p.categoria || []);
        populateSelect(`${prefix}_prioridade`, p.prioridade || []);
        populateSelect(`${prefix}_status`, p.status || []);
    });

    // Edit Modal
    populateSelect("edit_status", p.status || []);
    populateSelect("edit_prioridade", p.prioridade || []);

    // Set preview
    updateSlaBadgePreview("cad");
    updateSlaBadgePreview("modal");
}

function populateSelect(selectId, items, defaultPlaceholder = null) {
    const sel = document.getElementById(selectId);
    if (!sel) return;
    sel.innerHTML = "";
    if (defaultPlaceholder !== null) {
        const opt = document.createElement("option");
        opt.value = "";
        opt.textContent = defaultPlaceholder;
        sel.appendChild(opt);
    }
    items.forEach(it => {
        if (!it) return;
        const opt = document.createElement("option");
        opt.value = it;
        opt.textContent = it;
        sel.appendChild(opt);
    });
}

function updateSlaBadgePreview(prefix) {
    const prioEl = document.getElementById(`${prefix}_prioridade`);
    const previewEl = document.getElementById(`${prefix}_sla_preview`);
    if (!prioEl || !previewEl) return;

    const val = prioEl.value || "Média";
    const slaDays = (appParameters.sla_map && appParameters.sla_map[val]) || (val === "Crítica" ? 2 : val === "Alta" ? 5 : val === "Média" ? 10 : 15);
    previewEl.innerHTML = `<span>SLA: <strong>${slaDays} dias</strong></span> <span class="text-xs font-normal text-gray-500">Calculado via Excel</span>`;
}

// System Status
async function loadStatus() {
    try {
        const res = await fetch("/api/status");
        appStatus = await res.json();

        const fileNameEl = document.getElementById("header-file-name");
        if (fileNameEl && appStatus.file_path) {
            const parts = appStatus.file_path.split(/[\\/]/);
            fileNameEl.textContent = parts[parts.length - 1];
        }

        const pathEl = document.getElementById("status-filepath");
        if (pathEl) pathEl.textContent = appStatus.file_path || "-";

        const sizeEl = document.getElementById("status-filesize");
        if (sizeEl) sizeEl.textContent = `${appStatus.file_size_kb || 0} KB`;

        const mtimeEl = document.getElementById("status-filemtime");
        if (mtimeEl) mtimeEl.textContent = appStatus.last_modified || "-";

        const countEl = document.getElementById("status-backupcount");
        if (countEl) countEl.textContent = `${appStatus.backup_count || 0} versões`;
    } catch (e) {
        console.error("Error loading status:", e);
    }
}

// Main Refresh
async function refreshAllData() {
    const refreshBtn = document.getElementById("btn-refresh");
    if (refreshBtn) refreshBtn.classList.add("animate-spin");

    try {
        await Promise.all([loadRecords(), loadKpis(), loadStatus()]);
        showToast("Dados atualizados da planilha Excel com sucesso!", "success");
    } catch (e) {
        showToast("Erro ao sincronizar com a planilha: " + e.message, "error");
    } finally {
        if (refreshBtn) refreshBtn.classList.remove("animate-spin");
    }
}

// Load Records
async function loadRecords() {
    const res = await fetch("/api/records");
    const json = await res.json();
    if (json.success) {
        allRecords = json.data || [];
        applyFilters();
        renderRecentDashboardTable();
    }
}

// Load KPIs
async function loadKpis() {
    const res = await fetch("/api/kpis");
    const json = await res.json();
    if (json.success) {
        updateKpiDisplay(json.data);
    }
}

function updateKpiDisplay(k) {
    document.getElementById("kpi-total").textContent = k.total_ocorrencias || 0;
    document.getElementById("kpi-abertas").textContent = k.ocorrencias_abertas || 0;
    document.getElementById("kpi-concluidas").textContent = k.ocorrencias_concluidas || 0;
    document.getElementById("kpi-taxa-resolucao").textContent = `(${k.taxa_resolucao || 0}%)`;
    document.getElementById("kpi-sla-vencido").textContent = k.sla_vencido || 0;
    document.getElementById("kpi-sla-avencer").textContent = k.sla_a_vencer || 0;
    document.getElementById("kpi-valor-total").textContent = formatCurrency(k.valor_total_notas || 0);
    document.getElementById("kpi-tempo-aberto").textContent = `${k.tempo_medio_aberto || 0} d`;
    document.getElementById("kpi-tempo-resolucao").textContent = `${k.tempo_medio_resolucao || 0} d`;

    // Banner logic
    const banner = document.getElementById("urgent-banner");
    const bannerText = document.getElementById("urgent-banner-text");
    if (k.sla_vencido > 0) {
        banner?.classList.remove("hidden");
        if (bannerText) {
            bannerText.textContent = `${k.sla_vencido} ocorrência(s) ultrapassaram o prazo de atendimento do SLA!`;
        }
    } else {
        banner?.classList.add("hidden");
    }

    renderChartsWithData(k);
}

// Charts
function renderCharts() {
    if (appParameters) {
        // Redraw current
    }
}

function renderChartsWithData(k) {
    // Chart 1: Status Donut
    const statusCtx = document.getElementById("chartStatus")?.getContext("2d");
    if (statusCtx) {
        if (charts.status) charts.status.destroy();
        const labels = Object.keys(k.by_status || {});
        const data = Object.values(k.by_status || {});
        charts.status = new Chart(statusCtx, {
            type: "doughnut",
            data: {
                labels: labels.length ? labels : ["Sem dados"],
                datasets: [{
                    data: data.length ? data : [1],
                    backgroundColor: [
                        "#3b82f6", "#f59e0b", "#8b5cf6", "#f97316", "#10b981", "#64748b"
                    ],
                    borderWidth: 2,
                    borderColor: "#ffffff"
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } }
                },
                cutout: "68%"
            }
        });
    }

    // Chart 2: SLA Bar / Donut
    const slaCtx = document.getElementById("chartSla")?.getContext("2d");
    if (slaCtx) {
        if (charts.sla) charts.sla.destroy();
        const slaLabels = ["No prazo", "A vencer", "Vencido", "Concluído"];
        const slaData = slaLabels.map(l => (k.by_situacao_sla && k.by_situacao_sla[l]) || 0);
        charts.sla = new Chart(slaCtx, {
            type: "doughnut",
            data: {
                labels: slaLabels,
                datasets: [{
                    data: slaData,
                    backgroundColor: ["#10b981", "#f59e0b", "#ef4444", "#94a3b8"],
                    borderWidth: 2,
                    borderColor: "#ffffff"
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: "bottom", labels: { boxWidth: 12, font: { size: 11 } } }
                },
                cutout: "68%"
            }
        });
    }

    // Chart 3: Prioridade
    const prioCtx = document.getElementById("chartPrioridade")?.getContext("2d");
    if (prioCtx) {
        if (charts.prioridade) charts.prioridade.destroy();
        const prioLabels = ["Crítica", "Alta", "Média", "Baixa"];
        const prioData = prioLabels.map(l => (k.by_prioridade && k.by_prioridade[l]) || 0);
        charts.prioridade = new Chart(prioCtx, {
            type: "bar",
            data: {
                labels: prioLabels,
                datasets: [{
                    label: "Ocorrências",
                    data: prioData,
                    backgroundColor: ["#e11d48", "#ea580c", "#eab308", "#16a34a"],
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } },
                    x: { grid: { display: false } }
                }
            }
        });
    }

    // Chart 4: Categorias (Horizontal Bar)
    const catCtx = document.getElementById("chartCategoria")?.getContext("2d");
    if (catCtx) {
        if (charts.categoria) charts.categoria.destroy();
        const catLabels = Object.keys(k.by_categoria || {});
        const catData = Object.values(k.by_categoria || {});
        charts.categoria = new Chart(catCtx, {
            type: "bar",
            data: {
                labels: catLabels.length ? catLabels : ["Sem dados"],
                datasets: [{
                    label: "Volume",
                    data: catData.length ? catData : [0],
                    backgroundColor: "#6366f1",
                    borderRadius: 6
                }]
            },
            options: {
                indexAxis: "y",
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    x: { beginAtZero: true, ticks: { precision: 0 } },
                    y: { grid: { display: false } }
                }
            }
        });
    }

    // Chart 5: Setor Responsável
    const setorCtx = document.getElementById("chartSetor")?.getContext("2d");
    if (setorCtx) {
        if (charts.setor) charts.setor.destroy();
        const sLabels = Object.keys(k.by_setor_responsavel || {});
        const sData = Object.values(k.by_setor_responsavel || {});
        charts.setor = new Chart(setorCtx, {
            type: "bar",
            data: {
                labels: sLabels.length ? sLabels : ["Sem dados"],
                datasets: [{
                    label: "Pendências Atribuídas",
                    data: sData.length ? sData : [0],
                    backgroundColor: "#0ea5e9",
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } },
                    x: { grid: { display: false } }
                }
            }
        });
    }

    // Chart 6: Filiais
    const filialCtx = document.getElementById("chartFilial")?.getContext("2d");
    if (filialCtx) {
        if (charts.filial) charts.filial.destroy();
        const fLabels = Object.keys(k.by_filial || {});
        const fData = Object.values(k.by_filial || {});
        charts.filial = new Chart(filialCtx, {
            type: "bar",
            data: {
                labels: fLabels.length ? fLabels : ["Sem dados"],
                datasets: [{
                    label: "Ocorrências",
                    data: fData.length ? fData : [0],
                    backgroundColor: "#14b8a6",
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 } },
                    x: { grid: { display: false } }
                }
            }
        });
    }
}

// Filters & Table Rendering
function applyFilters() {
    const search = (document.getElementById("filter-search")?.value || "").toLowerCase().trim();
    const status = document.getElementById("filter-status")?.value;
    const prio = document.getElementById("filter-prioridade")?.value;
    const sla = document.getElementById("filter-sla")?.value;
    const filial = document.getElementById("filter-filial")?.value;

    filteredRecords = allRecords.filter(r => {
        if (status && r.status !== status) return false;
        if (prio && r.prioridade !== prio) return false;
        if (sla && r.situacao_sla !== sla) return false;
        if (filial && r.filial !== filial) return false;

        // Quick filter
        if (activeQuickFilter === "aberto" && (r.status === "Concluído" || r.status === "Cancelado")) return false;
        if (activeQuickFilter === "vencido" && r.situacao_sla !== "Vencido") return false;
        if (activeQuickFilter === "avencer" && r.situacao_sla !== "A vencer") return false;
        if (activeQuickFilter === "critica" && r.prioridade !== "Crítica") return false;
        if (activeQuickFilter === "concluido" && r.status !== "Concluído") return false;

        if (search) {
            const str = [
                r.id, r.responsavel, r.setor_responsavel, r.setor_impactado,
                r.filial, r.categoria_dor, r.descricao_problema, r.causa_raiz,
                r.impacto_negocio, r.nota_fiscal, r.numero_chamado, r.responsavel_solucao
            ].join(" ").toLowerCase();
            if (!str.includes(search)) return false;
        }

        return true;
    });

    renderTable();
}

function setQuickFilter(type) {
    activeQuickFilter = type;
    document.querySelectorAll(".quick-chip").forEach(el => el.classList.remove("active", "ring-2", "ring-blue-500"));
    const activeBtn = Array.from(document.querySelectorAll(".quick-chip")).find(el => {
        if (!type && el.textContent.includes("Todos")) return true;
        if (type === "aberto" && el.textContent.includes("Abertos")) return true;
        if (type === "vencido" && el.textContent.includes("Vencido")) return true;
        if (type === "avencer" && el.textContent.includes("A Vencer")) return true;
        if (type === "critica" && el.textContent.includes("Crítica")) return true;
        if (type === "concluido" && el.textContent.includes("Concluídos")) return true;
        return false;
    });
    if (activeBtn) activeBtn.classList.add("active", "ring-2", "ring-blue-500");
    applyFilters();
}

function filterByUrgency(val) {
    switchTab("consulta");
    const slaSelect = document.getElementById("filter-sla");
    if (slaSelect) slaSelect.value = val;
    setQuickFilter(val.toLowerCase());
}

function resetFilters() {
    document.getElementById("filter-search").value = "";
    document.getElementById("filter-status").value = "";
    document.getElementById("filter-prioridade").value = "";
    document.getElementById("filter-sla").value = "";
    document.getElementById("filter-filial").value = "";
    activeQuickFilter = "";
    applyFilters();
}

function renderTable() {
    const tbody = document.getElementById("records-table-body");
    const counter = document.getElementById("records-counter-text");
    const badgeCount = document.getElementById("badge-total-records");

    if (badgeCount) badgeCount.textContent = allRecords.length;
    if (counter) counter.textContent = `Exibindo ${filteredRecords.length} de ${allRecords.length} ocorrências`;

    if (!tbody) return;

    if (filteredRecords.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="px-6 py-12 text-center text-gray-500">
                    <i data-lucide="inbox" class="w-10 h-10 mx-auto text-gray-300 mb-2"></i>
                    <p class="font-medium text-sm text-gray-600">Nenhuma pendência encontrada com os filtros selecionados.</p>
                    <p class="text-xs text-gray-400 mt-1">Tente ajustar a busca ou limpe os filtros para ver todos os registros.</p>
                </td>
            </tr>
        `;
        lucide.createIcons();
        return;
    }

    tbody.innerHTML = filteredRecords.map(r => {
        const statusBadge = getStatusBadge(r.status);
        const slaBadge = getSlaBadge(r);
        const prioBadge = getPriorityBadge(r.prioridade);

        return `
            <tr class="hover:bg-gray-50 transition border-b border-gray-100">
                <td class="px-3.5 py-3 font-bold text-gray-800 text-center">#${r.id}</td>
                <td class="px-3.5 py-3 whitespace-nowrap">
                    <div class="font-semibold text-gray-900">${r.data_registro_br || "-"}</div>
                    <div class="text-[11px] text-gray-400">Ocorr: ${r.data_ocorrencia_br || "-"}</div>
                </td>
                <td class="px-3.5 py-3">
                    <div class="font-medium text-gray-900">${r.filial || "-"}</div>
                    <div class="text-[11px] text-blue-600 font-semibold">${r.categoria_dor || "-"}</div>
                </td>
                <td class="px-3.5 py-3">
                    <div class="font-medium text-gray-800">${r.setor_responsavel || "-"}</div>
                    <div class="text-[11px] text-gray-500">Resp: ${r.responsavel || "-"}</div>
                </td>
                <td class="px-3.5 py-3 max-w-xs">
                    <p class="font-medium text-gray-900 truncate" title="${escapeHtml(r.descricao_problema || '')}">${escapeHtml(r.descricao_problema || "-")}</p>
                    ${r.nota_fiscal ? `<span class="inline-flex items-center text-[10px] text-gray-500 bg-gray-100 px-1.5 py-0.5 rounded mt-0.5">NF ${r.nota_fiscal}${r.serie ? ` (Série ${r.serie})` : ''}</span>` : ''}
                </td>
                <td class="px-3.5 py-3 whitespace-nowrap">${prioBadge}</td>
                <td class="px-3.5 py-3 whitespace-nowrap">${statusBadge}</td>
                <td class="px-3.5 py-3 whitespace-nowrap">${slaBadge}</td>
                <td class="px-3.5 py-3 text-right font-medium text-gray-800 whitespace-nowrap">${formatCurrency(r.valor_notas)}</td>
                <td class="px-3.5 py-3 text-center whitespace-nowrap">
                    <div class="flex items-center justify-center space-x-1.5">
                        <button onclick="openDetailsModal(${r.id})" class="p-1.5 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded-lg transition" title="Ver detalhes completos">
                            <i data-lucide="eye" class="w-4 h-4"></i>
                        </button>
                        <button onclick="openEditModal(${r.id})" class="p-1.5 text-emerald-600 hover:text-emerald-800 hover:bg-emerald-50 rounded-lg transition" title="Atualizar / Tratar">
                            <i data-lucide="edit-3" class="w-4 h-4"></i>
                        </button>
                    </div>
                </td>
            </tr>
        `;
    }).join("");

    lucide.createIcons();
}

function renderRecentDashboardTable() {
    const tbody = document.getElementById("dashboard-recent-table");
    if (!tbody) return;

    const recent = allRecords.slice(-5).reverse();
    if (recent.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="p-4 text-center text-gray-400">Nenhum registro encontrado.</td></tr>`;
        return;
    }

    tbody.innerHTML = recent.map(r => `
        <tr class="hover:bg-gray-50 transition">
            <td class="px-4 py-2.5 font-bold text-gray-800">#${r.id}</td>
            <td class="px-4 py-2.5 text-gray-600">${r.data_registro_br}</td>
            <td class="px-4 py-2.5 font-medium text-gray-900">${r.filial} <span class="text-gray-400">/</span> <span class="text-blue-600 font-semibold">${r.categoria_dor}</span></td>
            <td class="px-4 py-2.5 text-gray-700">${r.setor_responsavel}</td>
            <td class="px-4 py-2.5 truncate max-w-xs text-gray-800" title="${escapeHtml(r.descricao_problema)}">${escapeHtml(r.descricao_problema)}</td>
            <td class="px-4 py-2.5">${getPriorityBadge(r.prioridade)}</td>
            <td class="px-4 py-2.5">${getStatusBadge(r.status)}</td>
            <td class="px-4 py-2.5">${getSlaBadge(r)}</td>
            <td class="px-4 py-2.5 text-right">
                <button onclick="openDetailsModal(${r.id})" class="text-xs text-blue-600 hover:underline font-semibold">Detalhes</button>
            </td>
        </tr>
    `).join("");

    lucide.createIcons();
}

// Badges Generator
function getStatusBadge(status) {
    if (!status) return `<span class="badge-status badge-aberto">Aberto</span>`;
    const s = status.toLowerCase();
    if (s.includes("conclu")) return `<span class="badge-status badge-concluido"><i data-lucide="check" class="w-3 h-3 mr-1"></i> Concluído</span>`;
    if (s.includes("cancel")) return `<span class="badge-status badge-cancelado">Cancelado</span>`;
    if (s.includes("terceiro")) return `<span class="badge-status badge-terceiros"><i data-lucide="clock" class="w-3 h-3 mr-1"></i> Aguardando</span>`;
    if (s.includes("tratat")) return `<span class="badge-status badge-tratativa">Em Tratativa</span>`;
    if (s.includes("análise") || s.includes("analise")) return `<span class="badge-status badge-analise">Em Análise</span>`;
    return `<span class="badge-status badge-aberto">${status}</span>`;
}

function getPriorityBadge(prio) {
    if (!prio) return `<span class="badge-status badge-prioridade-media">Média</span>`;
    const p = prio.toLowerCase();
    if (p.includes("crít") || p.includes("crit")) return `<span class="badge-status badge-prioridade-critica">Crítica (2d)</span>`;
    if (p.includes("alta")) return `<span class="badge-status badge-prioridade-alta">Alta (5d)</span>`;
    if (p.includes("baixa")) return `<span class="badge-status badge-prioridade-baixa">Baixa (15d)</span>`;
    return `<span class="badge-status badge-prioridade-media">Média (10d)</span>`;
}

function getSlaBadge(r) {
    const sit = r.situacao_sla;
    if (r.status === "Concluído") {
        return `<span class="badge-status badge-sla-concluido"><i data-lucide="check-check" class="w-3 h-3 mr-1"></i> Resolvido</span>`;
    }
    if (sit === "Vencido") {
        const days = r.dias_para_vencimento ? Math.abs(r.dias_para_vencimento) : "";
        return `<span class="badge-status badge-sla-vencido" title="Limite: ${r.data_limite_sla_br}"><i data-lucide="alert-octagon" class="w-3 h-3 mr-1"></i> Vencido (${days}d)</span>`;
    }
    if (sit === "A vencer") {
        const days = r.dias_para_vencimento ? r.dias_para_vencimento : "0";
        return `<span class="badge-status badge-sla-avencer" title="Limite: ${r.data_limite_sla_br}"><i data-lucide="clock" class="w-3 h-3 mr-1"></i> A vencer (${days}d)</span>`;
    }
    if (sit === "No prazo") {
        return `<span class="badge-status badge-sla-noprazo" title="Limite: ${r.data_limite_sla_br}"><i data-lucide="check" class="w-3 h-3 mr-1"></i> No prazo</span>`;
    }
    return `<span class="badge-status badge-sla-concluido">-</span>`;
}

// Modals
function openNewModal() {
    setDefaultFormDates();
    document.getElementById("modal-nova-pendencia")?.classList.remove("hidden");
    lucide.createIcons();
}

function closeNewModal() {
    document.getElementById("modal-nova-pendencia")?.classList.add("hidden");
}

function openDetailsModal(id) {
    const r = allRecords.find(item => item.id === id);
    if (!r) return;
    currentRecordInDetail = r;

    document.getElementById("detalhe-title").textContent = `Pendência #${r.id} - ${r.categoria_dor || 'Dor Operacional'}`;
    document.getElementById("detalhe-subtitle").textContent = `Registrado por ${r.responsavel || '-'} em ${r.data_registro_br || '-'}`;

    const content = document.getElementById("detalhes-content");
    content.innerHTML = `
        <!-- Top Status Bar in Details -->
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3 bg-gray-50 p-3.5 rounded-xl border border-gray-100">
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Status Atual</span>
                <div class="mt-0.5">${getStatusBadge(r.status)}</div>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Prioridade & SLA</span>
                <div class="mt-0.5">${getPriorityBadge(r.prioridade)}</div>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Situação do Prazo</span>
                <div class="mt-0.5">${getSlaBadge(r)}</div>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Dias em Aberto</span>
                <span class="text-sm font-bold text-gray-800 mt-0.5 block">${r.dias_aberto} dia(s)</span>
            </div>
        </div>

        <!-- Problem Description -->
        <div class="space-y-1">
            <span class="text-gray-500 font-bold uppercase text-[11px]">Descrição Completa do Problema</span>
            <div class="p-3 bg-blue-50/50 rounded-xl border border-blue-100 text-gray-900 leading-relaxed font-medium">
                ${escapeHtml(r.descricao_problema || "-")}
            </div>
        </div>

        <!-- Root Cause & Impact -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div class="p-3 bg-gray-50 rounded-xl border border-gray-100">
                <span class="text-gray-400 font-bold uppercase text-[10px] block">Causa Raiz</span>
                <span class="text-gray-800 font-medium">${escapeHtml(r.causa_raiz || "Não informada")}</span>
            </div>
            <div class="p-3 bg-gray-50 rounded-xl border border-gray-100">
                <span class="text-gray-400 font-bold uppercase text-[10px] block">Impacto no Negócio</span>
                <span class="text-gray-800 font-medium">${escapeHtml(r.impacto_negocio || "Não informado")}</span>
            </div>
        </div>

        <!-- Action Plan -->
        <div class="p-3.5 bg-emerald-50/40 rounded-xl border border-emerald-100 space-y-1">
            <div class="flex items-center justify-between">
                <span class="text-emerald-800 font-bold uppercase text-[11px] flex items-center gap-1">
                    <i data-lucide="compass" class="w-3.5 h-3.5"></i> Plano de Ação / Encaminhamento
                </span>
                <span class="text-[11px] text-gray-500 font-medium">Prazo: <strong>${r.prazo_br || "-"}</strong></span>
            </div>
            <p class="text-gray-800 text-xs leading-relaxed mt-1">${escapeHtml(r.plano_acao || "Nenhum plano de ação registrado até o momento.")}</p>
            <div class="flex justify-between items-center text-[10px] text-gray-500 pt-2 border-t border-emerald-100 mt-2">
                <span>Responsável pela Solução: <strong>${r.responsavel_solucao || "-"}</strong></span>
                <span>Data Conclusão: <strong>${r.data_conclusao_br || "-"}</strong></span>
            </div>
        </div>

        <!-- Operational and Fiscal Details -->
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3 bg-gray-50 p-3 rounded-xl border border-gray-100 text-gray-700">
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Filial</span>
                <span class="font-bold text-gray-800">${r.filial || "-"}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Setor Responsável</span>
                <span class="font-bold text-gray-800">${r.setor_responsavel || "-"}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Setor Impactado</span>
                <span class="font-medium text-gray-700">${r.setor_impactado || "-"}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Recorrente?</span>
                <span class="font-medium text-gray-700">${r.recorrente || "Não"}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Nota Fiscal / Série</span>
                <span class="font-medium text-gray-700">${r.nota_fiscal ? `NF ${r.nota_fiscal} (Série ${r.serie || 1})` : "-"}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Qtd. da Nota (kg)</span>
                <span class="font-medium text-gray-700">${r.qtd_nota_kg ? `${r.qtd_nota_kg} kg` : "-"}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Valor da Nota</span>
                <span class="font-bold text-emerald-700">${formatCurrency(r.valor_notas)}</span>
            </div>
            <div>
                <span class="text-gray-400 text-[10px] uppercase font-bold block">Nº do Chamado</span>
                <span class="font-mono text-gray-700">${r.numero_chamado || "-"}</span>
            </div>
        </div>

        <!-- Notes -->
        ${r.observacoes ? `
            <div class="p-3 bg-gray-50 rounded-xl border border-gray-100 text-xs">
                <span class="text-gray-400 font-bold uppercase text-[10px] block mb-1">Observações Complementares</span>
                <p class="text-gray-700">${escapeHtml(r.observacoes)}</p>
            </div>
        ` : ''}

        <!-- Timestamps -->
        <div class="flex justify-between items-center text-[10px] text-gray-400 pt-3 border-t border-gray-100">
            <span>Data Ocorrência: ${r.data_ocorrencia_br || "-"}</span>
            <span>Data Limite SLA: ${r.data_limite_sla_br || "-"}</span>
            <span>Última Atualização: ${r.ultima_atualizacao_br || "-"}</span>
        </div>
    `;

    document.getElementById("modal-detalhes")?.classList.remove("hidden");
    lucide.createIcons();
}

function closeDetailsModal() {
    document.getElementById("modal-detalhes")?.classList.add("hidden");
}

function openEditModalFromDetail() {
    if (currentRecordInDetail) {
        closeDetailsModal();
        openEditModal(currentRecordInDetail.id);
    }
}

function openEditModal(id) {
    const r = allRecords.find(item => item.id === id);
    if (!r) return;

    document.getElementById("edit_record_id").value = r.id;
    document.getElementById("edit-modal-title").textContent = `Atualizar Pendência #${r.id} - ${r.filial || ''}`;

    const statusEl = document.getElementById("edit_status");
    if (statusEl) statusEl.value = r.status || "Aberto";

    const prioEl = document.getElementById("edit_prioridade");
    if (prioEl) prioEl.value = r.prioridade || "Média";

    const prazoEl = document.getElementById("edit_prazo");
    if (prazoEl) prazoEl.value = r.prazo_iso || "";

    const concEl = document.getElementById("edit_data_conclusao");
    if (concEl) concEl.value = r.data_conclusao_iso || "";

    const respSolEl = document.getElementById("edit_responsavel_solucao");
    if (respSolEl) respSolEl.value = r.responsavel_solucao || "";

    const chamEl = document.getElementById("edit_numero_chamado");
    if (chamEl) chamEl.value = r.numero_chamado || "";

    const planoEl = document.getElementById("edit_plano_acao");
    if (planoEl) planoEl.value = r.plano_acao || "";

    const obsEl = document.getElementById("edit_observacoes");
    if (obsEl) obsEl.value = r.observacoes || "";

    document.getElementById("modal-editar")?.classList.remove("hidden");
    lucide.createIcons();
}

function closeEditModal() {
    document.getElementById("modal-editar")?.classList.add("hidden");
}

function toggleConclusionDate() {
    const statusVal = document.getElementById("edit_status")?.value;
    const concInput = document.getElementById("edit_data_conclusao");
    if (statusVal === "Concluído" && concInput && !concInput.value) {
        concInput.value = new Date().toISOString().split("T")[0];
    }
}

// Form Submission (Add Record)
async function handleFormSubmit(e, source) {
    e.preventDefault();
    const form = e.target;
    const submitBtn = form.querySelector('button[type="submit"]');
    const originalText = submitBtn ? submitBtn.innerHTML : "";

    const formData = new FormData(form);
    const payload = {};
    formData.forEach((value, key) => {
        payload[key] = value.trim();
    });

    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `<i data-lucide="loader" class="w-4 h-4 animate-spin"></i> Gravando no Excel...`;
        lucide.createIcons();
    }

    try {
        const res = await fetch("/api/records", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        const json = await res.json();
        if (json.success) {
            showToast(json.message || "Pendência cadastrada com sucesso no Excel!", "success");
            form.reset();
            setDefaultFormDates();
            if (source === "modal") {
                closeNewModal();
            } else {
                switchTab("consulta");
            }
            await refreshAllData();
        } else {
            showToast(json.message || "Erro ao salvar na planilha Excel.", "error");
        }
    } catch (err) {
        showToast("Erro na requisição: " + err.message, "error");
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalText;
            lucide.createIcons();
        }
    }
}

// Form Submission (Edit Record)
async function handleEditSubmit(e) {
    e.preventDefault();
    const form = e.target;
    const recordId = document.getElementById("edit_record_id")?.value;
    if (!recordId) return;

    const submitBtn = document.getElementById("btn-submit-edit");
    const originalText = submitBtn ? submitBtn.innerHTML : "";

    const formData = new FormData(form);
    const payload = {};
    formData.forEach((value, key) => {
        payload[key] = value.trim();
    });

    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `<i data-lucide="loader" class="w-4 h-4 animate-spin"></i> Atualizando Excel...`;
        lucide.createIcons();
    }

    try {
        const res = await fetch(`/api/records/${recordId}`, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        const json = await res.json();
        if (json.success) {
            showToast(json.message || "Pendência atualizada com sucesso no Excel!", "success");
            closeEditModal();
            await refreshAllData();
        } else {
            showToast(json.message || "Erro ao atualizar na planilha Excel.", "error");
        }
    } catch (err) {
        showToast("Erro na requisição: " + err.message, "error");
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalText;
            lucide.createIcons();
        }
    }
}

// Open Excel File Native
async function openExcelFile() {
    try {
        showToast("Solicitando abertura no Excel local...", "info");
        const res = await fetch("/api/open-excel", { method: "POST" });
        const json = await res.json();
        if (json.success) {
            showToast(json.message, "success");
        } else {
            showToast(json.message, "warning");
        }
    } catch (e) {
        showToast("Erro ao abrir Excel: " + e.message, "error");
    }
}

function escapeHtml(text) {
    if (!text) return "";
    return text.toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
