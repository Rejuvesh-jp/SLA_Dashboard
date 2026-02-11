/*************************
 * GLOBAL STATE & CONFIG
 *************************/
let allTickets = [];
let filteredTickets = [];
let chartInstances = {};
let sortConfig = { column: null, direction: 'asc' };
let paginationConfig = { currentPage: 1, pageSize: 25 };
let autoRefreshInterval = null;
let tableFilterState = {
    Priority: '',
    SLA: '',
    Team: '',
    'Assignment group': '',
    Created: ''
};
let previousTickets = new Set();
let closedTickets = [];

const AUTO_REFRESH_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
const API_BASE = '';

const TABLE_COLUMNS = [
    'Number',
    'Priority',
    'SLA',
    'Hours Outstanding',
    'Days Standing',
    'SLA Countdown',
    'Team',
    'Assignment group',
    'Created'
];

/*************************
 * INIT
 *************************/
document.addEventListener('DOMContentLoaded', async () => {
    await loadDashboardData();
    loadRecentlyClosedTickets();
    setupEventListeners();
    setupTableColumnFilters();
    startAutoRefresh();
});

/*************************
 * DATA LOADING
 *************************/
async function loadDashboardData(showLoadingState = true) {
    try {
        if (showLoadingState) showLoading();

        const res = await fetch(`${API_BASE}/api/tickets`);
        const result = await res.json();

        if (!result.success) throw new Error('API returned failure');

        // Hours Outstanding is calculated in real time on the frontend.
        // Ignore any "Hours Outstanding" value from Excel/API response.
        const currentTicketNumbers = new Set();
        const currentTicketsMap = new Map();
        
        allTickets = (result.data || []).map(ticket => {
            const { 'Hours Outstanding': _, ...ticketWithoutHours } = ticket;
            // Apply SLA breach logic if needed
            ticketWithoutHours.SLA = computeSLAStatus(ticketWithoutHours);
            currentTicketNumbers.add(ticketWithoutHours.Number);
            currentTicketsMap.set(ticketWithoutHours.Number, ticketWithoutHours);
            return ticketWithoutHours;
        });
        
        // Detect closed tickets (tickets that were in previousTickets but not in current list)
        if (previousTickets.size > 0) {
            previousTickets.forEach(ticketNumber => {
                if (!currentTicketNumbers.has(ticketNumber)) {
                    // Check if already tracked in closedTickets
                    const alreadyTracked = closedTickets.some(t => t.Number === ticketNumber);
                    if (alreadyTracked) return;
                    
                    // Find ticket in previous allTickets data (stored in window.previousTicketsData)
                    let closedTicket = null;
                    if (window.previousTicketsData && window.previousTicketsData.has(ticketNumber)) {
                        closedTicket = window.previousTicketsData.get(ticketNumber);
                    }
                    
                    if (closedTicket) {
                        closedTickets.push({
                            Number: closedTicket.Number || ticketNumber,
                            Priority: closedTicket.Priority || '',
                            Team: closedTicket.Team || '',
                            SLA: closedTicket.SLA || '',
                            closedAt: new Date().toISOString()
                        });
                    }
                }
            });
        }
        
        // Store current tickets data for next comparison
        window.previousTicketsData = new Map();
        currentTicketsMap.forEach((ticket, number) => {
            window.previousTicketsData.set(number, {
                Number: ticket.Number,
                Priority: ticket.Priority || '',
                Team: ticket.Team || '',
                SLA: ticket.SLA || ''
            });
        });
        
        // Update previousTickets for next comparison
        previousTickets = new Set(currentTicketNumbers);
        
        // Clean up closed tickets older than 24 hours and limit to 10 most recent
        const now = new Date();
        closedTickets = closedTickets
            .filter(t => {
                const closedDate = new Date(t.closedAt);
                const hoursAgo = (now - closedDate) / (1000 * 60 * 60);
                return hoursAgo <= 24;
            })
            .sort((a, b) => new Date(b.closedAt) - new Date(a.closedAt))
            .slice(0, 10);
        
        populateFilterOptions();
        applyFilters();
        refreshTableColumnFilters();
        await loadRecentlyClosedTickets();
        updateLastUpdated();

        console.log('Dashboard data loaded:', allTickets.length);
    } catch (err) {
        console.error(err);
        showError('Failed to load dashboard data');
    }
}

/*************************
 * EVENT LISTENERS
 *************************/
function setupEventListeners() {
    document.getElementById('searchInput')?.addEventListener('input', applyFilters);
    document.getElementById('filterPriority')?.addEventListener('change', applyFilters);
    document.getElementById('filterSLA')?.addEventListener('change', applyFilters);
    document.getElementById('filterTeam')?.addEventListener('change', applyFilters);
    document.getElementById('filterAssignmentGroup')?.addEventListener('change', applyFilters);
    document.getElementById('filterDateFrom')?.addEventListener('change', applyFilters);
    document.getElementById('filterDateTo')?.addEventListener('change', applyFilters);

    document.getElementById('pageSize')?.addEventListener('change', (e) => {
        paginationConfig.pageSize = parseInt(e.target.value) || 25;
        paginationConfig.currentPage = 1;
        updateDashboard();
    });

    document.querySelectorAll('.data-table th').forEach((th, idx) => {
        th.addEventListener('click', () => handleSort(idx));
    });
}

/*************************
 * FILTER DROPDOWNS
 *************************/
function populateFilterOptions() {
    // Get unique values directly from Excel column - no transformation
    const unique = (key) =>
        [...new Set(allTickets.map(t => t[key]).filter(Boolean))].sort();

    fillSelect('filterPriority', unique('Priority'));
    // SLA filter: Uses t["SLA"] directly from Excel - option.value === option.textContent === Excel value
    // First option is "All" (value=""), then unique Excel SLA values
    fillSelect('filterSLA', unique('SLA'));
    fillSelect('filterTeam', unique('Team'));
    fillSelect('filterAssignmentGroup', unique('Assignment group'));
}

function fillSelect(id, values, multi = false) {
    const el = document.getElementById(id);
    if (!el) return;

    // Clear existing options (except first "All" option for single-select)
    if (id === 'filterSLA' || id === 'filterPriority' || id === 'filterTeam' || id === 'filterAssignmentGroup') {
        // Keep first option (All/All Priorities/etc)
        const firstOption = el.options[0];
        el.innerHTML = '';
        if (firstOption) el.appendChild(firstOption);
    } else {
        el.innerHTML = '';
    }

    // Create options ensuring: option.value === option.textContent === Excel value
    values.forEach(v => {
        const option = document.createElement('option');
        option.value = v; // Exact Excel value
        option.textContent = v; // Exact Excel value (ensures value === textContent)
        el.appendChild(option);
    });
}

/*************************
 * KPI CALCULATION
 *************************/
function calculateKPIs(tickets) {
    let breached = 0;
    let pending = 0;

    tickets.forEach(t => {
        const status = t['SLA'] || '';
        // Use exact equality - no pattern matching
        if (status === 'SLA Breached') breached++;
        else if (status === 'Pending But Complaint') pending++;
    });

    return {
        total: tickets.length,
        breached,
        pending
    };
}

function updateKPIs(kpis) {
    document.getElementById('kpi-total').textContent = kpis.total;
    document.getElementById('kpi-breached').textContent = kpis.breached;
    document.getElementById('kpi-pending').textContent = kpis.pending;
}

/*************************
 * CHART DATA
 *************************/
function getSLAStatusColor(label) {
    const status = String(label || '');
    if (status === 'SLA Breached') return '#ef4444'; // Red
    if (status === 'Pending But Complaint') return '#10b981'; // Green
    return '#6b7280'; // Default gray
}

function calculateChartData(tickets) {
    const countBy = (key) =>
        tickets.reduce((acc, t) => {
            const v = t[key];
            if (v) {
                acc[v] = (acc[v] || 0) + 1;
            }
            return acc;
        }, {});

    // SLA Status chart - use exact Excel values
    const slaData = countBy('SLA');
    const slaLabels = Object.keys(slaData);
    const slaValues = Object.values(slaData);
    const slaColors = slaLabels.map(getSLAStatusColor);

    // Priority chart
    const priorityData = countBy('Priority');
    const priorityLabels = Object.keys(priorityData);
    const priorityValues = Object.values(priorityData);

    // Team chart
    const teamData = countBy('Team');
    const teamLabels = Object.keys(teamData);
    const teamValues = Object.values(teamData);

    // Aging buckets - use calculated Hours Outstanding
    const aging = { '0-4 hrs': 0, '4-8 hrs': 0, '8-24 hrs': 0, '>24 hrs': 0 };
    tickets.forEach(t => {
        const hoursOutstanding = calculateHoursOutstanding(t);
        const h = hoursOutstanding === '' ? 0 : parseFloat(hoursOutstanding) || 0;
        if (h < 4) aging['0-4 hrs']++;
        else if (h < 8) aging['4-8 hrs']++;
        else if (h < 24) aging['8-24 hrs']++;
        else aging['>24 hrs']++;
    });

    // Trend chart - Only count SLA Breached tickets
    const trendData = {};
    tickets.forEach(t => {
        if (t.Created && t.SLA === 'SLA Breached') {
            const date = new Date(t.Created);
            if (!isNaN(date.getTime())) {
                const dateStr = date.toISOString().split('T')[0];
                trendData[dateStr] = (trendData[dateStr] || 0) + 1;
            }
        }
    });
    const trendLabels = Object.keys(trendData).sort();
    const trendValues = trendLabels.map(d => trendData[d]);

    return {
        sla: { labels: slaLabels, values: slaValues, colors: slaColors },
        priority: { labels: priorityLabels, values: priorityValues },
        team: { labels: teamLabels, values: teamValues },
        aging: { labels: Object.keys(aging), values: Object.values(aging) },
        trend: { labels: trendLabels, values: trendValues }
    };
}

/*************************
 * CHART CREATION
 *************************/
function createDonutChart(canvasId, data) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }

    if (!data.labels.length) {
        document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'block');
        return;
    }
    document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'none');

    // Tickets by SLA Status ONLY: apply 3D-style depth (gradients + shadow + bevel) without changing data/labels.
    const isSLAStatusChart = canvasId === 'chart-sla-status';
    const baseColors = data.colors || data.labels.map(() => '#3b82f6');

    const clamp = (n, min, max) => Math.min(max, Math.max(min, n));
    const shadeHex = (hex, amt) => {
        const c = String(hex).replace('#', '');
        const num = parseInt(c, 16);
        const r = clamp((num >> 16) + amt, 0, 255);
        const g = clamp(((num >> 8) & 0xff) + amt, 0, 255);
        const b = clamp((num & 0xff) + amt, 0, 255);
        return `#${(r << 16 | g << 8 | b).toString(16).padStart(6, '0')}`;
    };

    // 3D shadow + bevel overlay (plugin is scoped ONLY to this chart instance)
    const slaDonut3D = {
        id: 'slaDonut3D',
        beforeDatasetDraw(chart, args) {
            if (args.index !== 0) return;
            const ctx2 = chart.ctx;
            ctx2.save();
            // Shadow only affects the donut arcs (restored immediately after)
            ctx2.shadowColor = 'rgba(0, 0, 0, 0.55)';
            ctx2.shadowBlur = 16;
            ctx2.shadowOffsetX = 0;
            ctx2.shadowOffsetY = 6;
        },
        afterDatasetDraw(chart, args) {
            if (args.index !== 0) return;
            const ctx2 = chart.ctx;
            ctx2.restore();

            // Subtle bevel/highlight across the donut ring for a 3D/premium look
            const meta = chart.getDatasetMeta(0);
            const first = meta && meta.data ? meta.data[0] : null;
            const p = first && first.getProps ? first.getProps(['x', 'y', 'innerRadius', 'outerRadius'], true) : first;
            if (!p || typeof p.x !== 'number' || typeof p.outerRadius !== 'number') return;

            ctx2.save();
            ctx2.globalCompositeOperation = 'source-atop';
            const g = ctx2.createRadialGradient(p.x, p.y, p.innerRadius, p.x, p.y, p.outerRadius);
            g.addColorStop(0, 'rgba(255, 255, 255, 0.16)');  // lighter inner edge
            g.addColorStop(0.55, 'rgba(255, 255, 255, 0.00)');
            g.addColorStop(1, 'rgba(0, 0, 0, 0.22)');        // darker outer edge
            ctx2.fillStyle = g;
            ctx2.beginPath();
            ctx2.arc(p.x, p.y, p.outerRadius, 0, Math.PI * 2);
            ctx2.closePath();
            ctx2.fill();
            ctx2.restore();
        }
    };

    chartInstances[canvasId] = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: data.labels,
            datasets: [{
                data: data.values,
                // 3D gradients (lighter toward center, darker toward edge) – SLA chart only
                backgroundColor: isSLAStatusChart
                    ? (context) => {
                        const i = context.dataIndex;
                        const base = baseColors[i] || '#3b82f6';
                        const el = context.element;
                        if (!el || typeof el.x !== 'number' || typeof el.outerRadius !== 'number') return base;
                        const grad = context.chart.ctx.createRadialGradient(el.x, el.y, el.innerRadius, el.x, el.y, el.outerRadius);
                        grad.addColorStop(0, shadeHex(base, 50));
                        grad.addColorStop(0.55, base);
                        grad.addColorStop(1, shadeHex(base, -70));
                        return grad;
                    }
                    : baseColors,
                borderWidth: isSLAStatusChart ? 3 : 2,
                borderColor: isSLAStatusChart ? '#0B0F14' : '#1F2937',
                hoverOffset: isSLAStatusChart ? 10 : 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            cutout: isSLAStatusChart ? '62%' : undefined,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#F9FAFB', font: { size: 12 } }
                },
                tooltip: {
                    backgroundColor: 'rgba(17, 24, 39, 0.88)',
                    titleColor: '#F9FAFB',
                    bodyColor: '#F9FAFB',
                    borderColor: 'rgba(255, 255, 255, 0.10)',
                    borderWidth: 1
                }
            }
        },
        plugins: isSLAStatusChart ? [slaDonut3D] : []
    });
}

function createBarChart(canvasId, data) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }

    if (!data.labels.length) {
        document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'block');
        return;
    }
    document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'none');

    chartInstances[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.labels,
            datasets: [{
                label: 'Count',
                data: data.values,
                backgroundColor: '#3b82f6',
                borderColor: '#2563eb',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: '#9CA3AF', stepSize: 1 },
                    grid: { color: '#1F2937' }
                },
                x: {
                    ticks: { color: '#9CA3AF' },
                    grid: { color: '#1F2937' }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(17, 24, 39, 0.88)',
                    titleColor: '#F9FAFB',
                    bodyColor: '#F9FAFB',
                    borderColor: 'rgba(255, 255, 255, 0.10)',
                    borderWidth: 1
                }
            }
        }
    });
}

// Tickets by Priority ONLY (chart-priority): custom colors + 3D-like depth
function createPriorityBarChart(canvasId, data) {
    const ctxEl = document.getElementById(canvasId);
    if (!ctxEl) return;

    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }

    if (!data.labels.length) {
        document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'block');
        return;
    }
    document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'none');

    const resolvePriorityKey = (label) => {
        const s = String(label || '');
        if (/priority\s*1/i.test(s) || /\bP?1\b/i.test(s) || /^\s*1\b/.test(s)) return 'P1';
        if (/priority\s*2/i.test(s) || /\bP?2\b/i.test(s) || /^\s*2\b/.test(s)) return 'P2';
        if (/priority\s*3/i.test(s) || /\bP?3\b/i.test(s) || /^\s*3\b/.test(s)) return 'P3';
        return 'OTHER';
    };

    const fillForIndex = (chart, index) => {
        const label = data.labels[index];
        const key = resolvePriorityKey(label);

        // Use per-bar geometry (when available) so gradients track actual bar height.
        const meta = chart.getDatasetMeta(0);
        const bar = meta && meta.data ? meta.data[index] : null;
        const props = bar && bar.getProps ? bar.getProps(['y', 'base'], true) : bar;
        const top = props && typeof props.y === 'number' ? props.y : (chart.chartArea?.top ?? 0);
        const bottom = props && typeof props.base === 'number' ? props.base : (chart.chartArea?.bottom ?? chart.height);

        // Requirement-specific colors for this chart (label-based, order-independent):
        // - Priority 1 – Critical → Red
        // - Priority 2 – High → Yellow
        // - Priority 3 – Moderate → Blue
        if (key === 'P1') {
            // Red (with subtle depth shading, still clearly "red")
            const g = chart.ctx.createLinearGradient(0, top, 0, bottom);
            g.addColorStop(0, '#fca5a5'); // lighter red highlight
            g.addColorStop(0.55, '#ef4444'); // base red
            g.addColorStop(1, '#b91c1c'); // deeper red
            return g;
        }
        if (key === 'P2') {
            // Yellow / Amber
            const g = chart.ctx.createLinearGradient(0, top, 0, bottom);
            g.addColorStop(0, '#fde047'); // yellow highlight
            g.addColorStop(0.55, '#f59e0b'); // amber
            g.addColorStop(1, '#b45309'); // deeper amber
            return g;
        }
        if (key === 'P3') {
            // Blue
            const g = chart.ctx.createLinearGradient(0, top, 0, bottom);
            g.addColorStop(0, '#93c5fd'); // lighter blue highlight
            g.addColorStop(0.55, '#3b82f6'); // base blue
            g.addColorStop(1, '#1d4ed8'); // deeper blue
            return g;
        }

        // All other priorities keep their existing color (same as default bar charts)
        return '#3b82f6';
    };

    const borderForLabel = (label) => {
        const key = resolvePriorityKey(label);
        if (key === 'P1') return '#7f1d1d';
        if (key === 'P2') return '#92400e';
        if (key === 'P3') return '#1e3a8a';
        return '#2563eb';
    };

    // Per-chart plugin to simulate 3D thickness + depth/shadow (does not affect other charts)
    const priorityBar3D = {
        id: 'priorityBar3D',
        beforeDatasetDraw(chart, args, pluginOptions) {
            if (args.index !== 0) return;
            const opts = pluginOptions || {};
            const ctx = chart.ctx;
            ctx.save();
            ctx.shadowColor = opts.shadowColor || 'rgba(0, 0, 0, 0.45)';
            ctx.shadowBlur = opts.shadowBlur ?? 10;
            ctx.shadowOffsetX = opts.shadowOffsetX ?? 6;
            ctx.shadowOffsetY = opts.shadowOffsetY ?? 4;
        },
        afterDatasetDraw(chart, args, pluginOptions) {
            if (args.index !== 0) return;
            const opts = pluginOptions || {};
            const ctx = chart.ctx;
            ctx.restore(); // remove shadow for the 3D faces

            const meta = chart.getDatasetMeta(args.index);
            const depthX = opts.depthX ?? 8;
            const depthY = opts.depthY ?? 6;

            ctx.save();
            for (let i = 0; i < meta.data.length; i++) {
                const bar = meta.data[i];
                const p = bar && bar.getProps ? bar.getProps(['x', 'y', 'base', 'width'], true) : bar;
                if (!p || typeof p.x !== 'number' || typeof p.width !== 'number') continue;

                const left = p.x - p.width / 2;
                const right = p.x + p.width / 2;
                const top = Math.min(p.y, p.base);
                const bottom = Math.max(p.y, p.base);

                // Right-side face (darker) to create thickness
                ctx.beginPath();
                ctx.moveTo(right, top);
                ctx.lineTo(right + depthX, top - depthY);
                ctx.lineTo(right + depthX, bottom - depthY);
                ctx.lineTo(right, bottom);
                ctx.closePath();
                ctx.fillStyle = 'rgba(0, 0, 0, 0.22)';
                ctx.fill();

                // Top face (slight highlight) for a 3D look
                ctx.beginPath();
                ctx.moveTo(left, top);
                ctx.lineTo(right, top);
                ctx.lineTo(right + depthX, top - depthY);
                ctx.lineTo(left + depthX, top - depthY);
                ctx.closePath();
                ctx.fillStyle = 'rgba(255, 255, 255, 0.10)';
                ctx.fill();
            }
            ctx.restore();
        }
    };

    chartInstances[canvasId] = new Chart(ctxEl, {
        type: 'bar',
        data: {
            labels: data.labels,
            datasets: [{
                label: 'Count',
                data: data.values,
                backgroundColor: (context) => fillForIndex(context.chart, context.dataIndex),
                borderColor: (context) => borderForLabel(data.labels[context.dataIndex]),
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: '#9CA3AF', stepSize: 1 },
                    grid: { color: '#1F2937' }
                },
                x: {
                    ticks: { color: '#9CA3AF' },
                    grid: { color: '#1F2937' }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(17, 24, 39, 0.88)',
                    titleColor: '#F9FAFB',
                    bodyColor: '#F9FAFB',
                    borderColor: 'rgba(255, 255, 255, 0.10)',
                    borderWidth: 1
                }
            }
        },
        plugins: [priorityBar3D]
    });
}

// Tickets by Team ONLY (chart-team): 3D look + stable team colors + in-chart toggle chips
function createTeamBarChart(canvasId, data, tickets) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    // Clean up ONLY this chart instance's handlers/state (stored on the previous Chart object)
    const prevChart = chartInstances[canvasId];
    if (prevChart && prevChart.$teamHandlers) {
        canvas.removeEventListener('click', prevChart.$teamHandlers.click);
        canvas.removeEventListener('mousemove', prevChart.$teamHandlers.move);
        canvas.style.cursor = 'default';
    }
    if (prevChart) {
        prevChart.destroy();
    }

    if (!data.labels.length) {
        document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'block');
        return;
    }
    document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'none');

    const labels = (data.labels || []).map(l => String(l));
    const allCounts = (data.values || []).map(v => Number(v) || 0);
    const ticketsArr = Array.isArray(tickets) ? tickets : [];
    const labelSet = new Set(labels);

    // --- Team filter mode (single source of truth) ---
    // chart._teamFilterMode = "OVERALL" | "BREACHED" | "PENDING"
    // IMPORTANT: exact Excel values (NO partial matches)
    // BREACHED => ticket.SLA === "SLA Breached"
    // PENDING  => ticket.SLA === "Pending But Complaint"
    const TEAM_FILTER_MODES = { OVERALL: 'OVERALL', BREACHED: 'BREACHED', PENDING: 'PENDING' };
    const computeTeamCounts = (teamFilterMode) => {
        const counts = new Map();
        for (let i = 0; i < ticketsArr.length; i++) {
            const t = ticketsArr[i] || {};
            const team = String(t.Team || '');
            if (!team || !labelSet.has(team)) continue;

            if (teamFilterMode === TEAM_FILTER_MODES.BREACHED) {
                if (t.SLA !== 'SLA Breached') continue;
            } else if (teamFilterMode === TEAM_FILTER_MODES.PENDING) {
                if (t.SLA !== 'Pending But Complaint') continue;
            }

            counts.set(team, (counts.get(team) || 0) + 1);
        }
        return labels.map(l => counts.get(l) || 0);
    };

    // --- Stable per-team colors (fixed map for known teams + deterministic fallback) ---
    const TEAM_COLORS = {
        "Infosec": "#e03131",
        "Analytics": "#1971c2",
        "SOC": "#2f9e44"
    };
    const clamp = (n, min, max) => Math.min(max, Math.max(min, n));
    const hashString = (str) => {
        // FNV-1a (stable across sessions)
        let h = 2166136261;
        for (let i = 0; i < str.length; i++) {
            h ^= str.charCodeAt(i);
            h = Math.imul(h, 16777619);
        }
        return h >>> 0;
    };
    const hslToHex = (h, s, l) => {
        s /= 100;
        l /= 100;
        const k = n => (n + h / 30) % 12;
        const a = s * Math.min(l, 1 - l);
        const f = n => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
        const toHex = x => Math.round(255 * x).toString(16).padStart(2, '0');
        return `#${toHex(f(0))}${toHex(f(8))}${toHex(f(4))}`;
    };
    const shadeHex = (hex, amt) => {
        const c = String(hex).replace('#', '');
        const num = parseInt(c, 16);
        const r = clamp((num >> 16) + amt, 0, 255);
        const g = clamp(((num >> 8) & 0xff) + amt, 0, 255);
        const b = clamp((num & 0xff) + amt, 0, 255);
        return `#${(r << 16 | g << 8 | b).toString(16).padStart(6, '0')}`;
    };
    const teamBaseColor = (teamLabel) => {
        const key = String(teamLabel);
        if (TEAM_COLORS[key]) return TEAM_COLORS[key];
        const h = hashString(key) % 360;
        return hslToHex(h, 72, 52);
    };

    // --- Gradient fill (lighter top, darker bottom) ---
    const fillForIndex = (chart, index) => {
        const base = teamBaseColor(labels[index]);

        // Prefer using the actual bar geometry so the gradient tracks the bar height.
        const meta = chart.getDatasetMeta(0);
        const bar = meta && meta.data ? meta.data[index] : null;
        const props = bar && bar.getProps ? bar.getProps(['y', 'base'], true) : bar;
        const top = props && typeof props.y === 'number' ? props.y : (chart.chartArea?.top ?? 0);
        const bottom = props && typeof props.base === 'number' ? props.base : (chart.chartArea?.bottom ?? chart.height);

        const g = chart.ctx.createLinearGradient(0, top, 0, bottom);
        g.addColorStop(0, shadeHex(base, 55));   // lighter top
        g.addColorStop(0.55, base);
        g.addColorStop(1, shadeHex(base, -55)); // darker bottom
        return g;
    };
    const borderForIndex = (index) => shadeHex(teamBaseColor(labels[index]), -70);

    // Overall mode (stacked): status colors (no per-team colors)
    const STATUS_BASE = {
        BREACHED: '#ef4444',            // red
        PENDING: '#f59e0b'              // yellow/amber
    };
    const fillForStatus = (context, baseHex) => {
        const chart = context.chart;
        const el = context.element;
        const y = el && typeof el.y === 'number' ? el.y : (chart.chartArea?.top ?? 0);
        const base = el && typeof el.base === 'number' ? el.base : (chart.chartArea?.bottom ?? chart.height);
        const top = Math.min(y, base);
        const bottom = Math.max(y, base);
        const g = chart.ctx.createLinearGradient(0, top, 0, bottom);
        g.addColorStop(0, shadeHex(baseHex, 55));   // lighter top
        g.addColorStop(0.55, baseHex);
        g.addColorStop(1, shadeHex(baseHex, -65)); // darker bottom
        return g;
    };

    // --- Chip UI + 3D extrusion are implemented as a per-chart plugin (scoped to this chart only) ---
    const roundRect = (ctx, x, y, w, h, r) => {
        const rr = Math.min(r, w / 2, h / 2);
        ctx.beginPath();
        ctx.moveTo(x + rr, y);
        ctx.arcTo(x + w, y, x + w, y + h, rr);
        ctx.arcTo(x + w, y + h, x, y + h, rr);
        ctx.arcTo(x, y + h, x, y, rr);
        ctx.arcTo(x, y, x + w, y, rr);
        ctx.closePath();
    };

    const modes = { ALL: 'ALL', BREACHED: 'BREACHED', PENDING: 'PENDING' };

    const team3DAndChips = {
        id: 'team3DAndChips',

        // 3D effect (shadow) — makes bars look lifted from the canvas
        beforeDatasetDraw(chart, args) {
            // In OVERALL (stacked) mode, apply to each stacked dataset; otherwise keep original behavior.
            if (chart._teamFilterMode !== TEAM_FILTER_MODES.OVERALL && args.index !== 0) return;
            const ctx = chart.ctx;
            ctx.save();
            ctx.shadowColor = 'rgba(0, 0, 0, 0.55)';
            ctx.shadowBlur = 14;
            ctx.shadowOffsetX = 8;
            ctx.shadowOffsetY = 6;
        },

        // 3D effect (right-side + top extrusion)
        afterDatasetDraw(chart, args) {
            // In OVERALL (stacked) mode, apply to each stacked dataset; otherwise keep original behavior.
            if (chart._teamFilterMode !== TEAM_FILTER_MODES.OVERALL && args.index !== 0) return;
            const ctx = chart.ctx;
            ctx.restore(); // stop shadow for extrusion faces and UI

            const meta = chart.getDatasetMeta(args.index);
            const depthX = 10;
            const depthY = 7;

            ctx.save();
            for (let i = 0; i < meta.data.length; i++) {
                const bar = meta.data[i];
                const p = bar && bar.getProps ? bar.getProps(['x', 'y', 'base', 'width'], true) : bar;
                if (!p || typeof p.x !== 'number' || typeof p.width !== 'number') continue;

                const left = p.x - p.width / 2;
                const right = p.x + p.width / 2;
                const top = Math.min(p.y, p.base);
                const bottom = Math.max(p.y, p.base);

                // Right face (darker) => thickness
                ctx.beginPath();
                ctx.moveTo(right, top);
                ctx.lineTo(right + depthX, top - depthY);
                ctx.lineTo(right + depthX, bottom - depthY);
                ctx.lineTo(right, bottom);
                ctx.closePath();
                ctx.fillStyle = 'rgba(0, 0, 0, 0.24)';
                ctx.fill();

                // Top face (highlight) => perspective
                ctx.beginPath();
                ctx.moveTo(left, top);
                ctx.lineTo(right, top);
                ctx.lineTo(right + depthX, top - depthY);
                ctx.lineTo(left + depthX, top - depthY);
                ctx.closePath();
                ctx.fillStyle = 'rgba(255, 255, 255, 0.12)';
                ctx.fill();
            }
            ctx.restore();
        }
    };

    const chart = new Chart(canvas, {
        type: 'bar',
        data: {
            labels,
            // Default state (OVERALL): one stacked bar per team split into status composition
            datasets: [{
                label: 'SLA Breached',
                data: computeTeamCounts(TEAM_FILTER_MODES.BREACHED),
                stack: 'status',
                backgroundColor: (context) => fillForStatus(context, STATUS_BASE.BREACHED),
                borderColor: shadeHex(STATUS_BASE.BREACHED, -85),
                borderWidth: 1,
                borderRadius: 6,
                // Thicker bars (3D look) — scoped to this chart only
                barThickness: 'flex',
                maxBarThickness: 46,
                categoryPercentage: 0.78,
                barPercentage: 0.92
            }, {
                label: 'Pending But Complaint',
                data: computeTeamCounts(TEAM_FILTER_MODES.PENDING),
                stack: 'status',
                backgroundColor: (context) => fillForStatus(context, '#10b981'),
                borderColor: shadeHex(STATUS_BASE.PENDING, -85),
                borderWidth: 1,
                borderRadius: 6,
                barThickness: 'flex',
                maxBarThickness: 46,
                categoryPercentage: 0.78,
                barPercentage: 0.92
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            // Animate bars upward from zero with smooth easing
            animation: {
                duration: 1200,
                easing: 'easeOutQuart'
            },
            animations: {
                y: { from: 0 }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: '#9CA3AF', stepSize: 1 },
                    grid: { color: '#1F2937' },
                    stacked: true
                },
                x: {
                    ticks: { color: '#9CA3AF' },
                    grid: { color: '#1F2937' },
                    stacked: true
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(17, 24, 39, 0.88)',
                    titleColor: '#F9FAFB',
                    bodyColor: '#F9FAFB',
                    borderColor: 'rgba(255, 255, 255, 0.10)',
                    borderWidth: 1,
                    callbacks: {
                        // In OVERALL mode, show status-specific labels; other modes keep existing "Count: <value>"
                        label: (context) => {
                            const val = context.formattedValue;
                            if (context.chart._teamFilterMode === TEAM_FILTER_MODES.OVERALL) {
                                return `${context.dataset.label}: ${val}`;
                            }
                            return `Count: ${val}`;
                        }
                    }
                }
            }
        },
        plugins: [team3DAndChips]
    });

    // All state is scoped ONLY to this Chart instance (no globals)
    chart._teamFilterMode = TEAM_FILTER_MODES.OVERALL;
    chart.$teamMode = modes.ALL;

    // --- 3 toggle buttons inside the Team chart container (no style changes here; only .active class) ---
    const teamCard = canvas.closest('.chart-card');
    // Remove any previous toggle container (and any older toggle containers, if present)
    teamCard?.querySelector('.team-chart-toggles')?.remove();

    // Build a right-aligned header row for title + buttons (flexbox via CSS)
    let headerRow = teamCard ? teamCard.querySelector('.team-chart-header') : null;
    if (teamCard && !headerRow) {
        const titleEl = teamCard.querySelector('h3');
        if (titleEl) {
            headerRow = document.createElement('div');
            headerRow.className = 'team-chart-header';
            teamCard.insertBefore(headerRow, titleEl);
            headerRow.appendChild(titleEl);
        }
    }

    let toggleBar = null;
    if (teamCard) {
        toggleBar = document.createElement('div');
        toggleBar.className = 'team-chart-toggles';

        const btnOverall = document.createElement('button');
        btnOverall.type = 'button';
        btnOverall.textContent = 'Overall';
        btnOverall.dataset.mode = TEAM_FILTER_MODES.OVERALL;

        const btnBreached = document.createElement('button');
        btnBreached.type = 'button';
        btnBreached.textContent = 'SLA Breached';
        btnBreached.dataset.mode = TEAM_FILTER_MODES.BREACHED;

        const btnPending = document.createElement('button');
        btnPending.type = 'button';
        btnPending.textContent = 'Pending But Complaint';
        btnPending.dataset.mode = TEAM_FILTER_MODES.PENDING;

        toggleBar.appendChild(btnOverall);
        toggleBar.appendChild(btnBreached);
        toggleBar.appendChild(btnPending);

        // Place toggles to the RIGHT side of the Team chart title area
        if (headerRow) headerRow.appendChild(toggleBar);
        else teamCard.insertBefore(toggleBar, canvas);

        chart.$teamButtons = { overall: btnOverall, breached: btnBreached, pending: btnPending };
    }

    const syncActiveButtons = () => {
        const btns = chart.$teamButtons;
        if (!btns) return;
        const m = chart._teamFilterMode || TEAM_FILTER_MODES.OVERALL;
        btns.overall.classList.toggle('active', m === TEAM_FILTER_MODES.OVERALL);
        btns.breached.classList.toggle('active', m === TEAM_FILTER_MODES.BREACHED);
        btns.pending.classList.toggle('active', m === TEAM_FILTER_MODES.PENDING);
    };

    const setTeamMode = (mode) => {
        chart._teamFilterMode = mode;

        // Overall button => stacked status composition (red + yellow) per team.
        // Other buttons => existing single-series view (unchanged behavior).
        if (mode === TEAM_FILTER_MODES.OVERALL) {
            chart.options.scales.x.stacked = true;
            chart.options.scales.y.stacked = true;
            chart.data.datasets = [{
                label: 'SLA Breached',
                data: computeTeamCounts(TEAM_FILTER_MODES.BREACHED),
                stack: 'status',
                backgroundColor: (context) => fillForStatus(context, STATUS_BASE.BREACHED),
                borderColor: shadeHex(STATUS_BASE.BREACHED, -85),
                borderWidth: 1,
                borderRadius: 6,
                barThickness: 'flex',
                maxBarThickness: 46,
                categoryPercentage: 0.78,
                barPercentage: 0.92
            }, {
                label: 'Pending But Complaint',
                data: computeTeamCounts(TEAM_FILTER_MODES.PENDING),
                stack: 'status',
                backgroundColor: (context) => fillForStatus(context, '#10b981'),
                borderColor: shadeHex(STATUS_BASE.PENDING, -85),
                borderWidth: 1,
                borderRadius: 6,
                barThickness: 'flex',
                maxBarThickness: 46,
                categoryPercentage: 0.78,
                barPercentage: 0.92
            }];
        } else {
            chart.options.scales.x.stacked = false;
            chart.options.scales.y.stacked = false;
            chart.data.datasets = [{
                label: 'Count',
                data: computeTeamCounts(mode),
                backgroundColor: (context) => fillForIndex(context.chart, context.dataIndex),
                borderColor: (context) => borderForIndex(context.dataIndex),
                borderWidth: 1,
                borderRadius: 6,
                barThickness: 'flex',
                maxBarThickness: 46,
                categoryPercentage: 0.78,
                barPercentage: 0.92
            }];
        }

        // Keep the in-canvas chip highlight state consistent (visuals unchanged)
        chart.$teamMode =
            mode === TEAM_FILTER_MODES.BREACHED ? modes.BREACHED :
            mode === TEAM_FILTER_MODES.PENDING ? modes.PENDING :
            modes.ALL;

        syncActiveButtons();
        chart.update();
    };

    // Default visual state
    syncActiveButtons();

    // Button click handlers (exact mapping; only one active at a time)
    if (chart.$teamButtons) {
        chart.$teamButtons.overall.addEventListener('click', () => setTeamMode(TEAM_FILTER_MODES.OVERALL));
        chart.$teamButtons.breached.addEventListener('click', () => setTeamMode(TEAM_FILTER_MODES.BREACHED));
        chart.$teamButtons.pending.addEventListener('click', () => setTeamMode(TEAM_FILTER_MODES.PENDING));
    }

    chartInstances[canvasId] = chart;
}

function createHorizontalBarChart(canvasId, data) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }

    if (!data.labels.length) {
        document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'block');
        return;
    }
    document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'none');

    // Aging Buckets ONLY (chart-aging): per-bucket 3D gradients + depth (no data/label changes)
    const isAgingBucketsChart = canvasId === 'chart-aging';
    const clamp = (n, min, max) => Math.min(max, Math.max(min, n));
    const shadeHex = (hex, amt) => {
        const c = String(hex).replace('#', '');
        const num = parseInt(c, 16);
        const r = clamp((num >> 16) + amt, 0, 255);
        const g = clamp(((num >> 8) & 0xff) + amt, 0, 255);
        const b = clamp((num & 0xff) + amt, 0, 255);
        return `#${(r << 16 | g << 8 | b).toString(16).padStart(6, '0')}`;
    };
    const normalizeBucketLabel = (label) => String(label || '').replace(/\u2013/g, '-').trim(); // en-dash -> hyphen
    const bucketBaseColor = (label) => {
        const s = normalizeBucketLabel(label);
        // REQUIRED COLOR MAPPING (EXACT bucket meaning)
        if (s === '>24 hrs') return '#ef4444';    // Red (critical)
        if (s === '8-24 hrs') return '#f59e0b';   // Yellow/amber (warning)
        if (s === '4-8 hrs') return '#22c55e';    // Moderate green
        if (s === '0-4 hrs') return '#10b981';    // Fresh green
        return '#3b82f6'; // fallback (should not be used for standard buckets)
    };
    const bucketFill = (chart, index, isHover = false) => {
        const label = data.labels[index];
        const base = bucketBaseColor(label);
        const meta = chart.getDatasetMeta(0);
        const bar = meta && meta.data ? meta.data[index] : null;
        const props = bar && bar.getProps ? bar.getProps(['y', 'height'], true) : bar;
        const top = props && typeof props.y === 'number' && typeof props.height === 'number'
            ? (props.y - props.height / 2)
            : (chart.chartArea?.top ?? 0);
        const bottom = props && typeof props.y === 'number' && typeof props.height === 'number'
            ? (props.y + props.height / 2)
            : (chart.chartArea?.bottom ?? chart.height);

        // Vertical gradient: lighter top -> darker bottom (3D depth shading)
        const g = chart.ctx.createLinearGradient(0, top, 0, bottom);
        const hi = isHover ? 70 : 55;
        const mid = isHover ? 15 : 0;
        const lo = isHover ? -50 : -65;
        g.addColorStop(0, shadeHex(base, hi));   // lighter top
        g.addColorStop(0.55, shadeHex(base, mid));
        g.addColorStop(1, shadeHex(base, lo));  // darker bottom
        return g;
    };

    const aging3D = {
        id: 'aging3D',
        beforeDatasetDraw(chart, args) {
            if (args.index !== 0) return;
            const c = chart.ctx;
            c.save();
            c.shadowColor = 'rgba(0, 0, 0, 0.50)';
            c.shadowBlur = 12;
            c.shadowOffsetX = 6;
            c.shadowOffsetY = 5;
        },
        afterDatasetDraw(chart, args) {
            if (args.index !== 0) return;
            const c = chart.ctx;
            c.restore(); // stop shadow for faces/highlights

            const meta = chart.getDatasetMeta(0);
            const depthX = 10;
            const depthY = 7;

            c.save();
            for (let i = 0; i < meta.data.length; i++) {
                const bar = meta.data[i];
                const p = bar && bar.getProps ? bar.getProps(['x', 'y', 'base', 'height'], true) : bar;
                if (!p || typeof p.x !== 'number' || typeof p.y !== 'number') continue;

                const halfH = typeof p.height === 'number' ? p.height / 2 : 0;
                const top = p.y - halfH;
                const bottom = p.y + halfH;
                const left = Math.min(p.base ?? 0, p.x);
                const right = Math.max(p.base ?? 0, p.x);

                // Right-side face (darker) for thickness
                c.beginPath();
                c.moveTo(right, top);
                c.lineTo(right + depthX, top - depthY);
                c.lineTo(right + depthX, bottom - depthY);
                c.lineTo(right, bottom);
                c.closePath();
                c.fillStyle = 'rgba(0, 0, 0, 0.22)';
                c.fill();

                // Top face (highlight) for bevel
                c.beginPath();
                c.moveTo(left, top);
                c.lineTo(right, top);
                c.lineTo(right + depthX, top - depthY);
                c.lineTo(left + depthX, top - depthY);
                c.closePath();
                c.fillStyle = 'rgba(255, 255, 255, 0.14)';
                c.fill();

                // Edge highlight (subtle) along the top edge of the bar
                c.beginPath();
                c.moveTo(left + 0.5, top + 0.5);
                c.lineTo(right - 0.5, top + 0.5);
                c.strokeStyle = 'rgba(255, 255, 255, 0.20)';
                c.lineWidth = 1;
                c.stroke();
            }
            c.restore();
        }
    };

    chartInstances[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.labels,
            datasets: [{
                label: 'Count',
                data: data.values,
                backgroundColor: isAgingBucketsChart
                    ? (context) => bucketFill(context.chart, context.dataIndex, false)
                    : '#3b82f6',
                hoverBackgroundColor: isAgingBucketsChart
                    ? (context) => bucketFill(context.chart, context.dataIndex, true)
                    : undefined,
                borderColor: isAgingBucketsChart
                    ? (context) => shadeHex(bucketBaseColor(data.labels[context.dataIndex]), -85)
                    : '#2563eb',
                borderWidth: 1,
                borderRadius: isAgingBucketsChart ? 6 : 0
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { color: '#9CA3AF', stepSize: 1 },
                    grid: { color: '#1F2937' }
                },
                y: {
                    ticks: { color: '#9CA3AF' },
                    grid: { color: '#1F2937' }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(17, 24, 39, 0.88)',
                    titleColor: '#F9FAFB',
                    bodyColor: '#F9FAFB',
                    borderColor: 'rgba(255, 255, 255, 0.10)',
                    borderWidth: 1
                }
            }
        },
        plugins: isAgingBucketsChart ? [aging3D] : []
    });
}

function createLineChart(canvasId, data) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    if (chartInstances[canvasId]) {
        chartInstances[canvasId].destroy();
    }

    if (!data.labels.length) {
        document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'block');
        return;
    }
    document.getElementById(canvasId + '-empty')?.style.setProperty('display', 'none');

    chartInstances[canvasId] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: data.labels,
            datasets: [{
                label: 'SLA Breached Tickets',
                data: data.values,
                borderColor: '#ef4444',
                backgroundColor: 'rgba(239, 68, 68, 0.1)',
                borderWidth: 2,
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { color: '#9CA3AF', stepSize: 1 },
                    grid: { color: '#1F2937' }
                },
                x: {
                    ticks: { color: '#9CA3AF' },
                    grid: { color: '#1F2937' }
                }
            },
            plugins: {
                legend: {
                    display: true,
                    labels: { color: '#F9FAFB' }
                },
                tooltip: {
                    backgroundColor: 'rgba(17, 24, 39, 0.88)',
                    titleColor: '#F9FAFB',
                    bodyColor: '#F9FAFB',
                    borderColor: 'rgba(255, 255, 255, 0.10)',
                    borderWidth: 1
                }
            }
        }
    });
}

/*************************
 * TABLE RENDERING
 *************************/
function renderTable(tickets) {
    const tbody = document.getElementById('ticketsTableBody');
    
    // Apply column filters (only affects table display, not charts/KPIs)
    let tableFilteredTickets = tickets.filter(t => {
        // Priority filter
        if (tableFilterState.Priority && t.Priority !== tableFilterState.Priority) return false;
        
        // SLA filter
        if (tableFilterState.SLA && t['SLA'] !== tableFilterState.SLA) return false;
        
        // Team filter
        if (tableFilterState.Team && t.Team !== tableFilterState.Team) return false;
        
        // Assignment group filter
        if (tableFilterState['Assignment group'] && t['Assignment group'] !== tableFilterState['Assignment group']) return false;
        
        // Created filter (date matching)
        if (tableFilterState.Created) {
            if (!t.Created) return false;
            const createdDate = new Date(t.Created);
            if (isNaN(createdDate.getTime())) return false;
            const ticketDate = createdDate.toISOString().split('T')[0];
            if (ticketDate !== tableFilterState.Created) return false;
        }
        
        return true;
    });
    
    if (!tableFilteredTickets.length) {
        tbody.innerHTML = `<tr><td colspan="9">No tickets found</td></tr>`;
        return;
    }

    const start = (paginationConfig.currentPage - 1) * paginationConfig.pageSize;
    const page = tableFilteredTickets.slice(start, start + paginationConfig.pageSize);

    tbody.innerHTML = page.map(t => {
        const sla = t['SLA'] || '';
        let rowClass = '';
        if (sla === 'SLA Breached') rowClass = 'row-breached';
        else if (sla === 'Pending But Complaint') rowClass = 'row-pending';
        
        const countdown = calculateSLACountdown(t);
        const daysStanding = calculateDaysStanding(t);
        
        return `
        <tr class="${rowClass}">
            <td>${t.Number || ''}</td>
            <td>${t.Priority || ''}</td>
            <td>${sla}</td>
            <td>${calculateHoursOutstanding(t)}</td>
            <td>${daysStanding}</td>
            <td class="${countdown.class}">${countdown.text}</td>
            <td>${t.Team || ''}</td>
            <td>${t['Assignment group'] || ''}</td>
            <td>${formatDate(t.Created)}</td>
        </tr>
    `;
    }).join('');

    updatePagination(tableFilteredTickets.length);
    document.getElementById('tableCount').textContent = `Showing ${tableFilteredTickets.length} tickets`;
}

/*************************
 * FILTERING & SORTING
 *************************/
function applyFilters() {
    const search = document.getElementById('searchInput')?.value.toLowerCase() || '';
    const priority = document.getElementById('filterPriority')?.value || '';
    const sla = document.getElementById('filterSLA')?.value || '';
    const team = document.getElementById('filterTeam')?.value || '';
    const assignmentGroup = document.getElementById('filterAssignmentGroup')?.value || '';
    const dateFrom = document.getElementById('filterDateFrom')?.value || '';
    const dateTo = document.getElementById('filterDateTo')?.value || '';

    filteredTickets = allTickets.filter(t => {
        // Search filter
        if (search) {
            const countdown = calculateSLACountdown(t);
            const daysStanding = calculateDaysStanding(t);
            const searchable = [
                t.Number, t.Priority, t['SLA'], countdown.text, daysStanding, t.Team, t['Assignment group'], t.Created
            ].map(v => String(v || '').toLowerCase()).join(' ');
            if (!searchable.includes(search)) return false;
        }

        // Priority filter
        if (priority && t.Priority !== priority) return false;

        // SLA filter
        if (sla && t['SLA'] !== sla) return false;

        // Team filter
        if (team && t.Team !== team) return false;

        // Assignment group filter
        if (assignmentGroup && t['Assignment group'] !== assignmentGroup) return false;

        // Date range filter
        if (dateFrom || dateTo) {
            if (!t.Created) return false;
            const createdDate = new Date(t.Created);
            if (isNaN(createdDate.getTime())) return false;
            const ticketDate = createdDate.toISOString().split('T')[0];
            if (dateFrom && ticketDate < dateFrom) return false;
            if (dateTo && ticketDate > dateTo) return false;
        }

        return true;
    });

    // Reset to first page when filtering
    paginationConfig.currentPage = 1;
    updateDashboard();
}

function handleSort(columnIndex) {
    const column = TABLE_COLUMNS[columnIndex];
    if (!column) return;

    if (sortConfig.column === column) {
        sortConfig.direction = sortConfig.direction === 'asc' ? 'desc' : 'asc';
    } else {
        sortConfig.column = column;
        sortConfig.direction = 'asc';
    }

    filteredTickets.sort((a, b) => {
        let x = a[column];
        let y = b[column];

        if (column === 'Hours Outstanding') {
            const hoursA = calculateHoursOutstanding(a);
            const hoursB = calculateHoursOutstanding(b);
            x = hoursA === '' ? 0 : parseFloat(hoursA) || 0;
            y = hoursB === '' ? 0 : parseFloat(hoursB) || 0;
        } else if (column === 'Days Standing') {
            const daysA = calculateDaysStanding(a);
            const daysB = calculateDaysStanding(b);
            x = daysA === 'NA' ? -1 : parseInt(daysA) || 0;
            y = daysB === 'NA' ? -1 : parseInt(daysB) || 0;
        } else if (column === 'SLA Countdown') {
            // Sort by time remaining (positive) or breached time (negative)
            const countdownA = calculateSLACountdown(a);
            const countdownB = calculateSLACountdown(b);
            if (countdownA.text === 'N/A') x = 999999;
            else if (countdownA.text.includes('Breached')) {
                // Extract breached time (negative for sorting)
                const match = countdownA.text.match(/(\d+)h\s*(\d+)m/);
                x = match ? -(parseInt(match[1]) * 60 + parseInt(match[2])) : -999999;
            } else {
                // Extract remaining time (positive for sorting)
                const match = countdownA.text.match(/(\d+)h\s*(\d+)m/);
                x = match ? (parseInt(match[1]) * 60 + parseInt(match[2])) : 0;
            }
            if (countdownB.text === 'N/A') y = 999999;
            else if (countdownB.text.includes('Breached')) {
                const match = countdownB.text.match(/(\d+)h\s*(\d+)m/);
                y = match ? -(parseInt(match[1]) * 60 + parseInt(match[2])) : -999999;
            } else {
                const match = countdownB.text.match(/(\d+)h\s*(\d+)m/);
                y = match ? (parseInt(match[1]) * 60 + parseInt(match[2])) : 0;
            }
        } else if (column === 'Created') {
            x = x ? new Date(x).getTime() : 0;
            y = y ? new Date(y).getTime() : 0;
        } else {
            x = String(x || '').toLowerCase();
            y = String(y || '').toLowerCase();
        }

        if (x < y) return sortConfig.direction === 'asc' ? -1 : 1;
        if (x > y) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
    });

    updateDashboard();
}

/*************************
 * DASHBOARD UPDATE
 *************************/
function updateDashboard() {
    const kpis = calculateKPIs(filteredTickets);
    updateKPIs(kpis);

    const chartData = calculateChartData(filteredTickets);
    createDonutChart('chart-sla-status', chartData.sla);
    createPriorityBarChart('chart-priority', chartData.priority);
    createTeamBarChart('chart-team', chartData.team, filteredTickets);
    createHorizontalBarChart('chart-aging', chartData.aging);
    createLineChart('chart-trend', chartData.trend);

    renderTable(filteredTickets);
}

/*************************
 * PAGINATION
 *************************/
function updatePagination(totalItems) {
    const totalPages = Math.ceil(totalItems / paginationConfig.pageSize);
    const currentPage = paginationConfig.currentPage;
    const paginationEl = document.getElementById('pagination');
    if (!paginationEl) return;

    if (totalPages <= 1) {
        paginationEl.innerHTML = '';
        return;
    }

    let html = '';
    if (currentPage > 1) {
        html += `<button onclick="changePage(${currentPage - 1})" class="page-btn">Previous</button>`;
    }
    html += `<span class="page-info">Page ${currentPage} of ${totalPages}</span>`;
    if (currentPage < totalPages) {
        html += `<button onclick="changePage(${currentPage + 1})" class="page-btn">Next</button>`;
    }
    paginationEl.innerHTML = html;
}

function changePage(page) {
    paginationConfig.currentPage = page;
    renderTable(filteredTickets);
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

/*************************
 * AUTO REFRESH
 *************************/
function startAutoRefresh() {
    if (autoRefreshInterval) clearInterval(autoRefreshInterval);
    autoRefreshInterval = setInterval(() => {
        loadDashboardData(false);
    }, AUTO_REFRESH_INTERVAL_MS);
}

/*************************
 * SLA STATUS COMPUTATION
 *************************/
function computeSLAStatus(ticket) {
    // SLA is computed ONLY on the frontend from Priority + Created.
    // Ignore any SLA value coming from Excel.

    if (!ticket || !ticket.Priority || !ticket.Created) return '';

    const createdDate = new Date(ticket.Created);
    if (isNaN(createdDate.getTime())) return '';

    // Extract leading digit from Priority: "1-Critical", "2-High", "3-Moderate"
    const priorityStr = String(ticket.Priority || '').trim();
    const match = priorityStr.match(/^(\d)/);
    if (!match) return '';

    const priorityDigit = parseInt(match[1], 10);
    let slaHours = 0;
    if (priorityDigit === 1) slaHours = 4;
    else if (priorityDigit === 2) slaHours = 12;
    else if (priorityDigit === 3) slaHours = 24;
    else return '';

    const elapsedHours = (Date.now() - createdDate.getTime()) / 3600000;
    if (!isFinite(elapsedHours)) return '';

    // Within SLA window => Pending But Complaint, otherwise => SLA Breached
    return elapsedHours >= slaHours ? 'SLA Breached' : 'Pending But Complaint';
}

/*************************
 * SLA COUNTDOWN CALCULATION
 *************************/
function calculateSLACountdown(ticket) {
    // Edge cases: Missing Priority or invalid Created date
    if (!ticket.Priority || !ticket.Created) {
        return { text: 'N/A', class: 'countdown-na' };
    }
    
    const createdDate = new Date(ticket.Created);
    if (isNaN(createdDate.getTime())) {
        return { text: 'N/A', class: 'countdown-na' };
    }
    
    // Extract first digit from Priority using regex
    const priorityStr = String(ticket.Priority || '');
    const match = priorityStr.match(/^(\d)/);
    if (!match) {
        return { text: 'N/A', class: 'countdown-na' };
    }
    
    const priorityDigit = parseInt(match[1], 10);
    
    // SLA HOURS MAPPING: 1 → 4 hours, 2 → 12 hours, 3 → 24 hours
    let slaHours = 0;
    if (priorityDigit === 1) slaHours = 4;
    else if (priorityDigit === 2) slaHours = 12;
    else if (priorityDigit === 3) slaHours = 24;
    else {
        return { text: 'N/A', class: 'countdown-na' };
    }
    
    // Countdown Calculation
    const now = new Date();
    const elapsedMs = now - createdDate;
    const elapsedHours = elapsedMs / 3600000;
    const remaining = slaHours - elapsedHours;
    
    // Format display text
    let text = '';
    let colorClass = '';
    
    if (remaining >= 0) {
        // Time remaining
        const remainingMinutes = Math.floor(remaining * 60);
        const hours = Math.floor(remainingMinutes / 60);
        const minutes = remainingMinutes % 60;
        text = `${hours}h ${minutes}m remaining`;
        
        // Color Logic: Remaining > 25% SLA → green, Remaining <= 25% SLA → amber
        const percentRemaining = (remaining / slaHours) * 100;
        if (percentRemaining > 25) {
            colorClass = 'countdown-green';
        } else {
            colorClass = 'countdown-amber';
        }
    } else {
        // Breached
        const breachedHours = Math.abs(remaining);
        const breachedMinutes = Math.floor(breachedHours * 60);
        const hours = Math.floor(breachedMinutes / 60);
        const minutes = breachedMinutes % 60;
        text = `Breached by ${hours}h ${minutes}m`;
        colorClass = 'countdown-red';
    }
    
    return { text, class: colorClass };
}

/*************************
 * UTILITIES
 *************************/
// Hours Outstanding is calculated in real time on the frontend.
// This function computes: (Current time - Created timestamp) in hours.
// Any "Hours Outstanding" value from Excel/API is ignored.
function calculateHoursOutstanding(ticket) {
    if (!ticket.Created) return '';
    
    const createdDate = new Date(ticket.Created);
    if (isNaN(createdDate.getTime())) return '';
    
    const now = new Date();
    const diffMs = now - createdDate;
    const diffHours = diffMs / (1000 * 60 * 60);
    
    return diffHours.toFixed(2);
}

// Days Standing is calculated in real time.
// (Current time - Created timestamp) in whole days (rounded down).
function calculateDaysStanding(ticket) {
    if (!ticket.Created) return 'NA';
    
    const createdDate = new Date(ticket.Created);
    if (isNaN(createdDate.getTime())) return 'NA';
    
    const now = new Date();
    const diffMs = now - createdDate;
    const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
    
    return diffDays < 0 ? 0 : diffDays;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatDate(dateStr) {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    return date.toLocaleString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function formatNumber(num) {
    return new Intl.NumberFormat('en-US').format(num);
}

function updateLastUpdated() {
    const el = document.getElementById('lastUpdated');
    if (el) {
        const now = new Date();
        el.textContent = `Last updated: ${now.toLocaleTimeString()}`;
    }
}

function refreshData() {
    loadDashboardData();
}

function clearAllFilters() {
    document.getElementById('searchInput').value = '';
    document.getElementById('filterPriority').value = '';
    document.getElementById('filterSLA').value = '';
    document.getElementById('filterTeam').value = '';
    document.getElementById('filterAssignmentGroup').value = '';
    document.getElementById('filterDateFrom').value = '';
    document.getElementById('filterDateTo').value = '';
    applyFilters();
}

function showLoading() {
    const tbody = document.getElementById('ticketsTableBody');
    if (tbody) {
        tbody.innerHTML = '<tr><td colspan="9" class="loading">Loading data...</td></tr>';
    }
}

function showError(message) {
    const tbody = document.getElementById('ticketsTableBody');
    if (tbody) {
        tbody.innerHTML = `<tr><td colspan="9" class="error">${escapeHtml(message)}</td></tr>`;
    }
}

/*************************
 * TABLE COLUMN FILTERS
 *************************/
function setupTableColumnFilters() {
    // Columns that should have filters (excluding "Number")
    const filterableColumns = ['Priority', 'SLA', 'Team', 'Assignment group', 'Created'];
    
    // Get all table headers
    const headers = document.querySelectorAll('.data-table thead tr th');
    
    headers.forEach((th, index) => {
        const columnName = TABLE_COLUMNS[index];
        if (!columnName || !filterableColumns.includes(columnName)) {
            return; // Skip "Number" and non-filterable columns
        }
        
        // Create filter dropdown
        const filterContainer = document.createElement('div');
        filterContainer.className = 'column-filter-container';
        
        const filterSelect = document.createElement('select');
        filterSelect.className = 'column-filter-select';
        filterSelect.setAttribute('data-column', columnName);
        filterSelect.title = `Filter by ${columnName}`;
        
        // Add "All" option
        const allOption = document.createElement('option');
        allOption.value = '';
        allOption.textContent = 'All';
        filterSelect.appendChild(allOption);
        
        // Populate unique values
        populateColumnFilterOptions(filterSelect, columnName);
        
        // Add change event listener
        filterSelect.addEventListener('change', (e) => {
            tableFilterState[columnName] = e.target.value;
            paginationConfig.currentPage = 1; // Reset to first page
            renderTable(filteredTickets);
        });
        
        filterContainer.appendChild(filterSelect);
        th.appendChild(filterContainer);
    });
}

function populateColumnFilterOptions(select, columnName) {
    // Get unique values from filteredTickets (not allTickets, to respect global filters)
    const uniqueValues = [...new Set(filteredTickets.map(t => {
        if (columnName === 'Created') {
            if (!t.Created) return null;
            const date = new Date(t.Created);
            if (isNaN(date.getTime())) return null;
            return date.toISOString().split('T')[0];
        }
        return t[columnName] || null;
    }).filter(v => v !== null && v !== ''))].sort();
    
    uniqueValues.forEach(value => {
        const option = document.createElement('option');
        option.value = value;
        option.textContent = value;
        select.appendChild(option);
    });
}

function refreshTableColumnFilters() {
    // Refresh filter options when data changes
    const filterSelects = document.querySelectorAll('.column-filter-select');
    filterSelects.forEach(select => {
        const columnName = select.getAttribute('data-column');
        const currentValue = select.value;
        
        // Clear options except "All"
        select.innerHTML = '<option value="">All</option>';
        
        // Repopulate
        populateColumnFilterOptions(select, columnName);
        
        // Restore previous selection if still valid
        if (currentValue && Array.from(select.options).some(opt => opt.value === currentValue)) {
            select.value = currentValue;
            tableFilterState[columnName] = currentValue;
        } else {
            select.value = '';
            tableFilterState[columnName] = '';
        }
    });
}

/*************************
 * CLOSED TICKETS PANEL
 *************************/
function updateClosedTicketsPanel(previousTickets, currentTickets) {
    // Compare previousTickets and currentTickets to find closed tickets
    if (!previousTickets || !currentTickets) {
        // If called without parameters, use existing closedTickets array
        renderClosedTicketsList();
        return;
    }
    
    const currentTicketNumbers = new Set(currentTickets.map(t => t.Number));
    
    // Find tickets that were in previousTickets but not in currentTickets
    previousTickets.forEach(prevTicket => {
        if (!currentTicketNumbers.has(prevTicket.Number)) {
            // Check if already tracked
            const alreadyTracked = closedTickets.some(t => t.Number === prevTicket.Number);
            if (!alreadyTracked) {
                closedTickets.push({
                    Number: prevTicket.Number || '',
                    Priority: prevTicket.Priority || '',
                    Team: prevTicket.Team || '',
                    SLA: prevTicket.SLA || '',
                    closedAt: new Date()
                });
            }
        }
    });
    
    // Keep only tickets closed in last 24 hours
    const now = new Date();
    closedTickets = closedTickets.filter(t => {
        const closedDate = new Date(t.closedAt);
        const hoursAgo = (now - closedDate) / (1000 * 60 * 60);
        return hoursAgo <= 24;
    });
    
    // Sort by most recent first and limit to max 10 items
    closedTickets.sort((a, b) => new Date(b.closedAt) - new Date(a.closedAt));
    closedTickets = closedTickets.slice(0, 10);
    
    // Render the list
    renderClosedTicketsList();
}

function renderClosedTicketsList() {
    const list = document.getElementById('closedTicketsList');
    if (!list) return;
    
    if (closedTickets.length === 0) {
        list.innerHTML = '<div class="closed-ticket-empty">No recently closed tickets</div>';
        return;
    }
    
    const now = new Date();
    list.innerHTML = closedTickets.map(ticket => {
        const closedDate = new Date(ticket.closedAt);
        const minutesAgo = Math.floor((now - closedDate) / (1000 * 60));
        const timeAgo = minutesAgo < 1 ? 'Just now' : `${minutesAgo} min ago`;
        
        return `
            <div class="closed-ticket-item">
                <div class="closed-ticket-row">
                    <span class="closed-ticket-number">${escapeHtml(ticket.Number || 'N/A')}</span>
                    <span class="closed-ticket-separator">|</span>
                    <span class="closed-ticket-team">${escapeHtml(ticket.Team || 'N/A')}</span>
                    <span class="closed-ticket-separator">|</span>
                    <span class="closed-ticket-time">Closed ${timeAgo}</span>
                </div>
            </div>
        `;
    }).join('');
}

function toggleClosedTicketsPanel() {
    const panel = document.getElementById('closedTicketsPanel');
    if (!panel) {
        console.error('closedTicketsPanel not found');
        return;
    }
    panel.classList.toggle('open');
}

// =============================
// Recently Closed Tickets
// =============================
async function loadRecentlyClosedTickets() {
    try {
        const res = await fetch("/api/recently-closed");
        const json = await res.json();

        const container = document.getElementById("closedTicketsList");
        if (!container) return;

        container.innerHTML = "";

        if (!json.success || !json.data || json.data.length === 0) {
            container.innerHTML = `<div class="closed-ticket-empty">No recently closed tickets</div>`;
            return;
        }

        json.data.forEach(ticket => {
            const div = document.createElement("div");
            div.className = "closed-ticket-item";

            div.innerHTML = `
                <div class="closed-ticket-number">${ticket.Number || "—"}</div>
                <div class="closed-ticket-time">
                    Closed at: ${ticket.closed_at || ""}
                </div>
            `;

            container.appendChild(div);
        });

    } catch (err) {
        console.error("Failed to load recently closed tickets", err);
    }
}