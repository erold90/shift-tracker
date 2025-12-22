/**
 * Shift Dashboard v5.1 - Con Impostazioni
 */

// State
let data = null;
let allGiornate = [];  // Usa giornate per straordinario corretto
let allLicenze = [];
let manualEntries = {};  // Dati manuali da localStorage
let chartMesi = null;
let currentCalendarMonth = new Date();
let currentServiziFilter = 'tutti';
let currentYear = new Date().getFullYear();
let availableArchives = [];

// LocalStorage keys
const MANUAL_ENTRIES_KEY = 'shift_dashboard_manual_entries';
const SETTINGS_KEY = 'shift_dashboard_settings';

// Default settings
const DEFAULT_SETTINGS = {
    giorniLicenzaAnnuale: 32,
    dataAssunzione: '2020-01-01',
    oreRecuperoIniziali: 0  // Ore non retribuite accumulate
};

// Constants
const MESI = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
const MESI_FULL = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
                   'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
const GIORNI = ['Dom', 'Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab'];

// Abbreviazioni assenze
function getAssenzaAbbr(dettaglio) {
    if (!dettaglio) return 'Riposo';
    const d = dettaglio.toLowerCase();
    if (d.includes('riposo settimanale')) return 'RS';
    if (d.includes('riposo festivita')) return 'RF';
    if (d.includes('recupero') && d.includes('riposo') && d.includes('settimanale')) return 'RRS';
    if (d.includes('recupero') && d.includes('riposo') && d.includes('festivita')) return 'RRF';
    if (d.includes('recupero') && d.includes('ore')) return 'RO';
    if (d.includes('licenza ordinaria')) return 'LO';
    if (d.includes('licenza straordinaria')) return 'LS';
    return 'Riposo';
}

// Abbreviazioni presenze
function getPresenzaAbbr(dettaglio) {
    if (!dettaglio) return 'Servizio';
    const d = dettaglio.toLowerCase();
    if (d.includes('esterno')) return 'Esterno';
    if (d.includes('operativ')) return 'Operativo';
    if (d.includes('istituzional')) return 'Istituz.';
    if (d.includes('pratiche') || d.includes('disbrigo') || d.includes('ufficio')) return 'Ufficio';
    return 'Servizio';
}

// Label breve assenza
function getAssenzaLabel(dettaglio) {
    if (!dettaglio) return 'Riposo';
    const d = dettaglio.toLowerCase();
    if (d.includes('riposo settimanale')) return 'Riposo Settimanale';
    if (d.includes('riposo festivita')) return 'Riposo Festivita\'';
    if (d.includes('recupero') && d.includes('riposo') && d.includes('settimanale')) return 'Rec. Riposo Sett.';
    if (d.includes('recupero') && d.includes('riposo') && d.includes('festivita')) return 'Rec. Riposo Fest.';
    if (d.includes('recupero') && d.includes('ore')) return 'Recupero Ore';
    if (d.includes('licenza ordinaria')) return 'Licenza Ordinaria';
    if (d.includes('licenza straordinaria')) return 'Licenza Straord.';
    return 'Riposo';
}

// Init
document.addEventListener('DOMContentLoaded', init);

async function init() {
    await loadData();
    setupEventListeners();
}

async function loadData(year = null) {
    try {
        let url = 'servizi.json';

        // Se è un anno archiviato, carica il file archivio
        if (year && year !== currentYear && availableArchives.includes(year)) {
            url = `archivio_${year}.json`;
        }

        const response = await fetch(url);
        data = await response.json();

        // Salva lista archivi disponibili
        availableArchives = data.archives || [];
        currentYear = data.anno || new Date().getFullYear();

        // Popola selettore anni
        populateYearSelect();

        // Carica dati manuali da localStorage
        loadManualEntries();

        // Usa giornate per dati corretti con straordinario aggregato
        allGiornate = data.giornate || [];
        allLicenze = data.licenze || [];

        // Merge con dati manuali (solo per anno corrente)
        if (!year || year === new Date().getFullYear()) {
            mergeManualEntries();
        }

        // Sort giornate by date (newest first)
        allGiornate.sort((a, b) => (b.data || '').localeCompare(a.data || ''));

        // Imposta il mese del calendario sull'anno selezionato
        if (year) {
            currentCalendarMonth = new Date(year, new Date().getMonth(), 1);
        }

        renderAll();
    } catch (error) {
        console.error('Error loading data:', error);
    }
}

function populateYearSelect() {
    const select = document.getElementById('yearSelect');
    if (!select) return;

    const thisYear = new Date().getFullYear();
    let options = `<option value="${thisYear}">${thisYear}</option>`;

    // Aggiungi anni archiviati
    availableArchives.forEach(year => {
        if (year !== thisYear) {
            options += `<option value="${year}">${year}</option>`;
        }
    });

    select.innerHTML = options;
    select.value = currentYear;
}

async function changeYear(year) {
    year = parseInt(year);
    if (year === currentYear) return;

    await loadData(year);
}

// === MANUAL ENTRIES MANAGEMENT ===

function loadManualEntries() {
    try {
        const stored = localStorage.getItem(MANUAL_ENTRIES_KEY);
        manualEntries = stored ? JSON.parse(stored) : {};
    } catch (e) {
        console.error('Error loading manual entries:', e);
        manualEntries = {};
    }
}

function saveManualEntries() {
    try {
        localStorage.setItem(MANUAL_ENTRIES_KEY, JSON.stringify(manualEntries));
    } catch (e) {
        console.error('Error saving manual entries:', e);
    }
}

function mergeManualEntries() {
    // Aggiungi le entry manuali a allGiornate
    for (const [dateStr, entry] of Object.entries(manualEntries)) {
        // Controlla se esiste già una giornata per questa data
        const existing = allGiornate.find(g => g.data === dateStr);

        if (!existing) {
            // Aggiungi come nuova giornata
            allGiornate.push(entry);
        }
        // Se esiste già, non sovrascrivere i dati dalle email
    }
}

function isManualEntry(dateStr) {
    return manualEntries.hasOwnProperty(dateStr);
}

function renderAll() {
    updateHeader();
    updateHeroStraordinario();
    updateMiniStats();
    renderProssimiServizi();
    renderCalendarGrid();
    renderChart();
    renderServizi('tutti');
    renderLicenze('pending');
    checkPendingLicenze();
    loadSettings();
}

function setupEventListeners() {
    // Year selector
    document.getElementById('yearSelect')?.addEventListener('change', (e) => {
        changeYear(e.target.value);
    });

    // Tab Navigation
    document.querySelectorAll('.nav-item').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const tabId = e.currentTarget.dataset.tab;
            switchTab(tabId);
        });
    });

    // Calendar navigation
    document.getElementById('prevMonth')?.addEventListener('click', () => {
        currentCalendarMonth.setMonth(currentCalendarMonth.getMonth() - 1);
        renderCalendarGrid();
    });
    document.getElementById('nextMonth')?.addEventListener('click', () => {
        currentCalendarMonth.setMonth(currentCalendarMonth.getMonth() + 1);
        renderCalendarGrid();
    });

    // Servizi filters
    document.querySelectorAll('#serviziFilters .tab').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('#serviziFilters .tab').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            renderServizi(e.target.dataset.filter);
        });
    });

    // Licenze filters
    document.querySelectorAll('#licenzeFilters .tab').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('#licenzeFilters .tab').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            renderLicenze(e.target.dataset.filter);
        });
    });

    // Export CSV
    document.getElementById('btnExport')?.addEventListener('click', exportCSV);

    // Modal
    document.querySelector('.modal-close')?.addEventListener('click', closeModal);
    document.getElementById('modal')?.addEventListener('click', (e) => {
        if (e.target.id === 'modal') closeModal();
    });
}

function switchTab(tabId) {
    // Update nav items
    document.querySelectorAll('.nav-item').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });

    // Update tab panes
    document.querySelectorAll('.tab-pane').forEach(pane => {
        pane.classList.toggle('active', pane.id === tabId);
    });
}

// === RENDER FUNCTIONS ===

function updateHeader() {
    if (data?.last_update) {
        const date = new Date(data.last_update);
        document.getElementById('lastUpdate').textContent =
            'Agg. ' + date.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
    }
}

function updateHeroStraordinario() {
    const now = new Date();
    const currentMonth = now.toISOString().slice(0, 7);
    const monthData = data?.stats?.per_mese?.[currentMonth] || {};

    const straordMese = monthData.ore_straordinario || 0;
    const oreMese = monthData.ore || 0;
    const giorniMese = monthData.giorni || 0;
    const scorteMese = monthData.turnazioni_esterne || 0;

    document.getElementById('straordMese').textContent = formatNumber(straordMese) + 'h';
    document.getElementById('meseLabel').textContent = MESI_FULL[now.getMonth()];
    document.getElementById('oreMese').textContent = formatNumber(oreMese);
    document.getElementById('giorniMese').textContent = giorniMese;
    document.getElementById('scorteMese').textContent = scorteMese;

    // Breakdown straordinario mensile
    const straordDiurno = monthData.straord_diurno || 0;
    const straordNotturno = monthData.straord_notturno || 0;
    const straordFestivoDiurno = monthData.straord_festivo_diurno || 0;
    const straordFestivoNotturno = monthData.straord_festivo_notturno || 0;

    document.getElementById('straordDiurno').textContent = formatNumber(straordDiurno) + 'h';
    document.getElementById('straordNotturno').textContent = formatNumber(straordNotturno) + 'h';
    document.getElementById('straordFestivo').textContent = formatNumber(straordFestivoDiurno) + 'h';
    document.getElementById('straordFestivoNott').textContent = formatNumber(straordFestivoNotturno) + 'h';
}

function updateMiniStats() {
    const stats = data?.stats || {};

    document.getElementById('giorniLavorati').textContent = stats.giorni_lavorati || 0;
    document.getElementById('straordTotale').textContent = formatNumber(stats.ore_straordinario || 0) + 'h';
    document.getElementById('giorniRiposo').textContent = stats.per_tipo?.ASSENZA?.count || 0;

    // Turnazioni esterne totali (scorte)
    document.getElementById('scorteTotale').textContent = stats.turnazioni_esterne || 0;
}

function renderProssimiServizi() {
    const container = document.getElementById('prossimiServizi');
    const today = new Date().toISOString().slice(0, 10);

    // Get future giornate
    const prossimi = allGiornate
        .filter(g => g.data >= today)
        .sort((a, b) => (a.data || '').localeCompare(b.data || ''))
        .slice(0, 5);

    if (prossimi.length === 0) {
        container.innerHTML = '<div class="empty-state">Nessun servizio programmato</div>';
        return;
    }

    container.innerHTML = prossimi.map(g => renderGiornataItem(g)).join('');
}

function renderGiornataItem(g) {
    const date = new Date(g.data);
    const today = new Date().toISOString().slice(0, 10);
    const isToday = g.data === today;

    // Determina tipo principale
    const turno = g.turni?.[0];
    const isAssenza = turno?.tipo === 'ASSENZA';
    const isLicenza = g.is_licenza;

    let tipoBadge = '';
    let tipoBadgeClass = '';
    let orario = '';
    let tipoClass = '';

    if (isLicenza) {
        tipoBadge = 'LO';
        tipoBadgeClass = 'riposo';
        tipoClass = 'riposo';
        orario = 'Licenza';
    } else if (isAssenza) {
        tipoBadge = getAssenzaAbbr(turno?.dettaglio);
        tipoBadgeClass = 'riposo';
        tipoClass = 'riposo';
        orario = getAssenzaLabel(turno?.dettaglio);
    } else if (turno) {
        tipoBadge = getPresenzaAbbr(turno.dettaglio);
        // Trova orario primo e ultimo turno
        const turni = g.turni || [];
        if (turni.length > 0) {
            const sorted = [...turni].sort((a, b) => (a.ora_inizio || '').localeCompare(b.ora_inizio || ''));
            const primo = sorted[0];
            const ultimo = sorted[sorted.length - 1];
            orario = `${primo.ora_inizio} - ${ultimo.ora_fine}`;
        }
    }

    const straordHours = g.ore_straordinario || 0;

    return `
        <div class="giorno-item ${tipoClass}" onclick="showGiornataDetails('${g.data}')">
            <div class="giorno-date">
                <span class="day">${date.getDate()}</span>
                <span class="weekday">${MESI[date.getMonth()]}</span>
            </div>
            <div class="giorno-info">
                <div class="giorno-tipo">
                    <span class="badge-tipo ${tipoBadgeClass}">${tipoBadge}</span>
                </div>
                <div class="giorno-orario">${orario}</div>
            </div>
            ${straordHours > 0 ? `
            <div class="giorno-straord">
                <span class="value">+${formatNumber(straordHours)}h</span>
                <span class="label">straord.</span>
            </div>
            ` : ''}
        </div>
    `;
}

// === CALENDAR GRID ===

function renderCalendarGrid() {
    const container = document.getElementById('calendarDays');
    const titleEl = document.getElementById('calendarMonthTitle');

    const year = currentCalendarMonth.getFullYear();
    const month = currentCalendarMonth.getMonth();

    titleEl.textContent = `${MESI_FULL[month]} ${year}`;

    // Get first day of month and total days
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const daysInMonth = lastDay.getDate();

    // Monday = 0, Sunday = 6 (European week)
    let startDay = firstDay.getDay() - 1;
    if (startDay < 0) startDay = 6;

    // Build giornate map for this month
    const monthStr = `${year}-${String(month + 1).padStart(2, '0')}`;
    const giornateMap = {};
    allGiornate
        .filter(g => g.data && g.data.startsWith(monthStr))
        .forEach(g => {
            giornateMap[g.data] = g;
        });

    // Add licenze approvate
    allLicenze
        .filter(l => l.stato === 'Approvata' && l.data_inizio && l.data_inizio.startsWith(monthStr))
        .forEach(l => {
            if (!giornateMap[l.data_inizio]) {
                giornateMap[l.data_inizio] = { data: l.data_inizio, is_licenza: true, tipo_licenza: l.tipo };
            } else {
                giornateMap[l.data_inizio].is_licenza = true;
            }
        });

    const today = new Date().toISOString().slice(0, 10);

    let html = '';

    // Empty cells before first day
    for (let i = 0; i < startDay; i++) {
        html += '<div class="cal-day empty"></div>';
    }

    // Days of month
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const giornata = giornateMap[dateStr];
        const isToday = dateStr === today;
        const isPast = dateStr < today;
        const isManual = isManualEntry(dateStr);

        let dayClass = '';
        let straordHtml = '';

        if (giornata) {
            if (giornata.is_licenza) {
                dayClass = 'licenza';
            } else if (giornata.turni?.[0]?.tipo === 'ASSENZA') {
                dayClass = 'riposo';
            } else if (giornata.ore_straordinario > 0) {
                dayClass = 'straord';
                straordHtml = `<span class="day-hours">+${formatNumber(giornata.ore_straordinario)}h</span>`;
            } else if (giornata.ore_totali > 0) {
                dayClass = 'servizio';
            }
            // Aggiungi classe manual se è un dato inserito manualmente
            if (isManual) {
                dayClass += ' manual';
            }
        } else if (isPast) {
            // Giorno passato senza dati: mostra come cliccabile per aggiungere
            dayClass = 'empty-add';
        }

        html += `
            <div class="cal-day ${dayClass} ${isToday ? 'today' : ''}" onclick="showDayDetail('${dateStr}')">
                <span class="day-num">${day}</span>
                ${straordHtml}
            </div>
        `;
    }

    container.innerHTML = html;
}

function showDayDetail(dateStr) {
    const giornata = allGiornate.find(g => g.data === dateStr);
    const licenza = allLicenze.find(l => l.stato === 'Approvata' && l.data_inizio === dateStr);
    const today = new Date().toISOString().slice(0, 10);
    const isPast = dateStr <= today;
    const isManual = isManualEntry(dateStr);

    const date = new Date(dateStr);
    document.getElementById('dayDetailTitle').textContent =
        `${date.getDate()} ${MESI_FULL[date.getMonth()]} ${date.getFullYear()}`;

    let content = '';

    // Se non ci sono dati e la data è passata, mostra pulsante per aggiungere
    if (!giornata && !licenza) {
        if (isPast) {
            content = `
                <div class="empty-state">Nessun dato per questo giorno</div>
                <button class="btn-add-day" onclick="openAddModal('${dateStr}')">
                    + Aggiungi Servizio
                </button>
            `;
        } else {
            content = '<div class="empty-state">Nessun servizio programmato</div>';
        }
        document.getElementById('dayDetailContent').innerHTML = content;
        document.getElementById('dayDetail').style.display = 'block';
        return;
    }

    if (licenza) {
        content += `
            <div class="detail-row">
                <span class="detail-label">Licenza</span>
                <span class="detail-value">${formatLicenzaTipo(licenza.tipo)}</span>
            </div>
        `;
    }

    if (giornata) {
        // Filtra solo turni attivi (escludi quelli eliminati/aggiornati)
        const turniAttivi = (giornata.turni || []).filter(t => t.stato !== 'eliminato');
        const isAssenza = turniAttivi[0]?.tipo === 'ASSENZA';

        if (isAssenza) {
            content += `
                <div class="detail-row">
                    <span class="detail-label">Tipo</span>
                    <span class="detail-value">${getAssenzaLabel(turniAttivi[0]?.dettaglio)}</span>
                </div>
            `;
        } else {
            // Mostra solo turni attivi (effettivi)
            turniAttivi.forEach((t, i) => {
                content += `
                    <div class="detail-row">
                        <span class="detail-label">${turniAttivi.length > 1 ? 'Turno ' + (i + 1) : 'Orario'}</span>
                        <span class="detail-value">${t.ora_inizio} - ${t.ora_fine} (${t.durata_ore}h)</span>
                    </div>
                `;
            });

            content += `
                <div class="detail-row">
                    <span class="detail-label">Ore Totali</span>
                    <span class="detail-value">${formatNumber(giornata.ore_totali || 0)}h</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Ore Ordinarie</span>
                    <span class="detail-value">${formatNumber(giornata.ore_ordinarie || 0)}h</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Straordinario</span>
                    <span class="detail-value ${giornata.ore_straordinario > 0 ? 'straord' : ''}">${formatNumber(giornata.ore_straordinario || 0)}h</span>
                </div>
            `;
        }

        // Se è un dato manuale, mostra opzione per eliminarlo
        if (isManual) {
            content += `
                <div class="detail-row">
                    <span class="detail-label">Fonte</span>
                    <span class="detail-value" style="color: var(--success);">Inserito manualmente</span>
                </div>
                <button class="btn-delete" onclick="deleteManualEntry('${dateStr}')">
                    Elimina dato manuale
                </button>
            `;
        }
    }

    document.getElementById('dayDetailContent').innerHTML = content;
    document.getElementById('dayDetail').style.display = 'block';
}

// === SERVIZI TAB ===

function renderServizi(filter = 'tutti') {
    const container = document.getElementById('serviziList');

    let filtered = [...allGiornate];

    if (filter === 'presenza') {
        filtered = filtered.filter(g => g.turni?.[0]?.tipo === 'PRESENZA');
    } else if (filter === 'assenza') {
        filtered = filtered.filter(g => g.turni?.[0]?.tipo === 'ASSENZA' || g.is_licenza);
    }

    // Update count
    document.getElementById('serviziCount').textContent = filtered.length;

    if (filtered.length === 0) {
        container.innerHTML = '<div class="empty-state">Nessun servizio</div>';
        return;
    }

    // Limit for performance
    const displayed = filtered.slice(0, 50);

    container.innerHTML = displayed.map(g => renderGiornataItem(g)).join('');
}

// === LICENZE TAB ===

function renderLicenze(filter = 'pending') {
    const container = document.getElementById('licenzeList');

    // Deduplicate
    const licenzeByKey = {};
    const statoPriority = { 'Approvata': 5, 'Rifiutata': 4, 'Annullata': 3, 'Validata': 2, 'Presentata': 1 };

    allLicenze.forEach(l => {
        if (!l.data_inizio) return;
        const key = `${l.tipo}_${l.data_inizio}`;
        const existing = licenzeByKey[key];
        if (!existing || (statoPriority[l.stato] || 0) > (statoPriority[existing.stato] || 0)) {
            licenzeByKey[key] = l;
        }
    });

    let filtered = Object.values(licenzeByKey);

    if (filter === 'pending') {
        filtered = filtered.filter(l => l.stato === 'Presentata' || l.stato === 'Validata');
    } else if (filter === 'approvate') {
        filtered = filtered.filter(l => l.stato === 'Approvata');
    } else if (filter === 'rifiutate') {
        filtered = filtered.filter(l => l.stato === 'Rifiutata' || l.stato === 'Annullata');
    }

    filtered.sort((a, b) => (b.data_inizio || '').localeCompare(a.data_inizio || ''));
    filtered = filtered.slice(0, 30);

    if (filtered.length === 0) {
        container.innerHTML = '<div class="empty-state">Nessuna licenza</div>';
        return;
    }

    container.innerHTML = filtered.map(l => {
        const icon = getStatoIcon(l.stato);

        // Mostra range date (dal - al) o singola data se sono uguali
        let dateStr = 'N/D';
        if (l.data_inizio) {
            const dataInizio = formatDateShort(l.data_inizio);
            const dataFine = l.data_fine ? formatDateShort(l.data_fine) : dataInizio;

            if (l.data_inizio === l.data_fine || !l.data_fine) {
                dateStr = dataInizio;
            } else {
                dateStr = `${dataInizio} → ${dataFine}`;
            }
        }

        let statusClass = 'pending';
        if (l.stato === 'Approvata') statusClass = 'approved';
        if (l.stato === 'Rifiutata' || l.stato === 'Annullata') statusClass = 'rejected';

        return `
            <div class="licenza-item">
                <span class="licenza-icon">${icon}</span>
                <div class="licenza-info">
                    <div class="licenza-tipo">${formatLicenzaTipo(l.tipo)}</div>
                    <div class="licenza-date">${dateStr}</div>
                </div>
                <span class="licenza-status ${statusClass}">${l.stato}</span>
            </div>
        `;
    }).join('');
}

function checkPendingLicenze() {
    const pendingByDate = {};
    allLicenze.forEach(l => {
        if ((l.stato === 'Presentata' || l.stato === 'Validata') && l.data_inizio) {
            const key = `${l.tipo}_${l.data_inizio}`;
            const hasResolved = allLicenze.some(x =>
                x.tipo === l.tipo &&
                x.data_inizio === l.data_inizio &&
                (x.stato === 'Approvata' || x.stato === 'Annullata' || x.stato === 'Rifiutata')
            );
            if (!hasResolved) {
                pendingByDate[key] = l;
            }
        }
    });

    const pending = Object.keys(pendingByDate).length;
    const alert = document.getElementById('alertLicenze');

    if (pending > 0) {
        document.getElementById('licenzePending').textContent = pending;
        alert.style.display = 'flex';
    } else {
        alert.style.display = 'none';
    }
}

// === CHART ===

function renderChart() {
    const ctx = document.getElementById('chartMesi')?.getContext('2d');
    if (!ctx || !data?.stats?.per_mese) return;

    const months = Object.keys(data.stats.per_mese).sort();
    const labels = months.map(m => MESI[parseInt(m.split('-')[1]) - 1]);
    const ordinarie = months.map(m => data.stats.per_mese[m].ore_ordinarie || 0);
    const straordinario = months.map(m => data.stats.per_mese[m].ore_straordinario || 0);

    if (chartMesi) chartMesi.destroy();

    chartMesi = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: 'Ordinarie',
                    data: ordinarie,
                    backgroundColor: '#3b82f6',
                    borderRadius: 4
                },
                {
                    label: 'Straordinario',
                    data: straordinario,
                    backgroundColor: '#f59e0b',
                    borderRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: { color: '#94a3b8', boxWidth: 12, padding: 10, font: { size: 10 } }
                }
            },
            scales: {
                x: {
                    stacked: true,
                    grid: { display: false },
                    ticks: { color: '#94a3b8', font: { size: 10 } }
                },
                y: {
                    stacked: true,
                    grid: { color: '#334155' },
                    ticks: { color: '#94a3b8', font: { size: 10 } }
                }
            }
        }
    });
}

// === MODAL ===

function showGiornataDetails(dateStr) {
    const giornata = allGiornate.find(g => g.data === dateStr);
    if (!giornata) return;

    const turni = giornata.turni || [];
    const isAssenza = turni[0]?.tipo === 'ASSENZA';
    const date = new Date(dateStr);

    document.getElementById('modalTitle').textContent =
        `${date.getDate()} ${MESI_FULL[date.getMonth()]} ${date.getFullYear()}`;

    let content = '';

    if (isAssenza) {
        content = `
            <div class="modal-row">
                <span class="modal-label">Tipo</span>
                <span class="modal-value">${getAssenzaLabel(turni[0]?.dettaglio)}</span>
            </div>
        `;
    } else {
        // Mostra turni
        turni.forEach((t, i) => {
            content += `
                <div class="modal-row">
                    <span class="modal-label">Turno ${turni.length > 1 ? (i + 1) : ''}</span>
                    <span class="modal-value">${t.ora_inizio} - ${t.ora_fine} (${t.durata_ore}h)</span>
                </div>
            `;
        });

        content += `
            <div class="modal-row">
                <span class="modal-label">Ore Totali</span>
                <span class="modal-value">${formatNumber(giornata.ore_totali || 0)}h</span>
            </div>
            <div class="modal-row">
                <span class="modal-label">Ore Ordinarie</span>
                <span class="modal-value">${formatNumber(giornata.ore_ordinarie || 0)}h</span>
            </div>
            <div class="modal-row">
                <span class="modal-label">Straordinario</span>
                <span class="modal-value" style="color: ${giornata.ore_straordinario > 0 ? '#f59e0b' : 'inherit'}">${formatNumber(giornata.ore_straordinario || 0)}h</span>
            </div>
        `;
    }

    document.getElementById('modalBody').innerHTML = content;
    document.getElementById('modal').classList.add('active');
}

function closeModal() {
    document.getElementById('modal').classList.remove('active');
}

// === UTILITIES ===

function formatNumber(num) {
    return Math.round(num * 10) / 10;
}

function formatDateShort(dateStr) {
    const date = new Date(dateStr);
    return `${date.getDate()} ${MESI[date.getMonth()]}`;
}

function formatLicenzaTipo(tipo) {
    const map = {
        'ordinaria': 'Licenza Ordinaria',
        'straordinaria': 'Licenza Straordinaria',
        'speciale': 'Licenza Speciale',
        'riposo_donatori': 'Riposo Donatori'
    };
    return map[tipo] || tipo;
}

function getStatoIcon(stato) {
    const icons = {
        'Presentata': '📋',
        'Validata': '✅',
        'Approvata': '🎉',
        'Annullata': '❌',
        'Rifiutata': '🚫'
    };
    return icons[stato] || '📋';
}

function exportCSV() {
    let csv = 'Data,Giorno,Tipo,Ore Totali,Ordinarie,Straordinario\n';

    allGiornate.forEach(g => {
        const date = new Date(g.data);
        const giorno = GIORNI[date.getDay()];
        const tipo = g.turni?.[0]?.tipo || 'N/D';
        csv += `${g.data},${giorno},${tipo},${g.ore_totali || 0},${g.ore_ordinarie || 0},${g.ore_straordinario || 0}\n`;
    });

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `servizi_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
}

// === ADD MANUAL ENTRY FUNCTIONS ===

function openAddModal(dateStr) {
    document.getElementById('addData').value = dateStr;

    const date = new Date(dateStr);
    document.getElementById('modalAddTitle').textContent =
        `Aggiungi - ${date.getDate()} ${MESI_FULL[date.getMonth()]} ${date.getFullYear()}`;

    // Reset form
    document.getElementById('addTipo').value = 'PRESENZA';
    document.getElementById('addDettaglio').value = 'Servizio esterno';
    document.getElementById('addAssenza').value = 'Riposo settimanale';
    document.getElementById('addOraInizio').value = '08:00';
    document.getElementById('addOraFine').value = '14:00';

    toggleFormFields();

    document.getElementById('modalAdd').classList.add('active');
}

function closeAddModal() {
    document.getElementById('modalAdd').classList.remove('active');
}

function toggleFormFields() {
    const tipo = document.getElementById('addTipo').value;
    const isAssenza = tipo === 'ASSENZA';

    document.getElementById('groupDettaglio').style.display = isAssenza ? 'none' : 'block';
    document.getElementById('groupAssenza').style.display = isAssenza ? 'block' : 'none';
    document.getElementById('groupOrari').style.display = isAssenza ? 'none' : 'block';

    // Aggiorna required
    document.getElementById('addOraInizio').required = !isAssenza;
    document.getElementById('addOraFine').required = !isAssenza;
}

function saveManualEntry(event) {
    event.preventDefault();

    const dateStr = document.getElementById('addData').value;
    const tipo = document.getElementById('addTipo').value;
    const isAssenza = tipo === 'ASSENZA';

    let dettaglio, oraInizio, oraFine, durataOre, oreOrdinarie, oreStraordinario;

    if (isAssenza) {
        dettaglio = document.getElementById('addAssenza').value;
        oraInizio = '00:00';
        oraFine = '23:59';
        durataOre = 0;
        oreOrdinarie = 0;
        oreStraordinario = 0;
    } else {
        dettaglio = document.getElementById('addDettaglio').value;
        oraInizio = document.getElementById('addOraInizio').value;
        oraFine = document.getElementById('addOraFine').value;

        // Calcola durata
        const [hI, mI] = oraInizio.split(':').map(Number);
        const [hF, mF] = oraFine.split(':').map(Number);
        let minutiInizio = hI * 60 + mI;
        let minutiFine = hF * 60 + mF;

        // Gestisce turni che superano la mezzanotte
        if (minutiFine < minutiInizio) {
            minutiFine += 24 * 60;
        }

        durataOre = (minutiFine - minutiInizio) / 60;
        oreOrdinarie = Math.min(durataOre, 6);
        oreStraordinario = Math.max(0, durataOre - 6);
    }

    // Crea l'oggetto giornata
    const giornata = {
        data: dateStr,
        turni: [{
            id: `${dateStr}_manual`,
            tipo: tipo,
            dettaglio: dettaglio,
            matricola: '000000',  // Placeholder - configurabile
            data: dateStr,
            ora_inizio: oraInizio,
            ora_fine: oraFine,
            durata_ore: Math.round(durataOre * 100) / 100,
            is_straordinario: oreStraordinario > 0,
            ore_ordinarie: Math.round(oreOrdinarie * 100) / 100,
            ore_straordinario: Math.round(oreStraordinario * 100) / 100,
            email_date: new Date().toISOString(),
            email_id: 'manual_entry',
            stato: 'attivo'
        }],
        ore_totali: Math.round(durataOre * 100) / 100,
        ore_ordinarie: Math.round(oreOrdinarie * 100) / 100,
        ore_straordinario: Math.round(oreStraordinario * 100) / 100,
        is_licenza: false,
        tipo_licenza: '',
        isManual: true
    };

    // Salva in localStorage
    manualEntries[dateStr] = giornata;
    saveManualEntries();

    // Ricarica dati e aggiorna UI
    closeAddModal();
    loadData();
}

function deleteManualEntry(dateStr) {
    if (confirm('Eliminare questo dato inserito manualmente?')) {
        delete manualEntries[dateStr];
        saveManualEntries();

        // Ricarica dati e aggiorna UI
        document.getElementById('dayDetail').style.display = 'none';
        loadData();
    }
}

// === SETTINGS FUNCTIONS ===

function loadSettings() {
    try {
        const stored = localStorage.getItem(SETTINGS_KEY);
        const settings = stored ? { ...DEFAULT_SETTINGS, ...JSON.parse(stored) } : DEFAULT_SETTINGS;

        // Popola i campi del form
        document.getElementById('giorniLicenzaAnnuale').value = settings.giorniLicenzaAnnuale;
        document.getElementById('dataAssunzione').value = settings.dataAssunzione;
        document.getElementById('oreRecuperoIniziali').value = settings.oreRecuperoIniziali || 0;

        // Calcola e mostra le informazioni
        updateSettingsInfo(settings);

        return settings;
    } catch (e) {
        console.error('Error loading settings:', e);
        return DEFAULT_SETTINGS;
    }
}

function saveSettings() {
    try {
        const settings = {
            giorniLicenzaAnnuale: parseInt(document.getElementById('giorniLicenzaAnnuale').value) || 32,
            dataAssunzione: document.getElementById('dataAssunzione').value || '2020-01-01',
            oreRecuperoIniziali: parseFloat(document.getElementById('oreRecuperoIniziali').value) || 0
        };

        localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
        updateSettingsInfo(settings);

        // Feedback visivo
        const btn = document.querySelector('.btn-save-settings');
        const originalText = btn.textContent;
        btn.textContent = 'Salvato!';
        btn.style.background = '#22c55e';
        setTimeout(() => {
            btn.textContent = originalText;
            btn.style.background = '';
        }, 1500);

    } catch (e) {
        console.error('Error saving settings:', e);
        alert('Errore durante il salvataggio');
    }
}

function updateSettingsInfo(settings) {
    const thisYear = new Date().getFullYear();

    // Calcola giorni di licenza usati quest'anno
    let giorniUsati = 0;

    // Conta le licenze ordinarie approvate per l'anno corrente
    // IMPORTANTE: deduplica per data_inizio (ci possono essere più record per la stessa licenza)
    if (allLicenze && allLicenze.length > 0) {
        const licenzeByDate = {};

        allLicenze
            .filter(l => l.tipo === 'ordinaria' && l.stato === 'Approvata' && l.data_inizio?.startsWith(String(thisYear)))
            .forEach(l => {
                // Usa data_inizio come chiave per deduplicare
                if (!licenzeByDate[l.data_inizio]) {
                    licenzeByDate[l.data_inizio] = l;
                }
            });

        // Ora conta i giorni dalle licenze uniche
        Object.values(licenzeByDate).forEach(l => {
            if (l.data_inizio && l.data_fine) {
                const start = new Date(l.data_inizio);
                const end = new Date(l.data_fine);
                const diffTime = Math.abs(end - start);
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
                giorniUsati += diffDays;
            } else if (l.data_inizio) {
                giorniUsati += 1;
            }
        });
    }

    const giorniRimanenti = settings.giorniLicenzaAnnuale - giorniUsati;

    // Aggiorna UI
    document.getElementById('giorniUsati').textContent = giorniUsati;
    document.getElementById('giorniRimanenti').textContent = giorniRimanenti;

    // Colore in base ai giorni rimanenti
    const rimanentiEl = document.getElementById('giorniRimanenti');
    if (giorniRimanenti <= 5) {
        rimanentiEl.style.color = '#ef4444';
    } else if (giorniRimanenti <= 10) {
        rimanentiEl.style.color = '#f59e0b';
    } else {
        rimanentiEl.style.color = '#22c55e';
    }

    // Calcola anni di servizio
    const assunzione = new Date(settings.dataAssunzione);
    const oggi = new Date();
    let anniServizio = oggi.getFullYear() - assunzione.getFullYear();

    // Aggiusta se non ha ancora raggiunto l'anniversario quest'anno
    if (oggi.getMonth() < assunzione.getMonth() ||
        (oggi.getMonth() === assunzione.getMonth() && oggi.getDate() < assunzione.getDate())) {
        anniServizio--;
    }

    document.getElementById('anniServizio').textContent = anniServizio;

    // Calcola ore a recupero
    const stats = data?.stats || {};
    const oreRecuperoIniziali = settings.oreRecuperoIniziali || 0;
    const oreRecuperateNonRetribuite = stats.ore_recuperate_non_retribuite || 0;
    const oreRecuperoRimanenti = Math.max(0, oreRecuperoIniziali - oreRecuperateNonRetribuite);

    document.getElementById('oreRecuperoUsate').textContent = oreRecuperateNonRetribuite + 'h';
    document.getElementById('oreRecuperoRimanenti').textContent = oreRecuperoRimanenti + 'h';

    // Colore in base alle ore rimanenti
    const oreRimanentiEl = document.getElementById('oreRecuperoRimanenti');
    if (oreRecuperoRimanenti <= 0) {
        oreRimanentiEl.style.color = '#22c55e';
    } else if (oreRecuperoRimanenti <= 12) {
        oreRimanentiEl.style.color = '#f59e0b';
    } else {
        oreRimanentiEl.style.color = '#ef4444';
    }

    // Nota licenza (dal 2026 conteggio attivo)
    const noteEl = document.getElementById('licenzaNote');
    if (thisYear >= 2026) {
        noteEl.textContent = `Conteggio basato su ${settings.giorniLicenzaAnnuale} giorni annuali.`;
    } else {
        noteEl.textContent = 'Il calcolo dei giorni rimanenti sarà attivo dal 2026.';
    }

    // Info app
    document.getElementById('lastSyncInfo').textContent = data?.last_update
        ? new Date(data.last_update).toLocaleString('it-IT')
        : '--';
    document.getElementById('giornateCount').textContent = allGiornate?.length || 0;
}

function clearCacheAndReload() {
    // Svuota la cache del browser per i dati
    if ('caches' in window) {
        caches.keys().then(names => {
            names.forEach(name => caches.delete(name));
        });
    }

    // Forza il ricaricamento senza cache
    window.location.reload(true);
}

function clearAllData() {
    if (confirm('Sei sicuro di voler eliminare tutti i dati inseriti manualmente?')) {
        localStorage.removeItem(MANUAL_ENTRIES_KEY);
        manualEntries = {};
        loadData();
        alert('Dati manuali eliminati');
    }
}

// Global functions for onclick handlers
window.showGiornataDetails = showGiornataDetails;
window.showDayDetail = showDayDetail;
window.openAddModal = openAddModal;
window.closeAddModal = closeAddModal;
window.toggleFormFields = toggleFormFields;
window.saveManualEntry = saveManualEntry;
window.deleteManualEntry = deleteManualEntry;
window.saveSettings = saveSettings;
window.clearCacheAndReload = clearCacheAndReload;
window.clearAllData = clearAllData;
