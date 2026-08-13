/**
 * IceStock Pro — Módulo RRHH (encapsulado)
 * Fuente: Tractchun/Datos/*.xlsx
 */
(function () {
    'use strict';

    const STORAGE_KEY = 'icestock_rrhh_v1';
    const TICKET_YEARS = ['2016', '2017', '2018', '2019', '2020'];

    let state = {
        agentes: { columns: [], rows: [] },
        empleados: { columns: [], rows: [] },
        puestos: { columns: [], rows: [] },
        tickets: {}
    };

    let patches = {
        agentes: { added: [], edited: {}, deleted: [] },
        empleados: { added: [], edited: {}, deleted: [] },
        puestos: { added: [], edited: {}, deleted: [] },
        tickets: {}
    };

    TICKET_YEARS.forEach(y => {
        patches.tickets[y] = { added: [], edited: {}, deleted: [] };
    });

    let activeTab = 'agentes';
    let ticketYear = '2020';
    let filteredRows = [];
    let currentPage = 1;
    let pageSize = 100;
    let editingContext = null;

    const TAB_CONFIG = {
        agentes: {
            label: 'Agentes',
            idField: 'ID Agente',
            newLabel: 'Nuevo Agente',
            filename: 'Agentes'
        },
        empleados: {
            label: 'Lista+Empleados',
            idField: 'ID Empleado',
            newLabel: 'Nuevo Empleado',
            filename: 'Lista_Empleados'
        },
        puestos: {
            label: 'Puestos',
            idField: 'ID Puesto',
            newLabel: 'Nuevo Puesto',
            filename: 'Puestos'
        },
        tickets: {
            label: 'Tickets',
            idField: 'ID Ticket',
            newLabel: 'Nuevo Ticket',
            filename: 'Tickets'
        }
    };

    function deepClone(obj) {
        return JSON.parse(JSON.stringify(obj));
    }

    function loadPatches() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return;
            const saved = JSON.parse(raw);
            if (saved.agentes) patches.agentes = { ...patches.agentes, ...saved.agentes };
            if (saved.empleados) patches.empleados = { ...patches.empleados, ...saved.empleados };
            if (saved.puestos) patches.puestos = { ...patches.puestos, ...saved.puestos };
            if (saved.tickets) {
                TICKET_YEARS.forEach(y => {
                    if (saved.tickets[y]) patches.tickets[y] = { ...patches.tickets[y], ...saved.tickets[y] };
                });
            }
        } catch (e) {
            console.warn('RRHH: no se pudieron cargar cambios guardados', e);
        }
    }

    function savePatches() {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(patches));
    }

    function mergeSection(baseRows, sectionPatch, idField) {
        const deletedSet = new Set(sectionPatch.deleted.map(String));
        const map = new Map();

        baseRows.forEach(row => {
            const id = String(row[idField]);
            if (!deletedSet.has(id)) map.set(id, deepClone(row));
        });

        Object.entries(sectionPatch.edited).forEach(([id, row]) => {
            if (!deletedSet.has(id)) map.set(id, deepClone(row));
        });

        sectionPatch.added.forEach(row => {
            const id = String(row[idField]);
            if (!deletedSet.has(id)) map.set(id, deepClone(row));
        });

        return Array.from(map.values());
    }

    function buildStateFromPreload() {
        const preload = window.RRHH_PRELOAD_DATA;
        if (!preload) {
            console.error('RRHH: datos precargados no encontrados');
            return;
        }

        state.agentes = {
            columns: preload.agentes.columns,
            rows: mergeSection(preload.agentes.rows, patches.agentes, 'ID Agente')
        };

        state.empleados = {
            columns: preload.empleados.columns,
            rows: mergeSection(preload.empleados.rows, patches.empleados, 'ID Empleado')
        };

        state.puestos = {
            columns: preload.puestos.columns,
            rows: mergeSection(preload.puestos.rows, patches.puestos, 'ID Puesto')
        };

        state.tickets = {};
        TICKET_YEARS.forEach(year => {
            const base = preload.tickets[year];
            state.tickets[year] = {
                columns: base.columns,
                rows: mergeSection(base.rows, patches.tickets[year], 'ID Ticket')
            };
        });
    }

    function getCurrentUser() {
        return typeof currentUser !== 'undefined' ? currentUser : null;
    }

    function canViewRRHH() {
        const user = getCurrentUser();
        return user && ['admin', 'supervisor'].includes(user.role);
    }

    function canEditRRHH() {
        const user = getCurrentUser();
        if (!user) return false;
        if (user.role === 'admin') return true;
        if (user.role === 'supervisor') return true;
        return false;
    }

    function canDeleteRRHH() {
        const user = getCurrentUser();
        return user && user.role === 'admin';
    }

    function notify(msg, type) {
        if (typeof showNotification === 'function') showNotification(msg, type);
    }

    function getActiveDataset() {
        if (activeTab === 'tickets') return state.tickets[ticketYear];
        return state[activeTab];
    }

    function getRecordId(row, tab) {
        const cfg = TAB_CONFIG[tab];
        return String(row[cfg.idField]);
    }

    function applyFilters(rows) {
        const q = (document.getElementById('rrhh-filter-search')?.value || '').trim().toLowerCase();
        const dataset = getActiveDataset();
        const columns = dataset.columns;

        let result = rows;

        if (activeTab === 'tickets') {
            const fecha = document.getElementById('rrhh-filter-fecha')?.value || '';
            const agente = document.getElementById('rrhh-filter-agente')?.value || '';
            const categoria = document.getElementById('rrhh-filter-categoria')?.value || '';
            const tipo = document.getElementById('rrhh-filter-tipo')?.value || '';
            const severidad = document.getElementById('rrhh-filter-severidad')?.value || '';
            const prioridad = document.getElementById('rrhh-filter-prioridad')?.value || '';

            if (fecha) {
                result = result.filter(r => String(r['Fecha'] || '').startsWith(fecha));
            }
            if (agente) {
                result = result.filter(r => String(r['ID Agente'] || '') === agente);
            }
            if (categoria) {
                const catCol = columns.find(c => c.startsWith('Categor')) || 'Categoría';
                result = result.filter(r => String(r[catCol] || '') === categoria);
            }
            if (tipo) {
                result = result.filter(r => String(r['Tipo'] || '') === tipo);
            }
            if (severidad) {
                result = result.filter(r => String(r['Severidad'] || '') === severidad);
            }
            if (prioridad) {
                result = result.filter(r => String(r['Prioridad'] || '') === prioridad);
            }
        }

        if (activeTab === 'agentes') {
            const nombre = (document.getElementById('rrhh-filter-nombre')?.value || '').trim().toLowerCase();
            if (nombre) result = result.filter(r => String(r['Nombre'] || '').toLowerCase().includes(nombre));
        }

        if (activeTab === 'empleados') {
            const turno = document.getElementById('rrhh-filter-turno')?.value || '';
            const planta = document.getElementById('rrhh-filter-planta')?.value || '';
            const apellido = (document.getElementById('rrhh-filter-apellido')?.value || '').trim().toLowerCase();
            if (turno) result = result.filter(r => String(r['Turno'] || '') === turno);
            if (planta) result = result.filter(r => String(r._planta || '') === planta);
            if (apellido) result = result.filter(r => String(r['Apellido'] || '').toLowerCase().includes(apellido));
        }

        if (activeTab === 'puestos') {
            const depto = document.getElementById('rrhh-filter-departamento')?.value || '';
            const tipoPuesto = document.getElementById('rrhh-filter-tipo-puesto')?.value || '';
            if (depto) result = result.filter(r => String(r['Departamento'] || '') === depto);
            if (tipoPuesto) result = result.filter(r => String(r['Tipo Puesto'] || '') === tipoPuesto);
        }

        if (q) {
            result = result.filter(row =>
                columns.some(col => String(row[col] ?? '').toLowerCase().includes(q))
            );
        }

        return result;
    }

    function populateFilterOptions() {
        const dataset = getActiveDataset();
        const rows = dataset.rows;

        const setSelectOptions = (id, values, allLabel) => {
            const el = document.getElementById(id);
            if (!el) return;
            const current = el.value;
            el.innerHTML = `<option value="">${allLabel}</option>`;
            values.sort((a, b) => String(a).localeCompare(String(b), 'es')).forEach(v => {
                const opt = document.createElement('option');
                opt.value = v;
                opt.textContent = v;
                el.appendChild(opt);
            });
            if ([...el.options].some(o => o.value === current)) el.value = current;
        };

        const unique = col => [...new Set(rows.map(r => r[col]).filter(v => v != null && v !== ''))];

        document.querySelectorAll('.rrhh-filter-group').forEach(g => g.classList.add('hidden'));
        document.getElementById('rrhh-filters-common')?.classList.remove('hidden');
        document.getElementById('rrhh-filters-tickets')?.classList.toggle('hidden', activeTab !== 'tickets');
        document.getElementById('rrhh-filters-agentes')?.classList.toggle('hidden', activeTab !== 'agentes');
        document.getElementById('rrhh-filters-empleados')?.classList.toggle('hidden', activeTab !== 'empleados');
        document.getElementById('rrhh-filters-puestos')?.classList.toggle('hidden', activeTab !== 'puestos');
        document.getElementById('rrhh-filter-year-wrap')?.classList.toggle('hidden', activeTab !== 'tickets');

        if (activeTab === 'tickets') {
            const catCol = columns.find(c => c.startsWith('Categor')) || 'Categoría';
            setSelectOptions('rrhh-filter-agente', unique('ID Agente'), 'Todos los agentes');
            setSelectOptions('rrhh-filter-categoria', unique(catCol), 'Todas las categorías');
            setSelectOptions('rrhh-filter-tipo', unique('Tipo'), 'Todos los tipos');
            setSelectOptions('rrhh-filter-severidad', unique('Severidad'), 'Todas las severidades');
            setSelectOptions('rrhh-filter-prioridad', unique('Prioridad'), 'Todas las prioridades');
        }
        if (activeTab === 'empleados') {
            setSelectOptions('rrhh-filter-turno', unique('Turno'), 'Todos los turnos');
            setSelectOptions('rrhh-filter-planta', unique('_planta'), 'Todas las plantas');
        }
        if (activeTab === 'puestos') {
            setSelectOptions('rrhh-filter-departamento', unique('Departamento'), 'Todos los departamentos');
            setSelectOptions('rrhh-filter-tipo-puesto', unique('Tipo Puesto'), 'Todos los tipos');
        }
    }

    function renderTable() {
        const dataset = getActiveDataset();
        const tbody = document.getElementById('rrhh-table-body');
        const thead = document.getElementById('rrhh-table-head');
        const countEl = document.getElementById('rrhh-record-count');
        if (!tbody || !thead || !dataset) return;

        filteredRows = applyFilters(dataset.rows);
        const total = filteredRows.length;
        const totalPages = Math.max(1, Math.ceil(total / pageSize));
        if (currentPage > totalPages) currentPage = totalPages;

        const start = (currentPage - 1) * pageSize;
        const pageRows = filteredRows.slice(start, start + pageSize);

        const displayColumns = dataset.columns.filter(c => c !== '_planta');
        thead.innerHTML = `<tr>${displayColumns.map(c => `<th class="text-left py-3 px-4 text-sm font-medium text-gray-500 whitespace-nowrap">${c}</th>`).join('')}${canEditRRHH() ? '<th class="text-center py-3 px-4 text-sm font-medium text-gray-500">Acciones</th>' : ''}</tr>`;

        tbody.innerHTML = pageRows.map(row => {
            const cells = displayColumns.map(c => `<td class="py-3 px-4 text-sm whitespace-nowrap">${row[c] ?? ''}</td>`).join('');
            const actions = canEditRRHH() ? `<td class="py-3 px-4 text-center whitespace-nowrap">
                <button class="text-primary hover:underline rrhh-edit-btn" data-id="${getRecordId(row, activeTab)}" title="Editar"><i class="fas fa-edit"></i></button>
                ${canDeleteRRHH() ? `<button class="text-red-500 hover:underline ml-2 rrhh-delete-btn" data-id="${getRecordId(row, activeTab)}" title="Eliminar"><i class="fas fa-trash"></i></button>` : ''}
            </td>` : '';
            return `<tr class="border-b border-gray-100 hover:bg-gray-50">${cells}${actions}</tr>`;
        }).join('') || `<tr><td colspan="${displayColumns.length + 1}" class="py-6 text-center text-gray-500">Sin registros</td></tr>`;

        if (countEl) {
            countEl.textContent = `${total.toLocaleString('es-PY')} registro(s) — Página ${currentPage} de ${totalPages}`;
        }

        document.getElementById('rrhh-page-info').textContent = `${start + 1}-${Math.min(start + pageSize, total)} de ${total}`;
        document.getElementById('rrhh-prev-page').disabled = currentPage <= 1;
        document.getElementById('rrhh-next-page').disabled = currentPage >= totalPages;
    }

    function exportFilteredCSV() {
        const dataset = getActiveDataset();
        const rows = applyFilters(dataset.rows);
        if (!rows.length) {
            notify('No hay datos para exportar', 'warning');
            return;
        }

        const columns = dataset.columns.filter(c => c !== '_planta');
        const BOM = '\uFEFF';
        const escape = val => {
            if (val == null) return '';
            const s = String(val);
            if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
            return s;
        };

        const lines = [columns.map(escape).join(',')];
        rows.forEach(row => lines.push(columns.map(c => escape(row[c])).join(',')));

        const cfg = TAB_CONFIG[activeTab];
        const suffix = activeTab === 'tickets' ? `_${ticketYear}` : '';
        const filename = `${cfg.filename}${suffix}_${new Date().toISOString().slice(0, 10)}.csv`;

        const blob = new Blob([BOM + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);

        if (typeof logAudit === 'function') {
            logAudit('export', `RRHH CSV exportado: ${cfg.label}${suffix} (${rows.length} registros)`);
        }
        notify(`CSV exportado (${rows.length} registros)`, 'success');
    }

    function openRecordModal(mode, record) {
        const dataset = getActiveDataset();
        const modal = document.getElementById('rrhh-record-modal');
        const form = document.getElementById('rrhh-record-form');
        const title = document.getElementById('rrhh-record-modal-title');
        if (!modal || !form) return;

        editingContext = { mode, tab: activeTab, year: ticketYear, originalId: record ? getRecordId(record, activeTab) : null };

        title.textContent = mode === 'edit'
            ? `Editar ${TAB_CONFIG[activeTab].label}`
            : TAB_CONFIG[activeTab].newLabel;

        const fieldsWrap = document.getElementById('rrhh-form-fields');
        fieldsWrap.innerHTML = '';

        dataset.columns.filter(c => c !== '_planta').forEach(col => {
            const div = document.createElement('div');
            const label = document.createElement('label');
            label.className = 'block text-sm font-medium text-gray-700 mb-1';
            label.textContent = col;
            const input = document.createElement('input');
            input.type = col === 'Fecha' ? 'date' : (col.includes('ID') || col === 'Año' || col === 'Mes' || col === 'Día' || col === 'Satisfacción' || col.includes('Días') ? 'text' : 'text');
            input.name = col;
            input.className = 'w-full border border-gray-300 rounded-lg px-3 py-2 rrhh-form-input';
            input.required = col.includes('ID') || col === 'Nombre';
            if (record && record[col] != null) {
                input.value = col === 'Fecha' && String(record[col]).length >= 10
                    ? String(record[col]).slice(0, 10)
                    : record[col];
            }
            div.appendChild(label);
            div.appendChild(input);
            fieldsWrap.appendChild(div);
        });

        if (activeTab === 'empleados') {
            const div = document.createElement('div');
            const label = document.createElement('label');
            label.className = 'block text-sm font-medium text-gray-700 mb-1';
            label.textContent = 'Planta (hoja origen)';
            const select = document.createElement('select');
            select.name = '_planta';
            select.className = 'w-full border border-gray-300 rounded-lg px-3 py-2';
            ['Planta 1', 'Planta 2'].forEach(p => {
                const opt = document.createElement('option');
                opt.value = p;
                opt.textContent = p;
                select.appendChild(opt);
            });
            select.value = record?._planta || 'Planta 1';
            div.appendChild(label);
            div.appendChild(select);
            fieldsWrap.appendChild(div);
        }

        modal.classList.remove('hidden');
    }

    function closeRecordModal() {
        document.getElementById('rrhh-record-modal')?.classList.add('hidden');
        editingContext = null;
    }

    function persistRecord(formData) {
        const cfg = TAB_CONFIG[activeTab];
        const idField = cfg.idField;
        const record = {};
        const dataset = getActiveDataset();

        dataset.columns.forEach(col => {
            if (col === '_planta') return;
            const val = formData.get(col);
            record[col] = val === '' ? null : val;
        });

        if (activeTab === 'empleados') {
            record._planta = formData.get('_planta') || 'Planta 1';
        }

        const id = String(record[idField] || '');
        if (!id) {
            notify('El identificador es obligatorio', 'error');
            return;
        }

        const patchSection = activeTab === 'tickets' ? patches.tickets[ticketYear] : patches[activeTab];
        const isEdit = editingContext?.mode === 'edit';
        const originalId = isEdit ? String(editingContext.originalId) : null;

        if (!isEdit && dataset.rows.some(r => getRecordId(r, activeTab) === id)) {
            notify('Ya existe un registro con ese identificador', 'error');
            return;
        }

        const preload = window.RRHH_PRELOAD_DATA;
        const baseRows = activeTab === 'tickets' ? preload.tickets[ticketYear].rows : preload[activeTab].rows;
        const inPreload = rid => baseRows.some(r => getRecordId(r, activeTab) === rid);
        const addedIndex = rid => patchSection.added.findIndex(r => getRecordId(r, activeTab) === rid);

        if (isEdit) {
            if (originalId !== id) {
                patchSection.deleted.push(originalId);
                delete patchSection.edited[originalId];
                const addedIdx = addedIndex(originalId);
                if (addedIdx >= 0) patchSection.added.splice(addedIdx, 1);
            }

            if (inPreload(isEdit && originalId !== id ? originalId : id) || patchSection.edited[originalId || id]) {
                patchSection.edited[id] = record;
            } else {
                const idx = addedIndex(originalId || id);
                if (idx >= 0) patchSection.added[idx] = record;
                else patchSection.added.push(record);
            }
        } else {
            patchSection.added.push(record);
        }

        savePatches();
        buildStateFromPreload();
        closeRecordModal();
        populateFilterOptions();
        renderTable();
        notify('Registro guardado correctamente', 'success');

        if (typeof logAudit === 'function') {
            logAudit(isEdit ? 'rrhh_updated' : 'rrhh_created', `RRHH ${cfg.label}: ${id}`);
        }
    }

    function deleteRecord(id) {
        Swal.fire({
            title: '¿Eliminar registro?',
            text: 'Esta acción no se puede deshacer',
            icon: 'warning',
            showCancelButton: true,
            confirmButtonText: 'Sí, eliminar',
            cancelButtonText: 'Cancelar'
        }).then(result => {
            if (!result.isConfirmed) return;

            let patchSection = activeTab === 'tickets' ? patches.tickets[ticketYear] : patches[activeTab];
            const idStr = String(id);

            patchSection.deleted.push(idStr);
            delete patchSection.edited[idStr];
            patchSection.added = patchSection.added.filter(r => getRecordId(r, activeTab) !== idStr);

            savePatches();
            buildStateFromPreload();
            populateFilterOptions();
            renderTable();
            notify('Registro eliminado', 'success');

            if (typeof logAudit === 'function') {
                logAudit('rrhh_deleted', `RRHH ${TAB_CONFIG[activeTab].label}: ${idStr}`);
            }
        });
    }

    function updateNewButtonLabel() {
        const label = document.getElementById('rrhh-new-record-label');
        if (label) label.textContent = TAB_CONFIG[activeTab].newLabel;
    }

    function bindEvents() {
        document.querySelectorAll('.rrhh-tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                activeTab = btn.dataset.rrhhTab;
                currentPage = 1;
                document.querySelectorAll('.rrhh-tab-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                updateNewButtonLabel();
                populateFilterOptions();
                renderTable();
            });
        });

        document.getElementById('rrhh-filter-year')?.addEventListener('change', e => {
            ticketYear = e.target.value;
            currentPage = 1;
            populateFilterOptions();
            renderTable();
        });

        ['rrhh-filter-search', 'rrhh-filter-fecha', 'rrhh-filter-agente', 'rrhh-filter-categoria',
            'rrhh-filter-tipo', 'rrhh-filter-severidad', 'rrhh-filter-prioridad', 'rrhh-filter-nombre',
            'rrhh-filter-turno', 'rrhh-filter-planta', 'rrhh-filter-apellido', 'rrhh-filter-departamento',
            'rrhh-filter-tipo-puesto'].forEach(id => {
            document.getElementById(id)?.addEventListener('input', () => { currentPage = 1; renderTable(); });
            document.getElementById(id)?.addEventListener('change', () => { currentPage = 1; renderTable(); });
        });

        document.getElementById('rrhh-filter-clear')?.addEventListener('click', () => {
            document.querySelectorAll('#section-rrhh select, #section-rrhh input[type=text], #section-rrhh input[type=date]').forEach(el => {
                if (el.id.startsWith('rrhh-filter')) el.value = '';
            });
            currentPage = 1;
            renderTable();
        });

        document.getElementById('rrhh-export-csv-btn')?.addEventListener('click', exportFilteredCSV);

        document.getElementById('rrhh-new-record-btn')?.addEventListener('click', () => {
            if (!canEditRRHH()) return;
            openRecordModal('new');
        });

        document.getElementById('rrhh-close-modal')?.addEventListener('click', closeRecordModal);
        document.getElementById('rrhh-cancel-form')?.addEventListener('click', closeRecordModal);

        document.getElementById('rrhh-record-form')?.addEventListener('submit', e => {
            e.preventDefault();
            if (!canEditRRHH()) return;
            persistRecord(new FormData(e.target));
        });

        document.getElementById('rrhh-table-body')?.addEventListener('click', e => {
            const editBtn = e.target.closest('.rrhh-edit-btn');
            const delBtn = e.target.closest('.rrhh-delete-btn');
            const dataset = getActiveDataset();
            if (editBtn) {
                const row = dataset.rows.find(r => getRecordId(r, activeTab) === editBtn.dataset.id);
                if (row) openRecordModal('edit', row);
            }
            if (delBtn && canDeleteRRHH()) deleteRecord(delBtn.dataset.id);
        });

        document.getElementById('rrhh-prev-page')?.addEventListener('click', () => {
            if (currentPage > 1) { currentPage--; renderTable(); }
        });
        document.getElementById('rrhh-next-page')?.addEventListener('click', () => {
            currentPage++; renderTable();
        });
        document.getElementById('rrhh-page-size')?.addEventListener('change', e => {
            pageSize = parseInt(e.target.value, 10) || 100;
            currentPage = 1;
            renderTable();
        });
    }

    function init() {
        if (!canViewRRHH()) return;

        loadPatches();
        buildStateFromPreload();

        const yearSelect = document.getElementById('rrhh-filter-year');
        if (yearSelect && yearSelect.options.length <= 1) {
            TICKET_YEARS.forEach(y => {
                const opt = document.createElement('option');
                opt.value = y;
                opt.textContent = y;
                yearSelect.appendChild(opt);
            });
            yearSelect.value = ticketYear;
        }

        document.getElementById('rrhh-new-record-btn').style.display = canEditRRHH() ? '' : 'none';
        updateNewButtonLabel();

        populateFilterOptions();
        renderTable();
    }

    bindEvents();
    window.IceStockRRHH = { init, canViewRRHH };
})();
