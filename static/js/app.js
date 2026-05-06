// State Management
const state = {
    user: null,
    currentSection: 'dashboard',
    records: [],
    fields: [],
    pagination: {
        total: 0,
        page: 1,
        pages: 0,
        perPage: 20
    },
    searchQuery: '',
    isLoading: false
};

// --- API Service ---
const api = {
    async fetch(url, options = {}) {
        try {
            const response = await fetch(url, {
                ...options,
                headers: {
                    'Content-Type': 'application/json',
                    ...options.headers
                }
            });
            if (response.status === 401 && !url.includes('/login')) {
                logout();
                return;
            }
            const data = await response.json();
            if (!response.ok) throw new Error(data.error || 'API Error');
            return data;
        } catch (error) {
            toast(error.message, 'error');
            throw error;
        }
    },
    login: (username, password) => api.fetch('/api/login', { method: 'POST', body: JSON.stringify({ username, password }) }),
    logout: () => api.fetch('/api/logout', { method: 'POST' }),
    getMe: () => api.fetch('/api/me'),
    getRecords: (search = '') => api.fetch(`/api/records?search=${encodeURIComponent(search)}`),
    saveRecord: (data, id = null) => api.fetch(id ? `/api/records/${id}` : '/api/records', { method: id ? 'PUT' : 'POST', body: JSON.stringify(data) }),
    deleteRecord: (id) => api.fetch(`/api/records/${id}`, { method: 'DELETE' }),
    updateRecordStatus: (id, active) => api.fetch(`/api/records/${id}/status`, { method: 'PUT', body: JSON.stringify({ active }) }),
    getFields: () => api.fetch('/api/fields'),
    saveField: (data, id = null) => api.fetch(id ? `/api/fields/${id}` : '/api/fields', { method: id ? 'PUT' : 'POST', body: JSON.stringify(data) }),
    deleteField: (id) => api.fetch(`/api/fields/${id}`, { method: 'DELETE' }),
    getConfig: (type) => api.fetch(`/api/config/${type}`),
    saveConfig: (type, data) => api.fetch(`/api/config/${type}`, { method: 'POST', body: JSON.stringify(data) }),
    getMapping: (type) => api.fetch(`/api/config/${type}/mapping`),
    updateMapping: (type, id, data) => api.fetch(`/api/config/${type}/mapping/${id}`, { method: 'PUT', body: JSON.stringify(data) }),
    getUsers: () => api.fetch('/api/users'),
    saveUser: (data) => api.fetch('/api/users', { method: 'POST', body: JSON.stringify(data) }),
    updateUserStatus: (id, active) => api.fetch(`/api/users/${id}/status`, { method: 'PUT', body: JSON.stringify({ active }) }),
    updateUserRole: (id, role) => api.fetch(`/api/users/${id}/role`, { method: 'PUT', body: JSON.stringify({ role }) }),
    resetPassword: (id, password) => api.fetch(`/api/users/${id}/reset-password`, { method: 'POST', body: JSON.stringify({ password }) }),
    getStats: () => api.fetch('/api/stats'),
    getAuditLogs: () => api.fetch('/api/audit-logs'),
    getRecords: (search = '', page = 1, perPage = 20) => api.fetch(`/api/records?search=${encodeURIComponent(search)}&page=${page}&per_page=${perPage}`),
    getAttachments: (recordId) => api.fetch(`/api/records/${recordId}/attachments`),
    uploadAttachment: (recordId, file) => {
        const fd = new FormData();
        fd.append('file', file);
        return api.fetch(`/api/records/${recordId}/attachments`, { method: 'POST', body: fd, headers: { 'Content-Type': undefined } });
    },
    deleteAttachment: (id) => api.fetch(`/api/attachments/${id}`, { method: 'DELETE' })
};

// --- UI Utilities ---
function toast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const t = document.createElement('div');
    t.className = `toast ${type}`;
    t.textContent = message;
    container.appendChild(t);
    setTimeout(() => {
        t.style.opacity = '0';
        setTimeout(() => t.remove(), 300);
    }, 3000);
}

function showModal(title, bodyHtml, footerHtml = '') {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-body').innerHTML = bodyHtml;
    document.getElementById('modal-footer').innerHTML = footerHtml;
    document.getElementById('modal-container').classList.add('active');
    lucide.createIcons();
}

function hideModal() {
    document.getElementById('modal-container').classList.remove('active');
}

// --- Navigation ---
function switchSection(sectionId) {
    state.currentSection = sectionId;
    document.querySelectorAll('.sidebar-nav li').forEach(li => {
        li.classList.toggle('active', li.dataset.section === sectionId);
    });
    
    const titles = {
        'dashboard': 'Panel Principal',
        'records': 'Registros Documentales',
        'generate-fuid': 'Generar FUID',
        'generate-rotulos': 'Generar Rótulos',
        'import': 'Importar desde Excel',
        'config': 'Configuración del Sistema',
        'audit': 'Registro de Actividad (Auditoría)',
        'users': 'Gestión de Usuarios'
    };
    
    document.getElementById('section-title').textContent = titles[sectionId] || 'Panel de Control';
    renderContent();
}

// --- Renderers ---
async function renderContent() {
    const area = document.getElementById('content-area');
    area.innerHTML = '<div class="loader">Cargando...</div>';
    
    switch (state.currentSection) {
        case 'dashboard':
            await renderDashboard();
            break;
        case 'records':
            await renderRecords();
            break;
        case 'generate-fuid':
            renderGenerateFUID();
            break;
        case 'generate-rotulos':
            renderGenerateRotulos();
            break;
        case 'import':
            renderImport();
            break;
        case 'config':
            await renderConfig();
            break;
        case 'audit':
            await renderAudit();
            break;
        case 'users':
            await renderUsers();
            break;
    }
    lucide.createIcons();
}

async function renderDashboard() {
    const stats = await api.getStats();
    
    let html = `
        <div class="dashboard-grid">
            <div class="stats-card glass animate-up">
                <div class="stats-icon" style="background: rgba(59, 130, 246, 0.1); color: #3b82f6;">
                    <i data-lucide="database"></i>
                </div>
                <div class="stats-info">
                    <span class="stats-label">Total Registros</span>
                    <h3 class="stats-value">${stats.total_records}</h3>
                </div>
            </div>
            
            <div class="stats-card glass animate-up" style="animation-delay: 0.1s">
                <div class="stats-icon" style="background: rgba(16, 185, 129, 0.1); color: #10b981;">
                    <i data-lucide="plus-circle"></i>
                </div>
                <div class="stats-info">
                    <span class="stats-label">Nuevos (7 días)</span>
                    <h3 class="stats-value">${stats.records_week}</h3>
                </div>
            </div>
            
            <div class="stats-card glass animate-up" style="animation-delay: 0.2s">
                <div class="stats-icon" style="background: rgba(245, 158, 11, 0.1); color: #f59e0b;">
                    <i data-lucide="users"></i>
                </div>
                <div class="stats-info">
                    <span class="stats-label">Usuarios</span>
                    <h3 class="stats-value">${stats.total_users}</h3>
                </div>
            </div>
            
            <div class="stats-card glass animate-up" style="animation-delay: 0.3s">
                <div class="stats-icon" style="background: rgba(139, 92, 246, 0.1); color: #8b5cf6;">
                    <i data-lucide="trending-up"></i>
                </div>
                <div class="stats-info">
                    <span class="stats-label">Más Activo</span>
                    <h3 class="stats-value" style="font-size: 1.2rem">${stats.active_user}</h3>
                </div>
            </div>
        </div>

        <div class="dashboard-footer" style="margin-top: 2rem; display: grid; grid-template-columns: 2fr 1fr; gap: 2rem;">
            <div class="recent-activity-card glass">
                <h3>Tendencia de Registros (6 Meses)</h3>
                <div style="height: 250px; width: 100%;">
                    <canvas id="recordsChart"></canvas>
                </div>
                <h3 style="margin-top: 1.5rem">Actividad Reciente</h3>
                <div class="activity-list">
                    ${stats.recent_activity.map(a => `
                        <div class="activity-item">
                            <div class="activity-time">${a.created_at.split(' ')[1]}</div>
                            <div class="activity-desc">
                                <strong>${a.username}</strong> ${a.action.replace(/_/g, ' ')} en ${a.module}
                                <small>${a.details || ''}</small>
                            </div>
                        </div>
                    `).join('')}
                    ${stats.recent_activity.length === 0 ? '<p>No hay actividad reciente</p>' : ''}
                </div>
            </div>
            <div class="welcome-card glass" style="display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; padding: 2rem;">
                <div class="welcome-img" style="font-size: 4rem; margin-bottom: 1rem;">📂</div>
                <h3>¡Hola, ${state.user.full_name}!</h3>
                <p>Bienvenido al Sistema de Gestión Documental. ¿Qué deseas hacer hoy?</p>
                <button class="btn btn-primary" style="margin-top: 1rem;" onclick="switchSection('records')">Ir a Registros</button>
            </div>
        </div>
    `;
    document.getElementById('content-area').innerHTML = html;
    lucide.createIcons();

    // Init Chart
    const chartEl = document.getElementById('recordsChart');
    if (chartEl) {
        new Chart(chartEl.getContext('2d'), {
            type: 'line',
            data: {
                labels: stats.chart_data.map(d => d.month),
                datasets: [{
                    label: 'Registros',
                    data: stats.chart_data.map(d => d.total),
                    borderColor: '#2563eb',
                    backgroundColor: 'rgba(37, 99, 235, 0.1)',
                    fill: true,
                    tension: 0.4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: {
                    y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' } },
                    x: { grid: { display: false } }
                }
            }
        });
    }
}

async function renderAudit() {
    const logs = await api.getAuditLogs();
    let html = `
        <div class="data-table-container">
            <table>
                <thead>
                    <tr>
                        <th>Fecha</th>
                        <th>Usuario</th>
                        <th>Acción</th>
                        <th>Módulo</th>
                        <th>ID Reg.</th>
                        <th>Detalles</th>
                    </tr>
                </thead>
                <tbody>
                    ${logs.map(log => `
                        <tr>
                            <td style="white-space: nowrap">${log.created_at}</td>
                            <td><strong>${log.username}</strong></td>
                            <td><span class="badge badge-${log.action.includes('delete') ? 'danger' : 'info'}">${log.action}</span></td>
                            <td>${log.module}</td>
                            <td>${log.record_id || '-'}</td>
                            <td style="font-size: 0.85rem">${log.details || ''}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    document.getElementById('content-area').innerHTML = html;
}

async function renderRecords() {
    const data = await api.getRecords(state.searchQuery, state.pagination.page, state.pagination.perPage);
    state.records = data.records;
    state.pagination.total = data.total;
    state.pagination.pages = data.pages;
    
    const fields = await api.getFields();
    state.fields = fields;
    
    const visibleFields = fields.filter(f => f.visible);
    
    let html = `
        <div class="data-table-container">
            <table>
                <thead>
                    <tr>
                        <th style="width: 100px; position: sticky; left: 0; background: var(--bg-card); z-index: 10;">Acciones</th>
                        <th style="width: 100px;">Estado</th>
                        ${visibleFields.map(f => `<th>${f.display_name}</th>`).join('')}
                        <th>Anexos</th>
                    </tr>
                </thead>
                <tbody>
                    ${state.records.map(rec => `
                        <tr>
                            <td class="actions-cell" style="position: sticky; left: 0; background: var(--bg-card); z-index: 5; border-right: 2px solid var(--border-color);">
                                <button class="btn-icon" title="Editar" onclick="editRecord(${rec.id})"><i data-lucide="edit-2"></i></button>
                                <button class="btn-icon" title="${rec.active ? 'Inactivar' : 'Activar'}" onclick="toggleRecordStatus(${rec.id}, ${rec.active ? 0 : 1})">
                                    <i data-lucide="${rec.active ? 'toggle-right' : 'toggle-left'}" style="color: ${rec.active ? 'var(--primary)' : 'var(--text-muted)'}"></i>
                                </button>
                                <button class="btn-icon" title="Eliminar" onclick="deleteRecord(${rec.id})"><i data-lucide="trash-2"></i></button>
                            </td>
                            <td>
                                <span class="badge badge-${rec.active ? 'success' : 'danger'}">
                                    ${rec.active ? 'Activo' : 'Inactivo'}
                                </span>
                            </td>
                            ${visibleFields.map(f => `<td>${rec[f.column_name] || ''}</td>`).join('')}
                            <td>
                                <div class="badge badge-info" style="cursor:pointer" onclick="openAttachmentsModal(${rec.id})">
                                    <i data-lucide="paperclip" style="width:12px; height:12px"></i> ${rec.attachments_count}
                                </div>
                            </td>
                        </tr>
                    `).join('')}
                    ${state.records.length === 0 ? '<tr><td colspan="100%" style="text-align:center; padding:2rem; color:var(--text-muted)">No se encontraron registros</td></tr>' : ''}
                </tbody>
            </table>
        </div>
        ${renderPagination()}
    `;
    document.getElementById('content-area').innerHTML = html;
    lucide.createIcons();
}

function renderPagination() {
    if (state.pagination.pages <= 1) return '';
    
    let html = `
        <div class="pagination">
            <button class="btn btn-sm" ${state.pagination.page === 1 ? 'disabled' : ''} onclick="changePage(${state.pagination.page - 1})">
                <i data-lucide="chevron-left"></i>
            </button>
            <span class="page-info">Página ${state.pagination.page} de ${state.pagination.pages}</span>
            <button class="btn btn-sm" ${state.pagination.page === state.pagination.pages ? 'disabled' : ''} onclick="changePage(${state.pagination.page + 1})">
                <i data-lucide="chevron-right"></i>
            </button>
        </div>
    `;
    return html;
}

window.changePage = (p) => {
    state.pagination.page = p;
    renderRecords();
};

function renderImport() {
    document.getElementById('content-area').innerHTML = `
        <div class="import-section">
            <div class="import-card glass">
                <h3>Migración desde Excel</h3>
                <p>Sube un archivo Excel (.xlsx) con una hoja llamada <strong>'Datos'</strong>. Las columnas nuevas se crearán automáticamente.</p>
                <div class="import-zone" id="drop-zone">
                    <i data-lucide="upload-cloud"></i>
                    <p>Haz clic para seleccionar o arrastra un archivo aquí</p>
                    <input type="file" id="file-input" accept=".xlsx" style="display:none">
                </div>
                <div id="import-result" class="import-result"></div>
            </div>
        </div>
    `;
    
    const zone = document.getElementById('drop-zone');
    const input = document.getElementById('file-input');
    zone.onclick = () => input.click();
    input.onchange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        const formData = new FormData();
        formData.append('file', file);
        
        const res = await api.fetch('/api/import', { method: 'POST', body: formData, headers: {'Content-Type': undefined} });
        if (res) {
            document.getElementById('import-result').innerHTML = `
                <div class="success-alert">
                    Importación completada: ${res.imported} registros cargados, ${res.created_fields} campos nuevos creados.
                </div>
            `;
            toast('Excel importado correctamente');
        }
    };
}

async function renderConfig() {
    const fields = await api.getFields();
    let html = `
        <div class="config-tabs">
            <div class="tabs-header">
                <button class="tab-btn active" data-tab="fields">Campos</button>
                <button class="tab-btn" data-tab="fuid">Configuración FUID</button>
                <button class="tab-btn" data-tab="rotulos">Configuración Rótulos</button>
            </div>
            <div class="tab-content" id="config-tab-content">
                <!-- Tab content dynamic -->
            </div>
        </div>
    `;
    document.getElementById('content-area').innerHTML = html;
    renderFieldsConfig(fields);
    
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.onclick = async () => {
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            const tab = btn.dataset.tab;
            if (tab === 'fields') renderFieldsConfig(await api.getFields());
            else if (tab === 'fuid') renderFuidConfig();
            else if (tab === 'rotulos') renderRotulosConfig();
        };
    });
}

function renderFieldsConfig(fields) {
    let html = `
        <div class="actions-bar" style="margin-bottom:1rem; display:flex; justify-content:flex-end">
            <button class="btn btn-primary" onclick="openFieldModal()"><i data-lucide="plus"></i> Nuevo Campo</button>
        </div>
        <div class="data-table-container">
            <table>
                <thead>
                    <tr>
                        <th>Nombre Columna</th>
                        <th>Etiqueta</th>
                        <th>Visible</th>
                        <th>Orden</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    ${fields.map(f => `
                        <tr>
                            <td>${f.column_name}</td>
                            <td>${f.display_name}</td>
                            <td>${f.visible ? 'Sí' : 'No'}</td>
                            <td>${f.display_order}</td>
                            <td class="actions-cell">
                                <button class="btn-icon" onclick="openFieldModal(${f.id})"><i data-lucide="edit-2"></i></button>
                                <button class="btn-icon" onclick="deleteField(${f.id})"><i data-lucide="trash-2"></i></button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    document.getElementById('config-tab-content').innerHTML = html;
    lucide.createIcons();
}

// --- FUID Generation ---
function renderGenerateFUID() {
    document.getElementById('content-area').innerHTML = `
        <div class="gen-card glass">
            <div class="gen-info">
                <h3>Generación de FUID</h3>
                <p>Se generará el Formato Único de Inventario Documental basado en los registros actuales.</p>
            </div>
            <div class="gen-form">
                <div class="form-group">
                    <label>Filtro de búsqueda (Opcional)</label>
                    <input type="text" id="fuid-search" placeholder="Filtro de registros..." class="form-control">
                </div>
                <div class="form-group">
                    <label>Formato</label>
                    <select id="fuid-format" class="form-control">
                        <option value="Excel">Excel (.xlsx)</option>
                        <option value="Word">Word (.docx)</option>
                        <option value="PDF">PDF (.pdf)</option>
                    </select>
                </div>
                <button class="btn btn-primary" onclick="generateFUID()">
                    <i data-lucide="download"></i> Generar FUID
                </button>
            </div>
        </div>
    `;
}

window.generateFUID = () => {
    const search = document.getElementById('fuid-search').value;
    const format = document.getElementById('fuid-format').value;
    window.location.href = `/api/generate/fuid?search=${encodeURIComponent(search)}&format=${format}`;
};

// --- Rótulos Generation ---
function renderGenerateRotulos() {
    document.getElementById('content-area').innerHTML = `
        <div class="gen-grid">
            <div class="gen-card glass">
                <h3>Rótulo de Carpeta</h3>
                <div class="form-group">
                    <label>Desde (Número de orden)</label>
                    <input type="text" id="rc-desde" class="form-control">
                </div>
                <div class="form-group">
                    <label>Hasta (Número de orden)</label>
                    <input type="text" id="rc-hasta" class="form-control">
                </div>
                <div class="form-group">
                    <label>Formato</label>
                    <select id="rc-format" class="form-control">
                        <option value="Word">Word (.docx)</option>
                        <option value="Excel">Excel (.xlsx)</option>
                        <option value="PDF">PDF (.pdf)</option>
                    </select>
                </div>
                <button class="btn btn-primary" onclick="generateRC()">Generar Rótulo</button>
            </div>
            
            <div class="gen-card glass">
                <h3>Rótulo de Caja</h3>
                <div class="form-group">
                    <label>Caja (Número)</label>
                    <input type="text" id="rj-caja" class="form-control">
                </div>
                <div class="form-group">
                    <label>Desde (Orden)</label>
                    <input type="text" id="rj-desde" class="form-control">
                </div>
                <div class="form-group">
                    <label>Hasta (Orden)</label>
                    <input type="text" id="rj-hasta" class="form-control">
                </div>
                <div class="form-group">
                    <label>Formato</label>
                    <select id="rj-format" class="form-control">
                        <option value="Word">Word (.docx)</option>
                        <option value="Excel">Excel (.xlsx)</option>
                        <option value="PDF">PDF (.pdf)</option>
                    </select>
                </div>
                <button class="btn btn-primary" onclick="generateRJ()">Generar Rótulo</button>
            </div>
        </div>
    `;
}

window.generateRC = () => {
    const d = document.getElementById('rc-desde').value;
    const h = document.getElementById('rc-hasta').value;
    const f = document.getElementById('rc-format').value;
    window.location.href = `/api/generate/rotulo-carpeta?desde=${d}&hasta=${h}&format=${f}`;
};

window.generateRJ = () => {
    const c = document.getElementById('rj-caja').value;
    const d = document.getElementById('rj-desde').value;
    const h = document.getElementById('rj-hasta').value;
    const f = document.getElementById('rj-format').value;
    window.location.href = `/api/generate/rotulo-caja?caja=${c}&desde=${d}&hasta=${h}&format=${f}`;
};

// --- Auth logic ---
async function login(e) {
    e.preventDefault();
    const user = document.getElementById('username').value;
    const pass = document.getElementById('password').value;
    try {
        const res = await api.login(user, pass);
        state.user = res.user;
        initApp();
    } catch (err) {}
}

async function logout() {
    await api.logout();
    state.user = null;
    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
    document.getElementById('login-screen').classList.add('active');
    document.getElementById('main-screen').classList.remove('active');
}

function initApp() {
    document.getElementById('login-screen').classList.remove('active');
    document.getElementById('main-screen').classList.add('active');
    document.getElementById('user-fullname').textContent = state.user.full_name;
    document.getElementById('user-role').textContent = state.user.role;
    document.getElementById('user-initials').textContent = state.user.username.substring(0, 2).toUpperCase();
    
    if (state.user.role === 'admin') {
        document.body.classList.add('is-admin');
    } else {
        document.body.classList.remove('is-admin');
    }
    
    switchSection('records');
}

window.openAttachmentsModal = async (recordId) => {
    const attachments = await api.getAttachments(recordId);
    let html = `
        <div class="attachments-list">
            ${attachments.map(a => `
                <div class="attachment-item glass" style="display:flex; justify-content:space-between; align-items:center; padding:0.75rem; margin-bottom:0.5rem; border-radius:var(--radius-md)">
                    <span>${a.filename}</span>
                    <div class="attachment-actions">
                        <a href="/api/attachments/${a.id}/download" class="btn-icon" title="Descargar"><i data-lucide="download"></i></a>
                        <button class="btn-icon" title="Eliminar" onclick="deleteAttachment(${recordId}, ${a.id})"><i data-lucide="trash-2"></i></button>
                    </div>
                </div>
            `).join('')}
            ${attachments.length === 0 ? '<p style="text-align:center; padding:1rem; color:var(--text-muted)">No hay anexos</p>' : ''}
        </div>
        <div style="margin-top:1.5rem; border-top:1px solid var(--border-color); padding-top:1rem">
            <input type="file" id="attach-file" style="display:none" onchange="uploadAttachment(${recordId})">
            <button class="btn btn-primary btn-block" onclick="document.getElementById('attach-file').click()">
                <i data-lucide="upload"></i> Subir Anexo
            </button>
        </div>
    `;
    showModal('Anexos del Registro', html, `
        <button class="btn" onclick="hideModal()">Cerrar</button>
    `);
};

window.uploadAttachment = async (recordId) => {
    const file = document.getElementById('attach-file').files[0];
    if (!file) return;
    await api.uploadAttachment(recordId, file);
    toast('Anexo subido');
    openAttachmentsModal(recordId);
    renderRecords(); // Update count in table
};

window.deleteAttachment = async (recordId, attachmentId) => {
    if (confirm('¿Eliminar este anexo?')) {
        await api.deleteAttachment(attachmentId);
        toast('Anexo eliminado');
        openAttachmentsModal(recordId);
        renderRecords();
    }
};

// --- Modals Handlers ---
window.editRecord = async (id) => {
    const rec = await api.fetch(`/api/records/${id}`);
    const fields = await api.getFields();
    
    let html = `<form id="record-form">`;
    fields.forEach(f => {
        html += `
            <div class="form-group">
                <label>${f.column_name}</label>
                <input type="text" name="${f.column_name}" value="${rec[f.column_name] || ''}" class="form-control">
            </div>
        `;
    });
    html += `</form>`;
    
    showModal('Editar Registro', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveRecord(${id})">Guardar</button>
    `);
};

window.saveRecord = async (id = null) => {
    const form = document.getElementById('record-form');
    const formData = new FormData(form);
    const data = {};
    formData.forEach((v, k) => data[k] = v);
    
    await api.saveRecord(data, id);
    hideModal();
    toast('Registro guardado');
    renderRecords();
};

window.deleteRecord = async (id) => {
    if (confirm('¿Estás seguro de eliminar este registro?')) {
        await api.deleteRecord(id);
        toast('Registro eliminado');
        renderRecords();
    }
};

window.toggleRecordStatus = async (id, active) => {
    await api.updateRecordStatus(id, active);
    toast(active ? 'Registro activado' : 'Registro inactivado');
    renderRecords();
};

// --- Fields Handlers ---
window.openFieldModal = async (id = null) => {
    let field = { column_name: '', display_name: '', visible: 1, display_order: 0, default_value: '' };
    if (id) {
        const fields = await api.getFields();
        field = fields.find(f => f.id === id);
    }
    
    let html = `
        <form id="field-form">
            <div class="form-group">
                <label>Nombre Columna</label>
                <input type="text" name="column_name" value="${field.column_name}" class="form-control" ${id ? 'disabled' : ''}>
            </div>
            <div class="form-group">
                <label>Etiqueta</label>
                <input type="text" name="display_name" value="${field.display_name}" class="form-control">
            </div>
            <div class="form-group">
                <label>Visible</label>
                <select name="visible" class="form-control">
                    <option value="1" ${field.visible ? 'selected' : ''}>Sí</option>
                    <option value="0" ${!field.visible ? 'selected' : ''}>No</option>
                </select>
            </div>
            <div class="form-group">
                <label>Orden</label>
                <input type="number" name="display_order" value="${field.display_order}" class="form-control">
            </div>
            <div class="form-group">
                <label>Valor Defecto</label>
                <input type="text" name="default_value" value="${field.default_value}" class="form-control">
            </div>
        </form>
    `;
    
    showModal(id ? 'Editar Campo' : 'Nuevo Campo', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveField(${id})">Guardar</button>
    `);
};

window.saveField = async (id = null) => {
    const form = document.getElementById('field-form');
    const data = {
        column_name: form.column_name ? form.column_name.value : '',
        display_name: form.display_name.value,
        visible: parseInt(form.visible.value),
        display_order: parseInt(form.display_order.value),
        default_value: form.default_value.value
    };
    await api.saveField(data, id);
    hideModal();
    toast('Campo guardado');
    renderFieldsConfig(await api.getFields());
};

window.deleteField = async (id) => {
    if (confirm('¿Eliminar este campo? Se borrarán los datos asociados.')) {
        await api.deleteField(id);
        toast('Campo eliminado');
        renderFieldsConfig(await api.getFields());
    }
};

// --- FUID Config ---
async function renderFuidConfig() {
    const header = await api.getConfig('fuid/header');
    const mapping = await api.getMapping('fuid');
    const fields = await api.getFields();
    
    let html = `
        <div class="config-grid" style="display:grid; grid-template-columns: 1fr 1.5fr; gap: 2rem">
            <div class="config-section">
                <h4>Encabezado FUID</h4>
                <form id="fuid-header-form">
                    ${Object.entries(header).map(([k, v]) => `
                        <div class="form-group">
                            <label>${k.replace(/_/g, ' ')}</label>
                            <input type="text" name="${k}" value="${v}" class="form-control">
                        </div>
                    `).join('')}
                    <button type="button" class="btn btn-primary" onclick="saveFuidHeader()">Guardar Encabezado</button>
                </form>
            </div>
            <div class="config-section">
                <h4>Mapeo de Columnas (Detalle)</h4>
                <div class="data-table-container">
                    <table>
                        <thead>
                            <tr>
                                <th>Campo FUID</th>
                                <th>Tipo</th>
                                <th>Origen</th>
                                <th>Acción</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${mapping.map(m => `
                                <tr>
                                    <td>${m.fuid_field}</td>
                                    <td>${m.mapping_type}</td>
                                    <td>${m.mapping_value}</td>
                                    <td>
                                        <button class="btn-icon" onclick="openMappingModal('fuid', ${m.id}, '${m.fuid_field}', '${m.mapping_type}', '${m.mapping_value}')">
                                            <i data-lucide="edit-2"></i>
                                        </button>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    `;
    document.getElementById('config-tab-content').innerHTML = html;
    lucide.createIcons();
}

window.saveFuidHeader = async () => {
    const form = document.getElementById('fuid-header-form');
    const formData = new FormData(form);
    const data = {};
    formData.forEach((v, k) => data[k] = v);
    await api.saveConfig('fuid/header', data);
    toast('Encabezado guardado');
};

// --- Mapping Modal ---
window.openMappingModal = async (type, id, fieldName, currentType, currentValue) => {
    const fields = await api.getFields();
    let html = `
        <form id="mapping-form">
            <div class="form-group">
                <label>Campo Destino</label>
                <input type="text" value="${fieldName}" class="form-control" disabled>
            </div>
            <div class="form-group">
                <label>Tipo de Mapeo</label>
                <select name="mapping_type" class="form-control" id="mapping-type-select">
                    <option value="field" ${currentType === 'field' ? 'selected' : ''}>Campo de Base de Datos</option>
                    <option value="fixed" ${currentType === 'fixed' ? 'selected' : ''}>Valor Fijo</option>
                    <option value="template" ${currentType === 'template' ? 'selected' : ''}>Plantilla {campo}</option>
                </select>
            </div>
            <div class="form-group">
                <label>Valor / Columna</label>
                <div id="mapping-value-container">
                    <!-- Dynamic based on type -->
                </div>
            </div>
        </form>
    `;
    
    showModal('Editar Mapeo', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveMapping('${type}', ${id})">Guardar</button>
    `);
    
    const container = document.getElementById('mapping-value-container');
    const select = document.getElementById('mapping-type-select');
    
    const updateValueInput = (t, v) => {
        if (t === 'field') {
            container.innerHTML = `
                <select name="mapping_value" class="form-control">
                    ${fields.map(f => `<option value="${f.column_name}" ${f.column_name === v ? 'selected' : ''}>${f.column_name}</option>`).join('')}
                </select>
            `;
        } else {
            container.innerHTML = `<input type="text" name="mapping_value" value="${v}" class="form-control">`;
        }
    };
    
    select.onchange = () => updateValueInput(select.value, '');
    updateValueInput(currentType, currentValue);
};

window.saveMapping = async (type, id) => {
    const form = document.getElementById('mapping-form');
    const data = {
        mapping_type: form.mapping_type.value,
        mapping_value: form.mapping_value.value
    };
    await api.updateMapping(type, id, data);
    hideModal();
    toast('Mapeo actualizado');
    if (type === 'fuid') renderFuidConfig();
    else if (type.includes('rotulo')) renderRotulosConfig();
};

// --- Rotulos Config ---
async function renderRotulosConfig() {
    const rc_cfg = await api.getConfig('rotulo-carpeta');
    const rc_map = await api.getMapping('rotulo-carpeta');
    
    let html = `
        <div class="config-grid" style="display:grid; grid-template-columns: 1fr 1fr; gap: 2rem">
            <div class="config-section">
                <h4>Rótulo Carpeta</h4>
                <form id="rc-config-form">
                    ${Object.entries(rc_cfg).map(([k, v]) => `
                        <div class="form-group">
                            <label>${k.replace(/_/g, ' ')}</label>
                            <input type="text" name="${k}" value="${v}" class="form-control">
                        </div>
                    `).join('')}
                    <button type="button" class="btn btn-primary" onclick="saveRcConfig()">Guardar Configuración</button>
                </form>
                <h5 style="margin-top:1.5rem">Mapeo Rótulo Carpeta</h5>
                <div class="data-table-container">
                    <table>
                        <thead>
                            <tr><th>Campo</th><th>Acción</th></tr>
                        </thead>
                        <tbody>
                            ${rc_map.map(m => `
                                <tr>
                                    <td>${m.rotulo_field}</td>
                                    <td>
                                        <button class="btn-icon" onclick="openMappingModal('rotulo-carpeta', ${m.id}, '${m.rotulo_field}', '${m.mapping_type}', '${m.mapping_value}')">
                                            <i data-lucide="edit-2"></i>
                                        </button>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
            <!-- Similar for Caja if needed -->
        </div>
    `;
    document.getElementById('config-tab-content').innerHTML = html;
    lucide.createIcons();
}

window.saveRcConfig = async () => {
    const form = document.getElementById('rc-config-form');
    const formData = new FormData(form);
    const data = {};
    formData.forEach((v, k) => data[k] = v);
    await api.saveConfig('rotulo-carpeta', data);
    toast('Configuración guardada');
};

// --- Users Management ---
async function renderUsers() {
    const users = await api.getUsers();
    let html = `
        <div class="actions-bar" style="margin-bottom:1rem; display:flex; justify-content:flex-end">
            <button class="btn btn-primary" onclick="openUserModal()"><i data-lucide="plus"></i> Nuevo Usuario</button>
        </div>
        <div class="data-table-container">
            <table>
                <thead>
                    <tr>
                        <th>Usuario</th>
                        <th>Nombre</th>
                        <th>Rol</th>
                        <th>Estado</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody>
                    ${users.map(u => `
                        <tr>
                            <td>${u.username}</td>
                            <td>${u.full_name}</td>
                            <td>
                                <span class="badge badge-outline">${u.role}</span>
                            </td>
                            <td>
                                <span class="badge badge-${u.active ? 'success' : 'danger'}">
                                    ${u.active ? 'Activo' : 'Inactivo'}
                                </span>
                            </td>
                            <td class="actions-cell">
                                <button class="btn-icon" title="Cambiar Rol" onclick="changeUserRole(${u.id}, '${u.role}')">
                                    <i data-lucide="shield"></i>
                                </button>
                                <button class="btn-icon" title="${u.active ? 'Inactivar' : 'Activar'}" onclick="toggleUserStatus(${u.id}, ${u.active ? 0 : 1})">
                                    <i data-lucide="${u.active ? 'toggle-right' : 'toggle-left'}" style="color: ${u.active ? 'var(--primary)' : 'var(--text-muted)'}"></i>
                                </button>
                                <button class="btn-icon" title="Resetear Clave" onclick="resetUserPassword(${u.id})">
                                    <i data-lucide="key"></i>
                                </button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
    document.getElementById('content-area').innerHTML = html;
    lucide.createIcons();
}

window.openUserModal = () => {
    let html = `
        <form id="user-form">
            <div class="form-group">
                <label>Usuario</label>
                <input type="text" name="username" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Nombre Completo</label>
                <input type="text" name="full_name" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Contraseña Inicial</label>
                <input type="password" name="password" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Rol</label>
                <select name="role" class="form-control">
                    <option value="normal">Usuario Normal</option>
                    <option value="admin">Administrador</option>
                </select>
            </div>
        </form>
    `;
    showModal('Nuevo Usuario', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveUser()">Crear</button>
    `);
};

window.saveUser = async () => {
    const form = document.getElementById('user-form');
    const formData = new FormData(form);
    const data = {};
    formData.forEach((v, k) => data[k] = v);
    await api.saveUser(data);
    hideModal();
    toast('Usuario creado');
    renderUsers();
};

window.toggleUserStatus = async (id, active) => {
    await api.updateUserStatus(id, active);
    toast('Estado actualizado');
    renderUsers();
};

window.resetUserPassword = (id) => {
    let html = `
        <div class="form-group">
            <label>Nueva Contraseña Temporal</label>
            <input type="text" id="reset-pass-val" class="form-control" placeholder="Ingrese la nueva clave">
        </div>
    `;
    showModal('Resetear Contraseña', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="confirmResetPassword(${id})">Confirmar</button>
    `);
};

window.confirmResetPassword = async (id) => {
    const pass = document.getElementById('reset-pass-val').value;
    if (!pass) {
        toast('Debe ingresar una contraseña', 'error');
        return;
    }
    await api.resetPassword(id, pass);
    hideModal();
    toast('Contraseña restablecida. El usuario deberá cambiarla al ingresar.');
};

window.changeUserRole = (id, currentRole) => {
    const newRole = currentRole === 'admin' ? 'normal' : 'admin';
    let html = `<p>¿Está seguro de cambiar el rol de este usuario a <strong>${newRole === 'admin' ? 'Administrador' : 'Usuario Normal'}</strong>?</p>`;
    showModal('Cambiar Rol', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="confirmChangeRole(${id}, '${newRole}')">Confirmar</button>
    `);
};

window.confirmChangeRole = async (id, newRole) => {
    await api.updateUserRole(id, newRole);
    hideModal();
    toast('Rol actualizado');
    renderUsers();
};

// Initialization
document.getElementById('login-form').onsubmit = login;
document.getElementById('logout-btn').onclick = logout;
document.getElementById('theme-toggle').onclick = () => {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    const icon = document.querySelector('#theme-toggle i');
    icon.setAttribute('data-lucide', isDark ? 'sun' : 'moon');
    document.querySelector('#theme-toggle span').textContent = isDark ? 'Modo Claro' : 'Modo Oscuro';
    lucide.createIcons();
};
document.getElementById('change-pass-btn').onclick = () => {
    let html = `
        <form id="change-pwd-form">
            <div class="form-group">
                <label>Contraseña Actual</label>
                <input type="password" name="current_password" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Nueva Contraseña</label>
                <input type="password" name="new_password" class="form-control" required>
            </div>
            <div class="form-group">
                <label>Confirmar Nueva Contraseña</label>
                <input type="password" name="confirm_password" class="form-control" required>
            </div>
        </form>
    `;
    showModal('Cambiar Contraseña', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveNewPassword()">Actualizar</button>
    `);
};

window.saveNewPassword = async () => {
    const form = document.getElementById('change-pwd-form');
    const data = {
        current_password: form.current_password.value,
        new_password: form.new_password.value,
        confirm_password: form.confirm_password.value
    };
    
    if (data.new_password !== data.confirm_password) {
        alert('Las contraseñas no coinciden');
        return;
    }
    
    try {
        await api.fetch('/api/change-password', { method: 'POST', body: JSON.stringify(data) });
        hideModal();
        toast('Contraseña actualizada correctamente');
    } catch (err) {}
};

document.getElementById('global-search').oninput = (e) => {
    state.searchQuery = e.target.value;
    if (state.currentSection === 'records') renderRecords();
};

document.querySelectorAll('.sidebar-nav li').forEach(li => {
    li.onclick = () => switchSection(li.dataset.section);
});

document.querySelector('.btn-close-modal').onclick = hideModal;

document.getElementById('btn-new-record').onclick = async () => {
    const fields = await api.getFields();
    let html = `<form id="record-form">`;
    fields.forEach(f => {
        html += `
            <div class="form-group">
                <label>${f.column_name}</label>
                <input type="text" name="${f.column_name}" value="${f.default_value || ''}" class="form-control">
            </div>
        `;
    });
    html += `</form>`;
    
    showModal('Nuevo Registro', html, `
        <button class="btn" onclick="hideModal()">Cancelar</button>
        <button class="btn btn-primary" onclick="saveRecord()">Crear</button>
    `);
};

// Check session on load
api.getMe().then(res => {
    if (res && res.user) {
        state.user = res.user;
        initApp();
    }
}).catch(() => {});

lucide.createIcons();
