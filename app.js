/**
 * CRM Óptica - Lógica Principal
 */

// --- DATOS ESTÁTICOS Y OPCIONES ---
const OPCIONES = {
    condiciones_visuales: [
        "Miopía", "Hipermetropía", "Astigmatismo", "Presbicia", "Emetropía",
        "Ambliopía", "Blefaritis", "Queratocono", "Conjuntivitis", "Pterigion",
        "Hemorragia conjuntival", "Estrabismo", "Orzuelo", "Ulcera", "Catarata", "Glaucoma"
    ],
    tipos_cita: [
        "Consulta integral", "Consulta integral + toma de presión", 
        "Consulta integral + topografía corneal", "Consulta diagnóstico ojo seco", 
        "Sesión tratamiento ojo seco", "Control", "Consulta especializada", 
        "Adaptación de lentes de contacto", "Prueba lentes de contacto", 
        "Control lentes de contacto", "Extracción de cuerpo extraño", 
        "Consultas externas", "Consulta niños + dilatación de pupila", "Rectificación"
    ],
    tipo_solucion: [
        "Gafas oftálmicas progresivas", "Gafas oftálmicas monofocales", 
        "Gafas oftálmicas bifocales", "Gafas oftálmicas ocupacionales", 
        "Gafas oftálmicas control miopía", "Medicamentos ojo seco", 
        "Tratamiento ojo seco", "Lentes de contacto blandos", 
        "Lentes de contacto rígidos", "Lentes de contacto especiales", 
        "Ortoqueratología", "Ninguna"
    ],
    laboratorios: [
        "Zeiss", "Geo", "Innovalab", "Servioptica", "Lentes VIP", 
        "World Visión", "Hoya", "P&O", "Ital-lent", "Cooper Vision", 
        "Univisual", "Ophta", "Ilab", "Restrepo comercial", 
        "Dempharma", "Botica", "Visión and Health"
    ],
    periodicidad: [
        "Diario", "Semanal", "Quincenal", "Mensual", 
        "Trimestral", "Semestral", "Anual", "Ninguna"
    ]
};

// --- ESTADO GLOBAL (Base de Datos en Memoria) ---
let DB = {
    clientes: [], // { id, tipo_doc, num_doc, nombres, apellidos, celular, correo, fecha_nac, genero, ocupacion, pais, ciudad, direccion, zona, estado_civil, eps, regimen, contacto_inicial, ref_doc, ref_parentesco, cond_iniciales }
    citas: [],    // { id, id_cliente, fecha, hora, profesional, motivo, cond_actuales, tipo_cita, estado }
    comercial: [] // { id_cita, adquirio, detalles: [{precio, tipo, lab, req_control, control_cual, periodicidad}], asesor, fecha_entrega, encuesta_enviada }
};

// --- VARIABLES GLOBALES ---
let pacienteActual = null; // Guardará el objeto paciente que se está consultando/creando
let citaActualID = null;
let calendar;
let charts = {};

// --- INICIALIZACIÓN ---
document.addEventListener('DOMContentLoaded', () => {
    initUI();
    loadFromLocalStorage();
    setupEventListeners();
});

function initUI() {
    // Llenar selects estáticos
    populateSelect('m1_tipo_doc', [
        {val: "CC", text: "Cédula de Ciudadanía"}, {val: "TI", text: "Tarjeta de Identidad"},
        {val: "CE", text: "Cédula de Extranjería"}, {val: "PA", text: "Pasaporte"}, {val: "RC", text: "Registro Civil"}
    ]);
    populateSelect('m2_tipo_cita', OPCIONES.tipos_cita.map(o => ({val: o, text: o})));
    populateSelect('dash_tipo_cita', OPCIONES.tipos_cita.map(o => ({val: o, text: o})));
    
    populateCheckboxes('m1_condiciones', OPCIONES.condiciones_visuales, 'cond_ini');
    populateCheckboxes('m2_condiciones_actuales', OPCIONES.condiciones_visuales, 'cond_act');
    populateSelect('s_laboratorio', OPCIONES.laboratorios.map(o => ({val: o, text: o})));
    populateSelect('s_periodicidad', OPCIONES.periodicidad.map(o => ({val: o, text: o})));
}

function populateSelect(id, options) {
    const select = document.getElementById(id);
    if (!select) return;
    options.forEach(opt => {
        let option = document.createElement('option');
        option.value = opt.val;
        option.textContent = opt.text;
        select.appendChild(option);
    });
}

function populateCheckboxes(containerId, options, namePrefix) {
    const container = document.getElementById(containerId);
    if (!container) return;
    options.forEach((opt, index) => {
        let div = document.createElement('div');
        div.className = 'checkbox-item';
        div.innerHTML = `
            <input type="checkbox" id="${namePrefix}_${index}" value="${opt}">
            <label for="${namePrefix}_${index}">${opt}</label>
        `;
        container.appendChild(div);
    });
}

// --- NAVEGACIÓN ---
function setupEventListeners() {
    // Sidebar Nav
    document.querySelectorAll('.nav-links li').forEach(item => {
        item.addEventListener('click', (e) => {
            document.querySelectorAll('.nav-links li').forEach(li => li.classList.remove('active'));
            item.classList.add('active');
            
            const targetId = item.getAttribute('data-target');
            document.querySelectorAll('.module').forEach(mod => mod.classList.add('hidden'));
            document.getElementById(targetId).classList.remove('hidden');
            
            document.getElementById('page-title').textContent = item.querySelector('span').textContent;

            // Trigger renders
            if (targetId === 'module-clients') renderClientes();
            if (targetId === 'module-appointments') setTimeout(() => { if(!calendar) initCalendar(); else calendar.render(); }, 100);
            if (targetId === 'module-commercial') renderComercial();
            if (targetId === 'module-dashboard') renderDashboard();
        });
    });

    // Excel Upload
    document.getElementById('btn-load-data').addEventListener('click', () => document.getElementById('file-upload').click());
    document.getElementById('file-upload').addEventListener('change', handleExcelUpload);
    document.getElementById('btn-export-data').addEventListener('click', exportToExcel);

    // Formulario Consulta
    document.getElementById('consulta-form').addEventListener('submit', handleConsulta);
    
    // Formulario Pt1 Logic
    document.getElementById('m1_contacto').addEventListener('change', (e) => {
        const container = document.getElementById('referido-container');
        if (e.target.value === 'Referido') {
            container.classList.remove('hidden');
        } else {
            container.classList.add('hidden');
        }
    });

    // Validación de máximo 4 checkboxes
    document.querySelectorAll('#m1_condiciones input[type="checkbox"]').forEach(cb => {
        cb.addEventListener('change', () => {
            const checked = document.querySelectorAll('#m1_condiciones input[type="checkbox"]:checked');
            if (checked.length > 4) {
                cb.checked = false;
                document.getElementById('condiciones-error').style.display = 'block';
            } else {
                document.getElementById('condiciones-error').style.display = 'none';
            }
        });
    });

    document.getElementById('btn-cancel-pt1').addEventListener('click', () => showFormConsulta());
    document.getElementById('maxi-pt1-form').addEventListener('submit', handleMaxiPt1);

    // Formulario Pt2 Logic
    document.getElementById('btn-cancel-pt2').addEventListener('click', () => showFormConsulta());
    document.getElementById('maxi-pt2-form').addEventListener('submit', handleMaxiPt2);

    // Búsquedas
    document.getElementById('search-clients').addEventListener('input', (e) => renderClientes(e.target.value));
    document.getElementById('search-commercial').addEventListener('input', (e) => renderComercial(e.target.value));

    // Modal Solución Comercial
    document.querySelector('.close-modal').addEventListener('click', () => document.getElementById('modal-solucion').classList.add('hidden'));
    document.getElementById('s_adquirio').addEventListener('change', (e) => {
        const container = document.getElementById('soluciones-container');
        if(e.target.value === 'Si') container.classList.remove('hidden');
        else container.classList.add('hidden');
    });
    
    document.getElementById('s_cantidad').addEventListener('input', renderSolucionItems);
    document.getElementById('s_req_control').addEventListener('change', (e) => {
        const c = document.getElementById('s_control_cual_container');
        if(e.target.value === 'Si') c.classList.remove('hidden');
        else c.classList.add('hidden');
    });
    
    document.getElementById('solucion-form').addEventListener('submit', handleSolucionSubmit);
    document.getElementById('btn-filter-dash').addEventListener('click', renderDashboard);
}

// --- LÓGICA DE FORMULARIOS ---
function showFormConsulta() {
    document.getElementById('form-consulta').classList.remove('hidden');
    document.getElementById('form-maxi-pt1').classList.add('hidden');
    document.getElementById('form-maxi-pt2').classList.add('hidden');
    document.getElementById('consulta-form').reset();
    pacienteActual = null;
}

function handleConsulta(e) {
    e.preventDefault();
    const tipoDoc = document.getElementById('c_tipo_doc').value;
    const numDoc = document.getElementById('c_num_doc').value.trim();

    // Buscar solo por num_doc ya que el Excel antiguo no tiene tipo de documento
    const paciente = DB.clientes.find(c => String(c.num_doc) === String(numDoc));

    if (paciente) {
        pacienteActual = paciente;
        showAlertModal(
            'Paciente Encontrado',
            `<p><strong>Nombre:</strong> ${paciente.nombres} ${paciente.apellidos}</p>
             <p><strong>Celular:</strong> ${paciente.celular}</p>
             <p class="mt-20">¿Desea realizar actualización de datos?</p>`,
            `<button class="btn btn-outline" onclick="goToPt2(false)">No, agendar cita</button>
             <button class="btn btn-primary" onclick="goToPt1(true)">Sí, actualizar</button>`
        );
    } else {
        pacienteActual = { tipo_doc: tipoDoc, num_doc: numDoc }; // Temporary new patient
        showAlertModal(
            'Paciente No Encontrado',
            `<p>Debe realizar la creación del cliente.</p>`,
            `<button class="btn btn-outline" onclick="closeAlertModal()">Cancelar</button>
             <button class="btn btn-primary" onclick="goToPt1(false)">Aceptar y Crear</button>`
        );
    }
}

function showAlertModal(title, body, footer) {
    document.getElementById('modal-alert-title').textContent = title;
    document.getElementById('modal-alert-body').innerHTML = body;
    document.getElementById('modal-alert-footer').innerHTML = footer;
    document.getElementById('modal-alert').classList.remove('hidden');
}

function closeAlertModal() {
    document.getElementById('modal-alert').classList.add('hidden');
}

function goToPt1(isUpdate) {
    closeAlertModal();
    document.getElementById('form-consulta').classList.add('hidden');
    document.getElementById('form-maxi-pt1').classList.remove('hidden');
    
    document.getElementById('m1_tipo_doc').value = pacienteActual.tipo_doc;
    document.getElementById('m1_num_doc').value = pacienteActual.num_doc;
    
    if (isUpdate) {
        document.getElementById('m1_nombres').value = pacienteActual.nombres || '';
        document.getElementById('m1_apellidos').value = pacienteActual.apellidos || '';
        document.getElementById('m1_celular').value = pacienteActual.celular || '';
        document.getElementById('m1_correo').value = pacienteActual.correo || '';
        document.getElementById('m1_fecha_nac').value = pacienteActual.fecha_nac || '';
        document.getElementById('m1_genero').value = pacienteActual.genero || '';
        document.getElementById('m1_ocupacion').value = pacienteActual.ocupacion || '';
        document.getElementById('m1_pais').value = pacienteActual.pais || 'Colombia';
        document.getElementById('m1_ciudad').value = pacienteActual.ciudad || '';
        document.getElementById('m1_direccion').value = pacienteActual.direccion || '';
        document.getElementById('m1_zona').value = pacienteActual.zona || '';
        document.getElementById('m1_estado_civil').value = pacienteActual.estado_civil || '';
        document.getElementById('m1_eps').value = pacienteActual.eps || '';
        document.getElementById('m1_regimen').value = pacienteActual.regimen || '';
        document.getElementById('m1_contacto').value = pacienteActual.contacto_inicial || '';
        
        if (pacienteActual.contacto_inicial === 'Referido') {
            document.getElementById('referido-container').classList.remove('hidden');
            document.getElementById('m1_ref_doc').value = pacienteActual.ref_doc || '';
            document.getElementById('m1_ref_parentesco').value = pacienteActual.ref_parentesco || '';
        }

        if (pacienteActual.cond_iniciales) {
            document.querySelectorAll('#m1_condiciones input').forEach(cb => {
                if (pacienteActual.cond_iniciales.includes(cb.value)) cb.checked = true;
                else cb.checked = false;
            });
        }
    } else {
        document.getElementById('maxi-pt1-form').reset();
        document.getElementById('m1_tipo_doc').value = pacienteActual.tipo_doc;
        document.getElementById('m1_num_doc').value = pacienteActual.num_doc;
        document.getElementById('referido-container').classList.add('hidden');
    }
}

function handleMaxiPt1(e) {
    e.preventDefault();
    const condCheckboxes = document.querySelectorAll('#m1_condiciones input:checked');
    if (condCheckboxes.length === 0 || condCheckboxes.length > 4) {
        document.getElementById('condiciones-error').style.display = 'block';
        return;
    }

    const clienteData = {
        id: pacienteActual.id || generateID(),
        tipo_doc: document.getElementById('m1_tipo_doc').value,
        num_doc: document.getElementById('m1_num_doc').value,
        nombres: document.getElementById('m1_nombres').value,
        apellidos: document.getElementById('m1_apellidos').value,
        celular: document.getElementById('m1_celular').value,
        correo: document.getElementById('m1_correo').value,
        fecha_nac: document.getElementById('m1_fecha_nac').value,
        genero: document.getElementById('m1_genero').value,
        ocupacion: document.getElementById('m1_ocupacion').value,
        pais: document.getElementById('m1_pais').value,
        ciudad: document.getElementById('m1_ciudad').value,
        direccion: document.getElementById('m1_direccion').value,
        zona: document.getElementById('m1_zona').value,
        estado_civil: document.getElementById('m1_estado_civil').value,
        eps: document.getElementById('m1_eps').value,
        regimen: document.getElementById('m1_regimen').value,
        contacto_inicial: document.getElementById('m1_contacto').value,
        ref_doc: document.getElementById('m1_contacto').value === 'Referido' ? document.getElementById('m1_ref_doc').value : 'N/A',
        ref_parentesco: document.getElementById('m1_contacto').value === 'Referido' ? document.getElementById('m1_ref_parentesco').value : 'N/A',
        cond_iniciales: Array.from(condCheckboxes).map(cb => cb.value),
        fecha_registro: new Date().toISOString().split('T')[0]
    };

    const idx = DB.clientes.findIndex(c => c.id === clienteData.id);
    if (idx >= 0) DB.clientes[idx] = clienteData;
    else DB.clientes.push(clienteData);

    pacienteActual = clienteData;
    saveToLocalStorage();
    goToPt2(true);
}

function goToPt2(fromPt1) {
    if (!fromPt1) closeAlertModal();
    document.getElementById('form-consulta').classList.add('hidden');
    document.getElementById('form-maxi-pt1').classList.add('hidden');
    document.getElementById('form-maxi-pt2').classList.remove('hidden');
    document.getElementById('maxi-pt2-form').reset();
    
    // Set default date to today, time to now
    const now = new Date();
    document.getElementById('m2_fecha').value = now.toISOString().split('T')[0];
    document.getElementById('m2_hora').value = now.toTimeString().substring(0,5);
}

function handleMaxiPt2(e) {
    e.preventDefault();
    const condCheckboxes = document.querySelectorAll('#m2_condiciones_actuales input:checked');
    
    const citaData = {
        id: generateID(),
        id_cliente: pacienteActual.id,
        fecha: document.getElementById('m2_fecha').value,
        hora: document.getElementById('m2_hora').value,
        profesional: document.getElementById('m2_profesional').value,
        motivo: document.getElementById('m2_motivo').value,
        cond_actuales: Array.from(condCheckboxes).map(cb => cb.value),
        tipo_cita: document.getElementById('m2_tipo_cita').value,
        estado: document.getElementById('m2_estado_cita').value
    };

    DB.citas.push(citaData);
    saveToLocalStorage();
    
    alert('Información guardada exitosamente.');
    showFormConsulta(); // Reset form
}

// --- MODULO CLIENTES ---
function renderClientes(searchTerm = '') {
    const tbody = document.querySelector('#table-clients tbody');
    tbody.innerHTML = '';
    
    let filterData = DB.clientes;
    if(searchTerm) {
        const s = searchTerm.toLowerCase();
        filterData = DB.clientes.filter(c => 
            (c.nombres + ' ' + c.apellidos).toLowerCase().includes(s) || 
            c.num_doc.includes(s)
        );
    }
    
    document.getElementById('total-clients').textContent = filterData.length;

    filterData.forEach(c => {
        let tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${c.tipo_doc} ${c.num_doc}</td>
            <td>${c.nombres}</td>
            <td>${c.apellidos}</td>
            <td>${c.celular}</td>
            <td>${c.correo}</td>
            <td><span class="badge badge-success">${c.contacto_inicial}</span></td>
            <td>${c.fecha_registro || 'N/A'}</td>
        `;
        tbody.appendChild(tr);
    });
}

// --- MODULO CALENDARIO (Gestión de Citas) ---
function initCalendar() {
    const calendarEl = document.getElementById('calendar');
    calendar = new FullCalendar.Calendar(calendarEl, {
        initialView: 'dayGridMonth',
        headerToolbar: {
            left: 'prev,next today',
            center: 'title',
            right: 'dayGridMonth,timeGridWeek,timeGridDay'
        },
        locale: 'es',
        events: getCalendarEvents()
    });
    calendar.render();
}

function getCalendarEvents() {
    return DB.citas.map(cita => {
        const cliente = DB.clientes.find(c => c.id === cita.id_cliente);
        const name = cliente ? `${cliente.nombres} ${cliente.apellidos}` : 'Desconocido';
        
        let color = '#3788d8'; // default
        if(cita.estado === 'Confirmada') color = '#10b981';
        if(cita.estado === 'Reprogramada') color = '#f59e0b';
        if(cita.estado === 'No asistida') color = '#ef4444';

        return {
            title: `${name} - ${cita.tipo_cita}`,
            start: `${cita.fecha}T${cita.hora}`,
            backgroundColor: color,
            extendedProps: { cita_id: cita.id }
        };
    });
}

// --- MODULO GESTIÓN COMERCIAL ---
function renderComercial(searchTerm = '') {
    const tbody = document.querySelector('#table-commercial tbody');
    tbody.innerHTML = '';
    
    // Sort by date desc
    let citasOrdenadas = [...DB.citas].sort((a,b) => new Date(b.fecha) - new Date(a.fecha));

    citasOrdenadas.forEach(cita => {
        const c = DB.clientes.find(cl => cl.id === cita.id_cliente);
        if(!c) return;

        if(searchTerm) {
            const s = searchTerm.toLowerCase();
            if(!((c.nombres + ' ' + c.apellidos).toLowerCase().includes(s) || c.num_doc.includes(s))) return;
        }

        const comercial = DB.comercial.find(com => com.id_cita === cita.id);
        const hasSolucion = comercial && comercial.adquirio === 'Si';
        
        // Calculate days since delivery
        let btnSurvey = '';
        if(hasSolucion && comercial.fecha_entrega) {
            const deliveryDate = new Date(comercial.fecha_entrega);
            const today = new Date();
            const diffTime = Math.abs(today - deliveryDate);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
            
            if(diffDays >= 7) {
                btnSurvey = `<button class="icon-btn icon-success" title="Enviar Encuesta de Satisfacción (7 Días)" onclick="enviarEncuesta('${cita.id}')">
                                <i class="fa-solid fa-clipboard-list"></i>
                             </button>`;
            }
        }

        let tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${c.tipo_doc} ${c.num_doc}</td>
            <td>${c.nombres}</td>
            <td>${c.apellidos}</td>
            <td>${c.celular}</td>
            <td>${cita.fecha}</td>
            <td>${cita.tipo_cita}</td>
            <td class="action-icons">
                <button class="icon-btn" title="Registrar Solución Visual" onclick="openSolucionModal('${cita.id}')">
                    <i class="fa-solid fa-glasses ${hasSolucion ? 'text-success' : ''}"></i>
                </button>
                ${btnSurvey}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function openSolucionModal(citaId) {
    document.getElementById('solucion-form').reset();
    document.getElementById('solucion_cita_id').value = citaId;
    document.getElementById('soluciones-container').classList.add('hidden');
    renderSolucionItems(); // Reset to 1

    const existing = DB.comercial.find(c => c.id_cita === citaId);
    if(existing) {
        document.getElementById('s_adquirio').value = existing.adquirio;
        if(existing.adquirio === 'Si') {
            document.getElementById('soluciones-container').classList.remove('hidden');
            document.getElementById('s_cantidad').value = existing.detalles.length;
            renderSolucionItems();
            
            // Llenar detalles
            existing.detalles.forEach((d, i) => {
                document.getElementById(`sol_precio_${i}`).value = d.precio;
                // Para simplificar, asumo 1 opcion por ahora
            });
            
            // Set values for general fields
            setMultipleSelectValues('s_laboratorio', existing.detalles[0]?.lab || []);
            document.getElementById('s_req_control').value = existing.detalles[0]?.req_control || '';
            document.getElementById('s_control_cual').value = existing.detalles[0]?.control_cual || '';
            setMultipleSelectValues('s_periodicidad', existing.detalles[0]?.periodicidad || []);
            document.getElementById('s_asesor').value = existing.asesor;
            document.getElementById('s_fecha_entrega').value = existing.fecha_entrega;
        }
    }

    document.getElementById('modal-solucion').classList.remove('hidden');
}

function renderSolucionItems() {
    let cant = parseInt(document.getElementById('s_cantidad').value) || 1;
    if(cant > 5) { cant = 5; document.getElementById('s_cantidad').value = 5; }
    
    const container = document.getElementById('items-solucion');
    container.innerHTML = '';
    
    for(let i=0; i<cant; i++) {
        let div = document.createElement('div');
        div.className = 'solucion-item';
        
        let typeOptions = OPCIONES.tipo_solucion.map(o => `<option value="${o}">${o}</option>`).join('');

        div.innerHTML = `
            <div class="solucion-item-header">Solución #${i+1}</div>
            <div class="form-group">
                <label>Precio $ *</label>
                <input type="number" id="sol_precio_${i}" required>
            </div>
            <div class="form-group">
                <label>Tipo de Solución Adquirida *</label>
                <select id="sol_tipo_${i}" required>
                    <option value="">Seleccione...</option>
                    ${typeOptions}
                </select>
            </div>
        `;
        container.appendChild(div);
    }
}

function handleSolucionSubmit(e) {
    e.preventDefault();
    const citaId = document.getElementById('solucion_cita_id').value;
    const adquirio = document.getElementById('s_adquirio').value;
    
    let comercialData = {
        id_cita: citaId,
        adquirio: adquirio,
        detalles: [],
        asesor: '',
        fecha_entrega: ''
    };

    if(adquirio === 'Si') {
        const cant = parseInt(document.getElementById('s_cantidad').value) || 1;
        for(let i=0; i<cant; i++) {
            comercialData.detalles.push({
                precio: document.getElementById(`sol_precio_${i}`).value,
                tipo: document.getElementById(`sol_tipo_${i}`).value,
                lab: getMultipleSelectValues('s_laboratorio'),
                req_control: document.getElementById('s_req_control').value,
                control_cual: document.getElementById('s_control_cual').value,
                periodicidad: getMultipleSelectValues('s_periodicidad')
            });
        }
        comercialData.asesor = document.getElementById('s_asesor').value;
        comercialData.fecha_entrega = document.getElementById('s_fecha_entrega').value;
    }

    const idx = DB.comercial.findIndex(c => c.id_cita === citaId);
    if(idx >= 0) DB.comercial[idx] = comercialData;
    else DB.comercial.push(comercialData);

    saveToLocalStorage();
    document.getElementById('modal-solucion').classList.add('hidden');
    renderComercial();
    alert("Datos comerciales guardados.");
}

function enviarEncuesta(citaId) {
    // Open Google Form in new tab
    const url = "https://docs.google.com/forms/d/e/1FAIpQLSeXRdQNnxHL6Fpxu74dOmmnylmCeU1k31NS2LGYlJsi_ev5xw/viewform?usp=header";
    window.open(url, "_blank");
    // Marcar como enviada
    const com = DB.comercial.find(c => c.id_cita === citaId);
    if(com) {
        com.encuesta_enviada = true;
        saveToLocalStorage();
    }
}

function getMultipleSelectValues(id) {
    const sel = document.getElementById(id);
    const result = [];
    for(let opt of sel.options) {
        if(opt.selected) result.push(opt.value);
    }
    return result;
}

function setMultipleSelectValues(id, valuesArray) {
    const sel = document.getElementById(id);
    for(let opt of sel.options) {
        if(valuesArray.includes(opt.value)) opt.selected = true;
        else opt.selected = false;
    }
}

// --- MODULO DASHBOARD ---
function renderDashboard() {
    const dIni = document.getElementById('dash_fecha_ini').value;
    const dFin = document.getElementById('dash_fecha_fin').value;
    const tCita = document.getElementById('dash_tipo_cita').value;

    // Filter citas
    let filteredCitas = DB.citas;
    if(dIni) filteredCitas = filteredCitas.filter(c => c.fecha >= dIni);
    if(dFin) filteredCitas = filteredCitas.filter(c => c.fecha <= dFin);
    if(tCita && tCita !== 'Todas') filteredCitas = filteredCitas.filter(c => c.tipo_cita === tCita);

    // KPIs
    document.getElementById('kpi-total-clients').textContent = DB.clientes.length;
    document.getElementById('kpi-asistidas').textContent = filteredCitas.filter(c => c.estado === 'Asistida').length;
    
    // Calculate Referidos
    const referidos = DB.clientes.filter(c => c.contacto_inicial === 'Referido').length;
    document.getElementById('kpi-referidos').textContent = referidos;

    // Frecuencia Chart
    const frecCounts = { "1":0, "2":0, "3":0, "4":0, "+4":0 };
    const clienteVisitas = {};
    filteredCitas.forEach(c => {
        clienteVisitas[c.id_cliente] = (clienteVisitas[c.id_cliente] || 0) + 1;
    });

    Object.values(clienteVisitas).forEach(count => {
        if(count === 1) frecCounts["1"]++;
        else if(count === 2) frecCounts["2"]++;
        else if(count === 3) frecCounts["3"]++;
        else if(count === 4) frecCounts["4"]++;
        else if(count > 4) frecCounts["+4"]++;
    });

    renderChart('chart-frecuencia', 'bar', ['1 Vez', '2 Veces', '3 Veces', '4 Veces', 'Más de 4'], Object.values(frecCounts), 'Frecuencia de Agendamiento');

    // Fidelización (Pie) - Simplified definition
    let fidelizados = 0; // > 3 visits
    let recurrentes = 0; // 2-3 visits
    let nuevos = 0; // 1 visit
    
    Object.values(clienteVisitas).forEach(count => {
        if(count === 1) nuevos++;
        else if(count > 1 && count <= 3) recurrentes++;
        else if(count > 3) fidelizados++;
    });

    renderChart('chart-fidelizacion', 'doughnut', ['Nuevos (1)', 'Recurrentes (2-3)', 'Fidelizados (>3)'], [nuevos, recurrentes, fidelizados], 'Tipo de Cliente');

    // Sociodemográfico (Edades)
    const edadesCounts = { "<18":0, "18-30":0, "31-50":0, "51-70":0, ">70":0 };
    DB.clientes.forEach(c => {
        if(!c.fecha_nac) return;
        const age = new Date().getFullYear() - new Date(c.fecha_nac).getFullYear();
        if(age < 18) edadesCounts["<18"]++;
        else if(age <= 30) edadesCounts["18-30"]++;
        else if(age <= 50) edadesCounts["31-50"]++;
        else if(age <= 70) edadesCounts["51-70"]++;
        else edadesCounts[">70"]++;
    });
    
    renderChart('chart-edades', 'bar', Object.keys(edadesCounts), Object.values(edadesCounts), 'Rango Etario');
}

function renderChart(canvasId, type, labels, data, label) {
    if(charts[canvasId]) charts[canvasId].destroy();
    
    const ctx = document.getElementById(canvasId).getContext('2d');
    charts[canvasId] = new Chart(ctx, {
        type: type,
        data: {
            labels: labels,
            datasets: [{
                label: label,
                data: data,
                backgroundColor: ['#2563eb', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444'],
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom' } }
        }
    });
}

// --- UTILIDADES Y MANEJO DE ARCHIVOS ---
function generateID() {
    return '_' + Math.random().toString(36).substr(2, 9);
}

function saveToLocalStorage() {
    localStorage.setItem('crm_optica_db', JSON.stringify(DB));
    updateStatusBadge();
    document.getElementById('btn-export-data').disabled = false;
}

function loadFromLocalStorage() {
    const data = localStorage.getItem('crm_optica_db');
    if (data) {
        DB = JSON.parse(data);
        updateStatusBadge();
        document.getElementById('btn-export-data').disabled = false;
    }
}

function updateStatusBadge() {
    const badge = document.getElementById('db-status-badge');
    badge.className = 'badge badge-success';
    badge.textContent = `Cargada (${DB.clientes.length} Clientes)`;
}

function showLoader() { document.getElementById('loader').classList.remove('hidden'); }
function hideLoader() { document.getElementById('loader').classList.add('hidden'); }

// --- EXCEL LOGIC (SheetJS) ---
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    showLoader();
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array', cellDates: true});
        
        try {
            // Verificar si es nuestro nuevo formato exportado
            if(workbook.SheetNames.includes('Clientes') || workbook.SheetNames.includes('Citas')) {
                if(workbook.SheetNames.includes('Clientes')) {
                    DB.clientes = XLSX.utils.sheet_to_json(workbook.Sheets['Clientes']).map(row => ({
                        ...row,
                        num_doc: String(row.num_doc)
                    }));
                }
                if(workbook.SheetNames.includes('Citas')) {
                    DB.citas = XLSX.utils.sheet_to_json(workbook.Sheets['Citas']).map(row => {
                        if (row.fecha instanceof Date) row.fecha = row.fecha.toISOString().split('T')[0];
                        return row;
                    });
                }
                if(workbook.SheetNames.includes('Comercial')) {
                    DB.comercial = XLSX.utils.sheet_to_json(workbook.Sheets['Comercial']).map(row => ({
                        ...row,
                        detalles: JSON.parse(row.detalles || '[]')
                    }));
                }
            } else {
                // Formato antiguo: Datos CRM 2.xlsx (Hoja "CRM" o similar)
                const legacySheetName = workbook.SheetNames.find(s => s.trim().toUpperCase() === 'CRM');
                if (legacySheetName) {
                    const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[legacySheetName], {raw: false, dateNF: 'yyyy-mm-dd'});
                    
                    DB.clientes = [];
                    DB.citas = [];
                    DB.comercial = [];

                    rawData.forEach((row, i) => {
                        let numDoc = String(row['NUMERO DOCUMENTO'] || row['NUMERO DE DOCUMENTO'] || `DESCONOCIDO_${i}`).trim();
                        
                        let cliente = DB.clientes.find(c => c.num_doc === numDoc);
                        if (!cliente) {
                            cliente = {
                                id: generateID(),
                                tipo_doc: 'CC', // Por defecto
                                num_doc: numDoc,
                                nombres: row['NOMBRE'] || '',
                                apellidos: row['APELLIDO'] || '',
                                fecha_nac: row['FECHA NACIMIENTO'] || '',
                                genero: row['SEXO'] || '',
                                eps: row['EPS'] || '',
                                celular: '',
                                correo: '',
                                cond_iniciales: row['CONDICION VISUAL INICIAL 1'] ? [row['CONDICION VISUAL INICIAL 1']] : []
                            };
                            DB.clientes.push(cliente);
                        }

                        if (row['FECHA DE CONSULTA'] || row['TIPO DE CITA']) {
                            let citaId = generateID();
                            let dConsulta = row['FECHA DE CONSULTA'];
                            
                            DB.citas.push({
                                id: citaId,
                                id_cliente: cliente.id,
                                fecha: dConsulta || '',
                                hora: '08:00',
                                profesional: row['PROFESIONAL'] || '',
                                tipo_cita: row['TIPO DE CITA'] || '',
                                estado: row['ESTADO DE LA CITA'] || 'Pendiente',
                                motivo: '',
                                cond_actuales: []
                            });

                            if (row['VENTA SOLUCION'] || row['ASESOR DE VENTA']) {
                                let ventaText = String(row['VENTA SOLUCION'] || '').toLowerCase();
                                let adquirio = (ventaText && ventaText !== 'no') ? 'Si' : 'No';
                                
                                DB.comercial.push({
                                    id_cita: citaId,
                                    adquirio: adquirio,
                                    detalles: [],
                                    asesor: row['ASESOR DE VENTA'] || '',
                                    fecha_entrega: ''
                                });
                            }
                        }
                    });
                }
            }
            
            saveToLocalStorage();
            alert("Base de datos cargada correctamente.");
        } catch(err) {
            console.error(err);
            alert("Error procesando el archivo. Asegúrese de que tenga la estructura correcta.");
        }
        hideLoader();
    };
    reader.readAsArrayBuffer(file);
}

function exportToExcel() {
    showLoader();
    
    const wb = XLSX.utils.book_new();
    
    // Sheet: Clientes
    const wsClientes = XLSX.utils.json_to_sheet(DB.clientes.map(c => ({
        ...c, cond_iniciales: c.cond_iniciales ? c.cond_iniciales.join(', ') : ''
    })));
    XLSX.utils.book_append_sheet(wb, wsClientes, "Clientes");
    
    // Sheet: Citas
    const wsCitas = XLSX.utils.json_to_sheet(DB.citas.map(c => ({
        ...c, cond_actuales: c.cond_actuales ? c.cond_actuales.join(', ') : ''
    })));
    XLSX.utils.book_append_sheet(wb, wsCitas, "Citas");
    
    // Sheet: Comercial
    const wsComercial = XLSX.utils.json_to_sheet(DB.comercial.map(c => ({
        ...c, detalles: JSON.stringify(c.detalles)
    })));
    XLSX.utils.book_append_sheet(wb, wsComercial, "Comercial");
    
    // Write and download
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Datos_CRM_${dateStr}.xlsx`);
    
    hideLoader();
}
