const app = firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// Constantes analíticas para la lectura del Excel de Microsoft Forms
const TECH_START = 18;
const TECH_END = 83;
const PEOPLE_START = 84;
const PEOPLE_END = 89;
const NPS_COL = 90;

let globalIdDocEvalua = '';
let globalCampoFaseEvalua = '';
let globalPadrinoFotoBase64 = '';
let cacheSnapshotEmpleadosLocal = null;

const miRol = localStorage.getItem('usuario_rol');
const cedulaLogueada = localStorage.getItem('usuario_cedula');

if (typeof emailjs !== "undefined") emailjs.init("JkzXt8TMr--AXQsiM");

if (!miRol) {
    window.location.href = 'login.html';
}

window.addEventListener('DOMContentLoaded', () => {
    ejecutarRestriccionDeRoles();

    // ✨ 🧙‍♂️ AUTOCUMPLEMENTADO INTELIGENTE POR CÉDULA
    const inputCedula = document.getElementById('padreApadrinadoId');
    if (inputCedula) {
        inputCedula.addEventListener('blur', async () => {
            const cedulaBuscada = inputCedula.value.trim();
            if (!cedulaBuscada) return;

            console.log(`🔍 Buscando datos preexistentes para la cédula: ${cedulaBuscada}`);

            try {
                let snapshot = await db.collection('empleados').where('cedula', '==', cedulaBuscada).limit(1).get();

                if (snapshot.empty && !isNaN(cedulaBuscada)) {
                    snapshot = await db.collection('empleados').where('cedula', '==', Number(cedulaBuscada)).limit(1).get();
                }

                if (!snapshot.empty) {
                    const empData = snapshot.docs[0].data();
                    console.log("🎯 ¡Empleado encontrado! Autocompletando campos...", empData);

                    if (document.getElementById('padreApadrinadoNombre')) document.getElementById('padreApadrinadoNombre').value = empData.nombre || '';
                    if (document.getElementById('padreApadrinadoCargo')) document.getElementById('padreApadrinadoCargo').value = empData.cargo || '';
                    if (document.getElementById('padreApadrinadoArea')) document.getElementById('padreApadrinadoArea').value = empData.area || '';
                    if (document.getElementById('padreApadrinadoRegion')) document.getElementById('padreApadrinadoRegion').value = empData.region || '';
                    if (document.getElementById('padreApadrinadoCorreo')) document.getElementById('padreApadrinadoCorreo').value = empData.correo || '';

                    inputCedula.style.borderColor = '#0f6e56';
                } else {
                    console.log("ℹ️ La cédula es nueva en el sistema, requiere registro manual.");
                    inputCedula.style.borderColor = '';
                }
            } catch (error) {
                console.error("❌ Error en el autocompletado inteligente:", error);
            }
        });
    }
});

function ejecutarRestriccionDeRoles() {
    const b1CertificarPadrino = document.querySelector('.grid-padrino > .card:nth-child(1)');
    const b2RegistrarIngreso = document.querySelector('.grid-padrino > .card:nth-child(2)');

    if (b1CertificarPadrino) b1CertificarPadrino.style.display = 'block';
    if (b2RegistrarIngreso) b2RegistrarIngreso.style.display = 'block';

    if (miRol === 'people') {
        if (b1CertificarPadrino) b1CertificarPadrino.style.display = 'none';
        const btnRegistrar = document.querySelector('.grid-padrino > .card:nth-child(2) .btn-action');
        if (btnRegistrar) btnRegistrar.style.display = 'none';
    } else if (miRol === 'lider' || miRol === 'jefe') {
        if (b1CertificarPadrino) b1CertificarPadrino.style.display = 'none';
    } else if (miRol === 'padrino') {
        if (b1CertificarPadrino) b1CertificarPadrino.style.display = 'none';
        if (b2RegistrarIngreso) b2RegistrarIngreso.style.display = 'none';
    }
}

// 📊 1. ESCUCHADOR DEL DASHBOARD DE INDICADORES GENERALES
db.collection('empleados').onSnapshot((snapshot) => {
    let ingresosMes = 0; let ingresosAnio = 0;
    let padrinosActivos = 0;
    let planesActivos = 0; let planesFinalizados = 0; let planesVencidos = 0;
    let tOnboardingTotal = 0; let tOnboardingComp = 0;
    let examenesRealizados = 0; let examenesPendientes = 0;

    let sumaNota7 = 0; let cant7 = 0;
    let sumaNota30 = 0; let cant30 = 0;
    let sumaNota90 = 0; let cant90 = 0;

    const hoy = new Date('2026-06-08T00:00:00');
    const setRegiones = new Set();
    const setMeses = new Set();

    snapshot.forEach((doc) => {
        const emp = doc.data();
        if (emp.region) setRegiones.add(emp.region.trim());
        if (emp.fecha_ingreso) setMeses.add(emp.fecha_ingreso.substring(0, 7));

        if (emp.es_padrino === true || emp.es_padrino === "true") {
            if (emp.padrino_estado !== "Inactivo") padrinosActivos++;
        }

        if (emp.es_apadrinado === true) {
            if (miRol === 'padrino' && String(emp.padrino_id) !== String(cedulaLogueada)) return;

            if (emp.estado_plan_padrino === "Plan Padrino Finalizado") {
                planesFinalizados++;
            } else {
                planesActivos++;
                if (emp.fechas_limite_evaluaciones) {
                    const f7 = new Date(emp.fechas_limite_evaluaciones.eval_7_dias + 'T00:00:00');
                    const f30 = new Date(emp.fechas_limite_evaluaciones.eval_30_dias + 'T00:00:00');
                    const f90 = new Date(emp.fechas_limite_evaluaciones.eval_90_dias + 'T00:00:00');

                    if (emp.etapa_actual === "Etapa Técnico" && hoy > f7) planesVencidos++;
                    else if (emp.etapa_actual === "Etapa Funcional" && hoy > f30) planesVencidos++;
                    else if (emp.etapa_actual === "Etapa Autónomo" && hoy > f90) planesVencidos++;
                }
            }

            if (emp.fecha_ingreso) {
                if (emp.fecha_ingreso.startsWith("2026-06")) ingresosMes++;
                if (emp.fecha_ingreso.startsWith("2026")) ingresosAnio++;
            }

            tOnboardingTotal += 2;
            if (emp.onboarding_dia1 && emp.onboarding_dia1.estado === "Completado") tOnboardingComp++;
            if (emp.onboarding_dia2 && emp.onboarding_dia2.estado === "Completado") tOnboardingComp++;

            if (emp.hitos) {
                if (emp.hitos.eval_tecnico_nota > 0) { sumaNota7 += emp.hitos.eval_tecnico_nota; cant7++; examenesRealizados++; } else { examenesPendientes++; }
                if (emp.hitos.eval_funcional_nota > 0) { sumaNota30 += emp.hitos.eval_funcional_nota; cant30++; examenesRealizados++; } else { examenesPendientes++; }
                if (emp.hitos.eval_autonomo_nota > 0) { sumaNota90 += emp.hitos.eval_autonomo_nota; cant90++; examenesRealizados++; } else { examenesPendientes++; }
            }
        }
    });

    poblarFiltrosEstrategicos(Array.from(setRegiones), Array.from(setMeses).sort());

    if (document.getElementById('dashIngresosMes')) document.getElementById('dashIngresosMes').textContent = ingresosMes;
    if (document.getElementById('dashIngresosAnio')) document.getElementById('dashIngresosAnio').textContent = ingresosAnio;
    if (document.getElementById('dashPadrinosActivos')) document.getElementById('dashPadrinosActivos').textContent = padrinosActivos;
    if (document.getElementById('dashPlanesActivos')) document.getElementById('dashPlanesActivos').textContent = planesActivos;
    if (document.getElementById('dashPlanesFinalizados')) document.getElementById('dashPlanesFinalizados').textContent = planesFinalizados;
    if (document.getElementById('dashPlanesVencidos')) document.getElementById('dashPlanesVencidos').textContent = planesVencidos;
    if (document.getElementById('countEvalRealizadas')) document.getElementById('countEvalRealizadas').textContent = examenesRealizados;
    if (document.getElementById('countEvalPendientes')) document.getElementById('countEvalPendientes').textContent = examenesPendientes;

    const pctOnboard = tOnboardingTotal > 0 ? Math.round((tOnboardingComp / tOnboardingTotal) * 100) : 0;
    const pct7 = cant7 > 0 ? Math.round(sumaNota7 / cant7) : 0;
    const pct30 = cant30 > 0 ? Math.round(sumaNota30 / cant30) : 0;
    const pct90 = cant90 > 0 ? Math.round(sumaNota90 / cant90) : 0;

    if (document.getElementById('barPctOnboarding')) document.getElementById('barPctOnboarding').textContent = pctOnboard + "%";
    if (document.getElementById('fillOnboarding')) document.getElementById('fillOnboarding').style.width = pctOnboard + "%";
    if (document.getElementById('barPct7Dias')) document.getElementById('barPct7Dias').textContent = pct7 + "%";
    if (document.getElementById('fill7Dias')) document.getElementById('fill7Dias').style.width = pct7 + "%";
    if (document.getElementById('barPct30Dias')) document.getElementById('barPct30Dias').textContent = pct30 + "%";
    if (document.getElementById('fill30Dias')) document.getElementById('fill30Dias').style.width = pct30 + "%";
    if (document.getElementById('barPct90Dias')) document.getElementById('barPct90Dias').textContent = pct90 + "%";
    if (document.getElementById('fill90Dias')) document.getElementById('fill90Dias').style.width = pct90 + "%";
});

function poblarFiltrosEstrategicos(regiones, meses) {
    const selectReg = document.getElementById('filter-region');
    const selectMes = document.getElementById('filter-mes');
    if (selectReg && selectReg.options.length <= 1) {
        regiones.forEach(r => selectReg.innerHTML += `<option value="${r}">${r}</option>`);
    }
    if (selectMes && selectMes.options.length <= 1) {
        meses.forEach(m => {
            const [anio, mes] = m.split('-');
            const l = new Date(+anio, +mes - 1, 1).toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
            selectMes.innerHTML += `<option value="${m}">${l.charAt(0).toUpperCase() + l.slice(1)}</option>`;
        });
    }
}

// 📸 2. PROCESAMIENTO DE IMAGEN CANVAS
function procesarFotoPadrinoLocal(input) {
    const file = input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        const img = new Image();
        img.onload = function () {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            canvas.width = 150; canvas.height = 150;
            ctx.drawImage(img, 0, 0, 150, 150);
            globalPadrinoFotoBase64 = canvas.toDataURL('image/jpeg', 0.85);
            if (document.getElementById('padrinoIconoDefault')) document.getElementById('padrinoIconoDefault').style.display = 'none';
            const imgElement = document.getElementById('padrinoImgPreview');
            if (imgElement) { imgElement.src = globalPadrinoFotoBase64; imgElement.style.display = 'block'; }
        };
        img.src = e.target.result;
    };
    reader.readAsDataURL(file);
}

// 🏅 3. CERTIFICACIÓN MANUAL DE TUTORES
async function convertirYConfigurarPadrino() {
    const cedulaPadrino = document.getElementById('padrePadrinoCedula').value.trim();
    const empresa = document.getElementById('padrePadrinoEmpresa').value.trim();
    const tiempo = document.getElementById('padrePadrinoTiempo').value.trim();
    const desempeno = document.getElementById('padrePadrinoDesempeno').value.trim();
    const estado = document.getElementById('padrePadrinoEstado').value;
    const correoTutor = document.getElementById('padrePadrinoCorreo').value.trim();
    const tTecnicas = document.getElementById('padrePadrinoTecnicas').value.trim();
    const tBlandas = document.getElementById('padrePadrinoBlandas').value.trim();

    if (!cedulaPadrino || !correoTutor) return alert("Cédula y Correo son obligatorios.");

    try {
        const snapshot = await db.collection('empleados').where('cedula', '==', cedulaPadrino).limit(1).get();
        if (snapshot.empty) return alert("⚠️ Colaborador no encontrado.");

        await db.collection('empleados').doc(snapshot.docs[0].id).update({
            es_padrino: true, padrino_certified: true, padrino_estado: estado || "Activo",
            padrino_calificacion_desempeño: Number(desempeno) || 100, empresa_padrino: empresa || "Bavaria AB InBev",
            tiempo_compania: tiempo || "N/A", correo: correoTutor,
            habilidades_tecnicas: tTecnicas ? tTecnicas.split(',').map(h => h.trim()) : [],
            habilidades_blandas: tBlandas ? tBlandas.split(',').map(h => h.trim()) : [],
            foto_url: globalPadrinoFotoBase64
        });
        alert("🏅 ¡Tutor Certificado!");
        location.reload();
    } catch (error) { console.error(error); }
}

// 🚀 4. REGISTRAR NUEVO INGRESO E HISTORIAL
async function vincularPlanPadrinoNuevo() {
    const cedula = document.getElementById('padreApadrinadoId').value.trim();
    const nombre = document.getElementById('padreApadrinadoNombre').value.trim();
    const cargo = document.getElementById('padreApadrinadoCargo').value.trim();
    const area = document.getElementById('padreApadrinadoArea').value.trim();
    const region = document.getElementById('padreApadrinadoRegion').value.trim();
    const fechaIngresoStr = document.getElementById('padreApadrinadoFechaIngreso').value;
    const jefe = document.getElementById('padreApadrinadoJefe').value.trim();
    const correo = document.getElementById('padreApadrinadoCorreo').value.trim();
    const empresaAnterior = document.getElementById('padreApadrinadoEmpresaAnt').value.trim();
    const etapa = document.getElementById('padreEtapa').value;
    const padrinoId = document.getElementById('selectPadrinoDinamico').value;

    if (!cedula || !nombre || !fechaIngresoStr || !padrinoId) return alert("⚠️ Campos incompletos.");

    try {
        const formatFecha = (d) => d.toISOString().split('T')[0];
        const fechaBase = new Date(fechaIngresoStr + 'T00:00:00');
        const f7 = new Date(fechaBase); f7.setDate(f7.getDate() + 7);
        const f30 = new Date(fechaBase); f30.setDate(f30.getDate() + 30);
        const f90 = new Date(fechaBase); f90.setDate(f90.getDate() + 90);

        const snapEmpleado = await db.collection('empleados').where('cedula', '==', cedula).limit(1).get();
        const datosEstructurados = {
            cedula: cedula, nombre: nombre, cargo: cargo || "Operativo", area: area || "Operaciones",
            region: region || "General", fecha_ingreso: fechaIngresoStr, jefe_directo: jefe || "N/A", correo: correo || "", rol: "empleado",
            password: String(cedula), es_apadrinado: true, padrino_id: padrinoId, fecha_asignacion_padrino: formatFecha(new Date()),
            estado_plan_padrino: "Activo", etapa_actual: etapa, empresa_anterior: empresaAnterior || "Ninguna",
            onboarding_dia1: { estado: "Pendiente", fecha_completado: "" }, onboarding_dia2: { estado: "Pendiente", fecha_completado: "" },
            fechas_limite_evaluaciones: { eval_7_dias: formatFecha(f7), eval_30_dias: formatFecha(f30), eval_90_dias: formatFecha(f90) },
            hitos: {
                fecha_contratacion: fechaIngresoStr, onboarding_safety_nota: 0, onboarding_people_nota: 0,
                eval_tecnico_nota: 0, eval_funcional_nota: 0, eval_autonomo_nota: 0
            }
        };

        if (!snapEmpleado.empty) {
            const docId = snapEmpleado.docs[0].id;
            const empPreexistente = snapEmpleado.docs[0].data();
            const actualizacionSegmentada = {
                nombre: nombre,
                cargo: cargo || empPreexistente.cargo || "Operativo",
                area: area || empPreexistente.area || "Operaciones",
                region: region || empPreexistente.region || "General",
                fecha_ingreso: fechaIngresoStr,
                jefe_directo: jefe || "N/A",
                correo: correo || empPreexistente.correo || "",
                es_apadrinado: true,
                padrino_id: padrinoId,
                fecha_asignacion_padrino: formatFecha(new Date()),
                estado_plan_padrino: "Activo",
                etapa_actual: etapa,
                empresa_anterior: empresaAnterior || "Ninguna",
                fechas_limite_evaluaciones: { eval_7_dias: formatFecha(f7), eval_30_dias: formatFecha(f30), eval_90_dias: formatFecha(f90) }
            };
            await db.collection('empleados').doc(docId).update(actualizacionSegmentada);
        } else {
            await db.collection('empleados').add(datosEstructurados);
        }

        await db.collection('items_cronograma').add({
            empleado_id: String(padrinoId), titulo: `🤝 Nuevo alumno asignado: ${nombre}`,
            ppt_drive_id: "#", qr_url: "#", fecha_limite: formatFecha(f7), completado: false, visto_ppt: true, fecha_completado: "", tipo_item: "notificacion_asignacion"
        });

        alert("🎯 ¡Colaborador registrado exitosamente!");

        const snapPadrino = await db.collection('empleados').where('cedula', '==', String(padrinoId)).limit(1).get();
        if (!snapPadrino.empty && typeof emailjs !== "undefined") {
            const dPadrino = snapPadrino.docs[0].data();
            if (dPadrino.correo) {
                const parametrosEmail = {
                    correo_padrino: String(dPadrino.correo), nombre_padrino: String(dPadrino.nombre),
                    email: String(dPadrino.correo), nombre_alumno: String(nombre),
                    cargo_alumno: String(cargo || "Auxiliar Logístico"), fecha_limite_7dias: String(formatFecha(f7))
                };
                await emailjs.send("service_dfb29wb", "template_ovdhmr3", parametrosEmail);
            }
        }
        location.reload();
    } catch (error) { console.error(error); }
}

// 👥 5. ESCUCHADOR EN VIVO DEL MENÚ DE SELECCIÓN Y PANEL DE PADRINOS
db.collection('empleados').onSnapshot((snapshot) => {
    const containerLista = document.getElementById('lista-padrinos-seleccionables');
    const selectDinamico = document.getElementById('selectPadrinoDinamico');
    const padrinosTotales = []; const padrinosActivosDisponibles = [];

    snapshot.forEach((doc) => {
        const datos = doc.data();
        if (datos.es_padrino === true || datos.es_padrino === "true") {
            padrinosTotales.push(datos);
            if (datos.padrino_estado !== "Inactivo") padrinosActivosDisponibles.push(datos);
        }
    });

    if (selectDinamico) {
        selectDinamico.innerHTML = padrinosActivosDisponibles.length === 0
            ? `<option value="">No hay padrinos activos</option>`
            : `<option value="">-- Selecciona un tutor calificado --</option>` +
            padrinosActivosDisponibles.map(p => `<option value="${p.cedula}">${p.nombre} (${p.cargo || 'Líder'})</option>`).join('');
    }

    if (containerLista) {
        containerLista.innerHTML = padrinosTotales.map(pad => `
      <div onclick="cargarPerfilDetalladoPadrinoCompleto('${pad.cedula}')" style="display:flex; align-items:center; gap:10px; background:white; padding:10px; border-radius:8px; border:1px solid #dde3e7; cursor:pointer; transition:all 0.2s;">
        <div style="width:32px; height:32px; border-radius:50%; background:#206987; color:white; display:flex; align-items:center; justify-content:center; font-weight:bold; font-size:11px;">${pad.nombre.substring(0, 2).toUpperCase()}</div>
        <div style="flex:1;">
          <div style="font-size:12px; font-weight:700; color:#1c2430;">${pad.nombre}</div>
          <div style="font-size:10px; color:#7a8f99;">${pad.cargo || 'Supervisor'}</div>
        </div>
        <i class="ti ti-chevron-right" style="color:#7a8f99; font-size:14px;"></i>
      </div>`).join('');
    }
});

// ⚡ 6. MONITOR DETALLADO DE HITOS CON ESCUCHADOR FIRESTORE
db.collection('empleados').where('es_apadrinado', '==', true).onSnapshot((snapshot) => {
    cacheSnapshotEmpleadosLocal = snapshot;
    renderizarMonitorColaboradores();
});

// ⚡ 6. MONITOR DETALLADO DE HITOS CON PAGINACIÓN SINCRONIZADA MAESTRA Y DIPLOMA ACCESIBLE
function renderizarMonitorColaboradores() {
    if (!cacheSnapshotEmpleadosLocal) return;
    const container = document.getElementById('lista-evaluaciones-pendientes');
    const onboardingContainer = document.getElementById('contenedor-monitor-onboarding');
    if (!container) return;
    container.innerHTML = '';

    const ITEMS_POR_PAGINA = 20;
    let paginaActual = window.paginaActualMonitor || 1;

    const regFiltro = document.getElementById('filter-region').value;
    const mesFiltro = document.getElementById('filter-mes').value;
    const textoBusqueda = document.getElementById('search-colaborador')?.value.toLowerCase().trim() || '';

    let todosLosEmpleados = [];
    cacheSnapshotEmpleadosLocal.forEach((doc) => {
        const emp = doc.data();
        const docId = doc.id;

        if (miRol === 'padrino' && String(emp.padrino_id) !== String(cedulaLogueada)) return;
        if (regFiltro && emp.region !== regFiltro) return;
        if (mesFiltro && (!emp.fecha_ingreso || !emp.fecha_ingreso.startsWith(mesFiltro))) return;

        todosLosEmpleados.push({ emp, docId });
    });

    if (textoBusqueda) {
        todosLosEmpleados = todosLosEmpleados.filter(({ emp }) =>
            (emp.nombre && emp.nombre.toLowerCase().includes(textoBusqueda)) ||
            String(emp.cedula).includes(textoBusqueda)
        );
    }

    const totalPaginas = Math.ceil(todosLosEmpleados.length / ITEMS_POR_PAGINA) || 1;
    if (paginaActual > totalPaginas) paginaActual = totalPaginas;

    const inicio = (paginaActual - 1) * ITEMS_POR_PAGINA;
    const paginaSegmentada = todosLosEmpleados.slice(inicio, inicio + ITEMS_POR_PAGINA);

    if (onboardingContainer) {
        onboardingContainer.innerHTML = '';
        paginaSegmentada.forEach(({ emp, docId }) => {
            const h = emp.hitos || {};
            const edo1 = (emp.onboarding_dia1 && emp.onboarding_dia1.estado) || "Pendiente";
            const edo2 = (emp.onboarding_dia2 && emp.onboarding_dia2.estado) || "Pendiente";

            const estadoSafety = h.onboarding_safety_nota ? (h.onboarding_safety_nota >= 100 ? '✅ Aprobado' : `❌ Reprobado (${h.onboarding_safety_nota}%)`) : null;
            const estadoPeople = h.onboarding_people_nota ? (h.onboarding_people_nota >= 80 ? '✅ Aprobado' : `❌ Reprobado (${h.onboarding_people_nota}%)`) : null;

            onboardingContainer.innerHTML += `
        <div class="onboard-row">
          <div style="flex:1; min-width:200px;">
            <strong>${emp.nombre}</strong>
            <div style="font-size:11px; color:#7a8f99;">Cargo: ${emp.cargo}</div>
            ${emp.fecha_ingreso ? `<div style="font-size:11px; color:#7a8f99;">Ingreso: ${emp.fecha_ingreso}</div>` : ''}
          </div>
          <div style="display:flex; flex-direction:column; gap:4px; align-items:flex-start; min-width:140px;">
            <span style="font-size:10px; font-weight:700; color:#206987;">DÍA 1 — SAFETY</span>
            ${estadoSafety ? `<span style="font-size:11px; font-weight:600; color:${estadoSafety.includes('Aprobado') ? '#0f6e56' : '#a32d2d'}">${estadoSafety}</span>` : `<span style="font-size:11px; color:#7a8f99;">⏳ Pendiente</span>`}
          </div>
          <div style="display:flex; flex-direction:column; gap:4px; align-items:flex-start; min-width:140px;">
            <span style="font-size:10px; font-weight:700; color:#3cbcae;">DÍA 2 — PEOPLE</span>
            ${estadoPeople ? `<span style="font-size:11px; font-weight:600; color:${estadoPeople.includes('Aprobado') ? '#0f6e56' : '#a32d2d'}">${estadoPeople}</span>` : `<span style="font-size:11px; color:#7a8f99;">⏳ Pendiente</span>`}
          </div>
        </div>`;
        });

        if (todosLosEmpleados.length > 0) {
            onboardingContainer.innerHTML += `
        <div style="text-align:center; padding:8px; font-size:11.5px; color:#7a8f99; font-weight:600; background:#f8fafb; border-radius:6px; border:1px solid #dde3e7;">
          Onboarding: Mostrando segmento de la Página ${paginaActual} de ${totalPaginas} &nbsp;·&nbsp; ${todosLosEmpleados.length} en inducción
        </div>`;
        }
    }

    const buildChip = (label, nota, fecha, limite, umbral = 60) => {
        if (nota === undefined || nota === null || nota === 0) {
            return `<div class="eval-chip chip-pending"><span class="chip-label">${label}</span><span class="chip-score">—</span><span class="chip-date">Pendiente</span></div>`;
        }
        let minAprobacion = umbral;
        if (label.includes("Safety")) minAprobacion = 100;
        else if (label.includes("People")) minAprobacion = 80;

        const aprobado = nota >= minAprobacion;
        return `<div class="eval-chip ${aprobado ? 'chip-pass' : 'chip-fail'}"><span class="chip-label">${label}</span><span class="chip-score">${nota}%</span><span class="chip-date">${fecha || '—'}</span></div>`;
    };

    paginaSegmentada.forEach(({ emp, docId }) => {
        const h = emp.hitos || {};
        const lim = emp.fechas_limite_evaluaciones || {};

        let notas90 = [];
        if (h.eval_autonomo_pre_nota) notas90.push(h.eval_autonomo_pre_nota);
        if (h.eval_autonomo_nota) notas90.push(h.eval_autonomo_nota);
        if (h.eval_autonomo_post_nota) notas90.push(h.eval_autonomo_post_nota);

        let notaUnificada90 = notas90.length > 0 ? Math.round(notas90.reduce((a, b) => a + b, 0) / notas90.length) : 0;
        let fecha90Final = h.eval_autonomo_nota_fecha || h.eval_autonomo_fecha || h.eval_autonomo_pre_fecha || h.eval_autonomo_post_fecha || "";

        // 🕵️‍♂️ REGLA MAESTRA DE APROBACIÓN DE TODAS LAS ETAPAS COMPLETA
        const safetyOk = h.onboarding_safety_nota >= 100;
        const peopleOk = h.onboarding_people_nota >= 80;
        const tecnicoOk = h.eval_tecnico_nota >= 60;
        const funcionalOk = h.eval_funcional_nota >= 60;
        const autonomoOk = notaUnificada90 >= 60;

        const todoCompletado = safetyOk && peopleOk && tecnicoOk && funcionalOk && autonomoOk;

        let totalNotasValidas = 0; let sumaTotal = 0;
        if (h.onboarding_safety_nota) { sumaTotal += h.onboarding_safety_nota; totalNotasValidas++; }
        if (h.onboarding_people_nota) { sumaTotal += h.onboarding_people_nota; totalNotasValidas++; }
        if (h.eval_tecnico_nota) { sumaTotal += h.eval_tecnico_nota; totalNotasValidas++; }
        if (h.eval_funcional_nota) { sumaTotal += h.eval_funcional_nota; totalNotasValidas++; }
        if (notaUnificada90) { sumaTotal += notaUnificada90; totalNotasValidas++; }
        const promedioAvanceActual = totalNotasValidas > 0 ? Math.round(sumaTotal / totalNotasValidas) : 0;

        // 🎓 BOTÓN MUTABLE INTELIGENTE SOLICITADO
        let botonCertificadoHtml = '';
        if (todoCompletado || emp.estado_plan_padrino === "Plan Padrino Finalizado") {
            botonCertificadoHtml = `<button class="status-btn active-comp" style="font-size:12px; padding:8px 14px; background:#e1f5ee; color:#0f6e56; border-color:#3cbcae;" onclick="imprimirCertificadoCompletoPDF('${docId}')">🎓 Descargar Certificado</button>`;
        } else {
            botonCertificadoHtml = `<button class="status-btn" style="font-size:12px; padding:8px 14px; background:#f1f3f5; color:#7a8f99; cursor:not-allowed;" disabled title="Disponible al completar el 100% de las fases">⏳ Fases en Progreso (${promedioAvanceActual}%)</button>`;
        }

        container.innerHTML += `
      <div class="collab-card" style="border-left-color: ${todoCompletado ? '#3cbcae' : '#206987'}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap: wrap; gap: 10px;">
          <div>
            <div class="collab-name">${emp.nombre}</div>
            <div class="collab-meta">
              CC ${emp.cedula} &nbsp;·&nbsp; 
              Regional: ${emp.region || 'General'} &nbsp;·&nbsp; 
              Grupo: <strong style="color: #206987;">${emp.grupo || 'General'}</strong> &nbsp;·&nbsp;
              Fase: <strong>${todoCompletado ? 'Graduado 🎓' : (emp.etapa_actual || 'Inducción')}</strong>
              ${emp.fecha_ingreso ? `&nbsp;·&nbsp; Ingreso: <strong>${emp.fecha_ingreso}</strong>` : ''}
            </div>
            <div style="margin-top:8px; font-size:11px; background:${todoCompletado ? '#e1f5ee' : '#fff8e1'}; border:1px solid ${todoCompletado ? '#3cbcae' : '#ffc404'}; color:${todoCompletado ? '#0f6e56' : '#8a6e00'}; padding:4px 10px; border-radius:6px; display:inline-block; font-weight:700;">
              ${todoCompletado ? '✓ Ciclo de Maduración Completado Exitosamente' : `⏳ Promedio Actual: ${promedioAvanceActual}%`}
            </div>
          </div>
          ${botonCertificadoHtml}
        </div>
        
        <div class="evals-row">
          ${buildChip("Safety (D1)", h.onboarding_safety_nota, h.onboarding_safety_fecha, emp.fecha_ingreso, 100)}
          ${buildChip("People (D2)", h.onboarding_people_nota, h.onboarding_people_fecha, emp.fecha_ingreso, 80)}
          ${buildChip("7 Días (Téc)", h.eval_tecnico_nota, h.eval_tecnico_fecha, lim.eval_7_dias, 60)}
          ${buildChip("30 Días (Fun)", h.eval_funcional_nota, h.eval_funcional_fecha, lim.eval_30_dias, 60)}
          ${buildChip("90 Días (Aut)", notaUnificada90, fecha90Final, lim.eval_90_dias, 60)}
        </div>

        <button class="toggle-detail" onclick="toggleDetalleContenedor('items-${docId}', this)">▾ Ver Detalle de Criterios</button>
        <div class="detail-panel" id="items-${docId}">
          <div class="detail-item">Hito 7 Días: ${h.eval_tecnico_nota || 0}/100</div>
          <div class="detail-item">Hito 30 Días: ${h.eval_funcional_nota || 0}/100</div>
          <div class="detail-item">Promedio Hito 90 Días: ${notaUnificada90 || 0}/100</div>
        </div>
      </div>`;
    });

    if (todosLosEmpleados.length > 0) {
        container.innerHTML += `
      <div style="display:flex; align-items:center; justify-content:center; gap:12px; padding:16px; margin-top:8px; width:100%;">
        <button onclick="cambiarPaginaMonitor(${paginaActual - 1})" ${paginaActual <= 1 ? 'disabled' : ''} style="padding:6px 14px; border-radius:6px; border:1px solid #dde3e7; cursor:pointer; background:white; font-weight:600; font-size:12px;">← Anterior</button>
        <span style="font-size:12px; color:#7a8f99; font-weight:600;">Página ${paginaActual} de ${totalPaginas} &nbsp;·&nbsp; ${todosLosEmpleados.length} colaboradores</span>
        <button onclick="cambiarPaginaMonitor(${paginaActual + 1})" ${paginaActual >= totalPaginas ? 'disabled' : ''} style="padding:6px 14px; border-radius:6px; border:1px solid #dde3e7; cursor:pointer; background:white; font-weight:600; font-size:12px;">Siguiente →</button>
      </div>`;
    }
}

function toggleDetalleContenedor(id, btn) {
    const panel = document.getElementById(id);
    if (!panel) return;
    panel.classList.toggle('open');
    btn.textContent = panel.classList.contains('open') ? '▴ Ocultar Detalle' : '▾ Ver Detalle de Criterios';
}

// 📂 7. RECONOCEDOR EXCEL CON PARSEO DE MATRICES SHEETJS
async function procesarCargaMasivaExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

            const rawRows = rows.slice(1);
            let totalProcesados = 0;

            for (const row of rawRows) {
                const ccColaborador = String(row[15] || '').trim();
                const fechaEjecucion = row[1] instanceof Date ? row[1].toISOString().split('T')[0] : new Date().toISOString().split('T')[0];

                if (!ccColaborador || ccColaborador === 'null') continue;

                const scores = scoreExcelDinamico(row);
                const techScore = scores.tech;
                const peopleScore = scores.people;
                const avgGeneral = techScore && peopleScore ? Math.round((techScore.pct + peopleScore.pct) / 2) : (techScore || peopleScore)?.pct || 0;

                let snap = await db.collection('empleados').where('cedula', '==', ccColaborador).limit(1).get();
                if (snap.empty) snap = await db.collection('empleados').where('cedula', '==', Number(ccColaborador)).limit(1).get();

                if (!snap.empty) {
                    const docId = snap.docs[0].id;
                    const emp = snap.docs[0].data();
                    const hitos = emp.hitos || {};
                    let up = {};

                    const cicloKey = resolverCicloExcel(row[17]);

                    if (cicloKey === '7') {
                        hitos.eval_tecnico_nota = avgGeneral;
                        hitos.eval_tecnico_fecha = fechaEjecucion;
                        hitos.onboarding_safety_nota = techScore?.pct || 100;
                        hitos.onboarding_safety_fecha = fechaEjecucion;
                        up.etapa_actual = "Etapa Funcional";
                    } else if (cicloKey === '30') {
                        hitos.eval_funcional_nota = avgGeneral;
                        hitos.eval_funcional_fecha = fechaEjecucion;
                        hitos.onboarding_people_nota = peopleScore?.pct || 100;
                        hitos.onboarding_people_fecha = fechaEjecucion;
                        up.etapa_actual = "Etapa Autónomo";
                    } else if (cicloKey === '90_pre') {
                        hitos.eval_autonomo_pre_nota = avgGeneral;
                        hitos.eval_autonomo_pre_fecha = fechaEjecucion;
                    } else if (cicloKey === '90') {
                        hitos.eval_autonomo_nota = avgGeneral;
                        hitos.eval_autonomo_fecha = fechaEjecucion;
                        if ((hitos.eval_autonomo_pre_nota || 0) > 0 && (hitos.eval_autonomo_post_nota || 0) > 0) up.estado_plan_padrino = "Plan Padrino Finalizado";
                    } else if (cicloKey === '90_post') {
                        hitos.eval_autonomo_post_nota = avgGeneral;
                        hitos.eval_autonomo_post_fecha = fechaEjecucion;
                        if ((hitos.eval_autonomo_pre_nota || 0) > 0 && (hitos.eval_autonomo_nota || 0) > 0) up.estado_plan_padrino = "Plan Padrino Finalizado";
                    } else if (cicloKey === 'retro') {
                        hitos.retro_people_nota = peopleScore?.pct || 0;
                        hitos.retro_nps = scores.nps || 0;
                        hitos.retro_fecha = fechaEjecucion;
                    }

                    up.hitos = hitos;
                    await db.collection('empleados').doc(docId).update(up);
                    totalProcesados++;
                }
            }
            alert(`🤖 PROCESAMIENTO COMPLETO:\n\nSe emparejaron ${totalProcesados} cuestionarios.`);
            location.reload();
        } catch (err) { console.error(err); }
    };
    reader.readAsArrayBuffer(file);
}

function scoreExcelDinamico(row) {
    let lastIdx = row.length - 1;
    while (lastIdx >= 0 && (row[lastIdx] === null || row[lastIdx] === undefined || row[lastIdx] === '')) lastIdx--;
    return { tech: scoreExcelRange(row, 18, lastIdx - 6), people: scoreExcelRange(row, lastIdx - 5, lastIdx - 1), nps: row[lastIdx] };
}

async function procesarOnboardingExcel(event, tipoOnboarding) {
    const file = event.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
            const headers = rows[0] || []; const rawRows = rows.slice(1);

            let idxTotalPuntos = headers.findIndex(h => {
                const txt = String(h || '').trim().toLowerCase();
                return txt === 'total de puntos' || txt === 'puntos' || txt === 'puntos totales';
            });
            let idxCedula = headers.findIndex(h => {
                const txt = String(h || '').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
                return txt.includes('identificacion') || txt.includes('cedula') || txt.includes('documento');
            });
            if (idxCedula === -1) idxCedula = 34;

            let totalProcesados = 0; const umbral = tipoOnboarding === 'safety' ? 100 : 80;

            for (const row of rawRows) {
                let cedulaRaw = String(row[idxCedula] || '').trim().replace(/['"\s.\-,]/g, '');
                if (!cedulaRaw || cedulaRaw === 'null' || cedulaRaw === '0') continue;

                let notaDirecta = idxTotalPuntos !== -1 ? Number(row[idxTotalPuntos] || 0) : Number(row[5] || 0);
                let pct = Math.min(100, Math.max(0, Math.round(notaDirecta)));
                const aprobado = pct >= umbral;
                const fechaEjecucion = row[1] instanceof Date ? row[1].toISOString().split('T')[0] : new Date().toISOString().split('T')[0];

                let snap = await db.collection('empleados').where('cedula', '==', cedulaRaw).limit(1).get();
                if (snap.empty) snap = await db.collection('empleados').where('cedula', '==', Number(cedulaRaw)).limit(1).get();

                if (!snap.empty) {
                    const docId = snap.docs[0].id; const hitos = snap.docs[0].data().hitos || {}; let up = {};
                    if (tipoOnboarding === 'safety') {
                        hitos.onboarding_safety_nota = pct; hitos.onboarding_safety_fecha = fechaEjecucion; hitos.onboarding_safety_aprobado = aprobado;
                        up.onboarding_dia1 = { estado: aprobado ? 'Completado' : 'Reprobado', fecha_completado: fechaEjecucion };
                    } else {
                        hitos.onboarding_people_nota = pct; hitos.onboarding_people_fecha = fechaEjecucion; hitos.onboarding_people_aprobado = aprobado;
                        up.onboarding_dia2 = { estado: aprobado ? 'Completado' : 'Reprobado', fecha_completado: fechaEjecucion };
                    }
                    up.hitos = hitos; await db.collection('empleados').doc(docId).set(up, { merge: true }); totalProcesados++;
                }
            }
            alert(`✅ Consolidación Onboarding exitosa.`); location.reload();
        } catch (err) { console.error(err); }
    };
    reader.readAsArrayBuffer(file);
}

function cambiarPaginaMonitor(nuevaPagina) {
    window.paginaActualMonitor = nuevaPagina;
    renderizarMonitorColaboradores();
}

function scoreExcelRange(row, start, end) {
    let total = 0, si = 0;
    for (let i = start; i <= end; i++) {
        const v = row[i];
        if (v !== null && v !== undefined && v !== '') { total++; if (String(v).toLowerCase().startsWith('si') || String(v).toLowerCase().startsWith('sí')) si++; }
    }
    return total > 0 ? { pct: Math.round((si / total) * 100), si, total } : null;
}

async function cambiarEstadoOnboarding(docId, campo, estado) {
    const hoyStr = new Date().toISOString().split('T')[0];
    await db.collection('empleados').doc(docId).update({ [campo]: { estado: estado, fecha_completado: estado === "Completado" ? hoyStr : "" } });
}



// 🎓 AMBIENTE B: ENLACE FIRESTORE REAL (Lee la data viva de Vladimir desde la nube)
// 🎓 FUNCIÓN MAESTRA CONECTADA A FIRESTORE — CORREGIDA CON TIEMPO DE ESPERA PARA EVITAR LIENZO EN BLANCO
async function imprimirCertificadoCompletoPDF(docId) {
    try {
        console.log(`🔍 Consultando Firebase para la cédula o documento: ${docId}`);

        let snap = await db.collection('empleados').doc(docId).get();
        let emp = snap.exists ? snap.data() : null;

        if (!emp) {
            const querySnap = await db.collection('empleados').where('cedula', '==', String(docId)).limit(1).get();
            if (!querySnap.empty) emp = querySnap.docs[0].data();
        }

        if (!emp) {
            return alert(`❌ ERROR: No se encontró registro real en Firebase para: "${docId}"`);
        }

        const h = emp.hitos || {};
        let notas90 = [];
        if (h.eval_autonomo_pre_nota) notas90.push(h.eval_autonomo_pre_nota);
        if (h.eval_autonomo_nota) notas90.push(h.eval_autonomo_nota);
        if (h.eval_autonomo_post_nota) notas90.push(h.eval_autonomo_post_nota);
        let n90Unificada = notas90.length > 0 ? Math.round(notas90.reduce((a, b) => a + b, 0) / notas90.length) : 0;

        const nSafety = h.onboarding_safety_nota || 0;
        const nPeople = h.onboarding_people_nota || 0;
        const n7      = h.eval_tecnico_nota       || 0;
        const n30     = h.eval_funcional_nota     || 0;
        const n90     = n90Unificada;

        let totalNotasValidas = 0; let sumaTotal = 0;
        if (nSafety > 0) { sumaTotal += nSafety; totalNotasValidas++; }
        if (nPeople > 0) { sumaTotal += nPeople; totalNotasValidas++; }
        if (n7 > 0)      { sumaTotal += n7;      totalNotasValidas++; }
        if (n30 > 0)     { sumaTotal += n30;     totalNotasValidas++; }
        if (n90 > 0)     { sumaTotal += n90;     totalNotasValidas++; }

        const promedioHistorico = totalNotasValidas > 0 ? Math.round(sumaTotal / totalNotasValidas) : 0;
        const distincionConceptual = promedioHistorico >= 90 ? "SOBRESALIENTE 🌟" : "SATISFACTORIO";

        let nombrePadrinoDoc = "Padrino Tutor Certificado";
        if (emp.padrino_id) {
            const tutorSnap = await db.collection('empleados').where('cedula', '==', String(emp.padrino_id)).limit(1).get();
            if (!tutorSnap.empty) nombrePadrinoDoc = tutorSnap.docs[0].data().nombre;
        }

        // 🏗️ Contenedor trampa: oculta al usuario pero html2canvas sí puede renderizarlo
        const contenedor = document.createElement('div');
        contenedor.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 0;
            height: 0;
            overflow: hidden;
            z-index: -1;
        `;

        const elementoDiploma = document.createElement('div');
        elementoDiploma.style.width = '800px';
        elementoDiploma.style.padding = '35px';
        elementoDiploma.style.background = '#ffffff';
        elementoDiploma.style.fontFamily = 'sans-serif';
        elementoDiploma.style.color = '#1c2430';

        elementoDiploma.innerHTML = `
          <div style="border: 8px double #206987; padding: 30px; border-radius: 12px; background: #fafbfc; box-sizing: border-box;">
            <div style="width: 100%; height: 6px; background: #3cbcae; margin-bottom: 20px;"></div>
            <div style="text-align: center; margin-bottom: 20px;">
              <span style="font-size: 11px; font-weight: 800; color: #206987; letter-spacing: 0.15em; text-transform: uppercase;">Portal People · Logística Conectamos Más</span>
              <h1 style="font-size: 24px; color: #206987; margin-top: 5px; font-weight: 800;">CERTIFICADO DE INDUCCIÓN Y MADURACIÓN</h1>
              <div style="width: 80px; height: 3px; background: #ffc404; margin: 8px auto;"></div>
            </div>
            <div style="text-align: center; margin-bottom: 25px;">
              <h2 style="font-size: 28px; color: #1c2430; font-weight: 700; margin: 0; text-transform: uppercase;">${emp.nombre}</h2>
              <p style="font-size: 12px; color: #7a8f99; margin-top: 4px;">Número de Identificación: <strong>${emp.cedula}</strong></p>
            </div>
            <p style="text-align: center; font-size: 13px; color: #1c2430; line-height: 1.6; max-width: 650px; margin: 0 auto 30px auto;">
              Por el avance de maduración registrado dentro de su ciclo de inducción y cargo de <strong>${emp.cargo || 'Operativo'}</strong>.
            </p>
            <div style="background: #ffffff; border: 1.5px solid #dde3e7; border-radius: 10px; padding: 15px; margin-bottom: 30px;">
              <div style="display: flex; justify-content: space-around; align-items: center; text-align: center;">
                <div><div style="font-size: 9px; color: #7a8f99;">SAFETY (D1)</div><div style="font-size: 14px; font-weight: 700;">${nSafety > 0 ? nSafety + '%' : '—'}</div></div>
                <div><div style="font-size: 9px; color: #7a8f99;">PEOPLE (D2)</div><div style="font-size: 14px; font-weight: 700;">${nPeople > 0 ? nPeople + '%' : '—'}</div></div>
                <div><div style="font-size: 9px; color: #7a8f99;">7 DÍAS (TÉC)</div><div style="font-size: 14px; font-weight: 700;">${n7 > 0 ? n7 + '%' : '—'}</div></div>
                <div><div style="font-size: 9px; color: #7a8f99;">30 DÍAS (FUN)</div><div style="font-size: 14px; font-weight: 700;">${n30 > 0 ? n30 + '%' : '—'}</div></div>
                <div><div style="font-size: 9px; color: #7a8f99;">90 DÍAS (AUT)</div><div style="font-size: 14px; font-weight: 700;">${n90 > 0 ? n90 + '%' : '—'}</div></div>
                <div style="background: #f4fbf9; padding: 6px 14px; border-radius: 8px;">
                  <div style="font-size: 9px; color: #0f6e56; font-weight: 700;">PROMEDIO</div>
                  <div style="font-size: 20px; font-weight: 800; color: #0f6e56;">${promedioHistorico}%</div>
                </div>
              </div>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: flex-end; padding: 0 10px;">
              <div style="text-align: center; width: 220px;">
                <div style="font-size: 12px; font-weight: 700; border-top: 1.5px solid #dde3e7; padding-top: 6px;">${nombrePadrinoDoc}</div>
              </div>
              <div style="text-align: center;">
                <div style="font-size: 11px; font-weight: 700; color: #0f6e56; background: #e1f5ee; padding: 5px 12px; border-radius: 6px;">${distincionConceptual}</div>
              </div>
              <div style="text-align: center; width: 220px;">
                <div style="font-size: 12px; font-weight: 700; color: #206987; border-top: 1.5px solid #dde3e7; padding-top: 6px;">Área de People</div>
              </div>
            </div>
          </div>`;

        contenedor.appendChild(elementoDiploma);
        document.body.appendChild(contenedor);

        setTimeout(async () => {
            try {
                const opciones = {
                    margin: 10,
                    filename: `Certificado_${emp.cedula}_REAL.pdf`,
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: {
                        scale: 2,
                        useCORS: true,
                        logging: false,
                        scrollX: 0,
                        scrollY: 0
                    },
                    jsPDF: { unit: 'mm', format: 'letter', orientation: 'landscape' }
                };

                await html2pdf().set(opciones).from(elementoDiploma).save();
                console.log("🎉 ¡Certificado Real descargado con éxito!");
            } catch (errPDF) {
                console.error("❌ Error generando el PDF:", errPDF);
            } finally {
                document.body.removeChild(contenedor); // siempre limpia el DOM
            }
        }, 600);

    } catch (e) {
        console.error("❌ Error crítico compilando el lienzo PDF:", e);
    }
}