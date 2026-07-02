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

// 🇨🇴 Matriz Oficial Unificada de Días Festivos en Colombia (2026 - 2027)
const FESTIVOS_COLOMBIA = [
    "2026-01-01", "2026-01-12", "2026-03-23", "2026-04-02", "2026-04-03",
    "2026-05-01", "2026-05-18", "2026-06-08", "2026-06-15", "2026-06-29",
    "2026-07-20", "2026-08-07", "2026-08-17", "2026-10-12", "2026-11-02",
    "2026-11-16", "2026-12-08", "2026-12-25",
    "2027-01-01", "2027-01-11", "2027-03-22", "2027-03-25", "2027-03-26",
    "2027-05-01", "2027-05-10", "2027-05-31", "2027-06-07", "2027-07-05",
    "2027-07-20", "2027-08-07", "2027-08-16", "2027-10-18", "2027-11-01",
    "2027-11-15", "2027-12-08", "2027-12-25"
];

function esDiaInhabilColombia(fechaObjeto) {
    if (fechaObjeto.getDay() === 0) return true; // Domingo
    const formatoISO = fechaObjeto.toISOString().split('T')[0];
    return FESTIVOS_COLOMBIA.includes(formatoISO);
}

function calcularFechaLimiteHabilesColombia(fechaIngresoStr, diasHabilesAsumidos) {
    if (!fechaIngresoStr || fechaIngresoStr === "—") return null;
    let fechaActual = new Date(fechaIngresoStr + 'T00:00:00');
    let diasContados = 0;
    while (diasContados < diasHabilesAsumidos) {
        fechaActual.setDate(fechaActual.getDate() + 1);
        if (!esDiaInhabilColombia(fechaActual)) {
            diasContados++;
        }
    }
    return fechaActual.toISOString().split('T')[0];
}

function calcularFechaLimiteCalendarioISO(fechaIngresoStr, dias) {
    if (!fechaIngresoStr || fechaIngresoStr === "—") return null;
    const d = new Date(fechaIngresoStr + 'T00:00:00');
    d.setDate(d.getDate() + dias);
    return d.toISOString().split('T')[0];
}

function calcularFechaBaseConNovedades(fechaIngresoStr, diasAcumulados) {
    const dias = parseInt(diasAcumulados) || 0;
    if (dias <= 0 || !fechaIngresoStr) return fechaIngresoStr;
    const d = new Date(fechaIngresoStr + 'T00:00:00');
    d.setDate(d.getDate() + dias);
    return d.toISOString().split('T')[0];
}

// 🧠 RESOLVEDOR DE CICLOS EN VIVO EXCEL
function resolverCicloExcel(valorCelda) {
    if (!valorCelda) return '7';
    const texto = String(valorCelda).toLowerCase();
    if (texto.includes('7') || texto.includes('técnico') || texto.includes('tecnico')) return '7';
    if (texto.includes('30') || texto.includes('funcional')) return '30';
    if (texto.includes('90') || texto.includes('autónomo') || texto.includes('autonomo')) {
        if (texto.includes('pre')) return '90_pre';
        if (texto.includes('post')) return '90_post';
        return '90';
    }
    return '7';
}

const miRol = localStorage.getItem('usuario_rol');
const cedulaLogueada = localStorage.getItem('usuario_cedula');

if (typeof emailjs !== "undefined") emailjs.init("JkzXt8TMr--AXQsiM");

if (!miRol) {
    window.location.href = 'login.html';
}

window.addEventListener('DOMContentLoaded', () => {
    ejecutarRestriccionDeRoles();

    const inputBusqueda = document.getElementById('search-colaborador');
    if (inputBusqueda) {
        inputBusqueda.addEventListener('input', () => {
            window.paginaActualMonitor = 1;
            renderizarMonitorColaboradores();
        });
    }

    const inputCedula = document.getElementById('padreApadrinadoId');
    if (inputCedula) {
        inputCedula.addEventListener('blur', async () => {
            const cedulaBuscada = inputCedula.value.trim();
            if (!cedulaBuscada) return;

            try {
                let snapshot = await db.collection('empleados').where('cedula', '==', cedulaBuscada).limit(1).get();
                if (snapshot.empty && !isNaN(cedulaBuscada)) {
                    snapshot = await db.collection('empleados').where('cedula', '==', Number(cedulaBuscada)).limit(1).get();
                }

                if (!snapshot.empty) {
                    const empData = snapshot.docs[0].data();
                    if (document.getElementById('padreApadrinadoNombre')) document.getElementById('padreApadrinadoNombre').value = empData.nombre || '';
                    if (document.getElementById('padreApadrinadoCargo')) document.getElementById('padreApadrinadoCargo').value = empData.cargo || '';
                    if (document.getElementById('padreApadrinadoArea')) document.getElementById('padreApadrinadoArea').value = empData.area || '';
                    if (document.getElementById('padreApadrinadoRegion')) document.getElementById('padreApadrinadoRegion').value = empData.region || '';
                    if (document.getElementById('padreApadrinadoCorreo')) document.getElementById('padreApadrinadoCorreo').value = empData.correo || '';

                    const campoFecha = document.getElementById('padreApadrinadoFechaIngreso');
                    if (campoFecha) {
                        campoFecha.value = empData.fecha_ingreso || (empData.hitos && empData.hitos.fecha_contratacion) || '';
                    }
                    inputCedula.style.borderColor = '#0f6e56';
                } else {
                    inputCedula.style.borderColor = '';
                }
            } catch (error) { console.error(error); }
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

let popupMostradoEnSesionActual = false;

db.collection('empleados').onSnapshot((snapshot) => {
    cacheSnapshotEmpleadosLocal = snapshot;

    let ingresosMes = 0; let ingresosAnio = 0;
    let padrinosActivos = 0;
    let planesActivos = 0; let planesFinalizados = 0; let planesVencidos = 0;
    let tOnboardingTotal = 0; let tOnboardingComp = 0;
    let examenesRealizados = 0; let examenesPendientes = 0;
    let sumaNota7 = 0; let sumaNota30 = 0; let sumaNota90 = 0; let totalApadrinados = 0;
    let totalEvaluadosOnboarding = 0; let arrayReprobadosHTML = [];
    let arrayVencidosHTML = [];
    let listadoVencidosPopupHTML = []; let listadoProximosPopupHTML = [];

    const padrinosTotales = [];
    const padrinosActivosDisponibles = [];
    const hoy = new Date();
    const setRegiones = new Set();
    const setMeses = new Set();

    const calcularFechaLimiteISO = (fechaIngresoStr, dias) => {
        if (!fechaIngresoStr || fechaIngresoStr === "—") return null;
        const d = new Date(fechaIngresoStr + 'T00:00:00');
        d.setDate(d.getDate() + dias);
        return d.toISOString().split('T')[0];
    };

    snapshot.forEach((doc) => {
        const emp = doc.data();
        if (emp.region) setRegiones.add(emp.region.trim());
        if (emp.fecha_ingreso) setMeses.add(emp.fecha_ingreso.substring(0, 7));

        if (emp.es_padrino === true || emp.es_padrino === "true") {
            padrinosTotales.push(emp);
            if (emp.padrino_estado !== "Inactivo") {
                padrinosActivosDisponibles.push(emp);
                padrinosActivos++;
            }
        }

        if (emp.es_apadrinado === true) {
            if (miRol === 'padrino' && String(emp.padrino_id) !== String(cedulaLogueada)) return;
            totalApadrinados++;

            const fBase = emp.fecha_ingreso || (emp.hitos && emp.hitos.fecha_contratacion) || "";
            // 🌟 AÑADIDO: fecha auxiliar solo para cálculos de plazos, NO se muestra ni se guarda como fecha_ingreso
            const fBaseParaPlazos = calcularFechaBaseConNovedades(fBase, emp.dias_incapacidad_acumulados);

            if (emp.estado_plan_padrino === "Plan Padrino Finalizado") {
                planesFinalizados++;
            } else {
                planesActivos++;
                if (emp.fechas_limite_evaluaciones) {
                    const f7 = new Date(emp.fechas_limite_evaluaciones.eval_7_dias + 'T00:00:00');
                    const f30 = new Date(emp.fechas_limite_evaluaciones.eval_30_dias + 'T00:00:00');
                    const f90 = new Date(emp.fechas_limite_evaluaciones.eval_90_dias + 'T00:00:00');

                    let etapaVencidaLabel = null; let fechaLimiteVencida = null;
                    if (emp.etapa_actual === "Etapa Técnico" && hoy > f7) { planesVencidos++; etapaVencidaLabel = "7 Días (Técnico)"; fechaLimiteVencida = emp.fechas_limite_evaluaciones.eval_7_dias; }
                    else if (emp.etapa_actual === "Etapa Funcional" && hoy > f30) { planesVencidos++; etapaVencidaLabel = "30 Días (Funcional)"; fechaLimiteVencida = emp.fechas_limite_evaluaciones.eval_30_dias; }
                    else if (emp.etapa_actual === "Etapa Autónomo" && hoy > f90) { planesVencidos++; etapaVencidaLabel = "90 Días (Autónomo)"; fechaLimiteVencida = emp.fechas_limite_evaluaciones.eval_90_dias; }

                    if (etapaVencidaLabel) {
                        const diasVencido = Math.floor((hoy - new Date(fechaLimiteVencida + 'T00:00:00')) / (1000 * 60 * 60 * 24));
                        arrayVencidosHTML.push(`
        <div class="reprobado-row-item">
            <div style="flex: 1;">
                <strong style="font-size:12px; color:#1c2430;">${emp.nombre}</strong>
                <div style="font-size:10.5px; color:#7a8f99;">CC ${emp.cedula} · CD: ${emp.region || 'General'}</div>
            </div>
            <div style="text-align: right;">
                <span style="font-size:10.5px; font-weight:700; color:#a32d2d; background:#fff0f0; padding:3px 8px; border-radius:4px; border:1px solid #f5c2c2;">
                    ⚠️ ${etapaVencidaLabel} · Venció: ${fechaLimiteVencida} (${diasVencido}d)
                </span>
            </div>
        </div>`);
                    }
                }
            }

            if (fBase) {
                if (fBase.startsWith("2026-06")) ingresosMes++;
                if (fBase.startsWith("2026")) ingresosAnio++;
            }

            tOnboardingTotal += 2;
            if (emp.onboarding_dia1 && emp.onboarding_dia1.estado === "Completado") tOnboardingComp++;
            if (emp.onboarding_dia2 && emp.onboarding_dia2.estado === "Completado") tOnboardingComp++;

            if (emp.hitos) {
                const nSafety = emp.hitos.onboarding_safety_nota || 0;
                const nPeople = emp.hitos.onboarding_people_nota || 0;
                const n7 = emp.hitos.eval_tecnico_nota || 0;
                const n30 = emp.hitos.eval_funcional_nota || 0;
                const n90 = emp.hitos.eval_autonomo_nota || 0;

                if (fBaseParaPlazos) { // 🌟 CAMBIADO: antes era "if (fBase)"
                    const limiteSafetyStr = calcularFechaLimiteHabilesColombia(fBaseParaPlazos, 0);       // 🌟 CAMBIADO
                    const limitePeopleStr = calcularFechaLimiteHabilesColombia(fBaseParaPlazos, 1);       // 🌟 CAMBIADO
                    const limiteTecnicoStr = calcularFechaLimiteHabilesColombia(fBaseParaPlazos, 7);      // 🌟 CAMBIADO
                    const limiteFuncionalStr = calcularFechaLimiteISO(fBaseParaPlazos, 32);               // 🌟 CAMBIADO
                    const limiteAutonomoStr = calcularFechaLimiteISO(fBaseParaPlazos, 92);                // 🌟 CAMBIADO

                    const hitosPlazos = [
                        { nombre: "Safety (D1)", nota: nSafety, limiteCalculado: limiteSafetyStr },
                        { nombre: "People (D2)", nota: nPeople, limiteCalculado: limitePeopleStr },
                        { nombre: "7 Días (Téc)", nota: n7, limiteCalculado: limiteTecnicoStr },
                        { nombre: "30 Días (Fun)", nota: n30, limiteCalculado: limiteFuncionalStr },
                        { nombre: "90 Días (Aut)", nota: n90, limiteCalculado: limiteAutonomoStr }
                    ];

                    hitosPlazos.forEach(p => {
                        const limiteVal = p.limiteCalculado;
                        if (!limiteVal) return;

                        const limiteFecha = new Date(limiteVal + 'T00:00:00');
                        const diferenciaMilisegundos = limiteFecha - hoy;
                        const diasRestantes = Math.ceil(diferenciaMilisegundos / (1000 * 60 * 60 * 24));

                        if (p.nota === 0) {
                            if (hoy > limiteFecha) {
                                listadoVencidosPopupHTML.push(`
                                <div class="popup-item-alerta popup-bg-rojo" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; border-radius: 8px; font-size: 12px; background: #fff5f5; border: 1px solid #f09595; margin-bottom: 5px;">
                                    <div><strong>${emp.nombre}</strong><br><span style="font-size:10px; color:#a32d2d;">Hito pendiente: ${p.nombre}</span></div>
                                    <div style="text-align:right; font-weight:700; color:#a32d2d;">Venció: ${limiteVal}</div>
                                </div>`);
                            } else if (diasRestantes >= 0 && diasRestantes <= 3) {
                                const textoDias = diasRestantes === 0 ? "¡Vence hoy!" : `Vence en ${diasRestantes} días`;
                                listadoProximosPopupHTML.push(`
                                <div class="popup-item-alerta popup-bg-amarillo" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; border-radius: 8px; font-size: 12px; background: #fff8e1; border: 1px solid #ffc404; margin-bottom: 5px;">
                                    <div><strong>${emp.nombre}</strong><br><span style="font-size:10px; color:#8a6e00;">Examen por rendir: ${p.nombre}</span></div>
                                    <div style="text-align:right; font-weight:700; color:#8a6e00;">${textoDias} (${limiteVal})</div>
                                </div>`);
                            }
                        }
                    });
                }

                let cargoFallas = [];
                if (nSafety > 0 && nSafety < 100) cargoFallas.push(`Safety (${nSafety}%)`);
                if (nPeople > 0 && nPeople < 80) cargoFallas.push(`People (${nPeople}%)`);
                if (nSafety > 0 || nPeople > 0) totalEvaluadosOnboarding++;

                if (cargoFallas.length > 0) {
                    arrayReprobadosHTML.push(`
                    <div class="reprobado-row-item">
                        <div style="flex: 1;">
                            <strong style="font-size:12px; color:#1c2430;">${emp.nombre}</strong>
                            <div style="font-size:10.5px; color:#7a8f99;">CC ${emp.cedula} · CD: ${emp.region || 'General'}</div>
                        </div>
                        <div style="text-align: right;">
                            <span style="font-size:10.5px; font-weight:700; color:#a32d2d; background:#fff0f0; padding:3px 8px; border-radius:4px; border:1px solid #f5c2c2;">❌ ${cargoFallas.join(' y ')}</span>
                        </div>
                    </div>`);
                }

                if (emp.hitos.eval_tecnico_nota > 0) { sumaNota7 += emp.hitos.eval_tecnico_nota; examenesRealizados++; } else { examenesPendientes++; }
                if (emp.hitos.eval_funcional_nota > 0) { sumaNota30 += emp.hitos.eval_funcional_nota; examenesRealizados++; } else { examenesPendientes++; }

                let n90Unificada = 0; let n90Arr = [];
                if (emp.hitos.eval_autonomo_pre_nota) n90Arr.push(emp.hitos.eval_autonomo_pre_nota);
                if (emp.hitos.eval_autonomo_nota) n90Arr.push(emp.hitos.eval_autonomo_nota);
                if (emp.hitos.eval_autonomo_post_nota) n90Arr.push(emp.hitos.eval_autonomo_post_nota);
                if (n90Arr.length > 0) n90Unificada = Math.round(n90Arr.reduce((a, b) => a + b, 0) / n90Arr.length);
                if (n90Unificada > 0) { sumaNota90 += n90Unificada; }
            }
        }
    });

    poblarFiltrosEstrategicos(Array.from(setRegiones), Array.from(setMeses).sort());

    const selectDinamico = document.getElementById('selectPadrinoDinamico');
    const containerLista = document.getElementById('lista-padrinos-seleccionables');

    if (selectDinamico) {
        selectDinamico.innerHTML = padrinosActivosDisponibles.length === 0
            ? `<option value="">No hay padrinos activos</option>`
            : `<option value="">-- Selecciona un tutor calificado --</option>` +
            padrinosActivosDisponibles.map(p => `<option value="${p.cedula}">${p.nombre} (${p.cargo || 'Líder'})</option>`).join('');
    }

    if (containerLista) {
        containerLista.innerHTML = padrinosTotales.length === 0
            ? `<div style="text-align:center; padding:15px; color:#7a8f99; font-size:12px;">No hay tutores registrados en el sistema aún.</div>`
            : padrinosTotales.map(pad => `
              <div onclick="cargarPerfilDetalladoPadrinoCompleto('${pad.cedula}')" style="display:flex; align-items:center; gap:10px; background:white; padding:10px; border-radius:8px; border:1px solid #dde3e7; cursor:pointer; transition:all 0.2s; margin-bottom: 6px;">
                <div style="width:32px; height:32px; border-radius:50%; background:#206987; color:white; display:flex; align-items:center; justify-content:center; font-weight:bold; font-size:11px; flex-shrink:0;">${pad.nombre.substring(0, 2).toUpperCase()}</div>
                <div style="flex:1; min-width:0;">
                  <div style="font-size:12px; font-weight:700; color:#1c2430; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${pad.nombre}</div>
                  <div style="font-size:10px; color:#7a8f99; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${pad.cargo || 'Supervisor'}</div>
                </div>
                <i class="ti ti-chevron-right" style="color:#7a8f99; font-size:14px; flex-shrink:0;"></i>
              </div>`).join('');
    }

    renderizarMonitorColaboradores();

    if (!popupMostradoEnSesionActual) {
        const modal = document.getElementById('modal-alertas-bienvenida');
        const contenedorVencidos = document.getElementById('popup-lista-vencidos');
        const contenedorProximos = document.getElementById('popup-lista-proximos');
        let tieneAlertasVivas = false;

        if (listadoVencidosPopupHTML.length > 0 && contenedorVencidos) {
            contenedorVencidos.innerHTML = listadoVencidosPopupHTML.join('');
            const seccionV = document.getElementById('popup-seccion-vencidos');
            if (seccionV) seccionV.style.display = 'flex';
            tieneAlertasVivas = true;
        }
        if (listadoProximosPopupHTML.length > 0 && contenedorProximos) {
            contenedorProximos.innerHTML = listadoProximosPopupHTML.join('');
            const seccionP = document.getElementById('popup-seccion-proximos');
            if (seccionP) seccionP.style.display = 'flex';
            tieneAlertasVivas = true;
        }
        if (tieneAlertasVivas && modal) modal.style.display = 'flex';
        popupMostradoEnSesionActual = true;
    }

    if (document.getElementById('dashTotalReprobados')) {
        document.getElementById('dashTotalReprobados').textContent = arrayReprobadosHTML.length;
        document.getElementById('dashTotalReprobados').style.background = arrayReprobadosHTML.length > 0 ? '#FCEBEB' : '#e1f5ee';
        document.getElementById('dashTotalReprobados').style.color = arrayReprobadosHTML.length > 0 ? '#a32d2d' : '#0f6e56';
    }
    if (document.getElementById('dashEvaluadosOnboarding')) document.getElementById('dashEvaluadosOnboarding').textContent = totalEvaluadosOnboarding;
    if (document.getElementById('lista-desplegable-reprobados')) {
        document.getElementById('lista-desplegable-reprobados').innerHTML = arrayReprobadosHTML.length === 0
            ? `<div style="text-align:center; padding:15px; color:#0f6e56; font-size:11.5px; font-weight:600; background:#e1f5ee; border-radius:6px;">✅ Todos los colaboradores se encuentran aprobados al día.</div>`
            : arrayReprobadosHTML.join('');
    }
    if (document.getElementById('lista-desplegable-vencidos')) {
    document.getElementById('lista-desplegable-vencidos').innerHTML = arrayVencidosHTML.length === 0
        ? `<div style="text-align:center; padding:15px; color:#0f6e56; font-size:11.5px; font-weight:600; background:#e1f5ee; border-radius:6px;">✅ No hay planes vencidos actualmente.</div>`
        : arrayVencidosHTML.join('');
}

    if (document.getElementById('dashIngresosMes')) document.getElementById('dashIngresosMes').textContent = ingresosMes;
    if (document.getElementById('dashIngresosAnio')) document.getElementById('dashIngresosAnio').textContent = ingresosAnio;
    if (document.getElementById('dashPadrinosActivos')) document.getElementById('dashPadrinosActivos').textContent = padrinosActivos;
    if (document.getElementById('dashPlanesActivos')) document.getElementById('dashPlanesActivos').textContent = planesActivos;
    if (document.getElementById('dashPlanesFinalizados')) document.getElementById('dashPlanesFinalizados').textContent = planesFinalizados;
    if (document.getElementById('dashPlanesVencidos')) document.getElementById('dashPlanesVencidos').textContent = planesVencidos;
    if (document.getElementById('countEvalRealizadas')) document.getElementById('countEvalRealizadas').textContent = examenesRealizados;
    if (document.getElementById('countEvalPendientes')) document.getElementById('countEvalPendientes').textContent = examenesPendientes;

    const denominadorReal = totalApadrinados || 1;
    const pctOnboard = tOnboardingTotal > 0 ? Math.round((tOnboardingComp / tOnboardingTotal) * 100) : 0;
    const pct7 = Math.round(sumaNota7 / denominadorReal);
    const pct30 = Math.round(sumaNota30 / denominadorReal);
    const pct90 = Math.round(sumaNota90 / denominadorReal);

    if (document.getElementById('barPctOnboarding')) document.getElementById('barPctOnboarding').textContent = pctOnboard + "%";
    if (document.getElementById('fillOnboarding')) document.getElementById('fillOnboarding').style.width = pctOnboard + "%";
    if (document.getElementById('barPct7Dias')) document.getElementById('barPct7Dias').textContent = pct7 + "%";
    if (document.getElementById('fill7Dias')) document.getElementById('fill7Dias').style.width = pct7 + "%";
    if (document.getElementById('barPct30Dias')) document.getElementById('barPct30Dias').textContent = pct30 + "%";
    if (document.getElementById('fill30Dias')) document.getElementById('fill30Dias').style.width = pct30 + "%";
    if (document.getElementById('barPct90Dias')) document.getElementById('barPct90Dias').textContent = pct90 + "%";
    if (document.getElementById('fill90Dias')) document.getElementById('fill90Dias').style.width = pct90 + "%";
});

function renderizarMonitorColaboradores() {
    if (!cacheSnapshotEmpleadosLocal) return;
    const container = document.getElementById('lista-evaluaciones-pendientes');
    const onboardingContainer = document.getElementById('contenedor-monitor-onboarding');
    if (!container) return;
    container.innerHTML = '';

    const ITEMS_POR_PAGINA = 20;
    let paginaActual = window.paginaActualMonitor || 1;

    const regFiltro = document.getElementById('filter-region').value;
    const mesDesde = document.getElementById('filter-mes-desde')?.value || '';
    const mesHasta = document.getElementById('filter-mes-hasta')?.value || '';
    const textoBusqueda = document.getElementById('search-colaborador')?.value.toLowerCase().trim() || '';

    let todosLosEmpleados = [];
    let empleadosHistorico = [];
    let totalEmpleadosFiltro = 0;

    cacheSnapshotEmpleadosLocal.forEach((doc) => {
        const emp = doc.data();
        const docId = doc.id;

        if (miRol === 'padrino' && String(emp.padrino_id) !== String(cedulaLogueada)) return;
        if (emp.es_apadrinado !== true) return;
        if (regFiltro && emp.region !== regFiltro) return;
        if (mesDesde || mesHasta) {
            const mesEmpleado = emp.fecha_ingreso ? emp.fecha_ingreso.substring(0, 7) : null;
            if (!mesEmpleado) return;
            if (mesDesde && mesEmpleado < mesDesde) return;
            if (mesHasta && mesEmpleado > mesHasta) return;
        }

        if (emp.es_apadrinado === true) {
            totalEmpleadosFiltro++;
        }

        if (emp.certificado_descargado === true) {
            empleadosHistorico.push({ emp, docId });
        } else {
            todosLosEmpleados.push({ emp, docId });
        }
    });

    if (document.getElementById('dashTotalEmpleados')) {
        document.getElementById('dashTotalEmpleados').textContent = totalEmpleadosFiltro;
    }

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

    const hoy = new Date();

    const calcularFechaLimiteReal = (fechaBaseStr, diasAsumidos) => {
        if (!fechaBaseStr || fechaBaseStr === "—") return null;
        const d = new Date(fechaBaseStr + 'T00:00:00');
        d.setDate(d.getDate() + diasAsumidos);
        return d.toISOString().split('T')[0];
    };

    // 🌟 CORRECCIÓN CRÍTICA DE AUDITORÍA: Forzar cumplimiento estricto del DÍA EXACTO LÍMITE
    const buildChipInteligente = (label, nota, fechaEjecucion, fechaBaseContratacion, diasDesfase, umbral) => {
        const limiteStr = (diasDesfase === 0 || diasDesfase === 1 || diasDesfase === 7)
            ? calcularFechaLimiteHabilesColombia(fechaBaseContratacion, diasDesfase)
            : calcularFechaLimiteReal(fechaBaseContratacion, diasDesfase);

        if (nota === undefined || nota === null || nota === 0) {
            const esVencido = limiteStr ? (hoy > new Date(limiteStr + 'T00:00:00')) : false;
            return `
                <div class="eval-chip ${esVencido ? 'chip-late-pending' : 'chip-pending'}">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">—</span>
                    <span class="chip-date">${esVencido ? '⚠️ VENCIDO' : '⏳ Día Exacto'}</span>
                    <span style="font-size:8.5px; opacity:0.8; font-weight:bold;">${limiteStr || '—'}</span>
                </div>`;
        }

        let minAprobacion = umbral;
        if (label.includes("Safety")) minAprobacion = 100;
        else if (label.includes("People")) minAprobacion = 80;

        const aprobóExamen = nota >= minAprobacion;

        if (!aprobóExamen) {
            return `
                <div class="eval-chip chip-fail">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">${nota}%</span>
                    <span class="chip-date">❌ Reprobado</span>
                    <span style="font-size:8.5px; opacity:0.8;">Requerido: ${limiteStr || '—'}</span>
                </div>`;
        }

        // 🌟 VALIDACIÓN DE DÍA EXACTO: La fecha de ejecución debe ser igual a la calculada
        const ejecutoEnElDiaExacto = (fechaEjecucion === limiteStr);

        if (ejecutoEnElDiaExacto) {
            return `
                <div class="eval-chip chip-pass">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">${nota}%</span>
                    <span class="chip-date">✓ Hecho a Tiempo</span>
                    <span style="font-size:8.5px; opacity:0.8;">Día exacto: ${fechaEjecucion || '—'}</span>
                </div>`;
        } else {
            // Se evalúa si lo hizo antes del tiempo reglamentario de maduración corporativa o si se pasó
            const esAdelantado = limiteStr ? (new Date(fechaEjecucion + 'T00:00:00') < new Date(limiteStr + 'T00:00:00')) : false;
            return `
                <div class="eval-chip chip-pass-late">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">${nota}%</span>
                    <span class="chip-date">${esAdelantado ? '⚠️ Hecho Antes' : '🎓 Fuera Plazo'}</span>
                    <span style="font-size:8.5px; opacity:0.8;">Hecho: ${fechaEjecucion || '—'} (Debió ser: ${limiteStr})</span>
                </div>`;
        }
    };

    if (onboardingContainer) {
        onboardingContainer.innerHTML = '';
        if (todosLosEmpleados.length === 0) {
            onboardingContainer.innerHTML = `<div style="text-align:center; padding:18px; color:#7a8f99; font-size:12px;">No hay colaboradores.</div>`;
        } else {
            const buildEstadoOnboarding = (nota, umbral, label) => {
                const notaNum = Number(nota) || 0;
                if (notaNum === 0) return `<div><div style="font-size:9.5px; font-weight:700; color:#7a8f99; text-transform:uppercase;">${label}</div><div style="font-size:12px; font-weight:700; color:#8a8f99;">⏳ Pendiente</div></div>`;
                if (notaNum >= umbral) return `<div><div style="font-size:9.5px; font-weight:700; color:#7a8f99; text-transform:uppercase;">${label}</div><div style="font-size:12px; font-weight:700; color:#0f6e56;">✅ Aprobado</div></div>`;
                return `<div><div style="font-size:9.5px; font-weight:700; color:#7a8f99; text-transform:uppercase;">${label}</div><div style="font-size:12px; font-weight:700; color:#a32d2d;">❌ Reprobado (${notaNum}%)</div></div>`;
            };

            onboardingContainer.innerHTML = todosLosEmpleados.map(({ emp }) => {
                const h = emp.hitos || {};
                const fechaBase = emp.fecha_ingreso || h.fecha_contratacion || '—';

                return `
                <div class="onboard-row">
                    <div style="min-width:220px; flex:1;">
                        <div style="font-size:13px; font-weight:700; color:#1c2430; text-transform:uppercase;">${emp.nombre}</div>
                        <div style="font-size:10.5px; color:#7a8f99;">Cargo: ${emp.cargo || 'Operativo'}<br>Ingreso: ${fechaBase}</div>
                    </div>
                    <div style="display:flex; gap:30px; align-items:center; flex-wrap:wrap;">
                        ${buildEstadoOnboarding(h.onboarding_safety_nota, 100, 'Día 1 — Safety')}
                        ${buildEstadoOnboarding(h.onboarding_people_nota, 80, 'Día 2 — People')}
                    </div>
                </div>`;
            }).join('');
        }
    }

    paginaSegmentada.forEach(({ emp, docId }) => {
        const h = emp.hitos || {};
        const fechaBase = emp.fecha_ingreso || h.fecha_contratacion || "";  // 👈 esta la borraste, hay que recuperarla
        const fechaBaseParaPlazos = calcularFechaBaseConNovedades(fechaBase, emp.dias_incapacidad_acumulados);  // 👈 esta es la nueva

        let notes90 = [];
        if (h.eval_autonomo_pre_nota) notes90.push(h.eval_autonomo_pre_nota);
        if (h.eval_autonomo_nota) notes90.push(h.eval_autonomo_nota);
        if (h.eval_autonomo_post_nota) notes90.push(h.eval_autonomo_post_nota);


        let notaUnificada90 = notes90.length > 0 ? Math.round(notes90.reduce((a, b) => a + b, 0) / notes90.length) : 0;
        let fecha90Final = h.eval_autonomo_nota_fecha || h.eval_autonomo_fecha || h.eval_autonomo_pre_fecha || h.eval_autonomo_post_fecha || "";

        let etapasRealizadasConteo = 0; let sumaNotasParaPromedio = 0;
        if (h.onboarding_safety_nota > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += h.onboarding_safety_nota; }
        if (h.onboarding_people_nota > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += h.onboarding_people_nota; }
        if (h.eval_tecnico_nota > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += h.eval_tecnico_nota; }
        if (h.eval_funcional_nota > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += h.eval_funcional_nota; }
        if (notaUnificada90 > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += notaUnificada90; }

        const porcentajeEjecucionReal = etapasRealizadasConteo * 20;
        const promedioCalificacionReal = etapasRealizadasConteo > 0 ? Math.round(sumaNotasParaPromedio / etapasRealizadasConteo) : 0;

        const safetyOk = h.onboarding_safety_nota >= 100;
        const peopleOk = h.onboarding_people_nota >= 80;
        const tecnicoOk = h.eval_tecnico_nota >= 60;
        const funcionalOk = h.eval_funcional_nota >= 60;
        const autonomoOk = notaUnificada90 >= 60;
        const todoCompletado = safetyOk && peopleOk && tecnicoOk && funcionalOk && autonomoOk;
        let botonCertificadoHtml = '';
        if (todoCompletado || emp.estado_plan_padrino === "Plan Padrino Finalizado") {
            botonCertificadoHtml = `<button class="status-btn active-comp" style="font-size:12px; padding:8px 14px; background:#e1f5ee; color:#0f6e56; border-color:#3cbcae;" onclick="imprimirCertificadoCompletoPDF('${docId}')">🎓 Certificado</button>`;
        } else {
            botonCertificadoHtml = `<button class="status-btn" style="font-size:12px; padding:8px 14px; background:#f1f3f5; color:#7a8f99; cursor:not-allowed;" disabled>⏳ En Progreso (${porcentajeEjecucionReal}%)</button>`;
        }

        // 🚨 CONFIGURACIÓN EXCLUSIVA: El Admin genera su botón de novedad en paralelo
        let botonNovedadAdminHtml = '';
        if (miRol === 'admin') {
            botonNovedadAdminHtml = `
              <button class="status-btn" style="font-size:12px; padding:8px 14px; background:#fff8e1; color:#8a6e00; border-color:#ffc404; margin-left: 6px; cursor:pointer; font-weight:700;" 
                      onclick="abrirModalNovedad('${emp.cedula}', '${emp.nombre}')">
                🚨 Novedad
              </button>`;
        }

        container.innerHTML += `
      <div class="collab-card" style="border-left-color: ${todoCompletado ? '#3cbcae' : '#206987'}">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap: wrap; gap: 10px;">
          <div>
            <div class="collab-name">${emp.nombre}</div>
            <div class="collab-meta">
              CC <strong>${emp.cedula}</strong> &nbsp;·&nbsp; 
              Cargo: <strong style="color:#206987;">${emp.cargo || 'Operativo'}</strong> &nbsp;·&nbsp; 
              Regional: ${emp.region || 'General'} &nbsp;·&nbsp; 
              Grupo: <strong>${emp.grupo || 'General'}</strong> &nbsp;·&nbsp;
              Fecha Contratación: <strong style="color:#0f6e56;">${fechaBase || '—'}</strong>
            </div>
            
            ${emp.estado_plan_padrino === "Empleado con Novedad" ? `
                <div style="margin-top:6px; font-size:10.5px; background:#fff5f5; border:1px solid #f09595; color:#a32d2d; padding:3px 8px; border-radius:4px; display:inline-block; font-weight:700;">
                  🤒 Empleado con Novedad Temporal (${emp.dias_incapacidad_acumulados || 0} días aplazados)
                </div>
            ` : ''}

            <div style="margin-top:8px; font-size:11px; background:#fff8e1; border:1px solid #ffc404; color:#8a6e00; padding:4px 10px; border-radius:6px; display:inline-block; font-weight:700;">
              🏅 Promedio Avance: ${promedioCalificacionReal}%
            </div>
          </div>
          
          <div style="display:flex; align-items:center; gap:4px; flex-wrap:wrap;">
             ${botonCertificadoHtml}     ${botonNovedadAdminHtml}     </div>
        </div>
        <div class="evals-row">
          ${buildChipInteligente("Safety (D1)", h.onboarding_safety_nota, h.onboarding_safety_fecha, fechaBaseParaPlazos, 0, 100)}
${buildChipInteligente("People (D2)", h.onboarding_people_nota, h.onboarding_people_fecha, fechaBaseParaPlazos, 1, 80)}
${buildChipInteligente("7 Días (Téc)", h.eval_tecnico_nota, h.eval_tecnico_fecha, fechaBaseParaPlazos, 7, 60)}
${buildChipInteligente("30 Días (Fun)", h.eval_funcional_nota, h.eval_funcional_fecha, fechaBaseParaPlazos, 32, 60)}
${buildChipInteligente("90 Días (Aut)", notaUnificada90, fecha90Final, fechaBaseParaPlazos, 92, 60)}
        </div>
      </div>`;
    });

    if (todosLosEmpleados.length > 0) {
        container.innerHTML += `
      <div style="display:flex; align-items:center; justify-content:center; gap:12px; padding:16px; margin-top:8px; width:100%;">
        <button onclick="cambiarPaginaMonitor(${paginaActual - 1})" ${paginaActual <= 1 ? 'disabled' : ''} style="padding:6px 14px; border-radius:6px; border:1px solid #dde3e7; cursor:pointer; background:white; font-weight:600; font-size:12px;">← Anterior</button>
        <span style="font-size:12px; color:#7a8f99; font-weight:600;">Página ${paginaActual} de ${totalPaginas} &nbsp;·&nbsp; ${todosLosEmpleados.length} colaboradores activos</span>
        <button onclick="cambiarPaginaMonitor(${paginaActual + 1})" ${paginaActual >= totalPaginas ? 'disabled' : ''} style="padding:6px 14px; border-radius:6px; border:1px solid #dde3e7; cursor:pointer; background:white; font-weight:600; font-size:12px;">Siguiente →</button>
      </div>`;
    }

    const historicoContainer = document.getElementById('lista-historico-certificados');
    if (historicoContainer) {
        historicoContainer.innerHTML = empleadosHistorico.length === 0
            ? `<div style="text-align:center; padding:15px; color:#7a8f99; font-size:12px; background:#f8fafb; border-radius:8px; border:1px dashed #dde3e7;">No se registran descargas históricas.</div>`
            : empleadosHistorico.map(({ emp }) => `<div style="display:flex; justify-content:space-between; align-items:center; background:#f4fbf9; border:1px solid #3cbcae; padding:10px; border-radius:8px; margin-bottom:6px; font-size:12px;"><div><strong>✓ ${emp.nombre}</strong> <span style="font-size:11px; color:#7a8f99;">(CC ${emp.cedula})</span></div><div style="font-size:11px; font-weight:700; color:#0f6e56; background:#e1f5ee; padding:2px 8px; border-radius:4px;">Certificado Concedido</div></div>`).join('');
    }
}

async function toggleOnboardingEstado(docId, campo, estaCompletadoActualmente) {
    try {
        const hoyStr = new Date().toISOString().split('T')[0];
        const nuevoEstado = estaCompletadoActualmente ? { estado: "Pendiente", fecha_completado: "" } : { estado: "Completado", fecha_completado: hoyStr };
        await db.collection('empleados').doc(docId).update({ [campo]: nuevoEstado });
    } catch (error) { console.error(error); }
}

function toggleDetalleContenedor(id, btn) {
    const panel = document.getElementById(id); if (!panel) return;
    panel.classList.toggle('open');
    btn.textContent = panel.classList.contains('open') ? '▴ Ocultar Detalle' : '▾ Ver Detalle de Criterios';
}

async function procesarCargaMasivaExcel(event) {
    const file = event.target.files[0]; if (!file) return;

    const banner = document.getElementById('banner-carga-excel');
    const bannerIcono = document.getElementById('banner-carga-icono');
    const bannerTexto = document.getElementById('banner-carga-texto');

    const mostrarBanner = (mensaje, tipo) => {
        if (!banner) return;
        banner.style.display = 'flex';
        if (tipo === 'cargando') {
            banner.style.background = '#fff8e1';
            banner.style.border = '1px solid #ffc404';
            banner.style.color = '#8a6e00';
            bannerIcono.className = 'ti ti-loader';
            bannerIcono.style.animation = 'spin 1s linear infinite';
        } else if (tipo === 'exito') {
            banner.style.background = '#e1f5ee';
            banner.style.border = '1px solid #3cbcae';
            banner.style.color = '#0f6e56';
            bannerIcono.className = 'ti ti-circle-check';
            bannerIcono.style.animation = 'none';
        } else {
            banner.style.background = '#fff5f5';
            banner.style.border = '1px solid #f09595';
            banner.style.color = '#a32d2d';
            bannerIcono.className = 'ti ti-alert-triangle';
            bannerIcono.style.animation = 'none';
        }
        bannerTexto.textContent = mensaje;
    };

    mostrarBanner(`⏳ Procesando "${file.name}"... no cierres esta ventana.`, 'cargando');

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

            const rawRows = rows.slice(1); let totalProcesados = 0;
            for (const row of rawRows) {
                let parsedRow = row; if (row.length === 1 && typeof row[0] === 'string') parsedRow = row[0].split(';');
                const ccColaborador = String(parsedRow[15] || '').trim().replace(/['"\s.\-,]/g, '');
                const fechaEjecucion = parsedRow[1] instanceof Date ? parsedRow[1].toISOString().split('T')[0] : new Date().toISOString().split('T')[0];
                if (!ccColaborador || ccColaborador === 'null' || ccColaborador === '0') continue;

                const scores = scoreExcelDinamico(parsedRow);
                const avgGeneral = scores.tech && scores.people ? Math.round((scores.tech.pct + scores.people.pct) / 2) : (scores.tech || scores.people)?.pct || 0;

                let snap = await db.collection('empleados').where('cedula', '==', ccColaborador).limit(1).get();
                if (snap.empty) snap = await db.collection('empleados').where('cedula', '==', Number(ccColaborador)).limit(1).get();

                if (!snap.empty) {
                    const docId = snap.docs[0].id; const emp = snap.docs[0].data(); const hitos = emp.hitos || {}; let up = {};
                    const cicloKey = resolverCicloExcel(parsedRow[17]);

                    if (cicloKey === '7') { hitos.eval_tecnico_nota = avgGeneral; hitos.eval_tecnico_fecha = fechaEjecucion; up.etapa_actual = "Etapa Funcional"; }
                    else if (cicloKey === '30') { hitos.eval_funcional_nota = avgGeneral; hitos.eval_funcional_fecha = fechaEjecucion; up.etapa_actual = "Etapa Autónomo"; }
                    else if (cicloKey === '90_pre') { hitos.eval_autonomo_pre_nota = avgGeneral; hitos.eval_autonomo_pre_fecha = fechaEjecucion; }
                    else if (cicloKey === '90') { hitos.eval_autonomo_nota = avgGeneral; hitos.eval_autonomo_fecha = fechaEjecucion; if (hitos.eval_autonomo_pre_nota > 0 && hitos.eval_autonomo_post_nota > 0) up.estado_plan_padrino = "Plan Padrino Finalizado"; }
                    else if (cicloKey === '90_post') { hitos.eval_autonomo_post_nota = avgGeneral; hitos.eval_autonomo_post_fecha = fechaEjecucion; if (hitos.eval_autonomo_pre_nota > 0 && hitos.eval_autonomo_nota > 0) up.estado_plan_padrino = "Plan Padrino Finalizado"; }

                    up.hitos = hitos; await db.collection('empleados').doc(docId).update(up); totalProcesados++;
                }
            }

            mostrarBanner(`🤖 Consolidación completa: ${totalProcesados} cuestionarios emparejados.`, 'exito');
            setTimeout(() => { location.reload(); }, 1800);
        } catch (err) {
            console.error(err);
            mostrarBanner('❌ Ocurrió un error al procesar el archivo. Revisa la consola.', 'error');
        }
    };
    reader.onerror = () => mostrarBanner('❌ No se pudo leer el archivo seleccionado.', 'error');
    reader.readAsArrayBuffer(file);
}

function scoreExcelDinamico(row) {
    let lastIdx = row.length - 1;
    while (lastIdx >= 0 && (row[lastIdx] === null || row[lastIdx] === undefined || row[lastIdx] === '')) lastIdx--;
    return { tech: scoreExcelRange(row, 18, lastIdx - 6), people: scoreExcelRange(row, lastIdx - 5, lastIdx - 1), nps: row[lastIdx] };
}

function cambiarPaginaMonitor(nuevaPagina) { window.paginaActualMonitor = nuevaPagina; renderizarMonitorColaboradores(); }
function scoreExcelRange(row, start, end) {
    let total = 0, si = 0;
    for (let i = start; i <= end; i++) { const v = row[i]; if (v !== null && v !== undefined && v !== '') { total++; if (String(v).toLowerCase().startsWith('si') || String(v).toLowerCase().startsWith('sí')) si++; } }
    return total > 0 ? { pct: Math.round((si / total) * 100), si, total } : null;
}

async function cargarPerfilDetalladoPadrinoCompleto(cedulaPadrino) {
    try {
        const fichaRoot = document.getElementById('ficha-detalle-padrino-root'); if (fichaRoot) fichaRoot.style.display = 'flex';
        const snapPad = await db.collection('empleados').where('cedula', '==', String(cedulaPadrino)).limit(1).get(); if (snapPad.empty) return;
        const pData = snapPad.docs[0].data();
        document.getElementById('viewPadName').textContent = pData.nombre;
        document.getElementById('viewPadCargo').textContent = pData.cargo || "Supervisor / Líder";
        document.getElementById('viewPadEmpresa').textContent = pData.empresa_padrino || "Bavaria AB InBev";
        document.getElementById('viewPadTiempo').textContent = pData.tiempo_compania || "N/A";
        if (document.getElementById('viewPadCorreo')) document.getElementById('viewPadCorreo').textContent = pData.correo || "Sin correo";

        const boxBadge = document.getElementById('viewPadBadgeEstado');
        if (boxBadge) boxBadge.innerHTML = pData.padrino_estado === "Inactivo" ? `<span>❌ Inactivo</span>` : `<span>🏅 Activo</span>`;

        const boxFoto = document.getElementById('contenedor-foto-padrino');
        if (boxFoto) {
            if (pData.foto_url) boxFoto.innerHTML = `<img src="${pData.foto_url}" style="width:85px; height:85px; object-fit:cover; border-radius:50%;">`;
            else boxFoto.innerHTML = `<div>${pData.nombre.substring(0, 2).toUpperCase()}</div>`;
        }

        const containerTecnicas = document.getElementById('viewPadTagsTecnicas'); if (containerTecnicas) containerTecnicas.innerHTML = '';
        if (pData.habilidades_tecnicas) pData.habilidades_tecnicas.forEach(h => { containerTecnicas.innerHTML += `<span>${h}</span>`; });
        const containerBlandas = document.getElementById('viewPadTagsBlandas'); if (containerBlandas) containerBlandas.innerHTML = '';
        if (pData.habilidades_blandas) pData.habilidades_blandas.forEach(h => { containerBlandas.innerHTML += `<span>${h}</span>`; });

        const muchachosSnapshot = await db.collection('empleados').where('padrino_id', '==', String(cedulaPadrino)).get();
        const containerMuchachos = document.getElementById('viewPadListaMuchachos'); const containerHistoricoNotas = document.getElementById('viewPadHistoricoNotasBody');
        if (containerMuchachos) containerMuchachos.innerHTML = ''; if (containerHistoricoNotas) containerHistoricoNotas.innerHTML = '';
        let cTecnico = 0; let cFuncional = 0; let cAutonomo = 0; let sumaNotasTutor = 0; let totalNotasContadas = 0;

        if (!muchachosSnapshot.empty) {
            muchachosSnapshot.forEach(docM => {
                const m = docM.data(); if (m.etapa_actual === "Etapa Funcional") cFuncional++; else if (m.etapa_actual === "Etapa Autónomo") cAutonomo++; else cTecnico++;
                containerMuchachos.innerHTML += `<div><div>${m.nombre}</div></div>`;
                const n7 = m.hitos?.eval_tecnico_nota || "---"; const n30 = m.hitos?.eval_funcional_nota || "---"; const n90 = m.hitos?.eval_autonomo_nota || "---";
                if (typeof n7 === "number") { sumaNotasTutor += n7; totalNotasContadas++; }
                containerHistoricoNotas.innerHTML += `<tr><td>${m.nombre}</td><td>${n7}</td><td>${n30}</td><td>${n90}</td><td>${m.estado_plan_padrino || 'Activo'}</td></tr>`;
            });
        }
        if (document.getElementById('cvCountTecnico')) document.getElementById('cvCountTecnico').textContent = cTecnico;
        if (document.getElementById('cvCountFuncional')) document.getElementById('cvCountFuncional').textContent = cFuncional;
        if (document.getElementById('cvCountAutonomo')) document.getElementById('cvCountAutonomo').textContent = cAutonomo;
    } catch (err) { console.error(err); }
}

async function imprimirCertificadoCompletoPDF(docId) {
    try {
        const docSnap = await db.collection('empleados').doc(docId).get(); if (!docSnap.exists) return;
        const emp = docSnap.data(); const h = emp.hitos || {};
        let nombrePadrino = "—";
        if (emp.padrino_id) { const snapPad = await db.collection('empleados').where('cedula', '==', String(emp.padrino_id)).limit(1).get(); if (!snapPad.empty) nombrePadrino = snapPad.docs[0].data().nombre; }
        const fechaHoyFormateada = new Date().toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });

        const certificadoHTML = `<div id="hoja-certificado-temporal" style="width:1000px; padding:60px; background:#fff; border:14px solid #206987;"><h2>CERTIFICADO DE FINALIZACIÓN</h2><h3>${emp.nombre}</h3><p>Padrino: ${nombrePadrino}</p><p>Fecha: ${fechaHoyFormateada}</p></div>`;
        const hojaTemporal = document.createElement('div'); hojaTemporal.innerHTML = certificadoHTML; document.body.appendChild(hojaTemporal);
        await html2pdf().set({ margin: 0, filename: `Certificado_${emp.nombre}.pdf`, html2canvas: { scale: 2 }, jsPDF: { format: 'a4', orientation: 'landscape' } }).from(hojaTemporal).save();
        document.body.removeChild(hojaTemporal);

        await db.collection('empleados').doc(docId).update({ certificado_descargado: true, fecha_certificado_descargado: new Date().toISOString().split('T')[0] });
        alert("🎓 ¡Certificado generado!");
    } catch (error) { console.error(error); }
}

function cambiarPestañaPadrino(targetPaneId, botonPresionado) {
    document.querySelectorAll('.padrino-section-pane').forEach(pane => {
        pane.classList.remove('pane-active');
    });
    document.querySelectorAll('.nav-tab-btn').forEach(btn => {
        btn.classList.remove('active-tab');
    });

    const seccionObjetivo = document.getElementById('padrino-pane-' + targetPaneId);
    if (seccionObjetivo) {
        seccionObjetivo.classList.add('pane-active');
    }
    if (botonPresionado) {
        botonPresionado.classList.add('active-tab');
    }

    if (targetPaneId === 'monitor' && typeof renderizarMonitorColaboradores === 'function') {
        renderizarMonitorColaboradores();
    }
}

function poblarFiltrosEstrategicos(regiones, meses) {
    const selectReg = document.getElementById('filter-region');
    const selectDesde = document.getElementById('filter-mes-desde');
    const selectHasta = document.getElementById('filter-mes-hasta');

    if (selectReg && selectReg.options.length <= 1) {
        regiones.forEach(r => {
            if (r) selectReg.innerHTML += `<option value="${r}">${r}</option>`;
        });
    }

    const formatearLabel = (m) => {
        const [anio, mes] = m.split('-');
        const l = new Date(+anio, +mes - 1, 1).toLocaleDateString('es-CO', { month: 'long', year: 'numeric' });
        return l.charAt(0).toUpperCase() + l.slice(1);
    };

    [selectDesde, selectHasta].forEach(select => {
        if (select && select.options.length <= 1) {
            meses.forEach(m => {
                if (m) select.innerHTML += `<option value="${m}">${formatearLabel(m)}</option>`;
            });
        }
    });
}

// 📸 Procesa la foto del padrino en Base64 antes de guardar
function procesarFotoPadrinoLocal(inputFile) {
    const file = inputFile.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        globalPadrinoFotoBase64 = e.target.result;
        const img = document.getElementById('padrinoImgPreview');
        const icon = document.getElementById('padrinoIconoDefault');
        if (img) { img.src = globalPadrinoFotoBase64; img.style.display = 'block'; }
        if (icon) icon.style.display = 'none';
    };
    reader.readAsDataURL(file);
}

// 🏅 Certifica a un colaborador antiguo como Padrino
async function convertirYConfigurarPadrino() {
    try {
        const cedula = document.getElementById('padrePadrinoCedula').value.trim();
        if (!cedula) { alert('⚠️ Debes ingresar la cédula del empleado.'); return; }

        const empresa = document.getElementById('padrePadrinoEmpresa').value.trim();
        const tiempo = document.getElementById('padrePadrinoTiempo').value.trim();
        const desempeno = document.getElementById('padrePadrinoDesempeno').value.trim();
        const estado = document.getElementById('padrePadrinoEstado').value;
        const correo = document.getElementById('padrePadrinoCorreo').value.trim();
        const tecnicas = document.getElementById('padrePadrinoTecnicas').value.split(',').map(s => s.trim()).filter(Boolean);
        const blandas = document.getElementById('padrePadrinoBlandas').value.split(',').map(s => s.trim()).filter(Boolean);

        const datosPadrino = {
            es_padrino: true,
            empresa_padrino: empresa,
            tiempo_compania: tiempo,
            padrino_desempeno: desempeno,
            padrino_estado: estado,
            habilidades_tecnicas: tecnicas,
            habilidades_blandas: blandas
        };
        if (correo) datosPadrino.correo = correo;
        if (globalPadrinoFotoBase64) datosPadrino.foto_url = globalPadrinoFotoBase64;

        let snap = await db.collection('empleados').where('cedula', '==', cedula).limit(1).get();
        if (snap.empty && !isNaN(cedula)) {
            snap = await db.collection('empleados').where('cedula', '==', Number(cedula)).limit(1).get();
        }

        if (snap.empty) {
            alert('⚠️ No se encontró ningún colaborador con esa cédula. Verifica el número.');
            return;
        }

        await db.collection('empleados').doc(snap.docs[0].id).update(datosPadrino);
        alert('🏅 Padrino certificado correctamente.');

        globalPadrinoFotoBase64 = '';
        ['padrePadrinoCedula', 'padrePadrinoEmpresa', 'padrePadrinoTiempo', 'padrePadrinoDesempeno', 'padrePadrinoCorreo', 'padrePadrinoTecnicas', 'padrePadrinoBlandas']
            .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
        const img = document.getElementById('padrinoImgPreview');
        const icon = document.getElementById('padrinoIconoDefault');
        if (img) { img.src = ''; img.style.display = 'none'; }
        if (icon) icon.style.display = 'block';

    } catch (error) {
        console.error(error);
        alert('❌ Ocurrió un error al certificar el padrino.');
    }
}

// 🧒 Registra un nuevo apadrinado y vincula su Plan Padrino con fechas límite calculadas
async function vincularPlanPadrinoNuevo() {
    try {
        const cedula = document.getElementById('padreApadrinadoId').value.trim();
        const nombre = document.getElementById('padreApadrinadoNombre').value.trim();
        const cargo = document.getElementById('padreApadrinadoCargo').value.trim();
        const area = document.getElementById('padreApadrinadoArea').value.trim();
        const region = document.getElementById('padreApadrinadoRegion').value.trim();
        const fechaIngreso = document.getElementById('padreApadrinadoFechaIngreso').value;
        const jefe = document.getElementById('padreApadrinadoJefe').value.trim();
        const correo = document.getElementById('padreApadrinadoCorreo').value.trim();
        const etapaInicial = document.getElementById('padreEtapa').value;
        const padrinoId = document.getElementById('selectPadrinoDinamico').value;

        if (!cedula || !nombre || !fechaIngreso) {
            alert('⚠️ Cédula, nombre y fecha de ingreso son obligatorios.');
            return;
        }
        if (!padrinoId) {
            alert('⚠️ Debes seleccionar un padrino certificado disponible.');
            return;
        }

        const calcularFechaLimiteCalendario = (fechaStr, dias) => {
            const d = new Date(fechaStr + 'T00:00:00');
            d.setDate(d.getDate() + dias);
            return d.toISOString().split('T')[0];
        };

        const fechasLimite = {
            eval_7_dias: calcularFechaLimiteHabilesColombia(fechaIngreso, 7),
            eval_30_dias: calcularFechaLimiteCalendario(fechaIngreso, 32),
            eval_90_dias: calcularFechaLimiteCalendario(fechaIngreso, 92)
        };

        // 🔍 1. Buscamos primero si el empleado ya existe para extraer y resguardar su progreso actual
        let snap = await db.collection('empleados').where('cedula', '==', cedula).limit(1).get();
        if (snap.empty && !isNaN(cedula)) {
            snap = await db.collection('empleados').where('cedula', '==', Number(cedula)).limit(1).get();
        }

        // Estructura por defecto por si llega a ser un ingreso 100% nuevo en la plataforma
        let hitosFinales = {
            fecha_contratacion: fechaIngreso,
            onboarding_safety_nota: 0,
            onboarding_people_nota: 0,
            eval_tecnico_nota: 0,
            eval_funcional_nota: 0,
            eval_autonomo_nota: 0
        };
        let onboarding1 = { estado: "Pendiente", fecha_completado: "" };
        let onboarding2 = { estado: "Pendiente", fecha_completado: "" };

        // 🛡️ PRESERVACIÓN INDESTRUCTIBLE DE HISTORIAL EXISTENTE
        if (!snap.empty) {
            const datosExistentes = snap.docs[0].data();

            // Si ya tenía notas cargadas de Microsoft Forms, las unificamos en lugar de borrarlas
            if (datosExistentes.hitos) {
                hitosFinales = {
                    ...datosExistentes.hitos,
                    fecha_contratacion: fechaIngreso // Mantiene sus notas pero actualiza la fecha base
                };
            }

            // Mantener los estados de completado del Onboarding Día 1 y Día 2
            if (datosExistentes.onboarding_dia1) onboarding1 = datosExistentes.onboarding_dia1;
            if (datosExistentes.onboarding_dia2) onboarding2 = datosExistentes.onboarding_dia2;
        }

        const datosApadrinado = {
            cedula, nombre, cargo, area, region,
            fecha_ingreso: fechaIngreso,
            jefe_directo: jefe,
            correo,
            es_apadrinado: true,
            padrino_id: padrinoId,
            etapa_actual: etapaInicial,
            estado_plan_padrino: "Plan Padrino Activo",
            fechas_limite_evaluaciones: fechasLimite,

            // Asignación de bloques validados
            hitos: hitosFinales,
            onboarding_dia1: onboarding1,
            onboarding_dia2: onboarding2
        };

        if (!snap.empty) {
            await db.collection('empleados').doc(snap.docs[0].id).update(datosApadrinado);
        } else {
            await db.collection('empleados').add(datosApadrinado);
        }

        alert('✅ Colaborador registrado y Plan Padrino activado.');

        ['padreApadrinadoId', 'padreApadrinadoNombre', 'padreApadrinadoCargo', 'padreApadrinadoArea', 'padreApadrinadoRegion', 'padreApadrinadoFechaIngreso', 'padreApadrinadoJefe', 'padreApadrinadoCorreo']
            .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });

    } catch (error) {
        console.error(error);
        alert('❌ Ocurrió un error al activar el plan padrino.');
    }
}

// 🔄 FUNCIÓN DE REASIGNACIÓN: Cambia de tutor resguardando las notas y el pasado en Firestore
async function reasignarNuevoPadrinoConHistorial(cedulaAlumno, nuevoPadrinoId) {
    if (!cedulaAlumno || !nuevoPadrinoId) {
        alert('⚠️ Cédula del alumno y selección de nuevo padrino son obligatorios.');
        return;
    }

    try {
        // 1. Buscamos el documento del alumno (Texto o Número)
        let snap = await db.collection('empleados').where('cedula', '==', cedulaAlumno).limit(1).get();
        if (snap.empty && !isNaN(cedulaAlumno)) {
            snap = await db.collection('empleados').where('cedula', '==', Number(cedulaAlumno)).limit(1).get();
        }

        if (snap.empty) {
            alert('❌ No se encontró al colaborador en la base de datos.');
            return;
        }

        const docRef = snap.docs[0].ref;
        const datosExistentes = snap.docs[0].data();

        // Extraemos o inicializamos el arreglo histórico para la trazabilidad del alumno
        let historialActual = datosExistentes.historial_padrinos || [];

        // 🛡️ RESPALDO EN CALIENTE: Si ya tenía un padrino asignado y es diferente, lo salvamos
        if (datosExistentes.padrino_id && datosExistentes.padrino_id !== nuevoPadrinoId) {

            let nombrePadrinoViejo = "Tutor Anterior";

            // Buscamos el nombre real del padrino saliente para registrarlo en texto claro
            const snapPadViejo = await db.collection('empleados').where('cedula', '==', datosExistentes.padrino_id).limit(1).get();
            if (!snapPadViejo.empty) {
                nombrePadrinoViejo = snapPadViejo.docs[0].data().nombre;
            }

            // Insertamos el registro de relevo al historial
            historialActual.push({
                padrino_id_anterior: datosExistentes.padrino_id,
                padrino_nombre_anterior: nombrePadrinoViejo,
                fecha_relevo: new Date().toISOString().split('T')[0], // Guarda el día de hoy
                motivo: "Reasignación por cambio operativo / salida de operación"
            });
        }

        // 2. Conservamos sus hitos intactos (notas de 7, 30, 90 días, Safety y People)
        let hitosProtegidos = datosExistentes.hitos || {
            fecha_contratacion: datosExistentes.fecha_ingreso || "",
            onboarding_safety_nota: 0, onboarding_people_nota: 0,
            eval_tecnico_nota: 0, eval_funcional_nota: 0, eval_autonomo_nota: 0
        };

        // 3. Empujamos la actualización a Firestore
        await docRef.update({
            padrino_id: nuevoPadrinoId,
            historial_padrinos: historialActual,
            hitos: hitosProtegidos, // Mantiene el 60% o avance actual
            estado_plan_padrino: "Plan Padrino Activo"
        });

        alert(`✅ Reasignación exitosa.\n\n${datosExistentes.nombre} pasó al control del nuevo tutor. Se guardó el registro histórico de su padrino anterior.`);

        // Si tienes un monitor abierto, lo refrescamos en caliente
        if (typeof renderizarMonitorColaboradores === 'function') renderizarMonitorColaboradores();

    } catch (error) {
        console.error("Error crítico en la reasignación con trazabilidad:", error);
        alert("❌ Error al procesar el cambio de padrino en Firestore.");
    }
}

// Interfaz: Abre el Modal
function abrirModalNovedad(cedula, nombre) {
    document.getElementById('nov-alumno-cedula').value = cedula;
    document.getElementById('nov-alumno-nombre').textContent = nombre;
    document.getElementById('nov-tipo').value = "";
    document.getElementById('nov-dias').value = "";
    document.getElementById('nov-wrapper-incapacidad').style.display = "none";
    document.getElementById('modal-novedad-padrino').style.display = "flex";
}

function cerrarModalNovedad() {
    document.getElementById('modal-novedad-padrino').style.display = "none";
}

function conmutarInputsNovedad() {
    const tipo = document.getElementById('nov-tipo').value;
    document.getElementById('nov-wrapper-incapacidad').style.display = (tipo === "incapacidad") ? "block" : "none";
}

// 🧠 PROCESADOR DE NOVEDADES OPERATIVAS CON RECALCULO MATEMÁTICO DE HITOS
async function guardarNovedadOperativaFirebase() {
    const cedula = document.getElementById('nov-alumno-cedula').value;
    const tipoNovedad = document.getElementById('nov-tipo').value;
    const diasIncapacidad = parseInt(document.getElementById('nov-dias').value) || 0;

    if (!tipoNovedad) return alert("⚠️ Selecciona el tipo de novedad.");
    if (tipoNovedad === "incapacidad" && diasIncapacidad <= 0) return alert("⚠️ Ingresa un número de días válido.");

    try {
        let snap = await db.collection('empleados').where('cedula', '==', cedula).limit(1).get();
        if (snap.empty && !isNaN(cedula)) {
            snap = await db.collection('empleados').where('cedula', '==', Number(cedula)).limit(1).get();
        }

        if (snap.empty) return alert("❌ No se encontró al colaborador.");

        const docRef = snap.docs[0].ref;
        const empData = snap.docs[0].data();

        // 🛡️ GUARD: evita aplicar novedades sobre alguien ya retirado
        if (empData.estado_plan_padrino === "Retirado" || empData.es_apadrinado === false) {
            return alert("⚠️ Este colaborador ya fue retirado. No se pueden aplicar más novedades sobre su registro.");
        }

        let updateData = {};

        if (tipoNovedad === "incapacidad") {
            // 🤒 ACUMULADOR DE DÍAS DE PRÓRROGA (No toca la fecha de ingreso real)
            const diasActualesAcumulados = parseInt(empData.dias_incapacidad_acumulados) || 0;
            const nuevoTotalDiasExtension = diasActualesAcumulados + diasIncapacidad;

            // Tomamos la fecha de ingreso original para recalcular los plazos reales movidos por la novedad
            const fIngresoOriginal = empData.fecha_ingreso || empData.hitos?.fecha_contratacion;
            if (!fIngresoOriginal) return alert("❌ El empleado no tiene una fecha base registrada.");

            // Calculamos una fecha virtual de referencia (Fecha ingreso original + días acumulados) para mover los plazos
            let fechaVirtualReferencia = new Date(fIngresoOriginal + 'T00:00:00');
            fechaVirtualReferencia.setDate(fechaVirtualReferencia.getDate() + nuevoTotalDiasExtension);
            const fechaVirtualStr = fechaVirtualReferencia.toISOString().split('T')[0];

            // Re-calculamos las fechas límite oficiales movidas por la incapacidad
            const nuevasFechasLimite = {
                eval_7_dias: calcularFechaLimiteHabilesColombia(fechaVirtualStr, 7),
                eval_30_dias: calcularFechaLimiteCalendarioISO(fechaVirtualStr, 32),
                eval_90_dias: calcularFechaLimiteCalendarioISO(fechaVirtualStr, 92)
            };

            updateData = {
                dias_incapacidad_acumulados: nuevoTotalDiasExtension,
                fechas_limite_evaluaciones: nuevasFechasLimite,
                estado_plan_padrino: "Empleado con Novedad", // 🌟 Bandera de estado para la etiqueta
                novedad_ultima_aplicada: `Incapacidad de ${diasIncapacidad} días registrada el ${new Date().toISOString().split('T')[0]}`
            };

        } else if (tipoNovedad === "retirado" || tipoNovedad === "Retiro Definitivo de la Empresa") {
            // ❌ LÓGICA DE RETIRO: Rompe los interruptores activos y congela para auditorías
            updateData = {
                es_apadrinado: false,            // 🌟 NUEVA LÍNEA: Crucial para removerlo de las listas activas de inmediato
                estado_plan_padrino: "Retirado", // Base para que se liste en retirados.html
                fecha_retiro_operacion: new Date().toISOString().split('T')[0],
                novedad_ultima_aplicada: `Retiro definitivo de la empresa aplicado el ${new Date().toISOString().split('T')[0]}`
            };
        }

        await docRef.update(updateData);
        alert("🤖 ¡Novedad procesada con éxito y base de datos sincronizada!");

        cerrarModalNovedad();

        // 🔄 Forzar actualización del monitor activo para que la tarjeta desaparezca al milisegundo
        if (typeof renderizarMonitorColaboradores === 'function') {
            renderizarMonitorColaboradores();
        }

    } catch (error) {
        console.error(error);
        alert("❌ Error al procesar la novedad en Firestore.");
    }
}

// 📊 EXPORTA A EXCEL TODOS LOS COLABORADORES CON EVALUACIONES VENCIDAS O REPROBADAS
function exportarVencidosYReprobadosExcel() {
    if (!cacheSnapshotEmpleadosLocal) {
        alert('⚠️ Los datos aún no han cargado. Espera un momento e intenta de nuevo.');
        return;
    }

    const hoy = new Date();
    const filasExportar = [];

    const calcularFechaLimiteReal = (fechaBaseStr, diasAsumidos) => {
        if (!fechaBaseStr || fechaBaseStr === "—") return null;
        const d = new Date(fechaBaseStr + 'T00:00:00');
        d.setDate(d.getDate() + diasAsumidos);
        return d.toISOString().split('T')[0];
    };

    // 🔎 Determina el estado de un hito puntual: Vencido, Reprobado, o null si está OK/pendiente-a-tiempo
    const evaluarEstadoHito = (label, nota, fechaBaseContratacion, diasDesfase, umbral) => {
        const limiteStr = (diasDesfase === 0 || diasDesfase === 1 || diasDesfase === 7)
            ? calcularFechaLimiteHabilesColombia(fechaBaseContratacion, diasDesfase)
            : calcularFechaLimiteReal(fechaBaseContratacion, diasDesfase);

        let minAprobacion = umbral;
        if (label.includes("Safety")) minAprobacion = 100;
        else if (label.includes("People")) minAprobacion = 80;

        if (nota === undefined || nota === null || nota === 0) {
            const esVencido = limiteStr ? (hoy > new Date(limiteStr + 'T00:00:00')) : false;
            if (esVencido) {
                return { estado: "VENCIDO (sin realizar)", limite: limiteStr, nota: "—" };
            }
            return null; // aún dentro de plazo, no se exporta
        }

        if (nota < minAprobacion) {
            return { estado: `REPROBADO (${nota}%)`, limite: limiteStr, nota: nota + "%" };
        }

        return null; // aprobado, no se exporta
    };

    cacheSnapshotEmpleadosLocal.forEach((doc) => {
        const emp = doc.data();
        if (emp.es_apadrinado !== true) return; // solo activos, igual que el monitor

        const h = emp.hitos || {};
        const fechaBase = emp.fecha_ingreso || h.fecha_contratacion || "";
        const fechaBaseParaPlazos = calcularFechaBaseConNovedades(fechaBase, emp.dias_incapacidad_acumulados);

        let notes90 = [];
        if (h.eval_autonomo_pre_nota) notes90.push(h.eval_autonomo_pre_nota);
        if (h.eval_autonomo_nota) notes90.push(h.eval_autonomo_nota);
        if (h.eval_autonomo_post_nota) notes90.push(h.eval_autonomo_post_nota);
        const notaUnificada90 = notes90.length > 0 ? Math.round(notes90.reduce((a, b) => a + b, 0) / notes90.length) : 0;

        const hitosAEvaluar = [
            { label: "Safety (D1)", nota: h.onboarding_safety_nota, dias: 0, umbral: 100 },
            { label: "People (D2)", nota: h.onboarding_people_nota, dias: 1, umbral: 80 },
            { label: "7 Días (Téc)", nota: h.eval_tecnico_nota, dias: 7, umbral: 60 },
            { label: "30 Días (Fun)", nota: h.eval_funcional_nota, dias: 32, umbral: 60 },
            { label: "90 Días (Aut)", nota: notaUnificada90, dias: 92, umbral: 60 }
        ];

        hitosAEvaluar.forEach(hito => {
            const resultado = evaluarEstadoHito(hito.label, hito.nota, fechaBaseParaPlazos, hito.dias, hito.umbral);
            if (resultado) {
                filasExportar.push({
                    "Nombre": emp.nombre || "",
                    "Cédula": emp.cedula || "",
                    "Cargo": emp.cargo || "Operativo",
                    "Regional": emp.region || "General",
                    "Grupo": emp.grupo || "General",
                    "Fecha Contratación": fechaBase || "—",
                    "Evaluación": hito.label,
                    "Estado": resultado.estado,
                    "Nota": resultado.nota,
                    "Fecha Límite": resultado.limite || "—",
                    "Días con Novedad": emp.dias_incapacidad_acumulados || 0,
                    "Estado Plan Padrino": emp.estado_plan_padrino || "Activo",
                    "Correo": emp.correo || "",
                    "Jefe Directo": emp.jefe_directo || ""
                });
            }
        });
    });

    if (filasExportar.length === 0) {
        alert('✅ No hay evaluaciones vencidas ni reprobadas en este momento. ¡Todo al día!');
        return;
    }

    // 📄 Construir el libro de Excel
    const ws = XLSX.utils.json_to_sheet(filasExportar);

    // Ajuste de ancho de columnas para mejor lectura
    ws['!cols'] = [
        { wch: 25 }, // Nombre
        { wch: 12 }, // Cédula
        { wch: 16 }, // Cargo
        { wch: 12 }, // Regional
        { wch: 10 }, // Grupo
        { wch: 14 }, // Fecha Contratación
        { wch: 14 }, // Evaluación
        { wch: 20 }, // Estado
        { wch: 8 },  // Nota
        { wch: 12 }, // Fecha Límite
        { wch: 12 }, // Días con Novedad
        { wch: 18 }, // Estado Plan Padrino
        { wch: 25 }, // Correo
        { wch: 20 }  // Jefe Directo
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Vencidos y Reprobados");

    const fechaHoyStr = hoy.toISOString().split('T')[0];
    XLSX.writeFile(wb, `Plan_Padrino_Vencidos_Reprobados_${fechaHoyStr}.xlsx`);
}

// Exportar al ámbito global
window.exportarVencidosYReprobadosExcel = exportarVencidosYReprobadosExcel;

// Globalizar funciones
window.abrirModalNovedad = abrirModalNovedad;
window.cerrarModalNovedad = cerrarModalNovedad;
window.conmutarInputsNovedad = conmutarInputsNovedad;
window.guardarNovedadOperativaFirebase = guardarNovedadOperativaFirebase;

// 🌟 EXPORTACIÓN OBLIGATORIA (Ponla al final de tu archivo junto a las otras window.X)
window.reasignarNuevoPadrinoConHistorial = reasignarNuevoPadrinoConHistorial;

// 🌟 EXPORTACIONES AL ÁMBITO GLOBAL (Evita los ReferenceError en el HTML)
window.cambiarPestañaPadrino = cambiarPestañaPadrino;
window.poblarFiltrosEstrategicos = poblarFiltrosEstrategicos;
window.convertirYConfigurarPadrino = convertirYConfigurarPadrino;
window.vincularPlanPadrinoNuevo = vincularPlanPadrinoNuevo;
window.renderizarMonitorColaboradores = renderizarMonitorColaboradores;
window.procesarCargaMasivaExcel = procesarCargaMasivaExcel;
window.imprimirCertificadoCompletoPDF = imprimirCertificadoCompletoPDF;
window.toggleOnboardingEstado = toggleOnboardingEstado;
window.procesarFotoPadrinoLocal = procesarFotoPadrinoLocal;