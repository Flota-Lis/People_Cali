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

     // 🔍 BUSCADOR EN VIVO: dispara el render cada vez que el usuario escribe
    const inputBusqueda = document.getElementById('search-colaborador');
    if (inputBusqueda) {
        inputBusqueda.addEventListener('input', () => {
            window.paginaActualMonitor = 1; // reinicia a la página 1 al buscar
            renderizarMonitorColaboradores();
        });
    }

    // ✨ 🧙‍♂️ AUTOCUMPLEMENTADO INTELIGENTE POR CÉDULA (CON CARGA DE FECHA AUTOMÁTICA)
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
                    
                    const campoFecha = document.getElementById('padreApadrinadoFechaIngreso');
                    if (campoFecha) {
                        campoFecha.value = empData.fecha_ingreso || (empData.hitos && empData.hitos.fecha_contratacion) || '';
                    }

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

// 📊 1 y 5. ESCUCHADOR CENTRALIZADO UNIFICADO DE PLATAFORMA (INDICADORES, POPUP PREVENTIVO Y FOLLETO DE PADRINOS)
let popupMostradoEnSesionActual = false;

db.collection('empleados').onSnapshot((snapshot) => {
    // 🌟 Sincronización mandatoria para el caché del monitor de colaboradores
    cacheSnapshotEmpleadosLocal = snapshot;

    let ingresosMes = 0; let ingresosAnio = 0;
    let padrinosActivos = 0;
    let planesActivos = 0; let planesFinalizados = 0; let planesVencidos = 0;
    let tOnboardingTotal = 0; let tOnboardingComp = 0;
    let examenesRealizados = 0; let examenesPendientes = 0;

    let sumaNota7 = 0;
    let sumaNota30 = 0;
    let sumaNota90 = 0;
    let totalApadrinados = 0;

    let totalEvaluadosOnboarding = 0;
    let arrayReprobadosHTML = [];
    
    // Arrays de memoria para alimentar el popup modal preventivo
    let listadoVencidosPopupHTML = [];
    let listadoProximosPopupHTML = [];

    // Colecciones temporales para armar el Folleto Digital de Tutores Certificados
    const padrinosTotales = []; 
    const padrinosActivosDisponibles = [];

    // ⏱️ Fecha fija de control de auditoría
    const hoy = new Date('2026-06-08T00:00:00');
    const setRegiones = new Set();
    const setMeses = new Set();

    // Calculador interno de fechas límite basado en días calendario reales
    const calcularFechaLimiteISO = (fechaIngresoStr, dias) => {
        if (!fechaIngresoStr) return null;
        const d = new Date(fechaIngresoStr + 'T00:00:00');
        d.setDate(d.getDate() + dias);
        return d.toISOString().split('T')[0];
    };

    // 🔄 UN SOLO RECORRIDO MAESTRO SOBRE LOS DOCUMENTOS VIVOS de FIRESTORE
    snapshot.forEach((doc) => {
        const emp = doc.data();
        
        // Extracción de datos estructurales para los filtros selectores
        if (emp.region) setRegiones.add(emp.region.trim());
        if (emp.fecha_ingreso) setMeses.add(emp.fecha_ingreso.substring(0, 7));

        // 🤝 EVALUACIÓN Y CAPTURA DINÁMICA DE PAQUETE DE PADRINOS
        if (emp.es_padrino === true || emp.es_padrino === "true") {
            padrinosTotales.push(emp);
            if (emp.padrino_estado !== "Inactivo") {
                padrinosActivosDisponibles.push(emp);
                padrinosActivos++;
            }
        }

        // 👥 EVALUACIÓN Y CAPTURA DINÁMICA DE APADRINADOS ACTIVOS
        if (emp.es_apadrinado === true) {
            if (miRol === 'padrino' && String(emp.padrino_id) !== String(cedulaLogueada)) return;

            totalApadrinados++; 

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
                const nSafety = emp.hitos.onboarding_safety_nota || 0;
                const nPeople = emp.hitos.onboarding_people_nota || 0;
                const n7 = emp.hitos.eval_tecnico_nota || 0;
                const n30 = emp.hitos.eval_funcional_nota || 0;
                const n90 = emp.hitos.eval_autonomo_nota || 0;

                const fBase = emp.fecha_ingreso || emp.hitos.fecha_contratacion || "";

                const hitosPlazos = [
                    { nombre: "Safety (D1)", nota: nSafety, dias: 0 },
                    { nombre: "People (D2)", nota: nPeople, dias: 1 },
                    { nombre: "7 Días (Téc)", nota: n7, dias: 9 },
                    { nombre: "30 Días (Fun)", nota: n30, dias: 32 },
                    { nombre: "90 Días (Aut)", nota: n90, dias: 92 }
                ];

                if (fBase) {
                    hitosPlazos.forEach(p => {
                        const limiteStr = calcularFechaLimiteISO(fBase, p.dias);
                        if (!limiteStr) return;
                        
                        const limiteFecha = new Date(limiteStr + 'T00:00:00');
                        const diferenciaMilisegundos = limiteFecha - hoy;
                        const diasRestantes = Math.ceil(diferenciaMilisegundos / (1000 * 60 * 60 * 24));

                        if (p.nota === 0) {
                            if (hoy > limiteFecha) {
                                listadoVencidosPopupHTML.push(`
                                    <div class="popup-item-alerta popup-bg-rojo" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; border-radius: 8px; font-size: 12px; background: #fff5f5; border: 1px solid #f09595; margin-bottom: 5px;">
                                        <div><strong>${emp.nombre}</strong><br><span style="font-size:10px; color:#a32d2d;">Hito pendiente: ${p.nombre}</span></div>
                                        <div style="text-align:right; font-weight:700; color:#a32d2d;">Venció: ${limiteStr}</div>
                                    </div>
                                `);
                            } else if (diasRestantes >= 0 && diasRestantes <= 3) {
                                const textoDias = diasRestantes === 0 ? "¡Vence hoy!" : `Vence en ${diasRestantes} días`;
                                listadoProximosPopupHTML.push(`
                                    <div class="popup-item-alerta popup-bg-amarillo" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; border-radius: 8px; font-size: 12px; background: #fff8e1; border: 1px solid #ffc404; margin-bottom: 5px;">
                                        <div><strong>${emp.nombre}</strong><br><span style="font-size:10px; color:#8a6e00;">Examen por rendir: ${p.nombre}</span></div>
                                        <div style="text-align:right; font-weight:700; color:#8a6e00;">${textoDias} (${limiteStr})</div>
                                    </div>
                                `);
                            }
                        }
                    });
                }

                let cargoFallas = [];
                if (nSafety > 0 && nSafety < 100) cargoFallas.push(`Safety (${nSafety}%)`);
                if (nPeople > 0 && nPeople < 80)  cargoFallas.push(`People (${nPeople}%)`);

                if (nSafety > 0 || nPeople > 0) totalEvaluadosOnboarding++;

                if (cargoFallas.length > 0) {
                    arrayReprobadosHTML.push(`
                        <div class="reprobado-row-item">
                            <div style="flex: 1;">
                                <strong style="font-size:12px; color:#1c2430;">${emp.nombre}</strong>
                                <div style="font-size:10.5px; color:#7a8f99;">CC ${emp.cedula} · CD: ${emp.region || 'General'}</div>
                            </div>
                            <div style="text-align: right;">
                                <span style="font-size:10.5px; font-weight:700; color:#a32d2d; background:#fff0f0; padding:3px 8px; border-radius:4px; border:1px solid #f5c2c2;">
                                    ❌ ${cargoFallas.join(' y ')}
                                </span>
                            </div>
                        </div>
                    `);
                }

                if (emp.hitos.eval_tecnico_nota > 0) { sumaNota7 += emp.hitos.eval_tecnico_nota; examenesRealizados++; } else { examenesPendientes++; }
                if (emp.hitos.eval_funcional_nota > 0) { sumaNota30 += emp.hitos.eval_funcional_nota; examenesRealizados++; } else { examenesPendientes++; }
                
                let n90Unificada = 0;
                let n90Arr = [];
                if (emp.hitos.eval_autonomo_pre_nota) n90Arr.push(emp.hitos.eval_autonomo_pre_nota);
                if (emp.hitos.eval_autonomo_nota) n90Arr.push(emp.hitos.eval_autonomo_nota);
                if (emp.hitos.eval_autonomo_post_nota) n90Arr.push(emp.hitos.eval_autonomo_post_nota);
                if (n90Arr.length > 0) n90Unificada = Math.round(n90Arr.reduce((a, b) => a + b, 0) / n90Arr.length);

                if (n90Unificada > 0) { sumaNota90 += n90Unificada; }
            }
        }
    });

    poblarFiltrosEstrategicos(Array.from(setRegiones), Array.from(setMeses).sort());

    // ⚡ INYECCIÓN DE RENDIMIENTO EN EL MENÚ SELECTOR Y EN EL FOLLETO CV DIGITAL DE PADRINOS
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

    // ⚡ RENDERIZADO INMEDIATO DE LA TABLA INTEGRAL DE AVANCE (MONITOR INDUCCIÓN)
    renderizarMonitorColaboradores();

    // ⚡ DISPARADOR DE AUDITORÍA AUTOMÁTICA DEL POPUP MODAL EN BIENVENIDA
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

        if (tieneAlertasVivas && modal) {
            modal.style.display = 'flex';
        }
        popupMostradoEnSesionActual = true;
    }

    // Inyectar datos analíticos en las tarjetas de control fijas del dashboard
    if (document.getElementById('dashTotalReprobados')) {
        document.getElementById('dashTotalReprobados').textContent = arrayReprobadosHTML.length;
        document.getElementById('dashTotalReprobados').style.background = arrayReprobadosHTML.length > 0 ? '#FCEBEB' : '#e1f5ee';
        document.getElementById('dashTotalReprobados').style.color = arrayReprobadosHTML.length > 0 ? '#a32d2d' : '#0f6e56';
    }
    if (document.getElementById('dashEvaluadosOnboarding')) {
        document.getElementById('dashEvaluadosOnboarding').textContent = totalEvaluadosOnboarding;
    }
    if (document.getElementById('lista-desplegable-reprobados')) {
        document.getElementById('lista-desplegable-reprobados').innerHTML = arrayReprobadosHTML.length === 0
            ? `<div style="text-align:center; padding:15px; color:#0f6e56; font-size:11.5px; font-weight:600; background:#e1f5ee; border-radius:6px;">✅ Todos los colaboradores se encuentran aprobados al día.</div>`
            : arrayReprobadosHTML.join('');
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

async function vincularPlanPadrinoNuevo() {
    const cedula = document.getElementById('padreApadrinadoId').value.trim();
    const nombre = document.getElementById('padreApadrinadoNombre').value.trim();
    const cargo = document.getElementById('padreApadrinadoCargo').value.trim();
    const area = document.getElementById('padreApadrinadoArea').value.trim();
    const region = document.getElementById('padreApadrinadoRegion').value.trim();
    const fechaIngresoStr = document.getElementById('padreApadrinadoFechaIngreso').value;
    const jefe = document.getElementById('padreApadrinadoJefe').value.trim();
    const correo = document.getElementById('padreApadrinadoCorreo').value.trim();
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
            estado_plan_padrino: "Activo", etapa_actual: etapa,
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

// ⚡ 6. MONITOR DETALLADO DE HITOS CON EXCLUSIÓN POR DESCARGA EFECTIVIZADA (REMOVIDO DETALLE CRITERIOS)
// ⚡ 6. MONITOR DE HITOS CON LOGICA DE EJECUCIÓN PURA (CADA ETAPA VALE 20%)
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
    let empleadosHistorico = [];

    cacheSnapshotEmpleadosLocal.forEach((doc) => {
        const emp = doc.data();
        const docId = doc.id;

        if (miRol === 'padrino' && String(emp.padrino_id) !== String(cedulaLogueada)) return;
        if (regFiltro && emp.region !== regFiltro) return;
        if (mesFiltro && (!emp.fecha_ingreso || !emp.fecha_ingreso.startsWith(mesFiltro))) return;

        if (emp.certificado_descargado === true) {
            empleadosHistorico.push({ emp, docId });
        } else {
            todosLosEmpleados.push({ emp, docId });
        }
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

    const hoy = new Date('2026-06-08T00:00:00');

    const calcularFechaLimiteReal = (fechaBaseStr, diasAsumidos) => {
        if (!fechaBaseStr) return null;
        const d = new Date(fechaBaseStr + 'T00:00:00');
        d.setDate(d.getDate() + diasAsumidos);
        return d.toISOString().split('T')[0];
    };

    const buildChipInteligente = (label, nota, fechaEjecucion, fechaBaseContratacion, diasDesfase, umbral) => {
        const limiteStr = calcularFechaLimiteReal(fechaBaseContratacion, diasDesfase);

        if (nota === undefined || nota === null || nota === 0) {
            const esVencido = limiteStr ? (hoy > new Date(limiteStr + 'T00:00:00')) : false;
            return `
                <div class="eval-chip ${esVencido ? 'chip-late-pending' : 'chip-pending'}">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">—</span>
                    <span class="chip-date">${esVencido ? '⚠️ VENCIDO' : '⏳ Límite'}</span>
                    <span style="font-size:8.5px; opacity:0.8; font-weight:bold;">${limiteStr || '—'}</span>
                </div>
            `;
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
                    <span style="font-size:8.5px; opacity:0.8;">Límite: ${limiteStr || '—'}</span>
                </div>
            `;
        }

        const ejecutoAPatirseDe = fechaEjecucion || fechaBaseContratacion;
        const ejecutoATiempo = limiteStr ? (new Date(ejecutoAPatirseDe + 'T00:00:00') <= new Date(limiteStr + 'T00:00:00')) : true;

        if (ejecutoATiempo) {
            return `
                <div class="eval-chip chip-pass">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">${nota}%</span>
                    <span class="chip-date">✓ A tiempo</span>
                    <span style="font-size:8.5px; opacity:0.8;">Hecho: ${fechaEjecucion || '—'}</span>
                </div>
            `;
        } else {
            return `
                <div class="eval-chip chip-pass-late">
                    <span class="chip-label">${label}</span>
                    <span class="chip-score">${nota}%</span>
                    <span class="chip-date">🎓 Fuera Plazo</span>
                    <span style="font-size:8.5px; opacity:0.8;">Hecho: ${fechaEjecucion || '—'}</span>
                </div>
            `;
        }
    };

    // ============================================================
    // 🆕 BLOQUE CORREGIDO: MONITOR DE ONBOARDING (D1 SAFETY / D2 PEOPLE)
    // Esta sección es 100% DE SOLO LECTURA: las notas vienen del Excel
    // que se carga en admin.html → procesarOnboardingExcel(), que escribe
    // en hitos.onboarding_safety_nota / hitos.onboarding_people_nota.
    // Aquí solo pintamos el estado resultante: Aprobado / Pendiente / Reprobado.
    // ============================================================
    if (onboardingContainer) {
        if (todosLosEmpleados.length === 0) {
            onboardingContainer.innerHTML = `
                <div style="text-align:center; padding:18px; color:#7a8f99; font-size:12px;">
                    No hay colaboradores que coincidan con el filtro actual.
                </div>`;
        } else {
            // Construye el estado visual de una columna (Safety o People)
            const buildEstadoOnboarding = (nota, umbral, label) => {
                const notaNum = Number(nota) || 0;

                if (notaNum === 0) {
                    return `
                        <div>
                            <div style="font-size:9.5px; font-weight:700; color:#7a8f99; text-transform:uppercase;">${label}</div>
                            <div style="font-size:12px; font-weight:700; color:#8a8f99; display:flex; align-items:center; gap:4px;">
                                ⏳ Pendiente
                            </div>
                        </div>`;
                }

                if (notaNum >= umbral) {
                    return `
                        <div>
                            <div style="font-size:9.5px; font-weight:700; color:#7a8f99; text-transform:uppercase;">${label}</div>
                            <div style="font-size:12px; font-weight:700; color:#0f6e56; display:flex; align-items:center; gap:4px;">
                                ✅ Aprobado
                            </div>
                        </div>`;
                }

                return `
                    <div>
                        <div style="font-size:9.5px; font-weight:700; color:#7a8f99; text-transform:uppercase;">${label}</div>
                        <div style="font-size:12px; font-weight:700; color:#a32d2d; display:flex; align-items:center; gap:4px;">
                            ❌ Reprobado (${notaNum}%)
                        </div>
                    </div>`;
            };

            onboardingContainer.innerHTML = todosLosEmpleados.map(({ emp }) => {
                const h = emp.hitos || {};
                const fechaBase = emp.fecha_ingreso || h.fecha_contratacion || '—';

                return `
                <div class="onboard-row">
                    <div style="min-width:220px; flex:1;">
                        <div style="font-size:13px; font-weight:700; color:#1c2430; text-transform:uppercase;">${emp.nombre}</div>
                        <div style="font-size:10.5px; color:#7a8f99;">
                            Cargo: ${emp.cargo || 'Operativo'}<br>
                            Ingreso: ${fechaBase}
                        </div>
                    </div>
                    <div style="display:flex; gap:30px; align-items:center; flex-wrap:wrap;">
                        ${buildEstadoOnboarding(h.onboarding_safety_nota, 100, 'Día 1 — Safety')}
                        ${buildEstadoOnboarding(h.onboarding_people_nota, 80, 'Día 2 — People')}
                    </div>
                </div>`;
            }).join('');
        }
    }
    // ============================================================
    // FIN BLOQUE ONBOARDING
    // ============================================================

    paginaSegmentada.forEach(({ emp, docId }) => {
        const h = emp.hitos || {};
        const fechaBase = emp.fecha_ingreso || h.fecha_contratacion || "";

        let notes90 = [];
        if (h.eval_autonomo_pre_nota) notes90.push(h.eval_autonomo_pre_nota);
        if (h.eval_autonomo_nota) notes90.push(h.eval_autonomo_nota);
        if (h.eval_autonomo_post_nota) notes90.push(h.eval_autonomo_post_nota);

        let notaUnificada90 = notes90.length > 0 ? Math.round(notes90.reduce((a, b) => a + b, 0) / notes90.length) : 0;
        let fecha90Final = h.eval_autonomo_nota_fecha || h.eval_autonomo_fecha || h.eval_autonomo_pre_fecha || h.eval_autonomo_post_fecha || "";

        // 🎯 LOGICA DE EJECUCIÓN PURA SOLICITADA: Contamos de 1 a 5 cuántas etapas ya TIENEN DATOS reales ingresados
        let etapasRealizadasConteo = 0;
        let sumaNotasParaPromedio = 0;

        if (h.onboarding_safety_nota > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += h.onboarding_safety_nota; }
        if (h.onboarding_people_nota > 0) { etapasRealizadasConteo++; sumaNotasParaPromedio += h.onboarding_people_nota; }
        if (h.eval_tecnico_nota > 0)       { etapasRealizadasConteo++; sumaNotasParaPromedio += h.eval_tecnico_nota; }
        if (h.eval_funcional_nota > 0)     { etapasRealizadasConteo++; sumaNotasParaPromedio += h.eval_funcional_nota; }
        if (notaUnificada90 > 0)           { etapasRealizadasConteo++; sumaNotasParaPromedio += notaUnificada90; }

        // Cada etapa realizada suma exactamente 20% al progreso global
        const porcentajeEjecucionReal = etapasRealizadasConteo * 20;

        // El promedio de calificación se calcula dividiendo solo entre las etapas efectivamente presentadas
        const promedioCalificacionReal = etapasRealizadasConteo > 0 ? Math.round(sumaNotasParaPromedio / etapasRealizadasConteo) : 0;

        const safetyOk = h.onboarding_safety_nota >= 100;
        const peopleOk = h.onboarding_people_nota >= 80;
        const tecnicoOk = h.eval_tecnico_nota >= 60;
        const funcionalOk = h.eval_funcional_nota >= 60;
        const autonomoOk = notaUnificada90 >= 60;
        const todoCompletado = safetyOk && peopleOk && tecnicoOk && funcionalOk && autonomoOk;

        let botonCertificadoHtml = '';
        if (todoCompletado || emp.estado_plan_padrino === "Plan Padrino Finalizado") {
            botonCertificadoHtml = `<button class="status-btn active-comp" style="font-size:12px; padding:8px 14px; background:#e1f5ee; color:#0f6e56; border-color:#3cbcae;" onclick="imprimirCertificadoCompletoPDF('${docId}')">🎓 Descargar Certificado</button>`;
        } else {
            botonCertificadoHtml = `<button class="status-btn" style="font-size:12px; padding:8px 14px; background:#f1f3f5; color:#7a8f99; cursor:not-allowed;" disabled>⏳ En Progreso (${porcentajeEjecucionReal}%)</button>`;
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
            
            <!-- 📊 TARJETA INFERIOR: Promedio de Calificación Real de los exámenes presentados -->
            <div style="margin-top:8px; font-size:11px; background:#fff8e1; border:1px solid #ffc404; color:#8a6e00; padding:4px 10px; border-radius:6px; display:inline-block; font-weight:700;">
              🏅 Promedio Avance: ${promedioCalificacionReal}%
            </div>
          </div>
          
          <!-- 🎯 INSIGNIA MAESTRA SUPERIOR DERECHA CORRECTA: Muestra la ejecución real por fases (20%, 40%, 60%, 80%) -->
          <div class="status-btn" style="background:#edf1f4; border-color:#cdd5db; color:#206987; font-weight:800; font-size:12px; pointer-events:none;">
             ⏳ Fases en Progreso (${porcentajeEjecucionReal}%)
          </div>
        </div>
        
        <div class="evals-row">
          ${buildChipInteligente("Safety (D1)", h.onboarding_safety_nota, h.onboarding_safety_fecha, fechaBase, 0, 100)}
          ${buildChipInteligente("People (D2)", h.onboarding_people_nota, h.onboarding_people_fecha, fechaBase, 1, 80)}
          ${buildChipInteligente("7 Días (Téc)", h.eval_tecnico_nota, h.eval_tecnico_fecha, fechaBase, 9, 60)}
          ${buildChipInteligente("30 Días (Fun)", h.eval_funcional_nota, h.eval_funcional_fecha, fechaBase, 32, 60)}
          ${buildChipInteligente("90 Días (Aut)", notaUnificada90, fecha90Final, fechaBase, 92, 60)}
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
            ? `<div style="text-align:center; padding:15px; color:#7a8f99; font-size:12px; background:#f8fafb; border-radius:8px; border:1px dashed #dde3e7;">No se registran descargas históricas el día de hoy.</div>`
            : empleadosHistorico.map(({ emp }) => `
                <div style="display:flex; justify-content:space-between; align-items:center; background:#f4fbf9; border:1px solid #3cbcae; padding:10px; border-radius:8px; margin-bottom:6px; font-size:12px;">
                    <div><strong>✓ ${emp.nombre}</strong> <span style="font-size:11px; color:#7a8f99;">(CC ${emp.cedula})</span></div>
                    <div style="font-size:11px; font-weight:700; color:#0f6e56; background:#e1f5ee; padding:2px 8px; border-radius:4px;">Certificado Concedido</div>
                </div>
            `).join('');
    }
}

// ============================================================
// 🆕 FUNCIÓN NUEVA: Marca/Desmarca Día 1 o Día 2 de onboarding en Firestore
// ============================================================
async function toggleOnboardingEstado(docId, campo, estaCompletadoActualmente) {
    try {
        const hoyStr = new Date().toISOString().split('T')[0];
        const nuevoEstado = estaCompletadoActualmente
            ? { estado: "Pendiente", fecha_completado: "" }
            : { estado: "Completado", fecha_completado: hoyStr };

        await db.collection('empleados').doc(docId).update({
            [campo]: nuevoEstado
        });
        // No hace falta location.reload(): el onSnapshot ya refresca el cache
        // y vuelve a llamar a renderizarMonitorColaboradores() automáticamente.
    } catch (error) {
        console.error("❌ Error actualizando estado de onboarding:", error);
        alert("No se pudo actualizar el estado. Revisa la consola.");
    }
}

function toggleDetalleContenedor(id, btn) {
    const panel = document.getElementById(id);
    if (!panel) return;
    panel.classList.toggle('open');
    btn.textContent = panel.classList.contains('open') ? '▴ Ocultar Detalle' : '▾ Ver Detalle de Criterios';
}

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
                let parsedRow = row;
                if (row.length === 1 && typeof row[0] === 'string') {
                    parsedRow = row[0].split(';');
                }

                const ccColaborador = String(parsedRow[15] || '').trim().replace(/['"\s.\-,]/g, '');
                const fechaEjecucion = parsedRow[1] instanceof Date ? parsedRow[1].toISOString().split('T')[0] : new Date().toISOString().split('T')[0];

                if (!ccColaborador || ccColaborador === 'null' || ccColaborador === '0') continue;

                const scores = scoreExcelDinamico(parsedRow);
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

                    const cicloKey = resolverCicloExcel(parsedRow[17]);

                    if (cicloKey === '7') {
                        hitos.eval_tecnico_nota = avgGeneral;
                        hitos.eval_tecnico_fecha = fechaEjecucion;
                        up.etapa_actual = "Etapa Funcional";
                    } else if (cicloKey === '30') {
                        hitos.eval_funcional_nota = avgGeneral;
                        hitos.eval_funcional_fecha = fechaEjecucion;
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
            alert(`🤖 CONSOLIDACIÓN COMPLETA:\n\nSe emparejaron ${totalProcesados} cuestionarios con la temporalidad estricta.`);
            location.reload();
        } catch (err) { console.error("Error cargando pilar maestro:", err); }
    };
    reader.readAsArrayBuffer(file);
}

function scoreExcelDinamico(row) {
    let lastIdx = row.length - 1;
    while (lastIdx >= 0 && (row[lastIdx] === null || row[lastIdx] === undefined || row[lastIdx] === '')) lastIdx--;
    return { tech: scoreExcelRange(row, 18, lastIdx - 6), people: scoreExcelRange(row, lastIdx - 5, lastIdx - 1), nps: row[lastIdx] };
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

async function cargarPerfilDetalladoPadrinoCompleto(cedulaPadrino) {
    try {
        const fichaRoot = document.getElementById('ficha-detalle-padrino-root');
        if (fichaRoot) fichaRoot.style.display = 'flex';
        const snapPad = await db.collection('empleados').where('cedula', '==', String(cedulaPadrino)).limit(1).get();
        if (snapPad.empty) return;
        const pData = snapPad.docs[0].data();
        document.getElementById('viewPadName').textContent = pData.nombre;
        document.getElementById('viewPadCargo').textContent = pData.cargo || "Supervisor / Líder";
        document.getElementById('viewPadEmpresa').textContent = pData.empresa_padrino || "Bavaria AB InBev";
        document.getElementById('viewPadTiempo').textContent = pData.tiempo_compania || "N/A";
        if (document.getElementById('viewPadCorreo')) document.getElementById('viewPadCorreo').textContent = pData.correo || "Sin correo registrado";

        const boxBadge = document.getElementById('viewPadBadgeEstado');
        if (boxBadge) boxBadge.innerHTML = pData.padrino_estado === "Inactivo" ? `<span style="background: #a32d2d; color: white; padding: 3px 10px; border-radius: 12px; font-size:9.5px;">❌ Inactivo</span>` : `<span style="background: #ffc404; color: #1c2430; padding: 3px 10px; border-radius: 12px; font-size:9.5px;">🏅 Activo</span>`;

        const boxFoto = document.getElementById('contenedor-foto-padrino');
        if (boxFoto) {
            if (pData.foto_url && pData.foto_url.trim() !== "") {
                boxFoto.innerHTML = `<img src="${pData.foto_url}" style="width:85px; height:85px; object-fit:cover; border-radius:50%; display:block;">`;
            } else {
                const iniciales = pData.nombre.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
                boxFoto.innerHTML = `<div style="font-weight:bold; font-size:24px; color:rgba(255,255,255,0.9);">${iniciales}</div>`;
            }
        }

        const containerTecnicas = document.getElementById('viewPadTagsTecnicas'); if (containerTecnicas) containerTecnicas.innerHTML = '';
        if (pData.habilidades_tecnicas && pData.habilidades_tecnicas.length > 0) { pData.habilidades_tecnicas.forEach(h => { if (containerTecnicas) containerTecnicas.innerHTML += `<span class="tag-skill tag-tecnica">${h}</span>`; }); } else { if (containerTecnicas) containerTecnicas.innerHTML = `<span>Sin competencias</span>`; }
        const containerBlandas = document.getElementById('viewPadTagsBlandas'); if (containerBlandas) containerBlandas.innerHTML = '';
        if (pData.habilidades_blandas && pData.habilidades_blandas.length > 0) { pData.habilidades_blandas.forEach(h => { if (containerBlandas) containerBlandas.innerHTML += `<span class="tag-skill tag-blanda">${h}</span>`; }); } else { if (containerBlandas) containerBlandas.innerHTML = `<span>Sin competencias</span>`; }

        const muchachosSnapshot = await db.collection('empleados').where('padrino_id', '==', String(cedulaPadrino)).get();
        const containerMuchachos = document.getElementById('viewPadListaMuchachos'); const containerHistoricoNotas = document.getElementById('viewPadHistoricoNotasBody');
        if (containerMuchachos) containerMuchachos.innerHTML = ''; if (containerHistoricoNotas) containerHistoricoNotas.innerHTML = '';
        let cTecnico = 0; let cFuncional = 0; let cAutonomo = 0; let sumaNotasTutor = 0; let totalNotasContadas = 0;

        if (!muchachosSnapshot.empty) {
            muchachosSnapshot.forEach(docM => {
                const m = docM.data(); let estiloTag = 'background: #fff8e1; color: #8a6e00; border: 1px solid #ffc404;'; let nombreEtapaSimple = 'Técnico';
                if (m.etapa_actual === "Etapa Funcional") { cFuncional++; estiloTag = 'background: #e1f5ee; color: #0f6e56; border: 1px solid #3cbcae;'; nombreEtapaSimple = 'Funcional'; } else if (m.etapa_actual === "Etapa Autónomo") { cAutonomo++; estiloTag = 'background: #e1f0f5; color: #206987; border: 1px solid #206987;'; nombreEtapaSimple = 'Autónomo'; } else { cTecnico++; }
                if (containerMuchachos) containerMuchachos.innerHTML += `<div style="display:flex; align-items:center; justify-content:between; background:#f8fafb; border:1px solid #eef1f3; padding:12px; border-radius:8px; width:100%;"><div style="flex:1;"><div style="font-size:12.5px; font-weight: 700;">${m.nombre}</div></div><div><span class="tag-skill" style="${estiloTag} font-size:9.5px; border-radius:12px;">${nombreEtapaSimple}</span></div></div>`;
                const n7 = (m.hitos && m.hitos.eval_tecnico_nota !== undefined) ? m.hitos.eval_tecnico_nota : "---"; const n30 = (m.hitos && m.hitos.eval_funcional_nota !== undefined) ? m.hitos.eval_funcional_nota : "---"; const n90 = (m.hitos && m.hitos.eval_autonomo_nota !== undefined) ? m.hitos.eval_autonomo_nota : "---"; const estadoPlanLabel = m.estado_plan_padrino || "Activo";
                if (typeof n7 === "number" && n7 > 0) { sumaNotasTutor += n7; totalNotasContadas++; } if (typeof n30 === "number" && n30 > 0) { sumaNotasTutor += n30; totalNotasContadas++; } if (typeof n90 === "number" && n90 > 0) { sumaNotasTutor += n90; totalNotasContadas++; }
                if (containerHistoricoNotas) containerHistoricoNotas.innerHTML += `<tr style="border-bottom: 1px solid #eef1f3;"><td style="padding: 8px; font-weight: 600; color: #1a3540;">${m.nombre}</td><td style="padding: 8px; text-align: center; font-weight: 700; color: ${n7 >= 70 ? '#0f6e56' : '#a32d2d'}">${n7}</td><td style="padding: 8px; text-align: center; font-weight: 700; color: ${n30 >= 70 ? '#0f6e56' : '#a32d2d'}">${n30}</td><td style="padding: 8px; text-align: center; font-weight: 700; color: ${n90 >= 70 ? '#0f6e56' : '#a32d2d'}">${n90}</td><td style="padding: 8px; text-align: center;"><span class="tag" style="font-size: 9px; ${estadoPlanLabel === 'Plan Padrino Finalizado' ? 'background:#e1f5ee; color:#0f6e56;' : 'background:#fff8e1; color:#8a6e00;'}">${estadoPlanLabel === 'Plan Padrino Finalizado' ? 'Graduado 🎓' : 'En proceso'}</span></td></tr>`;
            });
        } else {
            if (containerMuchachos) containerMuchachos.innerHTML = `<div style="text-align:center; color:#7a8f99; font-size:11px; padding:10px;">Sin alumnos asignados</div>`;
            if (containerHistoricoNotas) containerHistoricoNotas.innerHTML = `<tr><td colspan="5" style="text-align:center; color:#7a8f99; padding:10px;">Este tutor no registra histórico de notas.</td></tr>`;
        }
        const campoDesempenoVisual = document.getElementById('viewPadDesempeno'); if (campoDesempenoVisual) { if (totalNotasContadas > 0) { const promedioRealPadrino = Math.round(sumaNotasTutor / totalNotasContadas); campoDesempenoVisual.textContent = promedioRealPadrino + "%"; campoDesempenoVisual.style.color = promedioRealPadrino >= 85 ? "#0f6e56" : "#a32d2d"; } else { campoDesempenoVisual.textContent = "100%"; campoDesempenoVisual.style.color = "#0f6e56"; } }
        if (document.getElementById('cvCountTecnico')) document.getElementById('cvCountTecnico').textContent = cTecnico; if (document.getElementById('cvCountFuncional')) document.getElementById('cvCountFuncional').textContent = cFuncional; if (document.getElementById('cvCountAutonomo')) document.getElementById('cvCountAutonomo').textContent = cAutonomo;
    } catch (err) { console.error("Error cargando perfil del padrino:", err); }
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