/**
 * ================================================================
 * SISTEMA DE CAJA - PUNTO MEDICAL
 * Frontend JS — versión limpia y corregida
 * ================================================================
 *
 * SOBRE CORS CON GOOGLE APPS SCRIPT
 * ─────────────────────────────────
 * Apps Script bloquea fetch desde origen 'null' (archivo local file://)
 * cuando el navegador intenta LEER la respuesta.
 *
 * Estrategia aplicada:
 *  • ESCRITURAS (registrarCajaMenor, cierreCaja):
 *      mode: 'no-cors'  →  el request llega al servidor,
 *      no leemos respuesta pero se muestra éxito optimista.
 *
 *  • LECTURAS (consultarCaja):
 *      Requiere servir el HTML desde un servidor real (http://localhost
 *      o GitHub Pages / hosting). Desde file:// siempre fallará.
 *      → Solución rápida: usa la extensión "Live Server" de VS Code,
 *        o arrastra el index.html a GitHub Pages.
 *
 * CONFIGURACIÓN
 * ─────────────
 * Edita las constantes de abajo con tus URLs reales de Apps Script.
 * Un solo script (SCRIPT_URL) maneja TODO: acople, caja menor y consulta.
 * CAJA_MENOR_SCRIPT_URL se mantiene por compatibilidad pero ya no se usa.
 * ================================================================
 */

// ==================== CONFIGURACIÓN ====================
// ⚠️ PON AQUÍ TU URL REAL — es la misma para todas las acciones
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzSj8lQUCYXutHgKQYtrQohUCD-qexDykNP3NrP9TJr6HHpK099HYQNr1SEOymanQ7qHQ/exec';
const CAJA_MENOR_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxb7RDLMuxY1Flf2PCTKjqDIxy03T1JH-xcQ3AXQ2i25x1LAM9wUUqNrxZq2QJEEHWo/exec';
const PLANTILLA_GID = '1606540802'; // GID numérico de la hoja PLANTILLA

// ==================== ESTADO ====================
let movimientos = [];

// ==================== INICIALIZACIÓN ====================
document.addEventListener('DOMContentLoaded', function () {
    // Fecha de hoy en el campo de acople
    const today = new Date();
    const fechaInput = document.getElementById('fechaAcople');
    if (fechaInput) fechaInput.value = today.toISOString().split('T')[0];

    // Fecha en el header (elemento opcional)
    const currentDateEl = document.getElementById('currentDate');
    if (currentDateEl) {
        currentDateEl.textContent = today.toLocaleDateString('es-ES', {
            weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
        });
    }

    showTab('acople');
    actualizarListaMovimientos();
    actualizarResumen();

    // Formulario de Caja Menor
    const formCajaMenor = document.getElementById('cajaMenorForm');
    if (formCajaMenor) {
        formCajaMenor.addEventListener('submit', registrarCajaMenor);
    }
});

// ==================== NAVEGACIÓN DE TABS ====================
function showTab(tab) {
    // Ocultar todos los paneles
    ['formAcople', 'formConsulta', 'formCajaMenor', 'formConsultaCM'].forEach(id => {
        document.getElementById(id)?.classList.add('hidden');
    });

    // Resetear estilos de todos los botones
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('bg-primary', 'text-white', 'shadow-md');
        btn.classList.add('bg-white', 'text-gray-600');
    });

    // Activar el tab seleccionado
    const tabMap = {
        acople:     { form: 'formAcople',     btn: 'tabAcople' },
        consulta:   { form: 'formConsulta',   btn: 'tabConsulta' },
        consultacm: { form: 'formConsultaCM', btn: 'tabConsultaCM' },
        cajamenor:  { form: 'formCajaMenor',  btn: 'tabCajaMenor' }
    };

    const target = tabMap[tab];
    if (!target) return;

    document.getElementById(target.form)?.classList.remove('hidden');
    const btn = document.getElementById(target.btn);
    if (btn) {
        btn.classList.add('bg-primary', 'text-white', 'shadow-md');
        btn.classList.remove('bg-white', 'text-gray-600');
    }
}

// ==================== MOVIMIENTOS ====================
function agregarMovimiento() {
    const tipo = document.getElementById('tipoMovimiento').value;
    const monto = parseFloat(document.getElementById('montoMovimiento').value);
    const concepto = document.getElementById('conceptoMovimiento').value.trim();

    if (!monto || isNaN(monto) || monto <= 0) {
        showToast('Ingresa un monto válido', 'error');
        return;
    }
    if (!concepto) {
        showToast('Ingresa un concepto', 'error');
        return;
    }

    movimientos.push({ tipo, monto, concepto });
    document.getElementById('montoMovimiento').value = '';
    document.getElementById('conceptoMovimiento').value = '';
    actualizarListaMovimientos();
    actualizarResumen();
}

function eliminarMovimiento(idx) {
    movimientos.splice(idx, 1);
    actualizarListaMovimientos();
    actualizarResumen();
}

function actualizarListaMovimientos() {
    const lista = document.getElementById('listaMovimientos');
    if (!lista) return;

    if (movimientos.length === 0) {
        lista.innerHTML = `
            <div class="text-center py-8 text-gray-400">
                <i class="fas fa-inbox text-4xl mb-2 block"></i>
                <p>No hay movimientos registrados</p>
                <p class="text-sm">Agrega entradas o salidas arriba</p>
            </div>`;
        return;
    }

    lista.innerHTML = '';
    movimientos.forEach((mov, idx) => {
        const isEntrada = mov.tipo === 'entrada';
        const color = isEntrada ? 'text-green-600' : 'text-red-500';
        const signo = isEntrada ? '+' : '-';
        const icon = isEntrada ? 'fa-arrow-down' : 'fa-arrow-up';
        const badgeBg = isEntrada ? 'bg-green-100' : 'bg-red-100';

        const item = document.createElement('div');
        item.className = 'flex items-center gap-3 bg-gray-50 rounded-xl px-4 py-3 shadow-sm mb-2 border border-gray-100';
        item.innerHTML = `
            <span class="${badgeBg} rounded-full p-2">
                <i class="fas ${icon} ${color}"></i>
            </span>
            <span class="font-medium flex-1 text-gray-700">${mov.concepto}</span>
            <span class="${color} font-bold text-base">${signo} ${formatCOP(mov.monto)}</span>
            <button onclick="eliminarMovimiento(${idx})"
                class="ml-2 text-gray-300 hover:text-red-500 transition-colors p-1 rounded-lg hover:bg-red-50"
                title="Eliminar">
                <i class="fas fa-trash text-sm"></i>
            </button>`;
        lista.appendChild(item);
    });
}

function actualizarResumen() {
    let entradas = 0, salidas = 0;
    movimientos.forEach(m => {
        if (m.tipo === 'entrada') entradas += m.monto;
        else salidas += m.monto;
    });

    // Elementos opcionales en el HTML
    const el = id => document.getElementById(id);
    if (el('totalEntradasDia')) el('totalEntradasDia').textContent = formatCOP(entradas);
    if (el('totalSalidasDia')) el('totalSalidasDia').textContent = formatCOP(salidas);
    if (el('balanceDia')) el('balanceDia').textContent = formatCOP(entradas - salidas);
}

// ==================== CIERRE DE CAJA ====================
async function cierreCaja() {
    if (movimientos.length === 0) {
        showToast('No hay movimientos para guardar', 'warning');
        return;
    }

    const fecha = document.getElementById('fechaAcople').value;
    if (!fecha) {
        showToast('Selecciona una fecha', 'error');
        return;
    }

    setLoading('btnCierre', true, '<i class="fas fa-lock mr-2"></i>Realizar Cierre de Caja');

    try {
        // Se usa no-cors porque Apps Script no envía CORS headers en POST desde dominios externos.
        // No podemos leer la respuesta, pero si no hay excepción de red el request llegó.
        await fetch(SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({ action: 'acopleCaja', fecha, movimientos })
        });

        showToast('Cierre de caja realizado correctamente', 'success');
        movimientos = [];
        actualizarListaMovimientos();
        actualizarResumen();

        // Obtener el PDF como base64 desde el Apps Script (no requiere auth del usuario)
        try {
            const pdfRes = await fetch(`${SCRIPT_URL}?action=getPDFBase64`);
            const pdfData = await pdfRes.json();
            if (pdfData.success && pdfData.data) {
                const byteChars = atob(pdfData.data);
                const byteArray = new Uint8Array(byteChars.length);
                for (let i = 0; i < byteChars.length; i++) byteArray[i] = byteChars.charCodeAt(i);
                const blob = new Blob([byteArray], { type: 'application/pdf' });
                const blobUrl = URL.createObjectURL(blob);
                mostrarModalComprobante(blobUrl, true);
            } else {
                throw new Error(pdfData.message || 'Error al generar PDF');
            }
        } catch (pdfErr) {
            console.warn('[cierreCaja] No se pudo obtener PDF base64, usando enlace directo:', pdfErr);
            const pdfUrl = `https://docs.google.com/spreadsheets/d/1xoSAY47E2X9pAA7_hU6ZLzfHOaE2Br-kbeO4DXhuY1Q/export`
                + `?format=pdf&gid=${PLANTILLA_GID}&portrait=true&size=letter&fitw=true&range=A2:D20`;
            mostrarModalComprobante(pdfUrl, false);
        }

    } catch (err) {
        console.error('[cierreCaja] Error:', err);
        showToast('Error al guardar. Verifica tu conexión.', 'error');
    } finally {
        setLoading('btnCierre', false, '<i class="fas fa-lock mr-2"></i>Realizar Cierre de Caja');
    }
}

function mostrarModalComprobante(url, isBlob = false) {
    // Reusar modal si ya existe
    let modal = document.getElementById('modalComprobante');
    if (modal) {
        const link = modal.querySelector('a');
        link.href = url;
        if (isBlob) {
            link.setAttribute('download', 'comprobante-caja.pdf');
            link.removeAttribute('target');
        } else {
            link.setAttribute('target', '_blank');
            link.removeAttribute('download');
        }
        modal.classList.remove('hidden');
        return;
    }

    const downloadAttr = isBlob ? 'download="comprobante-caja.pdf"' : 'target="_blank"';
    modal = document.createElement('div');
    modal.id = 'modalComprobante';
    modal.className = 'fixed inset-0 bg-black/40 flex items-center justify-center z-50';
    modal.innerHTML = `
        <div class="bg-white p-8 rounded-2xl shadow-2xl text-center max-w-sm w-full mx-4">
            <div class="w-14 h-14 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <i class="fas fa-check-circle text-green-500 text-2xl"></i>
            </div>
            <h2 class="text-lg font-bold text-gray-800 mb-2">Cierre registrado</h2>
            <p class="text-gray-500 text-sm mb-6">El comprobante está listo para descargar.</p>
            <a href="${url}" ${downloadAttr}
               class="inline-block mb-4 px-6 py-3 bg-sky-500 hover:bg-sky-600 text-white rounded-xl font-semibold transition-colors">
                <i class="fas fa-file-pdf mr-2"></i>Descargar PDF
            </a><br>
            <button onclick="document.getElementById('modalComprobante').classList.add('hidden')"
                class="text-gray-400 hover:text-gray-600 text-sm mt-2 transition-colors">
                Cerrar
            </button>
        </div>`;
    document.body.appendChild(modal);
}

// ==================== CAJA MENOR ====================
async function registrarCajaMenor(event) {
    event.preventDefault();

    const fecha = document.getElementById('fechaCajaMenor').value;
    const detalle = document.getElementById('detalleCajaMenor').value;
    const nit = document.getElementById('nitCajaMenor').value.trim();
    const proveedor = document.getElementById('proveedorCajaMenor').value.trim();
    const factura = document.getElementById('facturaCajaMenor').value.trim();
    const valor = parseFloat(document.getElementById('valorCajaMenor').value);
    const observaciones = document.getElementById('obsCajaMenor').value.trim();

    if (!fecha || !detalle || !valor || isNaN(valor) || valor <= 0) {
        showToast('Completa los campos obligatorios (fecha, detalle y valor)', 'error');
        return;
    }

    const payload = { action: 'registrarCajaMenor', fecha, detalle, nit, proveedor, factura, valor, observaciones };

    const submitBtn = document.querySelector('#cajaMenorForm [type="submit"]');
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Guardando...';
    }

    try {
        // mode: 'no-cors' evita el error de CORS al abrir desde file://
        // El request llega al servidor aunque no podamos leer la respuesta.
        // Apps Script escribe en la hoja de cálculo correctamente.
        await fetch(CAJA_MENOR_SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });

        // Con no-cors asumimos éxito si no hubo excepción de red
        showToast('Registro guardado correctamente ✓', 'success');
        document.getElementById('cajaMenorForm').reset();

    } catch (err) {
        console.error('[registrarCajaMenor] Error de red:', err);
        showToast('Error de conexión. Verifica tu internet.', 'error');
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="fas fa-save mr-2"></i>Guardar Registro';
        }
    }
}

// ==================== CONSULTA DE CAJA ====================
async function consultarCaja() {
    const mes = document.getElementById('mesConsulta').value;
    const anio = document.getElementById('anioConsulta').value;

    if (!mes || !anio) {
        showToast('Selecciona mes y año', 'error');
        return;
    }

    // Las consultas GET necesitan leer la respuesta -> CORS falla desde file://
    // Solucion: servir el HTML con un servidor local (Live Server, http-server, etc.)
    if (window.location.protocol === 'file:') {
        mostrarAvisoServidor();
        return;
    }

    setLoading('btnConsulta', true, '<i class="fas fa-search mr-2"></i>Consultar Caja');

    try {
        const response = await fetch(
            `${SCRIPT_URL}?action=consultarCaja&mes=${encodeURIComponent(mes)}&anio=${encodeURIComponent(anio)}`
        );

        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        const data = await response.json();

        if (data.success) {
            mostrarResumen(data);
        } else {
            showToast(data.message || 'No se encontraron datos para ese período', 'warning');
            document.getElementById('resumenCaja')?.classList.add('hidden');
        }
    } catch (err) {
        console.error('[consultarCaja] Error:', err);
        const esCORS = err instanceof TypeError;
        showToast(
            esCORS
                ? 'Error CORS: abre el archivo con Live Server o desde hosting'
                : 'Error al consultar la caja',
            'error'
        );
    } finally {
        setLoading('btnConsulta', false, '<i class="fas fa-search mr-2"></i>Consultar Caja');
    }
}

/** Muestra aviso inline cuando se detecta origen file:// */
function mostrarAvisoServidor() {
    const contenedor = document.getElementById('resumenCaja');
    if (!contenedor) return;
    contenedor.classList.remove('hidden');
    contenedor.innerHTML = `
        <div class="bg-amber-50 border border-amber-300 rounded-xl p-6 flex gap-4 items-start">
            <div class="text-amber-500 text-2xl mt-0.5"><i class="fas fa-exclamation-triangle"></i></div>
            <div>
                <p class="font-bold text-amber-800 mb-1">Consulta no disponible en modo archivo local</p>
                <p class="text-amber-700 text-sm mb-3">
                    Las consultas necesitan leer datos del servidor, lo que requiere que el HTML se sirva
                    desde <strong>http://</strong> y no desde <code>file://</code>.
                    Los registros <strong>si se guardan correctamente</strong> aunque estes en local.
                </p>
                <p class="text-amber-700 text-sm font-semibold mb-1">Opciones para habilitar la consulta:</p>
                <ul class="text-amber-700 text-sm space-y-1 list-disc list-inside">
                    <li>En VS Code: instala <strong>Live Server</strong> &rarr; clic derecho sobre index.html &rarr; "Open with Live Server"</li>
                    <li>Con Node: <code class="bg-amber-100 px-1 rounded">npx http-server .</code> en la carpeta del proyecto</li>
                    <li>Sube los archivos a <strong>GitHub Pages</strong> o cualquier hosting estatico</li>
                </ul>
            </div>
        </div>`;
}

function mostrarResumen(data) {
    const el = id => document.getElementById(id);
    el('resumenCaja')?.classList.remove('hidden');
    if (el('totalEntradas')) el('totalEntradas').textContent = formatCOP(data.totalEntradas || 0);
    if (el('totalSalidas')) el('totalSalidas').textContent = formatCOP(data.totalSalidas || 0);
    if (el('saldoCaja')) el('saldoCaja').textContent = formatCOP(data.saldo || 0);
    if (el('cajaTotalMes')) el('cajaTotalMes').textContent = formatCOP(data.cajaTotalMes || 0);
    if (el('cajaFinalHoja')) el('cajaFinalHoja').textContent = formatCOP(data.sobraFalta || 0);

    const cajaFinal = (data.cajaTotalMes || 0) + (data.sobraFalta || 0);
    if (el('cajaFinalMes')) el('cajaFinalMes').textContent = formatCOP(cajaFinal);
    window._cajaFinalBase = cajaFinal;
}

// ==================== CONSULTA DE CAJA MENOR ====================
async function consultarCajaMenor(mes, anio) {
    try {
        const res = await fetch(`${CAJA_MENOR_SCRIPT_URL}?action=consultarCajaMenor&mes=${encodeURIComponent(mes)}&anio=${encodeURIComponent(anio)}`);
        const data = await res.json();
        if (data.success) {
            return data.registros;
        } else {
            showToast(data.message || 'Error al consultar caja menor', 'error');
            return [];
        }
    } catch (err) {
        console.error('[consultarCajaMenor] Error:', err);
        showToast('Error de conexión', 'error');
        return [];
    }
}

// ==================== PANEL CONSULTA CAJA MENOR ====================
let _datosCM = [];

async function consultarCajaMenorPanel() {
    const mes  = document.getElementById('mesCM').value;
    const anio = document.getElementById('anioCM').value;

    if (!mes || !anio) {
        showToast('Selecciona mes y año', 'error');
        return;
    }

    if (window.location.protocol === 'file:') {
        mostrarAvisoCM();
        return;
    }

    setLoading('btnConsultaCM', true, '<i class="fas fa-search mr-2"></i>Consultar Caja Menor');

    try {
        const response = await fetch(
            `${CAJA_MENOR_SCRIPT_URL}?action=consultarCajaMenor&mes=${encodeURIComponent(mes)}&anio=${encodeURIComponent(anio)}`
        );
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const data = await response.json();

        if (data.success) {
            _datosCM = data.registros || [];

            // Mostrar total caja (H93)
            const el = id => document.getElementById(id);
            el('totalCajaMenorMes').textContent = formatCOP(data.totalCaja || 0);

            // Poblar filtros con valores únicos
            const detalles    = [...new Set(_datosCM.map(r => r.detalle).filter(Boolean))].sort();
            const proveedores = [...new Set(_datosCM.map(r => r.proveedor).filter(Boolean))].sort();

            const filtroDetalle = el('filtroDetalleCM');
            filtroDetalle.innerHTML = '<option value="">Todos los detalles</option>' +
                detalles.map(d => `<option value="${d}">${d}</option>`).join('');

            const filtroProveedor = el('filtroProveedorCM');
            filtroProveedor.innerHTML = '<option value="">Todos los proveedores</option>' +
                proveedores.map(p => `<option value="${p}">${p}</option>`).join('');

            el('resultadoCM').classList.remove('hidden');
            filtrarTablaCM();
        } else {
            showToast(data.message || 'No se encontraron datos para ese período', 'warning');
            document.getElementById('resultadoCM')?.classList.add('hidden');
        }
    } catch (err) {
        console.error('[consultarCajaMenorPanel] Error:', err);
        const esCORS = err instanceof TypeError;
        showToast(
            esCORS
                ? 'Error CORS: abre el archivo con Live Server o desde hosting'
                : 'Error al consultar caja menor',
            'error'
        );
    } finally {
        setLoading('btnConsultaCM', false, '<i class="fas fa-search mr-2"></i>Consultar Caja Menor');
    }
}

function filtrarTablaCM() {
    const filtroDetalle   = document.getElementById('filtroDetalleCM')?.value || '';
    const filtroProveedor = document.getElementById('filtroProveedorCM')?.value || '';

    let datos = _datosCM;
    if (filtroDetalle)   datos = datos.filter(r => r.detalle   === filtroDetalle);
    if (filtroProveedor) datos = datos.filter(r => r.proveedor === filtroProveedor);

    const tabla = document.getElementById('tablaCM');
    if (!tabla) return;

    if (datos.length === 0) {
        tabla.innerHTML = `
            <div class="text-center py-10 text-gray-400">
                <i class="fas fa-inbox text-4xl mb-3 block"></i>
                <p class="font-medium">Sin registros para los filtros seleccionados</p>
            </div>`;
        return;
    }

    const totalFiltrado = datos.reduce((s, r) => s + (Number(r.valor) || 0), 0);

    tabla.innerHTML = `
        <div class="overflow-x-auto">
            <table class="w-full text-sm">
                <thead>
                    <tr class="bg-gray-50 border-b border-gray-200 text-left">
                        <th class="px-4 py-3 font-semibold text-gray-600">#</th>
                        <th class="px-4 py-3 font-semibold text-gray-600">Fecha</th>
                        <th class="px-4 py-3 font-semibold text-gray-600">Detalle</th>
                        <th class="px-4 py-3 font-semibold text-gray-600">Proveedor</th>
                        <th class="px-4 py-3 font-semibold text-gray-600">Factura</th>
                        <th class="px-4 py-3 font-semibold text-gray-600 text-right">Valor</th>
                    </tr>
                </thead>
                <tbody>
                    ${datos.map((r, i) => `
                        <tr class="border-b border-gray-100 hover:bg-sky-50 transition-colors">
                            <td class="px-4 py-3 text-gray-400">${r.numero || (i + 1)}</td>
                            <td class="px-4 py-3 text-gray-700">${formatFecha(r.fecha)}</td>
                            <td class="px-4 py-3">
                                <span class="bg-sky-100 text-sky-700 text-xs font-semibold px-2.5 py-1 rounded-full whitespace-nowrap">${r.detalle || '-'}</span>
                            </td>
                            <td class="px-4 py-3 text-gray-700">${r.proveedor || '-'}</td>
                            <td class="px-4 py-3 text-gray-500">${r.factura || '-'}</td>
                            <td class="px-4 py-3 text-right font-semibold text-gray-800">${formatCOP(r.valor || 0)}</td>
                        </tr>`).join('')}
                </tbody>
                <tfoot>
                    <tr class="bg-sky-50 font-bold border-t-2 border-sky-200">
                        <td colspan="5" class="px-4 py-3 text-sky-800">Total filtrado — ${datos.length} registro${datos.length !== 1 ? 's' : ''}</td>
                        <td class="px-4 py-3 text-right text-sky-800">${formatCOP(totalFiltrado)}</td>
                    </tr>
                </tfoot>
            </table>
        </div>`;
}

function formatFecha(fecha) {
    if (!fecha) return '-';
    // Apps Script puede devolver fecha como objeto Date serializado o string YYYY-MM-DD
    if (fecha instanceof Date) {
        return fecha.toLocaleDateString('es-CO', { day: '2-digit', month: '2-digit', year: 'numeric' });
    }
    const str = fecha.toString();
    // YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
        const [y, m, d] = str.split('T')[0].split('-');
        return `${d}/${m}/${y}`;
    }
    return str;
}

function mostrarAvisoCM() {
    const contenedor = document.getElementById('resultadoCM');
    if (!contenedor) return;
    contenedor.classList.remove('hidden');
    contenedor.innerHTML = `
        <div class="bg-amber-50 border border-amber-300 rounded-xl p-6 flex gap-4 items-start">
            <div class="text-amber-500 text-2xl mt-0.5"><i class="fas fa-exclamation-triangle"></i></div>
            <div>
                <p class="font-bold text-amber-800 mb-1">Consulta no disponible en modo archivo local</p>
                <p class="text-amber-700 text-sm mb-3">
                    Las consultas necesitan leer datos del servidor, lo que requiere que el HTML se sirva
                    desde <strong>http://</strong> y no desde <code>file://</code>.
                </p>
                <p class="text-amber-700 text-sm font-semibold mb-1">Opciones para habilitar la consulta:</p>
                <ul class="text-amber-700 text-sm space-y-1 list-disc list-inside">
                    <li>En VS Code: instala <strong>Live Server</strong> &rarr; clic derecho sobre index.html &rarr; "Open with Live Server"</li>
                    <li>Con Node: <code class="bg-amber-100 px-1 rounded">npx http-server .</code> en la carpeta del proyecto</li>
                </ul>
            </div>
        </div>`;
}

// ==================== UTILIDADES ====================

/** Formatea un número en formato colombiano: $1.234.567,89 */
function formatCOP(valor) {
    const n = Number(valor) || 0;
    const partes = n.toFixed(2).split('.');
    const miles = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `$${miles},${partes[1]}`;
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const styles = {
        success: { bg: 'bg-green-500', icon: 'fa-check-circle' },
        error: { bg: 'bg-red-500', icon: 'fa-times-circle' },
        warning: { bg: 'bg-amber-500', icon: 'fa-exclamation-circle' }
    };
    const s = styles[type] || styles.success;

    const toast = document.createElement('div');
    toast.className = `${s.bg} text-white px-5 py-4 rounded-xl shadow-lg flex items-center gap-3 min-w-[280px] max-w-xs transition-all duration-300`;
    toast.innerHTML = `<i class="fas ${s.icon} text-lg flex-shrink-0"></i><span class="font-medium text-sm leading-tight">${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(16px)';
        setTimeout(() => toast.remove(), 350);
    }, 4000);
}

/**
 * @param {string} buttonId   - ID del botón
 * @param {boolean} loading   - true = mostrar spinner, false = restaurar
 * @param {string} restoreHTML - HTML original del botón (para restaurar)
 */
function setLoading(buttonId, loading, restoreHTML) {
    const btn = document.getElementById(buttonId);
    if (!btn) return;
    btn.disabled = loading;
    btn.innerHTML = loading
        ? '<i class="fas fa-spinner fa-spin mr-2"></i>Procesando...'
        : restoreHTML;
}

// ==================== EXPOSICIÓN GLOBAL ====================
// Necesario para que los atributos onclick="..." del HTML funcionen
// cuando app.js se carga como módulo o en modo estricto.
window.showTab = showTab;
window.agregarMovimiento = agregarMovimiento;
window.eliminarMovimiento = eliminarMovimiento;
window.cierreCaja = cierreCaja;
window.consultarCaja = consultarCaja;
window.consultarCajaMenor = consultarCajaMenor;
window.consultarCajaMenorPanel = consultarCajaMenorPanel;
window.filtrarTablaCM = filtrarTablaCM;