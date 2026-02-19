// ==================== CONFIGURACIÓN ====================
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyRLanObEt_1rvGuZ4N74hH8R04bzSa1g_nZY_boeNbgQ-CY_ImL6iSkJkkzujuzqZf5A/exec';
const CAJA_MENOR_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwtbiRqpZ1ATxGX3JNlDUoIDUdmgx1HxnIX_vhVfDfZ_c4EBCQqzausPjR-3mPIhrO-2Q/exec';
const PLANTILLA_GID = '1606540802';

// ==================== ESTADO ====================
let movimientos = [];

// ==================== INICIALIZACIÓN ====================
document.addEventListener('DOMContentLoaded', function () {
    const today = new Date();
    document.getElementById('fechaAcople').value = today.toISOString().split('T')[0];

    // Fecha actual en header (si el elemento existe)
    const currentDateEl = document.getElementById('currentDate');
    if (currentDateEl) {
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        currentDateEl.textContent = today.toLocaleDateString('es-ES', options);
    }

    showTab('acople');
    actualizarListaMovimientos();
    actualizarResumen();

    // Vincular formulario de Caja Menor
    const formCajaMenor = document.getElementById('cajaMenorForm');
    if (formCajaMenor) {
        formCajaMenor.addEventListener('submit', registrarCajaMenor);
    }

    // Vincular ajuste de caja (si el elemento existe)
    const ajusteCajaEl = document.getElementById('ajusteCaja');
    if (ajusteCajaEl) {
        ajusteCajaEl.addEventListener('input', actualizarCajaFinal);
    }
});

// ==================== NAVEGACIÓN DE TABS ====================
function showTab(tab) {
    document.getElementById('formAcople').classList.add('hidden');
    document.getElementById('formConsulta').classList.add('hidden');
    document.getElementById('formCajaMenor').classList.add('hidden');

    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('bg-primary', 'text-white', 'shadow-md');
        btn.classList.add('bg-white', 'text-gray-600');
    });

    if (tab === 'acople') {
        document.getElementById('formAcople').classList.remove('hidden');
        document.getElementById('tabAcople').classList.add('bg-primary', 'text-white', 'shadow-md');
        document.getElementById('tabAcople').classList.remove('bg-white', 'text-gray-600');
    } else if (tab === 'consulta') {
        document.getElementById('formConsulta').classList.remove('hidden');
        document.getElementById('tabConsulta').classList.add('bg-primary', 'text-white', 'shadow-md');
        document.getElementById('tabConsulta').classList.remove('bg-white', 'text-gray-600');
    } else if (tab === 'cajamenor') {
        document.getElementById('formCajaMenor').classList.remove('hidden');
        document.getElementById('tabCajaMenor').classList.add('bg-primary', 'text-white', 'shadow-md');
        document.getElementById('tabCajaMenor').classList.remove('bg-white', 'text-gray-600');
    }
}

// ==================== AGREGAR / ELIMINAR MOVIMIENTO ====================
window.agregarMovimiento = function () {
    const tipo = document.getElementById('tipoMovimiento').value;
    const monto = parseFloat(document.getElementById('montoMovimiento').value);
    const concepto = document.getElementById('conceptoMovimiento').value.trim();

    if (!monto || !concepto) {
        showToast('Completa monto y concepto', 'error');
        return;
    }

    movimientos.push({ tipo, monto, concepto });
    document.getElementById('montoMovimiento').value = '';
    document.getElementById('conceptoMovimiento').value = '';
    actualizarListaMovimientos();
    actualizarResumen();
};

window.eliminarMovimiento = function (idx) {
    movimientos.splice(idx, 1);
    actualizarListaMovimientos();
    actualizarResumen();
};

function actualizarListaMovimientos() {
    const lista = document.getElementById('listaMovimientos');
    if (!lista) return;

    lista.innerHTML = '';

    if (movimientos.length === 0) {
        lista.innerHTML = `
            <div class="text-center py-8 text-gray-400">
                <i class="fas fa-inbox text-4xl mb-2"></i>
                <p>No hay movimientos registrados</p>
                <p class="text-sm">Agrega entradas o salidas arriba</p>
            </div>`;
        return;
    }

    movimientos.forEach((mov, idx) => {
        const color = mov.tipo === 'entrada' ? 'text-success' : 'text-danger';
        const signo = mov.tipo === 'entrada' ? '+' : '-';
        const icon = mov.tipo === 'entrada' ? 'fa-arrow-down' : 'fa-arrow-up';
        const item = document.createElement('div');
        item.className = 'flex items-center gap-3 bg-gray-50 rounded-xl px-4 py-3 shadow-sm mb-2';
        item.innerHTML = `
            <i class="fas ${icon} ${color} text-lg"></i>
            <span class="font-medium flex-1">${mov.concepto}</span>
            <span class="${color} font-bold text-lg">${signo} ${formatearNumeroColombiano(mov.monto)}</span>
            <button onclick="eliminarMovimiento(${idx})" class="ml-2 text-gray-400 hover:text-red-500 transition-colors">
                <i class="fas fa-trash"></i>
            </button>`;
        lista.appendChild(item);
    });
}

function actualizarResumen() {
    let entradas = 0, salidas = 0;
    movimientos.forEach(mov => {
        if (mov.tipo === 'entrada') entradas += mov.monto;
        else salidas += mov.monto;
    });

    // Solo actualiza si los elementos existen (opcionales en el HTML)
    const elEntradas = document.getElementById('totalEntradasDia');
    const elSalidas = document.getElementById('totalSalidasDia');
    const elBalance = document.getElementById('balanceDia');

    if (elEntradas) elEntradas.textContent = formatearNumeroColombiano(entradas);
    if (elSalidas) elSalidas.textContent = formatearNumeroColombiano(salidas);
    if (elBalance) elBalance.textContent = formatearNumeroColombiano(entradas - salidas);
}

// ==================== CIERRE DE CAJA ====================
window.cierreCaja = async function () {
    if (movimientos.length === 0) {
        showToast('No hay movimientos para guardar', 'warning');
        return;
    }

    const fecha = document.getElementById('fechaAcople').value;
    setLoading('btnCierre', true);

    try {
        // NOTA: se usa no-cors porque Google Apps Script no siempre permite CORS en POST.
        // Por eso no podemos leer la respuesta — se asume éxito si no hay excepción de red.
        await fetch(SCRIPT_URL, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'acopleCaja', fecha, movimientos })
        });

        showToast('Cierre de caja realizado', 'success');
        movimientos = [];
        actualizarListaMovimientos();
        actualizarResumen();

        const pdfUrl = `https://docs.google.com/spreadsheets/d/1xoSAY47E2X9pAA7_hU6ZLzfHOaE2Br-kbeO4DXhuY1Q/export?format=pdf&gid=${PLANTILLA_GID}&portrait=true&size=letter&fitw=true&range=A2:D20`;
        mostrarEnlaceComprobante(pdfUrl);
    } catch (e) {
        console.error('[Cierre Caja] Error:', e);
        showToast('Error al guardar', 'error');
    } finally {
        setLoading('btnCierre', false);
    }
};

function mostrarEnlaceComprobante(url) {
    let modal = document.getElementById('modalComprobante');
    if (modal) {
        modal.querySelector('a').href = url;
        modal.style.display = 'flex';
        return;
    }

    modal = document.createElement('div');
    modal.id = 'modalComprobante';
    Object.assign(modal.style, {
        position: 'fixed', top: '0', left: '0',
        width: '100vw', height: '100vh',
        background: 'rgba(0,0,0,0.4)',
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        zIndex: '9999'
    });
    modal.innerHTML = `
        <div style="background:white;padding:32px 24px;border-radius:16px;box-shadow:0 4px 32px #0002;text-align:center;max-width:350px">
            <h2 style="font-size:1.2rem;font-weight:600;margin-bottom:16px">Comprobante generado</h2>
            <a href="${url}" target="_blank" style="display:inline-block;margin-bottom:16px;padding:10px 20px;background:#0ea5e9;color:white;border-radius:8px;text-decoration:none;font-weight:600">
                Descargar comprobante PDF
            </a><br>
            <button id="cerrarModalComprobante" style="background:#eee;border:none;padding:8px 16px;border-radius:8px;font-size:1rem;cursor:pointer;margin-top:8px">Cerrar</button>
        </div>`;
    document.body.appendChild(modal);
    document.getElementById('cerrarModalComprobante').onclick = () => modal.remove();
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

    if (!fecha || !detalle || !valor) {
        showToast('Completa los campos obligatorios', 'error');
        return;
    }

    const payload = { action: 'registrarCajaMenor', fecha, detalle, nit, proveedor, factura, valor, observaciones };
    console.log('[Caja Menor] Payload:', payload);

    try {
        // Content-Type: text/plain evita el preflight OPTIONS que bloquea CORS en Apps Script
        const res = await fetch(CAJA_MENOR_SCRIPT_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify(payload)
        });
        const data = await res.json();
        console.log('[Caja Menor] Resultado:', data);

        if (data.success) {
            showToast('Registro guardado correctamente', 'success');
            document.getElementById('cajaMenorForm').reset();
        } else {
            showToast(data.message || 'Error al guardar', 'error');
        }
    } catch (err) {
        console.error('[Caja Menor] Error en POST:', err);
        showToast('Error de conexión', 'error');
    }
}

window.consultarCajaMenor = async function (mes, anio) {
    try {
        const res = await fetch(`${CAJA_MENOR_SCRIPT_URL}?action=consultarCajaMenor&mes=${mes}&anio=${anio}`);
        const data = await res.json();
        if (data.success) {
            return data.registros;
        } else {
            showToast(data.message || 'Error al consultar', 'error');
            return [];
        }
    } catch (err) {
        console.error('[Caja Menor] Error consulta:', err);
        showToast('Error de conexión', 'error');
        return [];
    }
};

// ==================== CONSULTA DE CAJA ====================
window.consultarCaja = async function () {
    const mes = document.getElementById('mesConsulta').value;
    const anio = document.getElementById('anioConsulta').value;
    setLoading('btnConsulta', true);

    try {
        console.log('[Consulta Caja] Solicitando:', { mes, anio });
        const response = await fetch(`${SCRIPT_URL}?action=consultarCaja&mes=${mes}&anio=${anio}`);
        const data = await response.json();
        console.log('[Consulta Caja] Respuesta:', data);

        if (data.success) {
            mostrarResumen(data);
        } else {
            showToast(data.message || 'No se encontraron datos', 'warning');
            document.getElementById('resumenCaja').classList.add('hidden');
        }
    } catch (error) {
        console.error('[Consulta Caja] Error:', error);
        showToast('Error al consultar la caja', 'error');
    } finally {
        setLoading('btnConsulta', false);
    }
};

function mostrarResumen(data) {
    document.getElementById('resumenCaja').classList.remove('hidden');
    document.getElementById('totalEntradas').textContent = formatearNumeroColombiano(data.totalEntradas || 0);
    document.getElementById('totalSalidas').textContent = formatearNumeroColombiano(data.totalSalidas || 0);
    document.getElementById('saldoCaja').textContent = formatearNumeroColombiano(data.saldo || 0);
    document.getElementById('cajaTotalMes').textContent = formatearNumeroColombiano(data.cajaTotalMes || 0);
    document.getElementById('cajaFinalHoja').textContent = formatearNumeroColombiano(data.sobraFalta || 0);

    const cajaFinal = Number(data.cajaTotalMes || 0) + Number(data.sobraFalta || 0);
    document.getElementById('cajaFinalMes').textContent = formatearNumeroColombiano(cajaFinal);

    // Guardar base para ajuste manual si se usa
    window._cajaFinalBase = cajaFinal;
}

// ==================== AJUSTE DE CAJA FINAL ====================
window.actualizarCajaFinal = function () {
    const ajuste = parseFloat(document.getElementById('ajusteCaja')?.value) || 0;
    const cajaFinal = (window._cajaFinalBase || 0) + ajuste;
    document.getElementById('cajaFinalMes').textContent = formatearNumeroColombiano(cajaFinal);
};

// ==================== UTILIDADES ====================
function formatearNumeroColombiano(valor) {
    const partes = Number(valor).toFixed(2).split('.');
    const miles = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return `$${miles},${partes[1]}`;
}

function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    const bgColor = type === 'success' ? 'bg-green-500' : type === 'error' ? 'bg-red-500' : 'bg-amber-500';
    const icon = type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-times-circle' : 'fa-exclamation-circle';
    toast.className = `${bgColor} text-white px-6 py-4 rounded-xl shadow-lg flex items-center gap-3 min-w-[300px] transition-all`;
    toast.innerHTML = `<i class="fas ${icon} text-xl"></i><span class="font-medium">${message}</span>`;
    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

function setLoading(buttonId, loading) {
    const button = document.getElementById(buttonId);
    if (!button) return;
    if (loading) {
        button.disabled = true;
        button.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Procesando...';
    } else {
        button.disabled = false;
        if (buttonId === 'btnCierre') {
            button.innerHTML = '<i class="fas fa-lock mr-2"></i>Realizar Cierre de Caja';
        } else if (buttonId === 'btnConsulta') {
            button.innerHTML = '<i class="fas fa-search mr-2"></i>Consultar Caja';
        }
    }
}