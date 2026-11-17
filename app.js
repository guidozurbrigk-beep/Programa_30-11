// Estado de la aplicación
const pedidosEsperados = new Map();
const codigosEscaneados = new Map();
let ultimoLeido = "";
let html5QrcodeScanner = null;

// Eventos principales
document.getElementById('input-excel').addEventListener('change', manejarCargaExcel);
document.getElementById('btn-finalizar').addEventListener('click', finalizarControl);

// 1. Procesar el Excel
function manejarCargaExcel(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const primeraHoja = workbook.SheetNames[0];
    const filas = XLSX.utils.sheet_to_json(workbook.Sheets[primeraHoja]);

    pedidosEsperados.clear();
    codigosEscaneados.clear();
    document.getElementById('resultados').innerHTML = "";

    filas.forEach(fila => {
      const rawId = fila.ID || fila.id || fila.Tracking || fila.Envio;
      if (rawId !== undefined && rawId !== null) {
        pedidosEsperados.set(String(rawId).trim(), fila);
      }
    });

    if (pedidosEsperados.size === 0) {
      alert("No se encontraron registros válidos. Verifica que exista una columna llamada 'ID'.");
      return;
    }

    document.getElementById('cant-esperados').innerText = pedidosEsperados.size;
    document.getElementById('cant-escaneados').innerText = "0";
    document.getElementById('seccion-escaner').style.display = "block";
    iniciarCamara();
  };
  reader.readAsArrayBuffer(file);
}

// 2. Activar cámara y procesar lectura
function iniciarCamara() {
  if (html5QrcodeScanner) return;
  html5QrcodeScanner = new Html5QrcodeScanner("reader", {
    fps: 15,
    qrbox: { width: 250, height: 250 }
  });
  html5QrcodeScanner.render(alEscanearExito);
}

function alEscanearExito(textoCrudo) {
  if (textoCrudo === ultimoLeido) return;
  ultimoLeido = textoCrudo;
  setTimeout(() => { ultimoLeido = ""; }, 2000);

  let idExtraido = textoCrudo.trim();
  let infoExtra = { raw: textoCrudo };

  try {
    const jsonML = JSON.parse(textoCrudo);
    if (jsonML.id) {
      idExtraido = String(jsonML.id).trim();
      infoExtra = jsonML;
    }
  } catch (err) {
    // Es un código de barras tradicional
  }

  mostrarFeedbackEscaneo(idExtraido, infoExtra);
}

function mostrarFeedbackEscaneo(idExtraido, infoExtra) {
  const feedback = document.getElementById('feedback-escaneo');
  feedback.style.display = "block";

  if (codigosEscaneados.has(idExtraido)) {
    feedback.style.backgroundColor = "#fff3cd";
    feedback.style.color = "#856404";
    feedback.innerText = `⚠️ Ya escaneado: ${idExtraido}`;
    return;
  }

  codigosEscaneados.set(idExtraido, infoExtra);
  document.getElementById('cant-escaneados').innerText = codigosEscaneados.size;

  if (navigator.vibrate) navigator.vibrate(120);

  if (pedidosEsperados.has(idExtraido)) {
    const datos = pedidosEsperados.get(idExtraido);
    const detalle = datos.Destinatario || datos.Direccion || "OK en lista";
    feedback.style.backgroundColor = "#d4edda";
    feedback.style.color = "#155724";
    feedback.innerText = `✅ ID: ${idExtraido} (${detalle})`;
  } else {
    feedback.style.backgroundColor = "#f8d7da";
    feedback.style.color = "#721c24";
    feedback.innerText = `🚨 ENVÍO EXTRA DETECTADO: ${idExtraido}`;
  }
}

// 3. Evaluar resultado al finalizar
function finalizarControl() {
  const faltantes = [];
  const extras = [];

  pedidosEsperados.forEach((datosFila, id) => {
    if (!codigosEscaneados.has(id)) faltantes.push({ id, datosFila });
  });

  codigosEscaneados.forEach((infoEscaneo, id) => {
    if (!pedidosEsperados.has(id)) extras.push({ id, infoEscaneo });
  });

  renderizarResultados(faltantes, extras);
}

function renderizarResultados(faltantes, extras) {
  const contenedor = document.getElementById('resultados');
  contenedor.innerHTML = "";

  if (faltantes.length === 0 && extras.length === 0) {
    contenedor.innerHTML = `
      <div class="card">
        <div class="alerta-ok">Los pedidos</div>
      </div>
    `;
    return;
  }

  let html = `<div class="card">`;

  if (extras.length > 0) {
    html += `
      <div class="alerta-error">
        <strong>Hay envíos extras en tu lista</strong>
        <ul>
          ${extras.map(item => {
            const detallesQR = item.infoEscaneo.receiver_id ? ` | Receptor ID: ${item.infoEscaneo.receiver_id}` : '';
            const cpQR = item.infoEscaneo.zip_code ? ` | CP: ${item.infoEscaneo.zip_code}` : '';
            return `<li><b>ID: ${item.id}</b>${cpQR}${detallesQR}</li>`;
          }).join('')}
        </ul>
      </div>
    `;
  }

  if (faltantes.length > 0) {
    html += `
      <div class="alerta-error">
        <strong>Faltan los siguientes envíos</strong>
        <ul>
          ${faltantes.map(item => {
            const resumenFila = Object.entries(item.datosFila)
              .filter(([clave]) => clave.toLowerCase() !== 'id')
              .map(([clave, valor]) => `${clave}: ${valor}`)
              .join(' - ');
            return `<li><b>ID: ${item.id}</b> (${resumenFila || 'Sin detalles adicionales'})</li>`;
          }).join('')}
        </ul>
      </div>
    `;
  }

  html += `</div>`;
  contenedor.innerHTML = html;
}