// Variables de memoria de la aplicación
let datosGlobales = [];    
let datosFiltrados = [];   

// Escucha el evento de carga de archivos
document.getElementById('excel_file').addEventListener('change', cargarExcel, false);

function cargarExcel(e) {
    const archivo = e.target.files[0];
    if(!archivo) return;

    const lector = new FileReader();
    lector.onload = function(event) {
        const datos = new Uint8Array(event.target.result);
        const libro = XLSX.read(datos, {type: 'array'});
        const hoja = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]]);
        procesarMatematicas(hoja);
    };
    lector.readAsArrayBuffer(archivo);
}

function formatMoneda(valor) {
    return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(valor);
}

function procesarMatematicas(datosRaw) {
    datosGlobales = [];

    datosRaw.forEach(registro => {
        const codFactura = String(registro.CodFactura || registro.Factura || 'SIN_CODIGO').toUpperCase().trim();
        const cedula = String(registro.Cedula || registro.Identificacion || 'N/A').toUpperCase().trim();
        const empleado = String(registro.Empleado || registro.Nombre || 'Desconocido').toUpperCase().trim();
        
        const salario = parseFloat(registro.SalarioBase || 0);
        const auxTrans = parseFloat(registro.AuxilioTransporte || 0);
        const otros = parseFloat(registro.OtrosDevengados || 0);
        
        const devengadoTotal = salario + auxTrans + otros;
        const baseVacacionesYSS = salario + otros;

        const prima = devengadoTotal * 0.0833;
        const cesantias = devengadoTotal * 0.0833;
        const intCesantias = prima * 0.12; 
        const vacaciones = baseVacacionesYSS * 0.0417;
        const totalProvisiones = prima + cesantias + intCesantias + vacaciones;

        const smmlvReferencia = 1500000; 
        const aplicaExencion = devengadoTotal < (smmlvReferencia * 10);
        
        const salud = aplicaExencion ? 0 : (baseVacacionesYSS * 0.085);
        const pension = baseVacacionesYSS * 0.12;
        const arl = baseVacacionesYSS * 0.00522; 
        const totalSegSocial = salud + pension + arl;

        const caja = baseVacacionesYSS * 0.04;
        const sena = aplicaExencion ? 0 : (baseVacacionesYSS * 0.02);
        const icbf = aplicaExencion ? 0 : (baseVacacionesYSS * 0.03);
        const totalParafiscales = caja + sena + icbf;

        const granTotal = totalProvisiones + totalSegSocial + totalParafiscales;

        datosGlobales.push({
            codFactura, cedula, empleado, devengadoTotal,
            prima, cesantias, intCesantias, vacaciones, totalProvisiones,
            pension, salud, arl, totalSegSocial,
            caja, sena, icbf, totalParafiscales,
            granTotal
        });
    });

    document.getElementById('filter_panel').classList.remove('hidden');
    document.getElementById('summary_row').classList.remove('hidden');
    document.getElementById('btn_exportar').classList.remove('hidden');
    
    aplicarFiltros(); 
}

function aplicarFiltros() {
    const filtroFac = document.getElementById("filtro_factura").value.toUpperCase().trim();
    const filtroGen = document.getElementById("filtro_general").value.toUpperCase().trim();

    datosFiltrados = datosGlobales.filter(row => {
        const coincideFactura = filtroFac === "" || row.codFactura.includes(filtroFac);
        const coincideGeneral = filtroGen === "" || row.cedula.includes(filtroGen) || row.empleado.includes(filtroGen);
        return coincideFactura && coincideGeneral;
    });

    renderizarTabla(datosFiltrados);
}

function renderizarTabla(datos) {
    let html = '';
    let sumDev = 0, sumProv = 0, sumSeg = 0, sumPar = 0, sumTot = 0;

    if (datos.length === 0) {
        html = `<tr><td colspan="15" class="p-8 text-center text-slate-500 font-sans">No se encontraron registros contables que coincidan con los criterios de búsqueda actuales.</td></tr>`;
    } else {
        datos.forEach(row => {
            sumDev += row.devengadoTotal;
            sumProv += row.totalProvisiones;
            sumSeg += row.totalSegSocial;
            sumPar += row.totalParafiscales;
            sumTot += row.granTotal;

            html += `
            <tr class="border-b border-slate-200 hover:bg-slate-100 transition-colors">
                <td class="p-2 border-r border-slate-200 font-sans font-bold text-slate-900">${row.codFactura}</td>
                <td class="p-2 border-r border-slate-200 font-sans">${row.cedula}</td>
                <td class="p-2 border-r border-slate-200 font-sans text-xs">${row.empleado}</td>
                <td class="p-2 text-right border-r border-slate-200 bg-slate-50 font-bold">${formatMoneda(row.devengadoTotal)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-blue-700 bg-blue-50/20">${formatMoneda(row.prima)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-blue-700 bg-blue-50/20">${formatMoneda(row.cesantias)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-blue-700 bg-blue-50/20">${formatMoneda(row.intCesantias)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-blue-700 bg-blue-50/20">${formatMoneda(row.vacaciones)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-indigo-700 bg-indigo-50/20">${formatMoneda(row.pension)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-indigo-700 bg-indigo-50/20">${formatMoneda(row.salud)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-indigo-700 bg-indigo-50/20">${formatMoneda(row.arl)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-teal-700 bg-teal-50/20">${formatMoneda(row.caja)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-teal-700 bg-teal-50/20">${formatMoneda(row.sena)}</td>
                <td class="p-2 text-right border-r border-slate-200 text-teal-700 bg-teal-50/20">${formatMoneda(row.icbf)}</td>
                <td class="p-2 text-right bg-emerald-50 text-emerald-800 font-bold">${formatMoneda(row.granTotal)}</td>
            </tr>`;
        });
    }

    document.getElementById('report_body').innerHTML = html;
    document.getElementById('lbl_mostrando_regs').innerText = datos.length;
    document.getElementById('sum_devengado').innerText = formatMoneda(sumDev);
    document.getElementById('sum_prestaciones').innerText = formatMoneda(sumProv);
    document.getElementById('sum_seguridad').innerText = formatMoneda(sumSeg);
    document.getElementById('sum_parafiscales').innerText = formatMoneda(sumPar);
    document.getElementById('sum_total').innerText = formatMoneda(sumTot);
}

function exportarReporteExcel() {
    const datosParaExcel = datosFiltrados.map(row => ({
        "CodFactura": row.codFactura,
        "Cédula": row.cedula,
        "Empleado": row.empleado,
        "Base Devengado": row.devengadoTotal,
        "Prima (8.33%)": row.prima,
        "Cesantías (8.33%)": row.cesantias,
        "Int. Cesantías (1%)": row.intCesantias,
        "Vacaciones (4.17%)": row.vacaciones,
        "Pensión (12%)": row.pension,
        "Salud (8.5%)": row.salud,
        "ARL I (0.522%)": row.arl,
        "Caja (4%)": row.caja,
        "SENA (2%)": row.sena,
        "ICBF (3%)": row.icbf,
        "Total Cargos Provisión": row.granTotal
    }));

    const hoja = XLSX.utils.json_to_sheet(datosParaExcel);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, "Consulta_Auditada");
    XLSX.writeFile(libro, "Auditoria_Provisiones_Detalle.xlsx");
}

// ==========================================
// LÓGICA DEL CHATBOT: ROMARIO IA
// ==========================================

function toggleRomario() {
    const chatWindow = document.getElementById('romario-window');
    if (chatWindow.classList.contains('hidden')) {
        chatWindow.classList.remove('hidden');
        document.getElementById('romario-input').focus();
    } else {
        chatWindow.classList.add('hidden');
    }
}

function handleRomarioEnter(event) {
    if (event.key === 'Enter') {
        sendRomarioMessage();
    }
}

function sendRomarioMessage() {
    const inputField = document.getElementById('romario-input');
    const userText = inputField.value.trim();
    if (!userText) return;

    // Agregar mensaje del usuario a la vista
    appendMessage(userText, 'user');
    inputField.value = '';

    // Simular tiempo de carga para ROMARIO IA
    setTimeout(() => {
        const response = getRomarioResponse(userText.toLowerCase());
        appendMessage(response, 'bot');
    }, 500);
}

function appendMessage(text, sender) {
    const messagesContainer = document.getElementById('romario-messages');
    const messageDiv = document.createElement('div');
    
    if (sender === 'user') {
        messageDiv.className = 'bg-orange-100 border border-orange-200 p-2 rounded-lg rounded-tr-none self-end max-w-[85%] shadow-sm text-slate-800';
    } else {
        messageDiv.className = 'bg-white border border-slate-200 p-2 rounded-lg rounded-tl-none self-start max-w-[85%] shadow-sm text-slate-800';
    }
    
    messageDiv.innerHTML = text;
    messagesContainer.appendChild(messageDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
}

// Base de Conocimiento de ROMARIO IA (Restringida solo a Nómina)
function getRomarioResponse(query) {
    // Reglas de nómina colombiana
    if (query.includes('prima')) {
        return "<strong>Prima de Servicios:</strong> Corresponde a 15 días de salario por cada semestre laborado. Su provisión mensual es del <strong>8.33%</strong> y se calcula sobre el Salario Base + Auxilio de Transporte + Otros Devengados.";
    } else if (query.includes('cesantia') || query.includes('cesantía')) {
        return "<strong>Cesantías:</strong> Es un ahorro para cuando el trabajador quede cesante. Equivale a un mes de salario por cada año trabajado. Su provisión es del <strong>8.33%</strong> mensual.";
    } else if (query.includes('interes') && query.includes('cesantia')) {
        return "<strong>Intereses sobre Cesantías:</strong> Corresponde al 12% anual sobre el valor acumulado de las cesantías. En provisión mensual, equivale al <strong>1.00%</strong>.";
    } else if (query.includes('vacacion') || query.includes('vacaciones')) {
        return "<strong>Vacaciones:</strong> El trabajador tiene derecho a 15 días hábiles de descanso remunerado por cada año de servicio. Su provisión es del <strong>4.17%</strong> y <em>no incluye el auxilio de transporte</em> en su base de cálculo.";
    } else if (query.includes('salud') || query.includes('pension') || query.includes('pensión')) {
        return "<strong>Seguridad Social:</strong> El empleador aporta el <strong>12%</strong> para Pensión y el <strong>8.5%</strong> para Salud. *Recuerda que aplica exención de Salud, SENA e ICBF (Ley 1607) para empleados que devenguen menos de 10 SMMLV.";
    } else if (query.includes('parafiscales') || query.includes('caja') || query.includes('sena') || query.includes('icbf')) {
        return "<strong>Parafiscales:</strong> Los aportes patronales son Caja de Compensación (4%), SENA (2%) e ICBF (3%). Las bases menores a 10 SMMLV están exoneradas de SENA e ICBF.";
    } else if (query.includes('transporte') || query.includes('auxilio')) {
        return "<strong>Auxilio de Transporte:</strong> Se paga a los trabajadores que devengan hasta 2 SMMLV. Hace base para el cálculo de Prima y Cesantías, pero NO para Vacaciones ni Seguridad Social.";
    } else if (query.includes('hola') || query.includes('saludo')) {
        return "¡Hola! Soy ROMARIO IA. Pregúntame sobre porcentajes de liquidación, primas, cesantías, vacaciones o aportes de ley.";
    } else if (query.includes('gracias')) {
        return "¡Con mucho gusto! Aquí estaré si tienes más dudas sobre cálculos laborales.";
    } else {
        // Bloqueo estricto para temas fuera de nómina
        return "⚠️ <strong>Lo siento.</strong> Soy una IA especializada de Asertempo y <em>solo estoy autorizado y entrenado para responder preguntas sobre nómina, provisiones, prestaciones sociales y seguridad social en Colombia</em>. Por favor, realiza una consulta relacionada a estos temas.";
    }
}
