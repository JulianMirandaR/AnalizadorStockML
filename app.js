// Elements
const fileMl = document.getElementById('file-ml');
const fileSys = document.getElementById('file-sys');
const statusMl = document.getElementById('status-ml');
const statusSys = document.getElementById('status-sys');
const btnProcess = document.getElementById('btn-process');
const resultsPanel = document.getElementById('results-panel');
const resultsBody = document.getElementById('results-body');
const btnDownload = document.getElementById('btn-download');

// Drop zones
const dropZoneMl = document.getElementById('drop-zone-ml');
const dropZoneSys = document.getElementById('drop-zone-sys');

// Stats data
const stats = {
    mlMayor: document.getElementById('stat-ml-mayor'),
    missingMl: document.getElementById('stat-missing-ml'),
    warning: document.getElementById('stat-warning'),
    other: document.getElementById('stat-other')
};

let dataMl = null;
let dataSys = null;
let finalResults = [];

// NEW TABS LOGIC
const tabBtns = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        tabBtns.forEach(b => b.classList.remove('active'));
        tabContents.forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
    });
});

// Drag and drop handlers
function setupDropZone(dropZone, fileInput, statusElement, type) {
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            if (type === 'ml' || type === 'sys') {
                handleFileSelect(fileInput, statusElement, type);
            } else {
                handleFileSelectRentabilidad(fileInput, statusElement, type);
            }
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length === 0) return;
        if (type === 'ml' || type === 'sys') {
            handleFileSelect(fileInput, statusElement, type);
        } else {
            handleFileSelectRentabilidad(fileInput, statusElement, type);
        }
    });
}

setupDropZone(dropZoneMl, fileMl, statusMl, 'ml');
setupDropZone(dropZoneSys, fileSys, statusSys, 'sys');

function handleFileSelect(input, statusElement, type) {
    if (input.files.length === 0) return;
    
    const file = input.files[0];
    statusElement.textContent = file.name;
    statusElement.classList.add('uploaded');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Select sheet based on file type
            let targetSheetName = workbook.SheetNames[0];
            
            // ML exports use the 'Publicaciones' sheet
            if (type === 'ml') {
                if (workbook.SheetNames.includes('Publicaciones')) {
                    targetSheetName = 'Publicaciones';
                } else if (workbook.SheetNames.length > 2) {
                    targetSheetName = workbook.SheetNames[2]; // Usually the 3rd sheet if name was changed
                } else if (workbook.SheetNames.length > 1) {
                    targetSheetName = workbook.SheetNames[1];
                }
            }
            
            const worksheet = workbook.Sheets[targetSheetName];
            
            // Convert to array of arrays
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            
            if (type === 'ml') {
                dataMl = parseMl(json);
            } else {
                dataSys = parseSys(json);
            }
            
            checkReady();
        } catch (error) {
            Swal.fire('Error', 'No se pudo leer el archivo Excel.', 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function checkReady() {
    if (dataMl && dataSys) {
        btnProcess.disabled = false;
    }
}

// parsing logic
// parsing logic
function parseMl(rows) {
    const map = {};
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        
        let sku = row[4];
        let stock = row[7];
        
        if (sku !== undefined && sku !== null && sku !== '') {
            // Remove single quotes, double quotes, and empty spaces
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            
            stock = parseInt(stock, 10);
            if (!isNaN(stock)) {
                map[sku] = stock;
            }
        }
    }
    return map;
}

function parseSys(rows) {
    const map = {};
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        
        let sku = row[3];
        let stock = row[9];
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            
            stock = parseInt(stock, 10);
            if (!isNaN(stock)) {
                map[sku] = stock;
            }
        }
    }
    return map;
}

btnProcess.addEventListener('click', () => {
    btnProcess.disabled = true;
    btnProcess.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
    
    setTimeout(() => {
        analizarDatos();
        renderTable();
        
        resultsPanel.classList.remove('hidden');
        
        btnProcess.innerHTML = '<i class="fa-solid fa-bolt"></i> Analizar y Comparar Stock';
        btnProcess.disabled = false;
        
        Swal.fire({
            title: '¡Análisis Completado!',
            text: `Se encontraron ${finalResults.length} inconsistencias.`,
            icon: 'success',
            confirmButtonColor: '#3b82f6'
        });
    }, 500); // Small delay for UX
});

function analizarDatos() {
    finalResults = [];
    
    let cMlMayor = 0;
    let cMissingMl = 0;
    let cWarning = 0;
    let cOther = 0;
    
    // Compare ML against System
    for (const rawSku in dataMl) {
        let originalMlStock = dataMl[rawSku];
        let skuForSystem = rawSku;
        let multiplier = 1;
        
        // Detectar si es un kit (ej: KITX2-32277)
        const kitMatch = rawSku.match(/^KITX(\d+)-(.*)/);
        if (kitMatch) {
            multiplier = parseInt(kitMatch[1], 10);
            skuForSystem = kitMatch[2]; // el SKU real en el sistema
        }
        
        const mlStockEq = originalMlStock * multiplier;
        
        if (dataSys.hasOwnProperty(skuForSystem)) {
            const sysStock = dataSys[skuForSystem];
            
            if (mlStockEq !== sysStock) {
                const diff = mlStockEq - sysStock;
                let motivo = '';
                let badgeClass = '';
                
                if (mlStockEq > sysStock && mlStockEq < 10 && sysStock < 10) {
                    motivo = 'Stock ML mayor que Sistema (Riesgo)';
                    badgeClass = 'bg-danger';
                    cMlMayor++;
                } else if (Math.abs(diff) <= 2 && mlStockEq < 8 && sysStock < 8) {
                    motivo = 'Diferencia de 2 unidades o menos (Advertencia)';
                    badgeClass = 'bg-warning';
                    cWarning++;
                } else if (mlStockEq < 4 && sysStock > 6) {
                    motivo = 'Actualizar ML (Sistema tiene más stock)';
                    badgeClass = 'bg-info';
                    cOther++;
                } else {
                    // No cumple con ninguna de las reglas prioritarias de corrección
                    continue;
                }
                
                finalResults.push({
                    SKU: rawSku,
                    'Stock ML': originalMlStock,
                    'Multiplier': multiplier, // para mostrarlo en la tabla
                    'Stock Sistema': sysStock,
                    Diferencia: diff,
                    Motivo: motivo,
                    _badgeClass: badgeClass
                });
            }
        } else {
            // Note: User didn't request this case, but good to have if needed. Omitting for exact compliance.
        }
    }
    
    // Compare System against ML (for missing items)
    for (const sku in dataSys) {
        const sysStock = dataSys[sku];
        
        // Rule 3: System has SKU with stock, but it's not in ML
        if (!dataMl.hasOwnProperty(sku) && sysStock > 0) {
            cMissingMl++;
            finalResults.push({
                SKU: sku,
                'Stock ML': 'No existe',
                'Stock Sistema': sysStock,
                Diferencia: -sysStock,
                Motivo: 'SKU no existe en ML (Publicar/Activar)',
                _badgeClass: 'bg-success'
            });
        }
    }
    
    // Update stats
    stats.mlMayor.textContent = cMlMayor;
    stats.missingMl.textContent = cMissingMl;
    stats.warning.textContent = cWarning;
    stats.other.textContent = cOther;
}

function renderTable() {
    resultsBody.innerHTML = '';
    
    // Sort by absolute difference from highest to lowest
    finalResults.sort((a, b) => {
        return Math.abs(b.Diferencia) - Math.abs(a.Diferencia);
    });
    
    finalResults.forEach(item => {
        const tr = document.createElement('tr');
        
        const mlDisplay = item.Multiplier > 1 
            ? `${item['Stock ML']} <span style="font-size: 0.8rem; color: var(--warning);">(x${item.Multiplier})</span>`
            : item['Stock ML'];
            
        tr.innerHTML = `
            <td><strong>${item.SKU}</strong></td>
            <td>${mlDisplay}</td>
            <td>${item['Stock Sistema']}</td>
            <td><span class="status-badge ${item._badgeClass}">${item.Motivo}</span></td>
        `;
        
        resultsBody.appendChild(tr);
    });
}

btnDownload.addEventListener('click', () => {
    // Preparar data para XLSX sin la propiedad _badgeClass
    const excelData = finalResults.map(item => {
        return {
            'SKU': item.SKU,
            'Stock Mercado Libre': item['Stock ML'],
            'Stock Sistema': item['Stock Sistema'],
            'Motivo de Corrección': item.Motivo
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    // Ajustar anchos
    const wscols = [
        {wch: 20}, // SKU
        {wch: 20}, // ML
        {wch: 20}, // Sys
        {wch: 50}  // Motivo
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Correcciones Stock");
    
    // Save file
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Conciliacion_Stock_ML_${dateStr}.xlsx`);
});

// --- NEW RENTABILIDAD LOGIC ---
const fileGuerrini = document.getElementById('file-guerrini');
const fileVentas = document.getElementById('file-ventas');
const statusGuerrini = document.getElementById('status-guerrini');
const statusVentas = document.getElementById('status-ventas');
const btnProcessRentabilidad = document.getElementById('btn-process-rentabilidad');
const resultsPanelRentabilidad = document.getElementById('results-panel-rentabilidad');
const resultsBodyRentabilidad = document.getElementById('results-body-rentabilidad');
const btnDownloadRentabilidad = document.getElementById('btn-download-rentabilidad');

const statGanancia = document.getElementById('stat-ganancia');
const statPerdida = document.getElementById('stat-perdida');

const dropZoneGuerrini = document.getElementById('drop-zone-guerrini');
const dropZoneVentas = document.getElementById('drop-zone-ventas');

let dataGuerrini = {};
let dataVentas = [];
let finalResultsRentabilidad = [];

setupDropZone(dropZoneGuerrini, fileGuerrini, statusGuerrini, 'guerrini');
setupDropZone(dropZoneVentas, fileVentas, statusVentas, 'ventas');

function handleFileSelectRentabilidad(input, statusElement, type) {
    if (input.files.length === 0) return;
    
    const file = input.files[0];
    statusElement.textContent = file.name;
    statusElement.classList.add('uploaded');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (type === 'guerrini') {
                dataGuerrini = {}; // Reset
                // Parse all sheets in Guerrini list
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                    parseGuerrini(json);
                });
            } else if (type === 'ventas') {
                // Ventas ML report typically uses the first active sheet
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                dataVentas = parseVentas(json);
            }
            
            if (Object.keys(dataGuerrini).length > 0 && dataVentas.length > 0) {
                btnProcessRentabilidad.disabled = false;
            }
        } catch (error) {
            Swal.fire('Error', 'No se pudo leer el archivo Excel.', 'error');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseGuerrini(rows) {
    // Col A (0) = SKU, Col H (7) = Precio Base
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 8) continue;
        
        let sku = row[0];
        let price = row[7];
        
        if (sku !== undefined && sku !== null && sku !== '') {
            sku = sku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            // Handle possibility of string formatted price (e.g., "$ 1,000.00")
            if(typeof price === 'string') {
               price = parseFloat(price.replace(/[^0-9,-]+/g, '').replace(',', '.'));
            } else {
               price = parseFloat(price);
            }
            
            if (!isNaN(price) && price > 0) {
                // Apple discounts: 35%, 25%, 8% => price * 0.65 * 0.75 * 0.92 = price * 0.4485
                let finalCost = price * 0.4485;
                dataGuerrini[sku] = finalCost;
            }
        }
    }
}

function parseVentas(rows) {
    const list = [];
    // Start reading from row 6 (index 5)
    // Col V (21) = SKU, Col S (18) = Sold Price, Col G (6) = Quantity
    for (let i = 5; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 25) continue;
        
        let dateRaw = row[1];    // B
        let paramSku = row[21]; // V
        let soldPrice = row[18]; // S
        let quantity = row[6];   // G
        let title = row[24];     // Y
        
        if (paramSku !== undefined && paramSku !== null && paramSku !== '') {
            let sku = paramSku.toString().replace(/^['"]+/, '').replace(/['"]+$/, '').trim().toUpperCase();
            
            if(typeof soldPrice === 'string') {
               soldPrice = parseFloat(soldPrice.replace(/[^0-9,-]+/g, '').replace(',', '.'));
            } else {
               soldPrice = parseFloat(soldPrice);
            }
            
            if(typeof quantity === 'string') {
               quantity = parseFloat(quantity.replace(/[^0-9,-]+/g, '').replace(',', '.'));
            } else {
               quantity = parseFloat(quantity);
            }
            
            // Prevent division by zero
            if (isNaN(quantity) || quantity <= 0) {
               quantity = 1;
            }
            
            if (!isNaN(soldPrice)) {
                // Divide the total sold price by the quantity to get the unit sold price
                let unitSoldPrice = soldPrice / quantity;
                
                // Ignorar ventas con montos muy bajos, suelen ser devoluciones
                if (unitSoldPrice > 1000) {
                    list.push({ 
                        rowIndex: i, 
                        dateRaw: dateRaw, 
                        sku: sku, 
                        title: title || 'Sin detalle', 
                        soldPrice: unitSoldPrice, 
                        quantity: quantity 
                    });
                }
            }
        }
    }
    return list;
}

btnProcessRentabilidad.addEventListener('click', () => {
    btnProcessRentabilidad.disabled = true;
    btnProcessRentabilidad.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';
    
    setTimeout(() => {
        analizarRentabilidad();
        renderTableRentabilidad();
        
        resultsPanelRentabilidad.classList.remove('hidden');
        
        btnProcessRentabilidad.innerHTML = '<i class="fa-solid fa-bolt"></i> Analizar Rentabilidad';
        btnProcessRentabilidad.disabled = false;
        
        Swal.fire({
            title: '¡Análisis de Rentabilidad Completado!',
            text: `Se analizaron ${finalResultsRentabilidad.length} ventas.`,
            icon: 'success',
            confirmButtonColor: '#3b82f6'
        });
    }, 500); 
});

function formatter(val) {
    return new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS' }).format(val);
}

function analizarRentabilidad() {
    finalResultsRentabilidad = [];
    let gananciaCount = 0;
    let perdidaCount = 0;

    dataVentas.forEach(venta => {
        // Find if base SKU is present (or support Kit variations just in case, but usually exact SKU match is fine for sales)
        let cost = dataGuerrini[venta.sku];
        
        // Let's also verify kit parsing like we did for stock, just in case they sell kits.
        const kitMatch = venta.sku.match(/^KITX(\d+)-(.*)/);
        if (kitMatch && cost === undefined) {
            let multiplier = parseInt(kitMatch[1], 10);
            let realSku = kitMatch[2];
            if(dataGuerrini[realSku] !== undefined) {
               cost = dataGuerrini[realSku] * multiplier;
            }
        }

        if (cost !== undefined) {
            const unitProfit = venta.soldPrice - cost;
            const totalDiferencia = unitProfit * venta.quantity;
            
            const estado = totalDiferencia > 0 ? 'Ganancia' : 'Pérdida';
            const badgeClass = totalDiferencia > 0 ? 'bg-success' : 'bg-danger';
            
            if (totalDiferencia > 0) gananciaCount++;
            else perdidaCount++;

            finalResultsRentabilidad.push({
                OriginalIndex: venta.rowIndex,
                Fecha: venta.dateRaw,
                SKU: venta.sku,
                Title: venta.title,
                CostoFinal: cost,
                PrecioVendido: venta.soldPrice,
                CantidadVendida: venta.quantity,
                Resultado: totalDiferencia,
                Estado: estado,
                _badgeClass: badgeClass
            });
        }
    });

    statGanancia.textContent = gananciaCount;
    statPerdida.textContent = perdidaCount;
}

function renderTableRentabilidad() {
    resultsBodyRentabilidad.innerHTML = '';
    
    // El reporte de ML originalmente suele venir ordenado por fecha de forma descendente.
    // Usamos el OriginalIndex para garantizar un orden cronológico exacto al del Excel madre,
    // garantizando que las fechas queden correlativas en orden.
    finalResultsRentabilidad.sort((a, b) => a.OriginalIndex - b.OriginalIndex);
    
    finalResultsRentabilidad.forEach(item => {
        const tr = document.createElement('tr');
        
        let displayDate = item.Fecha ? item.Fecha.toString() : '';
        // Si Excel envió la fecha como código serial (muy frecuente)
        if (typeof item.Fecha === 'number') {
            displayDate = new Date((item.Fecha - (25567 + 2)) * 86400 * 1000).toLocaleDateString('es-AR');
        } else if (displayDate.length > 10 && displayDate.includes('T')) {
            displayDate = displayDate.split('T')[0]; // Para fechas ISO reducidas
        } else if (displayDate.length > 15) {
            displayDate = displayDate.substring(0, 10); // Reducirlo para no ocupar mucho
        }
        
        tr.innerHTML = `
            <td><span style="font-size: 0.8rem; color: var(--text-secondary); white-space: nowrap;">${displayDate}</span></td>
            <td><strong>${item.SKU}</strong></td>
            <td><span style="font-size: 0.85rem; color: var(--text-secondary);">${item.Title}</span></td>
            <td>${formatter(item.CostoFinal)}</td>
            <td>${formatter(item.PrecioVendido)}</td>
            <td><strong>x${item.CantidadVendida}</strong></td>
            <td class="${item.Resultado > 0 ? 'text-success' : 'text-danger'}"><strong>${item.Resultado > 0 ? '+' : ''}${formatter(item.Resultado)}</strong></td>
            <td><span class="status-badge ${item._badgeClass}">${item.Estado}</span></td>
        `;
        
        resultsBodyRentabilidad.appendChild(tr);
    });
}

btnDownloadRentabilidad.addEventListener('click', () => {
    const excelData = finalResultsRentabilidad.map(item => {
        return {
            'Fecha Venta': item.Fecha,
            'SKU': item.SKU,
            'Detalle del Producto': item.Title,
            'Costo Guerrini (Por Artículo Publicado) ($)': item.CostoFinal,
            'Precio de Venta ML (Por Artículo Publicado) ($)': item.PrecioVendido,
            'Unidades de este SKU Vendidas': item.CantidadVendida,
            'Ganancia/Pérdida Total del Renglón ($)': item.Resultado,
            'Estado': item.Estado
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    const wscols = [
        {wch: 15}, // Fecha
        {wch: 25}, // SKU
        {wch: 45}, // Detalle
        {wch: 20}, // Costo
        {wch: 20}, // Venta
        {wch: 20}, // Cantidad
        {wch: 25}, // Resultado
        {wch: 15}  // Estado
    ];
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rentabilidad");
    
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Analisis_Rentabilidad_${dateStr}.xlsx`);
});
