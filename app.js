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

// Drag and drop handlers
function setupDropZone(dropZone, fileInput, statusElement) {
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
            handleFileSelect(fileInput, statusElement, dropZone.id === 'drop-zone-ml' ? 'ml' : 'sys');
        }
    });

    fileInput.addEventListener('change', () => {
        handleFileSelect(fileInput, statusElement, dropZone.id === 'drop-zone-ml' ? 'ml' : 'sys');
    });
}

setupDropZone(dropZoneMl, fileMl, statusMl);
setupDropZone(dropZoneSys, fileSys, statusSys);

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
