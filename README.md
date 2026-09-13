<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistema de Importación Masiva</title>
    <!-- Importar librerías necesarias -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>
    
    <!-- Fuente moderna -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    
    <style>
        :root {
            --cfe-green: #008040;
            --cfe-blue: #1B365D;
            --cfe-light: #f4f6f9;
            --text-color: #334155;
            --border-radius: 12px;
        }
        body {
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            background-color: var(--cfe-light);
            background-image: linear-gradient(135deg, #f4f6f9 0%, #e2e8f0 100%);
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            color: var(--text-color);
        }
        .container {
            background-color: #ffffff;
            padding: 48px 40px;
            border-radius: var(--border-radius);
            box-shadow: 0 10px 25px -5px rgba(0,0,0,0.1), 0 8px 10px -6px rgba(0,0,0,0.1);
            width: 100%;
            max-width: 480px;
            text-align: center;
            border-top: 5px solid var(--cfe-green);
            box-sizing: border-box;
        }
        h1 {
            color: var(--cfe-blue);
            font-size: 26px;
            font-weight: 700;
            margin-top: 0;
            margin-bottom: 12px;
            letter-spacing: -0.5px;
        }
        p {
            color: #64748b;
            font-size: 15px;
            margin-bottom: 28px;
            line-height: 1.5;
        }
        .input-group {
            margin-bottom: 24px;
            text-align: left;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: var(--cfe-blue);
            font-size: 14px;
        }
        input[type="password"], input[type="file"] {
            width: 100%;
            padding: 14px 16px;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            box-sizing: border-box;
            font-size: 15px;
            transition: all 0.2s ease;
            background-color: #f8fafc;
            color: var(--text-color);
            font-family: inherit;
        }
        input[type="password"]:focus, input[type="file"]:focus {
            outline: none;
            border-color: var(--cfe-green);
            box-shadow: 0 0 0 3px rgba(0, 128, 64, 0.15);
            background-color: #ffffff;
        }
        input[type="file"]::file-selector-button {
            background-color: var(--cfe-blue);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: 500;
            transition: background-color 0.2s;
            margin-right: 12px;
            font-family: inherit;
        }
        input[type="file"]::file-selector-button:hover {
            background-color: #122540;
        }
        button {
            background-color: var(--cfe-green);
            color: white;
            border: none;
            padding: 14px 20px;
            width: 100%;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 6px -1px rgba(0, 128, 64, 0.2);
            font-family: inherit;
        }
        button:hover {
            background-color: #006633;
            transform: translateY(-2px);
            box-shadow: 0 6px 10px -1px rgba(0, 128, 64, 0.3);
        }
        button:active {
            transform: translateY(0);
        }
        button:disabled {
            background-color: #94a3b8;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        #app-section, #progress-section {
            display: none;
        }
        .error {
            color: #b91c1c;
            background-color: #fef2f2;
            padding: 12px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            margin-top: 20px;
            border: 1px solid #fecaca;
            display: none;
        }
        .progress-bar-container {
            width: 100%;
            background-color: #e2e8f0;
            border-radius: 20px;
            margin-top: 28px;
            height: 10px;
            overflow: hidden;
            box-shadow: inset 0 1px 2px rgba(0,0,0,0.1);
        }
        .progress-bar {
            height: 100%;
            width: 0%;
            background-color: var(--cfe-green);
            background-image: linear-gradient(
                45deg, 
                rgba(255,255,255,0.15) 25%, 
                transparent 25%, 
                transparent 50%, 
                rgba(255,255,255,0.15) 50%, 
                rgba(255,255,255,0.15) 75%, 
                transparent 75%, 
                transparent
            );
            background-size: 1rem 1rem;
            transition: width 0.4s ease;
        }
        #status-text {
            margin-top: 16px;
            font-size: 14px;
            font-weight: 600;
            color: var(--cfe-blue);
        }
    </style>
</head>
<body>

<div class="container" id="login-section">
    <h1>SISTEMA DE PROCESAMIENTO MASIVO</h1>
    <p>Ingrese la contraseña para continuar con el proceso.</p>
    <div class="input-group">
        <input type="password" id="password" placeholder="Contraseña de acceso" onkeypress="if(event.key === 'Enter') checkPassword()">
    </div>
    <button onclick="checkPassword()">Ingresar</button>
    <p id="login-error" class="error">Contraseña incorrecta. Intente de nuevo.</p>
</div>

<div class="container" id="app-section">
    <h1>Procesador de Facturas</h1>
    <p>Sube un archivo <b>.zip</b> que contenga los archivos XML a procesar.</p>
    <div class="input-group">
        <input type="file" id="zip-file" accept=".zip">
    </div>
    <button id="btn-procesar" onclick="processZip()">Procesar y Generar Excel</button>
    
    <div id="progress-section">
        <div class="progress-bar-container">
            <div class="progress-bar" id="progress-bar"></div>
        </div>
        <p id="status-text">Procesando... 0%</p>
    </div>
</div>

<script>
    // 1. Verificación de Contraseña (HHMM)
    function checkPassword() {
        const input = document.getElementById('password').value.trim();
        const now = new Date();
        
        let hh = now.getHours().toString().padStart(2, '0');
        let mm = now.getMinutes().toString().padStart(2, '0');
        
        let pass1 = hh + mm;
        let pass2 = parseInt(pass1, 10).toString(); // Remueve ceros a la izquierda
        if (pass2 === '') pass2 = '0';

        if (input === pass1 || input === pass2) {
            document.getElementById('login-section').style.display = 'none';
            document.getElementById('app-section').style.display = 'block';
        } else {
            const err = document.getElementById('login-error');
            err.style.display = 'block';
            setTimeout(() => err.style.display = 'none', 3000);
        }
    }

    // Función auxiliar para obtener valores de XML ignorando namespaces si es necesario
    function getVal(xmlDoc, tagName) {
        let elements = xmlDoc.getElementsByTagName(tagName);
        if (elements.length > 0) return elements[0].textContent.trim();
        
        // Fallback por si hay namespaces
        let all = xmlDoc.getElementsByTagName("*");
        for (let i = 0; i < all.length; i++) {
            if (all[i].localName === tagName) {
                return all[i].textContent.trim();
            }
        }
        return "";
    }

    // 2. Procesamiento del ZIP
    async function processZip() {
        const fileInput = document.getElementById('zip-file');
        if (!fileInput.files.length) {
            alert("Por favor, selecciona un archivo .zip");
            return;
        }

        const btn = document.getElementById('btn-procesar');
        btn.disabled = true;
        document.getElementById('progress-section').style.display = 'block';

        const file = fileInput.files[0];
        const jszip = new JSZip();
        
        try {
            const zip = await jszip.loadAsync(file);
            const xmlFiles = [];
            
            // Filtrar solo XMLs
            zip.forEach((relativePath, zipEntry) => {
                if (!zipEntry.dir && relativePath.toLowerCase().endsWith('.xml')) {
                    xmlFiles.push(zipEntry);
                }
            });

            if (xmlFiles.length === 0) {
                alert("No se encontraron archivos XML en el ZIP.");
                btn.disabled = false;
                document.getElementById('progress-section').style.display = 'none';
                return;
            }

            const total = xmlFiles.length;
            const registros = [];
            const parser = new DOMParser();
            let omitidosDB1 = 0;

            for (let i = 0; i < total; i++) {
                const xmlString = await xmlFiles[i].async("string");
                const xmlDoc = parser.parseFromString(xmlString, "application/xml");
                
                try {
                    const tarifaElem = getVal(xmlDoc, 'TARIFA_REG');
                    const tarifa = tarifaElem ? tarifaElem.toUpperCase() : "DESCONOCIDA";

                    // Omitir si la tarifa es DB1
                    if (tarifa === "DB1") {
                        omitidosDB1++;
                        console.log(`Archivo ignorado (Tarifa DB1): ${xmlFiles[i].name}`);
                    } else {
                        let fila = {
                            "Archivo_XML": xmlFiles[i].name.split('/').pop(),
                            "Nombre": getVal(xmlDoc, 'NOMBRE'),
                            "Direccion": getVal(xmlDoc, 'DIRECC'),
                            "Calle": getVal(xmlDoc, 'CALLE1'),
                            "Colonia": getVal(xmlDoc, 'COLONIA'),
                            "CP": getVal(xmlDoc, 'CODIGO_POSTAL'),
                            "No de servicio": getVal(xmlDoc, 'RPU'),
                            "Tarifa": tarifa,
                            "No De medidor": getVal(xmlDoc, 'NUMMED1'),
                            "Multiplicador": getVal(xmlDoc, 'MULTI1'),
                            "DEMANDA CONTRATADA kW": getVal(xmlDoc, 'CARGA_CONTRATADA'),
                            "CARGA CONECTADA kW": getVal(xmlDoc, 'CARGA_CONECTADA'),
                            "TOTAL A PAGAR": getVal(xmlDoc, 'IMPTOTAL'),
                            "Inicio de factura": getVal(xmlDoc, 'FECDESDE'),
                            "termino de factura": getVal(xmlDoc, 'FECHASTA'),
                            "Factor de potencia %": getVal(xmlDoc, 'FacPot'),

                            // Inicializar todos los campos posibles para asegurar el orden
                            "kWh base": "", "kWh intermedia": "", "kWh punta": "",
                            "kW base": "", "kW intermedia": "", "kW punta": "", "KWMax": "", "kVArh": "",
                            "kWh_lec_Actual": "", "kWh_lec_Anterior": "", "kWh_Diferencias": "", "kWh_Totales": "",
                            "kW _lec_Actual": "", "kW _lec_Anterior": "", "kW _Diferencias": "", "kW _Totales": "",
                            "kVArh _lec_Actual": "", "kVArh _lec_Anterior": "", "kVArh _Diferencias": "", "kVArh _Totales": "",
                            
                            "Suministro": getVal(xmlDoc, 'IMPTE_TOT_REG_1'),
                            "Distribución": getVal(xmlDoc, 'IMPTE_TOT_REG_2'),
                            "Transmisión": getVal(xmlDoc, 'IMPTE_TOT_REG_3'),
                            "CENACE": getVal(xmlDoc, 'IMPTE_TOT_REG_4'),
                            "Generación_B": "", "Generación_I": "", "Generación_P": "",
                            "Energía": "", "Capacidad": "", "SCnMEM(1)": "", "Total_costos": ""
                        };

                        if (tarifa === "GDMTH") {
                            fila["kWh base"] = getVal(xmlDoc, 'CONSUMO3F');
                            fila["kWh intermedia"] = getVal(xmlDoc, 'CONSUMO2F');
                            fila["kWh punta"] = getVal(xmlDoc, 'CONSUMO1F');
                            fila["kW base"] = getVal(xmlDoc, 'DEMANDA3P');
                            fila["kW intermedia"] = getVal(xmlDoc, 'DEMANDA2P');
                            fila["kW punta"] = getVal(xmlDoc, 'DEMANDA1P');
                            fila["KWMax"] = getVal(xmlDoc, 'DEMANDA3P');
                            fila["kVArh"] = getVal(xmlDoc, 'KVARH');

                            fila["Generación_B"] = getVal(xmlDoc, 'IMPTE_TOT_REG_5');
                            fila["Generación_I"] = getVal(xmlDoc, 'IMPTE_TOT_REG_6');
                            fila["Generación_P"] = getVal(xmlDoc, 'IMPTE_TOT_REG_7');
                            fila["Capacidad"] = getVal(xmlDoc, 'IMPTE_TOT_REG_8');
                            fila["SCnMEM(1)"] = getVal(xmlDoc, 'IMPTE_TOT_REG_9');

                            const vals = [fila["Suministro"], fila["Distribución"], fila["Transmisión"], fila["CENACE"],
                                          fila["Generación_B"], fila["Generación_I"], fila["Generación_P"], fila["Capacidad"], fila["SCnMEM(1)"]];
                            
                            let sum = vals.reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
                            fila["Total_costos"] = sum.toString();

                        } else if (["PDBT", "GDBT", "GDMTO"].includes(tarifa)) {
                            fila["kWh_lec_Actual"] = getVal(xmlDoc, 'LECACT1');
                            fila["kWh_lec_Anterior"] = getVal(xmlDoc, 'LECANT1');
                            fila["kWh_Diferencias"] = getVal(xmlDoc, 'DIFLEC1');
                            fila["kWh_Totales"] = getVal(xmlDoc, 'DIFLEC1');

                            fila["kW _lec_Actual"] = getVal(xmlDoc, 'LECACT4');
                            fila["kW _lec_Anterior"] = getVal(xmlDoc, 'LECANT4');
                            fila["kW _Diferencias"] = getVal(xmlDoc, 'DIFLEC4');
                            fila["kW _Totales"] = getVal(xmlDoc, 'DIFLEC4');

                            fila["kVArh _lec_Actual"] = getVal(xmlDoc, 'LECACT5');
                            fila["kVArh _lec_Anterior"] = getVal(xmlDoc, 'LECANT5');
                            fila["kVArh _Diferencias"] = getVal(xmlDoc, 'DIFLEC5');
                            fila["kVArh _Totales"] = getVal(xmlDoc, 'DIFLEC5');

                            fila["Energía"] = getVal(xmlDoc, 'IMPTE_TOT_REG_5');
                            fila["Capacidad"] = getVal(xmlDoc, 'IMPTE_TOT_REG_6');
                            fila["SCnMEM(1)"] = getVal(xmlDoc, 'IMPTE_TOT_REG_7');

                            const vals = [fila["Suministro"], fila["Distribución"], fila["Transmisión"], fila["CENACE"],
                                          fila["Energía"], fila["Capacidad"], fila["SCnMEM(1)"]];
                            
                            let sum = vals.reduce((acc, val) => acc + (parseFloat(val) || 0), 0);
                            fila["Total_costos"] = sum.toString();
                        }

                        // Conceptos dinámicos
                        let conceptos = Array.from(xmlDoc.getElementsByTagName('*')).filter(el => el.localName && el.localName.startsWith('Concepto') && el.localName !== 'Conceptos');
                        let importes = Array.from(xmlDoc.getElementsByTagName('*')).filter(el => el.localName && el.localName.startsWith('Importe') && el.localName !== 'Importes');
                        
                        if (conceptos.length > 0 && importes.length > 0) {
                            for(let c = 1; c <= 20; c++) {
                                let cText = getVal(xmlDoc, `Concepto${c}`);
                                let iText = getVal(xmlDoc, `Importe${c}`);
                                
                                if (!cText && !iText) continue;
                                
                                let nombre_col = cText ? cText.replace(/\?/g, '').trim() : `Concepto_${c}`;
                                if (!nombre_col) nombre_col = `Concepto_${c}`;
                                
                                fila[nombre_col] = iText || "0";
                            }
                        }

                        registros.push(fila);
                    }

                } catch (err) {
                    console.error("Error en archivo " + xmlFiles[i].name, err);
                }

                // Actualizar progreso
                let pct = Math.round(((i + 1) / total) * 100);
                document.getElementById('progress-bar').style.width = pct + '%';
                document.getElementById('status-text').innerText = `Procesando... ${i + 1}/${total} (${pct}%)`;
            }

            // Generar Excel
            if (registros.length > 0) {
                document.getElementById('status-text').innerText = "Generando archivo Excel...";
                await generarExcel(registros);
                document.getElementById('status-text').innerText = `¡Éxito! Se procesaron ${registros.length} facturas. (Se omitieron ${omitidosDB1} con tarifa DB1).`;
            } else {
                document.getElementById('status-text').innerText = `No se generó Excel. Se omitieron ${omitidosDB1} archivos (Tarifa DB1) y no hubo otros registros.`;
            }

        } catch (e) {
            console.error(e);
            alert("Hubo un error procesando el archivo ZIP.");
        } finally {
            btn.disabled = false;
        }
    }

    // 3. Generación del Excel (estilo profesional con ExcelJS)
    async function generarExcel(registros) {
        if (!registros || registros.length === 0) return;

        const workbook = new ExcelJS.Workbook();
        // Configurar freeze pane C2 (1 row, 2 cols frozen)
        const worksheet = workbook.addWorksheet('Facturas_CFE', {
            views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }]
        });

        // Obtener todas las columnas (cabeceras únicas manteniendo un orden)
        let headersSet = new Set();
        registros.forEach(row => {
            Object.keys(row).forEach(key => headersSet.add(key));
        });
        const headers = Array.from(headersSet);
        
        // Agregar fila de cabecera
        worksheet.addRow(headers);
        worksheet.getRow(1).height = 28;

        // Estilos de cabecera
        worksheet.getRow(1).eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF1B365D' }
            };
            cell.font = {
                name: 'Calibri',
                size: 10,
                bold: true,
                color: { argb: 'FFFFFFFF' }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
            };
        });

        // Agregar datos
        registros.forEach((row, rowIndex) => {
            const rowData = headers.map(header => row[header] !== undefined ? row[header] : "");
            const excelRow = worksheet.addRow(rowData);
            excelRow.height = 20;

            excelRow.eachCell((cell) => {
                cell.font = { name: 'Calibri', size: 9.5 };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                    left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                    bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
                    right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
                };

                let val = cell.value;
                if (typeof val === 'string' && val.trim() !== '') {
                    // Checar si es número (permite un punto decimal)
                    let isNum = !isNaN(val) && !isNaN(parseFloat(val));
                    if (isNum) {
                        cell.value = parseFloat(val);
                        cell.numFmt = '#,##0.00';
                        cell.alignment = { horizontal: 'right', vertical: 'middle' };
                    } else {
                        cell.alignment = { horizontal: 'left', vertical: 'middle' };
                    }
                } else {
                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                }
            });
        });

        // Autoajustar ancho de columnas
        worksheet.columns.forEach(column => {
            let maxLen = 12; // Mínimo
            column.eachCell({ includeEmpty: false }, cell => {
                let colLen = cell.value ? cell.value.toString().length : 0;
                if (colLen > maxLen) {
                    maxLen = colLen;
                }
            });
            column.width = maxLen + 3;
        });

        // Escribir y descargar
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        saveAs(blob, "Facturas_CFE_Procesadas.xlsx");
    }
</script>

</body>
</html>
