/**
 * PNGP Dashboard Pro - Com Gráficos Nativos no Excel
 * @author Claude Enhanced with ExcelJS
 */

Chart.register(ChartDataLabels);

let database = { port: [], mat: [] };
let charts = {};
let currentMateria = 'PORT';
let currentPage = 1;
const rowsPerPage = 20;

const COLORS = {
    concluido: '#10b981',
    naoConcluido: '#ef4444',
    naoIniciado: '#f6be00',
    portugues: '#004a8d',
    matematica: '#f7941d'
};

async function processData() {
    const inputs = {
        csvP: document.getElementById('csvPort').files[0],
        csvM: document.getElementById('csvMat').files[0],
        xlsP: document.getElementById('xlsxPort').files[0],
        xlsM: document.getElementById('xlsxMat').files[0]
    };

    if (!Object.values(inputs).every(f => f)) {
        alert("Favor selecionar os 4 arquivos.");
        return;
    }

    const geralP = await readExcel(inputs.xlsP);
    const geralM = await readExcel(inputs.xlsM);

    Papa.parse(inputs.csvP, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => {
            database.port = mergeData(res.data, geralP);
            Papa.parse(inputs.csvM, {
                header: true,
                skipEmptyLines: true,
                complete: (resM) => {
                    database.mat = mergeData(resM.data, geralM);
                    initDashboard();
                }
            });
        }
    });
}

function mergeData(prog, geral) {
    return prog.filter(r => r.Nome).map(row => {
        const email = (row['Endereço de e-mail'] || "").toLowerCase().trim();
        const info = geral.find(g => (g['Endereço de e-mail'] || "").toLowerCase().trim() === email);
        return {
            ...row,
            uf: info ? info['UF Residência'] || 'Não Informado' : 'Não Informado',
            idade: info ? calculateAge(info['Data de nascimento']) : 0
        };
    });
}

function calculateAge(dateStr) {
    if (!dateStr) return 0;
    const birth = new Date(dateStr);
    const age = new Date().getFullYear() - birth.getFullYear();
    return isNaN(age) ? 0 : age;
}

function readExcel(file) {
    return new Promise(resolve => {
        const reader = new FileReader();
        reader.onload = async e => {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(e.target.result);
            const ws = wb.worksheets[0];
            const data = [];
            const headers = [];
            
            ws.getRow(1).eachCell(cell => headers.push(cell.value));
            
            for (let i = 2; i <= ws.rowCount; i++) {
                const row = ws.getRow(i);
                const obj = {};
                headers.forEach((header, idx) => {
                    obj[header] = row.getCell(idx + 1).value;
                });
                if (obj[headers[0]]) data.push(obj);
            }
            resolve(data);
        };
        reader.readAsArrayBuffer(file);
    });
}

function initDashboard() {
    document.getElementById('dashContent').classList.remove('hidden');
    document.getElementById('filterContainer').classList.remove('hidden');
    document.getElementById('btnExport').classList.remove('hidden');
    
    const allData = [...database.port, ...database.mat];
    const ufs = [...new Set(allData.map(d => d.uf))].filter(u => u !== 'Não Informado').sort();
    const select = document.getElementById('filterUF');
    select.innerHTML = '<option value="ALL">Brasil (Visão Geral)</option>';
    ufs.forEach(uf => select.innerHTML += `<option value="${uf}">${uf}</option>`);

    updateDashboard();
}

function updateDashboard() {
    const uf = document.getElementById('filterUF').value;
    const filter = d => uf === "ALL" || d.uf === uf;
    const fPort = database.port.filter(filter);
    const fMat = database.mat.filter(filter);
    const combined = [...fPort, ...fMat];

    document.getElementById('kpiTotal').innerText = combined.length;
    document.getElementById('kpiPort').innerText = fPort.length;
    document.getElementById('kpiMat').innerText = fMat.length;
    const avg = combined.reduce((acc, c) => acc + c.idade, 0) / (combined.length || 1);
    document.getElementById('kpiIdade').innerText = avg.toFixed(1);

    renderCharts(fPort, fMat);
    currentPage = 1;
    renderTable();
}

function renderTable() {
    const uf = document.getElementById('filterUF').value;
    const search = document.getElementById('tableSearch').value.toLowerCase();
    
    let filtered = (currentMateria === 'PORT' ? database.port : database.mat)
        .filter(d => uf === "ALL" || d.uf === uf)
        .filter(d => d.Nome.toLowerCase().includes(search) || d['Endereço de e-mail'].toLowerCase().includes(search));

    const totalPages = Math.ceil(filtered.length / rowsPerPage);
    const start = (currentPage - 1) * rowsPerPage;
    const pageData = filtered.slice(start, start + rowsPerPage);

    const units = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];

    document.getElementById('dataTableBody').innerHTML = pageData.map(d => `
        <tr class="modern-row">
            <td class="font-bold text-slate-900">${d.Nome}</td>
            <td class="text-xs text-slate-500">${d['Endereço de e-mail']}</td>
            <td class="text-center text-[#004a8d] font-bold">${d.uf}</td>
            <td class="text-center font-bold">${d.idade}</td>
            
            ${units.map(u => {
                const dateKey = `${u} - Data de conclusão`;
                const dateValue = d[dateKey] ? d[dateKey].split(' ')[0] : '-';
                return `
                    <td>${d[u] || 'Não iniciado'}</td>
                    <td class="text-[11px] text-slate-400 font-medium">${dateValue}</td>
                `;
            }).join('')}
            
            <td class="font-bold text-[#004a8d]">${d['Curso concluído'] ? 'Concluído' : 'Pendente'}</td>
        </tr>
    `).join('');

    document.getElementById('pageLabel').innerText = `Página ${currentPage} de ${totalPages || 1} | ${filtered.length} registros encontrados`;
}

function toggleTable(m) {
    currentMateria = m;
    currentPage = 1;
    document.getElementById('btnTabPort').classList.toggle('active', m === 'PORT');
    document.getElementById('btnTabMat').classList.toggle('active', m === 'MAT');
    renderTable();
}

function prevPage() { if (currentPage > 1) { currentPage--; renderTable(); } }
function nextPage() { 
    const uf = document.getElementById('filterUF').value;
    const search = document.getElementById('tableSearch').value.toLowerCase();
    const filtered = (currentMateria === 'PORT' ? database.port : database.mat)
        .filter(d => uf === "ALL" || d.uf === uf)
        .filter(d => d.Nome.toLowerCase().includes(search) || d['Endereço de e-mail'].toLowerCase().includes(search));
    
    if (currentPage * rowsPerPage < filtered.length) { currentPage++; renderTable(); } 
}

function renderCharts(port, mat) {
    const uts = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    Object.values(charts).forEach(c => c.destroy());

    charts.ut = new Chart(document.getElementById('chartUT'), {
        type: 'bar',
        data: {
            labels: uts,
            datasets: [
                { label: 'Português', data: uts.map(u => countUT(port, u)), backgroundColor: COLORS.portugues },
                { label: 'Matemática', data: uts.map(u => countUT(mat, u)), backgroundColor: COLORS.matematica }
            ]
        },
        options: { 
            responsive: true, maintainAspectRatio: false,
            plugins: { datalabels: { anchor: 'end', align: 'top', font: { weight: 'bold', size: 12 } } }
        }
    });

    const getStats = (data) => {
        const res = { "Concluído": 0, "Não concluído": 0, "Não iniciado": 0 };
        data.forEach(r => uts.forEach(u => { if(res[r[u]] !== undefined) res[r[u]]++; }));
        return res;
    };

    const sP = getStats(port);
    const sM = getStats(mat);

    const pieOptions = {
        responsive: true, maintainAspectRatio: false,
        plugins: { 
            datalabels: { 
                color: '#fff', font: { weight: 'bold', size: 12 },
                formatter: (v, ctx) => {
                    const sum = ctx.chart.data.datasets[0].data.reduce((a,b) => a+b, 0);
                    return v > 0 ? `${v}\n(${(v*100/sum).toFixed(1)}%)` : '';
                }
            }
        }
    };

    charts.pPort = new Chart(document.getElementById('piePort'), {
        type: 'pie',
        data: { labels: Object.keys(sP), datasets: [{ data: Object.values(sP), backgroundColor: [COLORS.concluido, COLORS.naoConcluido, COLORS.naoIniciado] }]},
        options: pieOptions
    });
    charts.pMat = new Chart(document.getElementById('pieMat'), {
        type: 'pie',
        data: { labels: Object.keys(sM), datasets: [{ data: Object.values(sM), backgroundColor: [COLORS.concluido, COLORS.naoConcluido, COLORS.naoIniciado] }]},
        options: pieOptions
    });

    const createStack = (id, data) => new Chart(document.getElementById(id), {
        type: 'bar',
        data: {
            labels: uts,
            datasets: [
                { label: 'Concluído', data: uts.map(u => data.filter(d => d[u] === 'Concluído').length), backgroundColor: COLORS.concluido },
                { label: 'Não Concluído', data: uts.map(u => data.filter(d => d[u] === 'Não concluído').length), backgroundColor: COLORS.naoConcluido },
                { label: 'Não Iniciado', data: uts.map(u => data.filter(d => (d[u] || 'Não iniciado') === 'Não iniciado').length), backgroundColor: COLORS.naoIniciado }
            ]
        },
        options: { 
            responsive: true, maintainAspectRatio: false,
            scales: { x: { stacked: true }, y: { stacked: true } },
            plugins: { datalabels: { color: '#fff', font: { size: 11, weight: 'bold' } } }
        }
    });

    charts.sPort = createStack('statusPort', port);
    charts.sMat = createStack('statusMat', mat);
}

function countUT(data, ut) { return data.filter(d => d[ut] && d[ut] !== "Não iniciado").length; }

/**
 * ========================================
 * EXPORTAÇÃO PARA EXCEL COM GRÁFICOS NATIVOS
 * ========================================
 */
async function exportToExcel() {
    const uf = document.getElementById('filterUF').value;
    const search = document.getElementById('tableSearch').value.toLowerCase();
    const filter = d => uf === "ALL" || d.uf === uf;
    
    const fPort = database.port.filter(filter).filter(d => 
        d.Nome.toLowerCase().includes(search) || d['Endereço de e-mail'].toLowerCase().includes(search)
    );
    const fMat = database.mat.filter(filter).filter(d => 
        d.Nome.toLowerCase().includes(search) || d['Endereço de e-mail'].toLowerCase().includes(search)
    );

    const wb = new ExcelJS.Workbook();
    wb.creator = 'PNGP Analytics';
    wb.created = new Date();
    
    // ABA 1: PORTUGUÊS
    await createTableSheet(wb, 'Português', fPort);
    
    // ABA 2: MATEMÁTICA
    await createTableSheet(wb, 'Matemática', fMat);
    
    // ABA 3: DASHBOARD COM GRÁFICOS
    await createDashboardSheet(wb, fPort, fMat);
    
    // Gera e baixa o arquivo
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PNGP_Analytics_${uf === 'ALL' ? 'Brasil' : uf}_${new Date().toISOString().split('T')[0]}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
}

async function createTableSheet(wb, sheetName, data) {
    const ws = wb.addWorksheet(sheetName);
    const units = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    
    // Cabeçalhos
    const headers = ['Nome', 'E-mail', 'UF', 'Idade'];
    units.forEach(u => {
        headers.push(u);
        headers.push(`${u} - Data`);
    });
    headers.push('Status Geral');
    
    ws.addRow(headers);
    
    // Estilizar cabeçalho
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF004a8d' }
    };
    ws.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    ws.getRow(1).height = 25;
    
    // Dados
    data.forEach(row => {
        const rowData = [
            row.Nome,
            row['Endereço de e-mail'],
            row.uf,
            row.idade
        ];
        
        units.forEach(u => {
            rowData.push(row[u] || 'Não iniciado');
            const dateKey = `${u} - Data de conclusão`;
            rowData.push(row[dateKey] ? row[dateKey].split(' ')[0] : '-');
        });
        
        rowData.push(row['Curso concluído'] ? 'Concluído' : 'Pendente');
        ws.addRow(rowData);
    });
    
    // Larguras das colunas
    ws.getColumn(1).width = 35; // Nome
    ws.getColumn(2).width = 40; // E-mail
    ws.getColumn(3).width = 8;  // UF
    ws.getColumn(4).width = 10; // Idade
    for (let i = 5; i <= 16; i++) {
        ws.getColumn(i).width = 15;
    }
    ws.getColumn(17).width = 18; // Status Geral
    
    // Bordas e cores alternadas
    for (let i = 2; i <= ws.rowCount; i++) {
        ws.getRow(i).eachCell(cell => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
        });
        if (i % 2 === 0) {
            ws.getRow(i).fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFF8FAFC' }
            };
        }
    }
}

async function createDashboardSheet(wb, portData, matData) {
    const ws = wb.addWorksheet('Dashboard');
    const uts = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    
    let currentRow = 1;
    
    // === SEÇÃO 1: KPIs ===
    ws.mergeCells(`A${currentRow}:D${currentRow}`);
    ws.getCell(`A${currentRow}`).value = 'INDICADORES PRINCIPAIS';
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 14, color: { argb: 'FF004a8d' } };
    ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
    currentRow += 2;
    
    ws.addRow(['Métrica', 'Português', 'Matemática', 'Total']);
    ws.getRow(currentRow).font = { bold: true };
    ws.getRow(currentRow).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' }
    };
    currentRow++;
    
    ws.addRow(['Total de Alunos', portData.length, matData.length, portData.length + matData.length]);
    currentRow++;
    
    const avgPort = (portData.reduce((acc, c) => acc + c.idade, 0) / (portData.length || 1)).toFixed(1);
    const avgMat = (matData.reduce((acc, c) => acc + c.idade, 0) / (matData.length || 1)).toFixed(1);
    const avgTotal = ((portData.reduce((acc, c) => acc + c.idade, 0) + matData.reduce((acc, c) => acc + c.idade, 0)) / ((portData.length + matData.length) || 1)).toFixed(1);
    ws.addRow(['Média de Idade', avgPort, avgMat, avgTotal]);
    currentRow += 2;
    
    // === SEÇÃO 2: ALUNOS POR UT ===
    ws.mergeCells(`A${currentRow}:D${currentRow}`);
    ws.getCell(`A${currentRow}`).value = 'ALUNOS POR UNIDADE DE TRABALHO';
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 14, color: { argb: 'FF004a8d' } };
    ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
    currentRow += 2;
    
    const utTableStart = currentRow;
    ws.addRow(['Unidade', 'Português', 'Matemática', 'Total']);
    ws.getRow(currentRow).font = { bold: true };
    ws.getRow(currentRow).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' }
    };
    currentRow++;
    
    uts.forEach(ut => {
        const countPort = countUT(portData, ut);
        const countMat = countUT(matData, ut);
        ws.addRow([ut, countPort, countMat, countPort + countMat]);
        currentRow++;
    });
    
    // Dados prontos para criação manual de gráfico no Excel
    
    currentRow += 2;
    
    // === SEÇÃO 3: STATUS GERAL - PORTUGUÊS ===
    ws.mergeCells(`A${currentRow}:D${currentRow}`);
    ws.getCell(`A${currentRow}`).value = 'STATUS GERAL - PORTUGUÊS';
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 14, color: { argb: 'FF004a8d' } };
    ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
    currentRow += 2;
    
    const statsPort = getStatsForExcel(portData, uts);
    const portPieStart = currentRow;
    ws.addRow(['Status', 'Quantidade', 'Percentual']);
    ws.getRow(currentRow).font = { bold: true };
    ws.getRow(currentRow).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' }
    };
    currentRow++;
    
    ws.addRow(['Concluído', statsPort['Concluído'], `${((statsPort['Concluído'] / (statsPort.total || 1)) * 100).toFixed(1)}%`]);
    currentRow++;
    ws.addRow(['Não Concluído', statsPort['Não concluído'], `${((statsPort['Não concluído'] / (statsPort.total || 1)) * 100).toFixed(1)}%`]);
    currentRow++;
    ws.addRow(['Não Iniciado', statsPort['Não iniciado'], `${((statsPort['Não iniciado'] / (statsPort.total || 1)) * 100).toFixed(1)}%`]);
    currentRow++;
    
    // Dados prontos para criação manual de gráfico de pizza no Excel
    
    currentRow += 2;
    
    // === SEÇÃO 4: STATUS GERAL - MATEMÁTICA ===
    ws.mergeCells(`A${currentRow}:D${currentRow}`);
    ws.getCell(`A${currentRow}`).value = 'STATUS GERAL - MATEMÁTICA';
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 14, color: { argb: 'FFf7941d' } };
    ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
    currentRow += 2;
    
    const statsMat = getStatsForExcel(matData, uts);
    const matPieStart = currentRow;
    ws.addRow(['Status', 'Quantidade', 'Percentual']);
    ws.getRow(currentRow).font = { bold: true };
    ws.getRow(currentRow).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' }
    };
    currentRow++;
    
    ws.addRow(['Concluído', statsMat['Concluído'], `${((statsMat['Concluído'] / (statsMat.total || 1)) * 100).toFixed(1)}%`]);
    currentRow++;
    ws.addRow(['Não Concluído', statsMat['Não concluído'], `${((statsMat['Não concluído'] / (statsMat.total || 1)) * 100).toFixed(1)}%`]);
    currentRow++;
    ws.addRow(['Não Iniciado', statsMat['Não iniciado'], `${((statsMat['Não iniciado'] / (statsMat.total || 1)) * 100).toFixed(1)}%`]);
    currentRow++;
    
    // Dados prontos para criação manual de gráfico de pizza no Excel
    
    currentRow += 2;
    
    // === SEÇÃO 5: STATUS POR UT - PORTUGUÊS ===
    ws.mergeCells(`A${currentRow}:E${currentRow}`);
    ws.getCell(`A${currentRow}`).value = 'STATUS POR UNIDADE - PORTUGUÊS';
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 14, color: { argb: 'FF004a8d' } };
    ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
    currentRow += 2;
    
    const portStackStart = currentRow;
    ws.addRow(['Unidade', 'Concluído', 'Não Concluído', 'Não Iniciado']);
    ws.getRow(currentRow).font = { bold: true };
    ws.getRow(currentRow).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' }
    };
    currentRow++;
    
    uts.forEach(ut => {
        const conc = portData.filter(d => d[ut] === 'Concluído').length;
        const naoc = portData.filter(d => d[ut] === 'Não concluído').length;
        const naoi = portData.filter(d => (d[ut] || 'Não iniciado') === 'Não iniciado').length;
        ws.addRow([ut, conc, naoc, naoi]);
        currentRow++;
    });
    
    // Dados prontos para criação manual de gráfico de barras empilhadas no Excel
    
    currentRow += 2;
    
    // === SEÇÃO 6: STATUS POR UT - MATEMÁTICA ===
    ws.mergeCells(`A${currentRow}:E${currentRow}`);
    ws.getCell(`A${currentRow}`).value = 'STATUS POR UNIDADE - MATEMÁTICA';
    ws.getCell(`A${currentRow}`).font = { bold: true, size: 14, color: { argb: 'FFf7941d' } };
    ws.getCell(`A${currentRow}`).alignment = { horizontal: 'center' };
    currentRow += 2;
    
    const matStackStart = currentRow;
    ws.addRow(['Unidade', 'Concluído', 'Não Concluído', 'Não Iniciado']);
    ws.getRow(currentRow).font = { bold: true };
    ws.getRow(currentRow).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDBEAFE' }
    };
    currentRow++;
    
    uts.forEach(ut => {
        const conc = matData.filter(d => d[ut] === 'Concluído').length;
        const naoc = matData.filter(d => d[ut] === 'Não concluído').length;
        const naoi = matData.filter(d => (d[ut] || 'Não iniciado') === 'Não iniciado').length;
        ws.addRow([ut, conc, naoc, naoi]);
        currentRow++;
    });
    
    // Dados prontos para criação manual de gráfico de barras empilhadas no Excel
    
    // Larguras das colunas
    ws.getColumn(1).width = 20;
    ws.getColumn(2).width = 15;
    ws.getColumn(3).width = 15;
    ws.getColumn(4).width = 15;
    ws.getColumn(5).width = 15;
}

function getStatsForExcel(data, uts) {
    const res = { "Concluído": 0, "Não concluído": 0, "Não iniciado": 0, total: 0 };
    data.forEach(r => {
        uts.forEach(u => {
            const status = r[u] || 'Não iniciado';
            if(res[status] !== undefined) {
                res[status]++;
                res.total++;
            }
        });
    });
    return res;
}