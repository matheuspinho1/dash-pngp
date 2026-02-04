Chart.register(ChartDataLabels);

let database = { port: [], mat: [] };
let charts = {};
let currentMateria = 'PORT';
let currentPage = 1;
const rowsPerPage = 20;

const COLORS = {
    concluido: '#004a8d',
    naoConcluido: '#f7941d',
    naoIniciado: '#f6be00',
    portugues: '#004a8d',
    matematica: '#f7941d'
};

// --- PROCESSAMENTO DE DADOS ---

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
            cidade: info ? info['Cidade'] || 'Não Informado' : 'Não Informado',
            idade: info ? calculateAge(info['Data de nascimento']) : 0
        };
    });
}

function calculateAge(dateStr) {
    if (!dateStr) return 0;
    const birth = new Date(dateStr);
    const age = 2026 - birth.getFullYear(); 
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
                headers.forEach((header, idx) => { obj[header] = row.getCell(idx + 1).value; });
                if (obj[headers[0]]) data.push(obj);
            }
            resolve(data);
        };
        reader.readAsArrayBuffer(file);
    });
}

// --- CONTROLE DA INTERFACE (HTML) ---

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
            <td class="text-center text-slate-600 font-medium">${d.cidade}</td>
            <td class="text-center font-bold">${d.idade}</td>
            ${units.map(u => {
                const dateKey = `${u} - Data de conclusão`;
                const dateValue = d[dateKey] ? d[dateKey].split(' ')[0] : '-';
                return `<td>${d[u] || 'Não iniciado'}</td><td class="text-[11px] text-slate-400 font-medium">${dateValue}</td>`;
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

// --- GRÁFICOS ---

function renderCharts(port, mat) {
    const uts = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    Object.values(charts).forEach(c => { if(c) c.destroy(); });

    const getStats = (data) => {
        const res = { "Concluído": 0, "Não concluído": 0 };
        data.forEach(r => uts.forEach(u => { if(res[r[u]] !== undefined) res[r[u]]++; }));
        return res;
    };

    const pieOptions = {
        responsive: true, maintainAspectRatio: false,
        animation: false,
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
        data: { labels: Object.keys(getStats(port)), datasets: [{ data: Object.values(getStats(port)), backgroundColor: [COLORS.concluido, COLORS.naoConcluido] }]},
        options: pieOptions
    });
    charts.pMat = new Chart(document.getElementById('pieMat'), {
        type: 'pie',
        data: { labels: Object.keys(getStats(mat)), datasets: [{ data: Object.values(getStats(mat)), backgroundColor: [COLORS.concluido, COLORS.naoConcluido] }]},
        options: pieOptions
    });

    const createStack = (id, data) => new Chart(document.getElementById(id), {
        type: 'bar',
        data: {
            labels: uts,
            datasets: [
                { label: 'Concluído', data: uts.map(u => data.filter(d => d[u] === 'Concluído').length), backgroundColor: COLORS.concluido },
                { label: 'Não Concluído', data: uts.map(u => data.filter(d => d[u] === 'Não concluído').length), backgroundColor: COLORS.naoConcluido }
            ]
        },
        options: { 
            responsive: true, maintainAspectRatio: false,
            animation: false,
            scales: { x: { stacked: true }, y: { stacked: true } },
            plugins: { datalabels: { color: '#fff', font: { size: 11, weight: 'bold' } } }
        }
    });

    charts.sPort = createStack('statusPort', port);
    charts.sMat = createStack('statusMat', mat);
}

// --- EXPORTAÇÃO EXECUTIVA PARA EXCEL ---

async function exportToExcel() {
    const uf = document.getElementById('filterUF').value;
    const filter = d => uf === "ALL" || d.uf === uf;
    const fPort = database.port.filter(filter);
    const fMat = database.mat.filter(filter);

    const chartImages = {
        piePort: charts.pPort.toBase64Image(),
        pieMat: charts.pMat.toBase64Image(),
        statusPort: charts.sPort.toBase64Image(),
        statusMat: charts.sMat.toBase64Image()
    };

    const wb = new ExcelJS.Workbook();
    wb.creator = 'PNGP Analytics';
    
    await createTableSheet(wb, 'Português', fPort);
    await createTableSheet(wb, 'Matemática', fMat);
    await createDashboardSheet(wb, fPort, fMat, chartImages, uf);
    
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PNGP_Analytics_${uf === 'ALL' ? 'Brasil' : uf}_2026.xlsx`;
    a.click();
}

async function createTableSheet(wb, sheetName, data) {
    const ws = wb.addWorksheet(sheetName);
    const units = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    
    ws.columns = [
        { header: 'Nome', key: 'nome', width: 35 },
        { header: 'E-mail', key: 'email', width: 40 },
        { header: 'UF', key: 'uf', width: 10 },
        { header: 'Cidade', key: 'cidade', width: 25 },
        { header: 'Idade', key: 'idade', width: 10 },
        { header: 'Teste inicial', key: 'ti', width: 18 },
        { header: 'TI - Data', key: 'ti_d', width: 15 },
        { header: 'Unidade 1', key: 'u1', width: 18 },
        { header: 'U1 - Data', key: 'u1_d', width: 15 },
        { header: 'Unidade 2', key: 'u2', width: 18 },
        { header: 'U2 - Data', key: 'u2_d', width: 15 },
        { header: 'Unidade 3', key: 'u3', width: 18 },
        { header: 'U3 - Data', key: 'u3_d', width: 15 },
        { header: 'Unidade 4', key: 'u4', width: 18 },
        { header: 'U4 - Data', key: 'u4_d', width: 15 },
        { header: 'Unidade 5', key: 'u5', width: 18 },
        { header: 'U5 - Data', key: 'u5_d', width: 15 },
        { header: 'Status Geral', key: 'status', width: 20 }
    ];

    const headerRow = ws.getRow(1);
    headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF004a8d' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    data.forEach(row => {
        const rowData = [row.Nome, row['Endereço de e-mail'], row.uf, row.cidade, row.idade];
        units.forEach(u => {
            rowData.push(row[u] || 'Não iniciado');
            const dateKey = `${u} - Data de conclusão`;
            rowData.push(row[dateKey] ? row[dateKey].split(' ')[0] : '-');
        });
        rowData.push(row['Curso concluído'] ? 'Concluído' : 'Pendente');
        ws.addRow(rowData);
    });

    for (let i = 2; i <= ws.rowCount; i++) {
        const row = ws.getRow(i);
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (colNumber <= 18) {
                cell.border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
                if (i % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
            }
        });
    }
}

async function createDashboardSheet(wb, portData, matData, imgs, ufNome) {
    const ws = wb.addWorksheet('Dashboard');
    
    ws.getColumn(1).width = 30; ws.getColumn(6).width = 30;
    [2,3,4,7,8,9].forEach(c => ws.getColumn(c).width = 18);

    ws.mergeCells('A1:I1');
    const mainTitle = ws.getCell('A1');
    mainTitle.value = `PNGP ANALYTICS - RELATÓRIO EXECUTIVO | ${ufNome === 'ALL' ? 'BRASIL' : ufNome}`;
    mainTitle.font = { bold: true, size: 20, color: { argb: 'FF004A8D' } };
    mainTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 45;

    let curRow = 3;
    ws.mergeCells(`A${curRow}:I${curRow}`);
    ws.getCell(`A${curRow}`).value = 'RESUMO GERAL DE MÉTRICAS';
    ws.getCell(`A${curRow}`).font = { bold: true, size: 12 };
    ws.getCell(`A${curRow}`).alignment = { horizontal: 'center' };
    curRow++;

    const header = ws.addRow(['Métrica', 'Português', 'Matemática', 'Total Geral']);
    header.eachCell(c => { c.font = { bold: true }; c.alignment = { horizontal: 'center' }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } }; });
    curRow++;

    const avgPort = (portData.reduce((acc, c) => acc + c.idade, 0) / (portData.length || 1)).toFixed(1);
    const avgMat = (matData.reduce((acc, c) => acc + c.idade, 0) / (matData.length || 1)).toFixed(1);
    const avgTotal = ((portData.reduce((acc, c) => acc + c.idade, 0) + matData.reduce((acc, c) => acc + c.idade, 0)) / ((portData.length + matData.length) || 1)).toFixed(1);

    const r1 = ws.addRow(['Total de Matrículas', portData.length, matData.length, portData.length + matData.length]);
    const r2 = ws.addRow(['Média de Idade', avgPort, avgMat, avgTotal]);
    [r1, r2].forEach(r => { r.alignment = { horizontal: 'center' }; curRow++; });
    
    curRow += 2;

    const titles = [
        { label: 'VISÃO GERAL: PORTUGUÊS', col: 1, color: 'FF004A8D' },
        { label: 'VISÃO GERAL: MATEMÁTICA', col: 6, color: 'FFF7941D' }
    ];

    titles.forEach(t => {
        const cell = ws.getCell(curRow, t.col);
        ws.mergeCells(curRow, t.col, curRow, t.col + 3);
        cell.value = t.label;
        cell.font = { bold: true, color: { argb: t.color } };
        cell.alignment = { horizontal: 'center' };
    });

    curRow++;
    ws.addImage(wb.addImage({ base64: imgs.piePort, extension: 'png' }), { tl: { col: 0, row: curRow - 1 }, ext: { width: 500, height: 380 } });
    ws.addImage(wb.addImage({ base64: imgs.pieMat, extension: 'png' }), { tl: { col: 5, row: curRow - 1 }, ext: { width: 500, height: 380 } });

    curRow += 21;

    const barTitles = [
        { label: 'STATUS POR UNIDADE: PORTUGUÊS', col: 1, color: 'FF004A8D' },
        { label: 'STATUS POR UNIDADE: MATEMÁTICA', col: 6, color: 'FFF7941D' }
    ];

    barTitles.forEach(t => {
        const cell = ws.getCell(curRow, t.col);
        ws.mergeCells(curRow, t.col, curRow, t.col + 3);
        cell.value = t.label;
        cell.font = { bold: true, color: { argb: t.color } };
        cell.alignment = { horizontal: 'center' };
    });

    curRow++;
    ws.addImage(wb.addImage({ base64: imgs.statusPort, extension: 'png' }), { tl: { col: 0, row: curRow - 1 }, ext: { width: 500, height: 350 } });
    ws.addImage(wb.addImage({ base64: imgs.statusMat, extension: 'png' }), { tl: { col: 5, row: curRow - 1 }, ext: { width: 500, height: 350 } });
}

