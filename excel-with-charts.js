/**
 * Alternativa usando SheetJS + Chart.js para gerar imagens de gráficos
 * Requer instalação: npm install xlsx canvas chart.js
 */

// Função alternativa para criar Excel com gráficos como imagens
async function exportToExcelWithCharts() {
    const uf = document.getElementById('filterUF').value;
    const filter = d => uf === "ALL" || d.uf === uf;
    
    const fPort = database.port.filter(filter);
    const fMat = database.mat.filter(filter);

    // Criar workbook usando SheetJS
    const wb = XLSX.utils.book_new();
    
    // Criar planilhas de dados
    createSheetJSTable(wb, 'Português', fPort);
    createSheetJSTable(wb, 'Matemática', fMat);
    createSheetJSDashboard(wb, fPort, fMat);
    
    // Gerar e baixar
    XLSX.writeFile(wb, `PNGP_Analytics_${uf === 'ALL' ? 'Brasil' : uf}_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function createSheetJSTable(wb, sheetName, data) {
    const units = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    
    const wsData = [
        ['Nome', 'E-mail', 'UF', 'Idade', ...units.flatMap(u => [u, `${u} - Data`]), 'Status Geral']
    ];
    
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
        wsData.push(rowData);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

function createSheetJSDashboard(wb, portData, matData) {
    const uts = ["Teste inicial", "Unidade 1", "Unidade 2", "Unidade 3", "Unidade 4", "Unidade 5"];
    
    const wsData = [
        ['INDICADORES PRINCIPAIS'],
        [],
        ['Métrica', 'Português', 'Matemática', 'Total'],
        ['Total de Alunos', portData.length, matData.length, portData.length + matData.length],
        [],
        ['ALUNOS POR UNIDADE DE TRABALHO'],
        [],
        ['Unidade', 'Português', 'Matemática', 'Total']
    ];
    
    uts.forEach(ut => {
        const countPort = countUT(portData, ut);
        const countMat = countUT(matData, ut);
        wsData.push([ut, countPort, countMat, countPort + countMat]);
    });
    
    wsData.push([]);
    wsData.push(['STATUS GERAL - PORTUGUÊS']);
    wsData.push([]);
    
    const statsPort = getStatsForExcel(portData, uts);
    wsData.push(['Status', 'Quantidade', 'Percentual']);
    wsData.push(['Concluído', statsPort['Concluído'], `${((statsPort['Concluído'] / (statsPort.total || 1)) * 100).toFixed(1)}%`]);
    wsData.push(['Não Concluído', statsPort['Não concluído'], `${((statsPort['Não concluído'] / (statsPort.total || 1)) * 100).toFixed(1)}%`]);
    wsData.push(['Não Iniciado', statsPort['Não iniciado'], `${((statsPort['Não iniciado'] / (statsPort.total || 1)) * 100).toFixed(1)}%`]);
    
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'Dashboard');
}