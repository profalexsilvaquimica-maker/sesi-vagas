/* Exportacao para Excel, sem modificar os registros na nuvem. */
(function () {
  'use strict';
  let carregando = null;
  function bibliotecaExcel() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!carregando) {
      carregando = new Promise(function (resolve, reject) {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
        script.onload = () => window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('ExcelJS nao inicializou.'));
        script.onerror = () => reject(new Error('Nao foi possivel carregar a biblioteca de Excel. Verifique sua internet.'));
        document.head.appendChild(script);
      }).catch(err => { carregando = null; throw err; });
    }
    return carregando;
  }
  const azul = 'FF003366', azulClaro = 'FFEEF4FF', verdeClaro = 'FFE6F6EF';
  function configurarTabela(ws, colunas) {
    ws.columns = colunas.map(c => ({ header: c[0], key: c[1], width: c[2] }));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: colunas.length } };
    ws.getRow(1).height = 26;
    ws.getRow(1).eachCell(c => {
      c.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: azul } };
      c.alignment = { vertical: 'middle', wrapText: true };
    });
  }
  function linhaTotal(ws, cols, ultimo) {
    const r = ws.addRow([]);
    r.getCell(1).value = 'TOTAL';
    cols.forEach(col => {
      r.getCell(col).value = { formula: 'SUM(' + ws.getColumn(col).letter + '2:' + ws.getColumn(col).letter + ultimo + ')' };
      r.getCell(col).numFmt = '#,##0';
    });
    r.eachCell({ includeEmpty: true }, cell => {
      cell.font = { bold: true, color: { argb: azul } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: azulClaro } };
    });
  }
  function finalizarTabela(ws) {
    ws.eachRow({ includeEmpty: false }, (r, n) => {
      if (n !== 1) {
        r.height = 20;
        r.eachCell(c => { c.alignment = { vertical: 'middle' }; });
        if (n % 2 === 0) r.eachCell(c => {
          if (!c.fill || c.fill.pattern !== 'solid') c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF6F9FD' } };
        });
      }
    });
  }
  function dataExcel(valor) {
    if (!valor) return null;
    const match = String(valor).match(/^(\d{4})-(\d{2})-(\d{2})/);
    return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12) : String(valor);
  }
  window.exportarVagasExcel = async function (dados) {
    const ExcelJS = await bibliotecaExcel();
    const turmas = [...dados.turmas].sort((a, b) => a.nome.localeCompare(b.nome, 'pt-BR', { numeric: true }));
    const ressalvas = [...(dados.ressalvas || [])];
    const wb = new ExcelJS.Workbook();
    wb.creator = 'SESI Ananindeua — Controle de Vagas';
    wb.created = new Date();

    const ws = wb.addWorksheet('Turmas 2027');
    configurarTabela(ws, [
      ['Turma', 'nome', 36], ['Série', 'serie', 18], ['Segmento', 'segmento', 22],
      ['Turno', 'turno', 16], ['Alunos 2026', 'base', 16], ['Vagas 2027', 'vagas', 16],
      ['Matriculados 2027', 'matriculados', 21], ['Disponíveis', 'disponiveis', 17],
      ['Ocupação', 'ocupacao', 15], ['Variação 2026–27', 'variacao', 20]
    ]);
    turmas.forEach((t, i) => {
      const n = i + 2, base = Number(t.alunos2026) || 0, vagas = Number(t.vagas2027) || 0, mat = Number(t.matriculados2027) || 0;
      const row = ws.addRow([t.nome, t.serie || '', t.segmento, t.turno, base, vagas, mat,
        { formula: 'F' + n + '-G' + n, result: vagas - mat },
        { formula: 'IF(F' + n + '=0,0,G' + n + '/F' + n + ')', result: vagas ? mat / vagas : 0 },
        { formula: 'F' + n + '-E' + n, result: vagas - base }
      ]);
      row.getCell(9).numFmt = '0.0%';
      [5, 6, 7, 8, 10].forEach(col => row.getCell(col).numFmt = '#,##0');
      if (vagas - mat < 0) row.getCell(8).font = { bold: true, color: { argb: 'FFB91C1C' } };
    });
    if (turmas.length) linhaTotal(ws, [5, 6, 7, 8, 10], turmas.length + 1);
    finalizarTabela(ws);

    const base = wb.addWorksheet('Base 2026');
    configurarTabela(base, [
      ['Turma', 'nome', 36], ['Série', 'serie', 18], ['Segmento', 'segmento', 22],
      ['Turno', 'turno', 16], ['Alunos 2026', 'alunos', 18]
    ]);
    turmas.forEach(t => base.addRow([t.nome, t.serie || '', t.segmento, t.turno, Number(t.alunos2026) || 0]));
    if (turmas.length) linhaTotal(base, [5], turmas.length + 1);
    finalizarTabela(base);

    const cmp = wb.addWorksheet('Comparativo');
    configurarTabela(cmp, [
      ['Turma', 'nome', 36], ['Segmento', 'segmento', 22], ['Turno', 'turno', 16],
      ['Alunos 2026', 'base', 17], ['Vagas 2027', 'vagas', 17],
      ['Diferença de vagas', 'diferenca', 21], ['Matriculados 2027', 'mat', 21],
      ['Disponíveis 2027', 'livres', 20]
    ]);
    turmas.forEach((t, i) => {
      const n = i + 2, v = Number(t.vagas2027) || 0, a = Number(t.alunos2026) || 0, m = Number(t.matriculados2027) || 0;
      cmp.addRow([t.nome, t.segmento, t.turno, a, v,
        { formula: 'E' + n + '-D' + n, result: v - a }, m,
        { formula: 'E' + n + '-G' + n, result: v - m }]);
    });
    if (turmas.length) linhaTotal(cmp, [4, 5, 6, 7, 8], turmas.length + 1);
    finalizarTabela(cmp);

    const res = wb.addWorksheet('Ressalvas');
    configurarTabela(res, [
      ['Nome do aluno', 'aluno', 36], ['Turma', 'turma', 36], ['Série', 'serie', 18],
      ['Documento entregue', 'doc', 22], ['Data da solicitação', 'data', 23]
    ]);
    ressalvas.forEach(r => {
      const row = res.addRow([r.nome_aluno || '', r.turma_nome || r.turma_id || '',
        r.serie || '', r.documento_entregue === true ? 'SIM' : r.documento_entregue === false ? 'NÃO' : r.documento_entregue || '',
        dataExcel(r.data_solicitacao)]);
      if (row.getCell(5).value instanceof Date) row.getCell(5).numFmt = 'dd/mm/yyyy';
    });
    finalizarTabela(res);

    const resumo = wb.addWorksheet('Resumo Geral', { views: [{ state: 'frozen', ySplit: 2 }] });
    resumo.columns = [{ width: 33 }, { width: 22 }];
    resumo.mergeCells('A1:B1');
    resumo.getCell('A1').value = 'SESI ANANINDEUA | CONTROLE DE VAGAS 2027';
    resumo.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14 };
    resumo.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: azul } };
    resumo.getRow(1).height = 31;
    const n = turmas.length + 1;
    const totalVagas = turmas.reduce((s,t) => s + (Number(t.vagas2027) || 0), 0);
    const totalMat = turmas.reduce((s,t) => s + (Number(t.matriculados2027) || 0), 0);
    const totalBase = turmas.reduce((s,t) => s + (Number(t.alunos2026) || 0), 0);
    const rows = [
      ['Total de turmas', turmas.length],
      ['Total de vagas 2027', { formula: "'Turmas 2027'!F" + (n + 1), result: totalVagas }],
      ['Matriculados 2027', { formula: "'Turmas 2027'!G" + (n + 1), result: totalMat }],
      ['Vagas disponíveis', { formula: 'B3-B4', result: totalVagas - totalMat }],
      ['Ocupação', { formula: 'IF(B3=0,0,B4/B3)', result: totalVagas ? totalMat / totalVagas : 0 }],
      ['Base 2026', totalBase],
      ['Meta configurada (' + dados.metaTipo + ')', Number(dados.metaValor) || 0],
      ['Meta atingida', dados.metaTipo === 'alunos' ? (Number(dados.metaValor) ? totalMat / Number(dados.metaValor) : 0) : (totalVagas ? totalMat / totalVagas : 0)],
      ['Total de ressalvas', ressalvas.length],
      ['Exportado em', new Date().toLocaleString('pt-BR')]
    ];
    rows.forEach(row => resumo.addRow(row));
    [6, 9].forEach(r => resumo.getCell('B' + r).numFmt = '0.0%');
    resumo.getColumn(2).alignment = { horizontal: 'right' };
    resumo.getRow(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: verdeClaro } };
    resumo.getRow(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: verdeClaro } };
    resumo.eachRow((r, i) => { if (i > 1) { r.height = 23; r.getCell(1).font = { bold: true, color: { argb: azul } }; } });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'SESI_Controle_de_Vagas_2027_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  };
})();
