/* Sincronizacao persistente das turmas. Nao altera as tabelas existentes. */
(function () {
  'use strict';
  const STORAGE_KEY = 'sesi_turmas_fila_v1';
  window.criarSincronizadorVagas = function (client, obterTurma, avisar) {
    let fila = {};
    try { fila = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') || {}; }
    catch (_) { fila = {}; }
    const pendingIds = new Set(Object.keys(fila));
    let enviando = false;
    let timer = null;
    let retryTimer = null;

    function persistir() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(fila));
      pendingIds.clear();
      Object.keys(fila).forEach(id => pendingIds.add(id));
    }
    function mapear(t) {
      return {
        id: t.id, nome: t.nome, turno: t.turno, segmento: t.segmento,
        serie: t.serie || '', vagas_2026: Number(t.alunos2026) || 0,
        vagas_2027: Number(t.vagas2027) || 0,
        matriculados: Number(t.matriculados2027) || 0
      };
    }
    function pendente(id) { return pendingIds.has(String(id)); }
    function avisarPendente() {
      avisar('⏳ Salvando automaticamente...', pendingIds.size + ' turma(s) aguardando confirmacao do Supabase.',
        'bg-yellow-100 text-yellow-800');
    }
    function agendar(ms) {
      clearTimeout(timer);
      timer = setTimeout(processar, ms == null ? 350 : ms);
    }
    function colocar(id, item) {
      fila[String(id)] = item;
      persistir();
      avisarPendente();
      agendar();
    }
    function matricula(t) {
      const anterior = fila[String(t.id)];
      if (anterior && anterior.tipo === 'upsert') {
        colocar(t.id, { tipo: 'upsert', dados: mapear(t) });
      } else {
        colocar(t.id, { tipo: 'matricula', dados: { matriculados: Number(t.matriculados2027) || 0 } });
      }
    }
    function editar(t) { colocar(t.id, { tipo: 'upsert', dados: mapear(t) }); }
    function excluir(id) { colocar(id, { tipo: 'delete' }); }

    async function processar() {
      clearTimeout(timer);
      if (enviando || pendingIds.size === 0) return;
      if (!navigator.onLine) {
        avisar('⚠️ Sem internet — alteracoes preservadas',
          pendingIds.size + ' alteracao(oes) guardada(s) neste navegador. Serao enviadas ao reconectar.',
          'bg-yellow-100 text-yellow-800');
        return;
      }
      enviando = true;
      let erros = 0, ultimoErro = '';
      try {
        const itens = Object.entries(fila);
        for (const [id, item] of itens) {
          if (!fila[id]) continue;
          const snapshot = JSON.stringify(item);
          try {
            if (item.tipo === 'delete') {
              const { error } = await client.from('turmas').delete().eq('id', id);
              if (error) throw error;
            } else if (item.tipo === 'upsert') {
              const { data, error } = await client.from('turmas')
                .upsert(item.dados, { onConflict: 'id' }).select('id');
              if (error) throw error;
              if (!data || data.length === 0) throw new Error('Supabase nao confirmou a gravacao da turma.');
            } else {
              const { data, error } = await client.from('turmas')
                .update(item.dados).eq('id', id).select('id');
              if (error) throw error;
              if (!data || data.length === 0) {
                const turma = obterTurma(id);
                if (!turma) throw new Error('Turma nao encontrada para sincronizar.');
                const insert = await client.from('turmas')
                  .upsert(mapear(turma), { onConflict: 'id' }).select('id');
                if (insert.error) throw insert.error;
                if (!insert.data || insert.data.length === 0) {
                  throw new Error('Supabase nao confirmou a inclusao da turma.');
                }
              }
            }
            // Se houve outro clique durante o envio, mantenha a alteracao mais recente na fila.
            if (JSON.stringify(fila[id]) === snapshot) {
              delete fila[id];
              persistir();
            }
          } catch (err) {
            erros++;
            ultimoErro = (err && err.message) ? err.message : String(err);
          }
        }
      } finally {
        enviando = false;
      }
      if (erros) {
        avisar('❌ Falha ao sincronizar — dados locais preservados',
          pendingIds.size + ' alteracao(oes) pendente(s). Nova tentativa automatica.',
          'bg-red-100 text-red-800', ultimoErro);
        clearTimeout(retryTimer);
        retryTimer = setTimeout(processar, 15000);
      } else if (pendingIds.size) {
        avisarPendente();
        agendar(250);
      } else {
        clearTimeout(retryTimer);
        avisar('☁️ Salvo automaticamente na nuvem',
          'Todas as alteracoes das turmas foram confirmadas pelo Supabase.',
          'bg-green-100 text-green-800');
      }
    }
    window.addEventListener('online', () => { if (pendingIds.size) agendar(150); });
    window.addEventListener('beforeunload', persistir);
    if (pendingIds.size) agendar(1000);
    return { pendingIds, pendente, matricula, editar, excluir, processar, mapear };
  };
})();
