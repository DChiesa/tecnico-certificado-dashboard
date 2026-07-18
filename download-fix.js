/* Correcao do carregamento Supabase - Tecnico Certificado */
(function () {
  'use strict';

  const pausa = (ms = 80) => new Promise(resolve => setTimeout(resolve, ms));

  function status(tipo, texto) {
    if (typeof setStatus === 'function') setStatus(tipo, texto);
    else {
      const box = document.getElementById('autoStatus');
      const label = document.getElementById('autoStatusText');
      if (box) box.className = 'auto-status ' + tipo;
      if (label) label.textContent = texto;
    }
  }

  async function fetchComTimeout(url, timeoutMs = 60000) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      return await fetch(url, {
        method: 'GET',
        cache: 'no-store',
        redirect: 'follow',
        signal: controller.signal
      });
    } finally {
      clearTimeout(timer);
    }
  }

  window.download = async function downloadCorrigido(path, label) {
    const bucket = window.APP_CONFIG?.BUCKET || 'bases-tecnico';
    const caminho = path || window.APP_CONFIG?.ARQUIVO_ATUAL || 'atual/tecnico_certificado_atual.xlsx';

    try {
      status('loading', '1/5 Autorizando acesso ao arquivo...');
      await pausa();

      const { data: signedData, error: signedError } = await supa()
        .storage
        .from(bucket)
        .createSignedUrl(caminho, 120);

      if (signedError) throw signedError;
      if (!signedData?.signedUrl) throw new Error('O Supabase não retornou a URL temporária do arquivo.');

      status('loading', '2/5 Baixando o Excel...');
      await pausa();

      const response = await fetchComTimeout(signedData.signedUrl, 60000);
      if (!response.ok) throw new Error(`Falha HTTP ${response.status} ao baixar o Excel.`);

      const contentType = response.headers.get('content-type') || '';
      const buffer = await response.arrayBuffer();
      if (!buffer.byteLength) throw new Error('O arquivo baixado está vazio.');

      const bytes = new Uint8Array(buffer);
      if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
        throw new Error(`O conteúdo recebido não é um XLSX válido (${contentType || 'tipo desconhecido'}).`);
      }

      status('loading', `3/5 Excel recebido (${(buffer.byteLength / 1024).toFixed(1)} KB). Lendo planilha...`);
      await pausa(120);

      const workbook = XLSX.read(buffer, {
        type: 'array',
        cellDates: true,
        dense: true
      });

      if (!workbook.SheetNames?.length) throw new Error('O Excel não possui planilhas legíveis.');

      status('loading', '4/5 Processando indicadores...');
      await pausa(120);

      const total = processWorkbook(workbook, label || `${caminho} · Supabase`);

      status('ok', `5/5 Concluído: ${Number(total || 0).toLocaleString('pt-BR')} registros · ${new Date().toLocaleString('pt-BR')}`);
      return total;
    } catch (error) {
      console.error('Falha no carregamento corrigido:', error);
      const mensagem = error?.name === 'AbortError'
        ? 'Tempo excedido ao baixar o Excel. Verifique a conexão e tente novamente.'
        : (error?.message || String(error));
      status('error', 'Falha: ' + mensagem);
      throw error;
    }
  };

  console.info('Correção de download Supabase carregada.');
})();
