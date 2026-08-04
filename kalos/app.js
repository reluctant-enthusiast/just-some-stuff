(() => {
  'use strict';
  const root = document.querySelector('#app');
  const parts = [
    './app.part.00a', './app.part.00b', './app.part.01', './app.part.02', './app.part.03',
    './app.part.04a', './app.part.04b', './app.part.04c',
    './app.part.05a', './app.part.05b', './app.part.05c'
  ];
  const fail = (error) => {
    console.error(error);
    const panel = document.querySelector('#error-panel');
    const message = document.querySelector('#error-message');
    if (panel && message) {
      panel.hidden = false;
      message.textContent = error instanceof Error ? error.message : String(error);
    } else if (root) {
      root.innerHTML = '<main style="max-width:760px;margin:40px auto;padding:24px;color:#f2ede2;font-family:Georgia,serif"><h1>The chronicle could not open</h1><p>' + String(error instanceof Error ? error.message : error) + '</p><p><a href="./legacy-v1.html" style="color:#d9b77e">Open the legacy prototype</a></p></main>';
    }
  };
  (async () => {
    try {
      if (typeof DecompressionStream !== 'function') throw new Error('This browser does not support the compressed narrative build.');
      const responses = await Promise.all(parts.map(url => fetch(url, { cache: 'no-store' })));
      const failed = responses.find(response => !response.ok);
      if (failed) throw new Error('Could not load the narrative engine (' + failed.status + ').');
      const encoded = (await Promise.all(responses.map(response => response.text()))).join('').replace(/\s+/g, '');
      const binary = atob(encoded);
      const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));
      const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
      const source = await new Response(stream).text();
      const script = document.createElement('script');
      script.textContent = source;
      document.head.append(script);
    } catch (error) {
      fail(error);
    }
  })();
})();
