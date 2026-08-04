(() => {
  'use strict';
  const root = document.querySelector('#app');
  const parts = [
    './app.part.00a',
    './app.part.00b',
    './app.part.01',
    './app.part.02',
    './app.part.03',
    './app.part.04',
    './app.part.05'
  ];

  const showFailure = (error) => {
    console.error(error);
    if (!root) return;
    root.innerHTML = `
      <main class="fatal-panel" role="alert">
        <h1>The chronicle could not open</h1>
        <p>${error instanceof Error ? error.message : String(error)}</p>
        <p>This hosted build needs a current Safari, Chrome, or Firefox browser.</p>
        <button type="button" onclick="location.reload()">Try again</button>
      </main>`;
  };

  (async () => {
    try {
      if (typeof DecompressionStream !== 'function') {
        throw new Error('This browser does not provide the required gzip decompression support.');
      }
      const responses = await Promise.all(parts.map((url) => fetch(url, { cache: 'no-store' })));
      const failed = responses.find((response) => !response.ok);
      if (failed) throw new Error(`Could not load the narrative engine (${failed.status}).`);
      const encoded = (await Promise.all(responses.map((response) => response.text())))
        .join('')
        .replace(/\s+/g, '');
      const binary = atob(encoded);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
      const decompressed = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
      const source = await new Response(decompressed).text();
      const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
      const script = document.createElement('script');
      script.src = url;
      script.onload = () => URL.revokeObjectURL(url);
      script.onerror = () => {
        URL.revokeObjectURL(url);
        showFailure(new Error('The narrative engine could not start.'));
      };
      document.head.append(script);
    } catch (error) {
      showFailure(error);
    }
  })();
})();
