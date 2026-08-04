(async()=>{try{
const parts=await Promise.all(['./bundle-0.txt','./bundle-1.txt','./bundle-2.txt','./bundle-3.txt','./bundle-4.txt'].map(async url=>{const r=await fetch(url);if(!r.ok)throw new Error(`Could not load ${url}`);return r.text();}));
if(!('DecompressionStream' in window))throw new Error('This browser is too old for the compressed story engine.');
const raw=atob(parts.join('').trim());
const bytes=Uint8Array.from(raw,c=>c.charCodeAt(0));
const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
const code=await new Response(stream).text();
(0,eval)(code);window.dispatchEvent(new Event('DOMContentLoaded'));
}catch(error){const root=document.getElementById('app');root.innerHTML='<main class="fatal"><h1>The story could not wake.</h1><p>'+String(error.message||error)+'</p><p><a href="./legacy-v1.html" style="color:#d9b77e">Open the legacy prototype</a></p></main>';}})();