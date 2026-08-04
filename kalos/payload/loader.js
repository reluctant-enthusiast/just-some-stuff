(async()=>{
  const status=document.getElementById("status");
  const error=document.getElementById("error");
  const track=document.querySelector(".track");
  try{
    if (!("DecompressionStream" in globalThis)) throw new Error("DecompressionStream unavailable");
    const b64=globalThis.__KALOS_PAYLOAD||"";
    if(!b64) throw new Error("Narrative payload missing");
    const raw=atob(b64);
    const bytes=new Uint8Array(raw.length);
    for(let i=0;i<raw.length;i++) bytes[i]=raw.charCodeAt(i);
    const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
    const html=await new Response(stream).text();
    globalThis.__KALOS_PAYLOAD="";
    document.open(); document.write(html); document.close();
  }catch(err){
    console.error(err);
    if(track) track.style.display="none";
    if(status) status.textContent="The chronicle could not open.";
    if(error) error.style.display="block";
  }
})();
