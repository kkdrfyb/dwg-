importScripts('dxf-parser.min.js');

self.onmessage = function(e){
  const { id, op, text, pairs } = e.data || {};
  try{
    const parser = new DxfParser();
    const dxf = parser.parseSync(text);
    const entities = (dxf.entities || []).map(ent => {
      let content = '';
      if (ent.type === 'TEXT' || ent.type === 'MTEXT' || ent.type === 'ATTRIB') content = ent.text || '';
      else if (ent.type === 'INSERT') content = ent.name || '';
      return { type: ent.type, layer: ent.layer || '', text: content };
    }).filter(e => e.text);
    postMessage({ id, ok: true, entities });
  }catch(err){
    postMessage({ id, ok: false, error: String(err) });
  }
}
