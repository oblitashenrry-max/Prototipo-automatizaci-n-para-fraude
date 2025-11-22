/********* CONFIG *********/
const SPREADSHEET_ID = '1be6jjOD8rMWhT6t5bybvZO1MKBo_53doIy3X3_g9sV8';
const BASE_SHEET_GID   = 0;           // Base de pólizas
const RESP_SHEET_GID   = 1967832507;  // Registros Siniestros
const DRIVE_ROOT_FOLDER_ID = '1Ih3TNtfchbqBtVa1YtEp2aYMUG5SGxP5';
const INTERNAL_RECIPIENTS = [];       // opcional
const DUP_WINDOW_DAYS = 30;           // ventana para duplicados

/********* VISUAL / EMOJIS *********/
const E = {
  NOK:'🔴',
  WARN:'⚠️',
  OK:'✅',
  IMG:'📎',
  LIC:'💳',
  PRES:'🧾',
  CAR:'🚘',
  CLOCK:'⏱️',
  COPY:'🗂️',
  MONEY:'💰'
};

/********* UTILIDADES *********/
function norm(s){
  return String(s||'')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .trim()
    .toLowerCase();
}
function tz_(){
  return Session.getScriptTimeZone();
}
function parseDateSmart_(v){
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : v;
  if (typeof v === 'number'){
    const d = new Date(v);
    return isNaN(d) ? null : d;
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m){
    const d = new Date(+m[1], +m[2]-1, +m[3]);
    return isNaN(d) ? null : d;
  }
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m){
    const d = new Date(+m[3], +m[2]-1, +m[1]);
    return isNaN(d) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function isValidEmail_(s){
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s||''));
}
function isValidPhone_(s){
  const d = String(s||'').replace(/\D/g,'');
  return d.length>=9 && d.length<=12;
}
function normalizePlate_(s){
  return String(s||'')
    .toUpperCase()
    .replace(/\s+/g,'')
    .replace(/O/g,'0')
    .replace(/[—–]/g,'-')
    .replace(/[^A-Z0-9-]/g,'');
}
function plausiblePlate_(s){
  return /^[A-Z0-9]{3}-[A-Z0-9]{3}$/.test(normalizePlate_(s));
}
function extractBestPlate_(text){
  const t = normalizePlate_(text);
  const r = /[A-Z0-9]{3}-[A-Z0-9]{3}/g;
  const f = t.match(r)||[];
  return f[0]||'';
}
function levenshtein_(a,b){
  a=String(a); b=String(b);
  const m=[];
  let i,j;
  for(i=0;i<=b.length;i++){ m[i]=[i]; }
  for(j=0;j<=a.length;j++){ m[0][j]=j; }
  for(i=1;i<=b.length;i++){
    for(j=1;j<=a.length;j++){
      m[i][j]=Math.min(
        m[i-1][j]+1,
        m[i][j-1]+1,
        m[i-1][j-1]+(a[j-1]===b[i-1]?0:1)
      );
    }
  }
  return m[b.length][a.length];
}

/********* PARSER MONTOS *********/
function moneyParse_(s){
  const txt=String(s||'');
  const re=/((S\/|USD|US\$|\$)\s*[\d\.,]+)/ig;
  let m,best=0,mon='';
  while((m=re.exec(txt))!==null){
    let num=m[0]
      .replace(/[^\d.,]/g,'')
      .replace(/\.(?=\d{3}\b)/g,'')
      .replace(',', '.');
    const v=parseFloat(num);
    if(!isNaN(v)&&v>best){
      best=v;
      mon=(m[2]||'$').replace(/\s/g,'');
    }
  }
  return {amount:best,currency:mon||'$'};
}
function parseBudgetTotal_(txt){
  const lines = String(txt||'').split(/\r?\n/)
    .map(l=>l.trim())
    .filter(l=>l);
  let bestAmount = 0;
  let bestCurrency = '$';
  for (let i = 0; i < lines.length; i++){
    const line = lines[i];
    const normLine = line.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase();
    if (!/\bTOTAL\b/.test(normLine)) continue;

    const curMatch = line.match(/(S\/|USD|US\$|\$)/i);
    if (curMatch){
      bestCurrency = curMatch[1].toUpperCase().replace(/\s/g,'');
    }

    const numMatches = Array.from(line.matchAll(/(-?\d[\d.,]*)/g));
    if (!numMatches.length) continue;
    let rawNum = numMatches[numMatches.length-1][1];
    let cleaned = rawNum.replace(/[^0-9,.\-]/g,'');

    if (cleaned.indexOf(',') !== -1 && cleaned.indexOf('.') !== -1){
      cleaned = cleaned
        .replace(/\.(?=\d{3}(\D|$))/g,'')
        .replace(',', '.');
    } else if (cleaned.indexOf(',') !== -1){
      if (cleaned.length - cleaned.lastIndexOf(',') <= 3){
        cleaned = cleaned.replace(',', '.');
      } else {
        cleaned = cleaned.replace(/,/g,'');
      }
    } else {
      if (cleaned.length - cleaned.lastIndexOf('.') > 3){
        cleaned = cleaned.replace(/\./g,'');
      }
    }

    const val = parseFloat(cleaned);
    if (!isNaN(val) && val > bestAmount){
      bestAmount = val;
    }
  }
  return { amount: bestAmount, currency: bestCurrency || '$' };
}

/********* CONFIG DINÁMICA: SOLO THRESHOLDS *********/
function getConfig_(){
  const out = { thresholds: { ROJO:70, AMBAR:40 } };
  try{
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh=ss.getSheetByName('Config');
    if (!sh) return out;
    const values=sh.getDataRange().getValues();
    if(values.length<2) return out;
    const hdr = values[0].map(x=>String(x||'').trim().toLowerCase());
    const iReg=hdr.indexOf('regla'), iP=hdr.indexOf('peso');
    for (let r=1; r<values.length; r++){
      const row=values[r];
      const regla=String(row[iReg]||'').trim();
      const peso=parseFloat(row[iP]);
      if (!regla || isNaN(peso)) continue;
      if (regla==='THRESHOLD_ROJO') out.thresholds.ROJO=peso;
      if (regla==='THRESHOLD_AMBAR') out.thresholds.AMBAR=peso;
    }
  }catch(e){}
  return out;
}

/********* ORDEN DE COLUMNAS (LAYOUT CORTO) *********/
const REQUIRED_HEADERS = [
  'Poliza','Marca','Modelo','Patente','Revisado',
  'Nomb','Apell','TipoDoc','NroDoc','Email','Fono','ValBase',
  'TipoSin','FecOcc','HrOcc','Relato','SevRel','UrlRel',
  'Parte','Ter','TagsRel',
  'Adj','UrlFolder',
  'OCRsrc','PatOCR','A_PatOCR',
  'MontoPres','A_Pres',
  'Prior','Asig','Next',
  'A_Lic',
  'JSON','DupRef','Resumen',
  'A_Vig','A_Fec','A_Contacto','A_PatForm','A_Adj','Adj_Falt','A_Ter','A_Parte','A_Dup',
  'Score','Semaf'
];

/********* HOJAS *********/
function getSheetByGid_(ssid,gid){
  const ss=SpreadsheetApp.openById(ssid);
  const sh=ss.getSheets().find(s=>s.getSheetId()===gid);
  if(!sh) throw new Error('No encuentro la hoja con GID: '+gid);
  return sh;
}
function ensureHeaders_(sheet){
  const lastCol=sheet.getLastColumn();
  const have=sheet.getLastRow()>=1 && lastCol>=1;
  let headers=have?sheet.getRange(1,1,1,lastCol).getValues()[0]:[];
  if(!headers.length || (headers.length===1 && String(headers[0]).trim()==='')){
    sheet.getRange(1,1,1,REQUIRED_HEADERS.length).setValues([REQUIRED_HEADERS]);
    sheet.setFrozenRows(1);
    return REQUIRED_HEADERS.slice();
  }
  const existing=headers.map(h=>String(h||'').trim());
  const toAdd=REQUIRED_HEADERS.filter(h=>!existing.includes(h));
  if(toAdd.length){
    sheet.getRange(1,existing.length+1,1,toAdd.length).setValues([toAdd]);
    headers=existing.concat(toAdd);
  }
  return headers;
}
function getSheetByGidIdSafe_(){
  return getSheetByGid_(SPREADSHEET_ID, RESP_SHEET_GID);
}

/********* BASE POLIZAS *********/
const COL_ALIASES = {
  Nombre_asegurado:['nombre_asegurado','nombre','nombres'],
  Apellido_asegurado:['apellido_asegurado','apellidos','apellido'],
  Poliza:['poliza','nro_poliza','numero_poliza','póliza','n° póliza'],
  Fecha_inicio:['fecha_inicio','inicio','vigencia_desde','desde'],
  Fecha_fin:['fecha_fin','fin','vigencia_hasta','hasta'],
  Estado:['estado','estatus'],
  Tipo_documento:['tipo_documento','tipo doc','tipo'],
  Numero_documento:['numero_documento','dni','documento','nro_documento'],
  Celular:['celular','telefono','teléfono'],
  email_asegurado:['email_asegurado','correo','email','correo_electronico'],
  Marca:['marca'],
  Modelo:['modelo'],
  Patente:['patente','placa','matricula']
};
function getBaseData_(){
  const sh=getSheetByGid_(SPREADSHEET_ID,BASE_SHEET_GID);
  const values=sh.getDataRange().getValues();
  if(values.length<2) throw new Error('La base de pólizas está vacía.');
  const headers=values.shift().map(h=>norm(h));
  const idx={};
  Object.keys(COL_ALIASES).forEach(k=>{
    const cands=COL_ALIASES[k].map(norm);
    idx[k]=headers.findIndex(h=>cands.includes(h));
  });
  return {idx:idx,rows:values};
}
function toObject_(row,idx){
  const o={};
  Object.keys(idx).forEach(k=>{
    o[k]=(idx[k]>=0?row[idx[k]]:'');
  });
  return o;
}
function findRecord(payload){
  const data=getBaseData_(); const idx=data.idx; const rows=data.rows; const s=v=>norm(v);
  return rows.find(r=>{
    return (payload.poliza && idx.Poliza>=0 && s(r[idx.Poliza])===s(payload.poliza)) ||
           (payload.numeroDoc && idx.Numero_documento>=0 && s(r[idx.Numero_documento])===s(payload.numeroDoc)) ||
           (payload.patente && idx.Patente>=0 && s(r[idx.Patente])===s(payload.patente));
  });
}
function validateAgainstBase(payload){
  const base=getBaseData_(); const idx=base.idx;
  const match=findRecord(payload);
  if(!match){
    return {
      found:false,
      msg:'No se encuentra en la base.',
      record:null,
      softWarnings:[],
      reasons:['NO_ENCONTRADO'],
      baseStatus:'NO_ENCONTRADO'
    };
  }
  const rec=toObject_(match,idx); const s=v=>norm(v);
  const softWarnings=[];
  if(payload.marca && s(payload.marca)!==s(rec.Marca)) softWarnings.push('Marca distinta a la base.');
  if(payload.modelo && s(payload.modelo)!==s(rec.Modelo)) softWarnings.push('Modelo distinto a la base.');
  if(payload.patente && s(payload.patente)!==s(rec.Patente)) softWarnings.push('Patente distinta a la base.');
  if(payload.nombre && s(payload.nombre)!==s(rec.Nombre_asegurado)) softWarnings.push('Nombre distinto.');
  if(payload.apellido && s(payload.apellido)!==s(rec.Apellido_asegurado)) softWarnings.push('Apellido distinto.');

  const reasons=[];
  const fi=parseDateSmart_(rec.Fecha_inicio);
  const ff=parseDateSmart_(rec.Fecha_fin);
  const now=new Date();
  let vig='VIGENTE';
  const fechasInvalidas=!fi||!ff;
  const fueraR=!fechasInvalidas && !(now>=fi && now<=ff);
  const estado=String(rec.Estado||'').trim().toUpperCase();
  const estadoNoActivo=(estado && !/ACTIV[OA]|VIGENT/.test(estado));

  if(fechasInvalidas||fueraR||estadoNoActivo){
    vig='NO_VIGENTE';
    if(fechasInvalidas) reasons.push('Fechas inválidas en base.');
    if(!fechasInvalidas&&fueraR) reasons.push('Fuera de rango.');
    if(estadoNoActivo) reasons.push('Estado en base: '+estado+'.');
  }

  return {
    found:true,
    record:{
      Nombre_asegurado:rec.Nombre_asegurado,
      Apellido_asegurado:rec.Apellido_asegurado,
      Poliza:rec.Poliza,
      Tipo_documento:rec.Tipo_documento,
      Numero_documento:rec.Numero_documento,
      Celular:rec.Celular,
      email_asegurado:rec.email_asegurado,
      Marca:rec.Marca,
      Modelo:rec.Modelo,
      Patente:rec.Patente,
      Vigencia:vig
    },
    softWarnings:softWarnings,
    reasons:reasons,
    baseStatus:(vig==='VIGENTE'?'VALIDADO':'NO_VIGENTE')
  };
}

/********* DRIVE HELPERS *********/
function getIdFromDriveUrl_(url){
  const m=String(url).match(/[-\w]{25,}/);
  return m?m[0]:null;
}
function ensureOrCreateChildFolder_(rootFolder, name){
  const it = rootFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return rootFolder.createFolder(name);
}

/********* OCR (Drive Avanzado) *********/
function ocrFileTextById_(fileId,lang){
  if(!fileId) return '';
  const resource={ title:'tmp-ocr-'+fileId, mimeType:MimeType.GOOGLE_DOCS };
  const params={ ocr:true, ocrLanguage:lang||'es' };
  const copied=Drive.Files.copy(resource,fileId,params);
  const docId=copied.id; Utilities.sleep(800);
  const text=DocumentApp.openById(docId).getBody().getText();
  Drive.Files.remove(docId);
  return text||'';
}

/********* Parser Licencia PERÚ *********/
function parsePeruLicenseOCR_(rawText){
  const text = rawText || '';
  const plain = text.replace(/\s+/g,' ').trim();

  let lic = '';
  const reLic1 = /N[°ºo]?\s*de\s*Licencia[:\s]*([A-Z]\s*\d{8,9})/i;
  const reLic2 = /\b([A-Z]\s*\d{8,9})\b/;
  let m = plain.match(reLic1);
  if (m) lic = m[1];
  else{
    const m2 = plain.match(reLic2);
    if (m2) lic = m2[1];
  }
  lic = (lic||'').replace(/\s+/g,'').toUpperCase().replace(/O/g,'0');

  const lines = text.split(/\r?\n/).map(l=>l.trim()).filter(l=>l);

  function extractAroundLabel(labelRegex){
    for (let i = 0; i < lines.length; i++){
      const line = lines[i];
      if (!labelRegex.test(line)) continue;

      let tail = line.replace(labelRegex, '').replace(/^[\s:\-]+/,'').trim();
      if (!tail || /licencia\s+de\s+conducir/i.test(tail) || /republica\s+del?\s+peru/i.test(tail)){
        const next = lines[i+1] || '';
        if (next && !/Apellidos?|Nombres?|Nro\s*de\s*Licencia/i.test(next)){
          tail = next.trim();
        }
      }

      tail = tail.replace(/Nombres?.*/i,'');
      tail = tail.replace(/Apellidos?.*/i,'');
      tail = tail.replace(/Nro\s*de\s*Licencia.*/i,'');

      tail = tail.replace(/LICENCIA\s+DE\s+CONDUCIR/gi,'');
      tail = tail.replace(/REPUBLICA\s+DEL?\s+PERU/gi,'');

      tail = tail.replace(/\s+/g,' ').trim();
      return tail;
    }
    return '';
  }

  let ap = extractAroundLabel(/Apellidos?\s*[:\-]*/i);
  let no = extractAroundLabel(/Nombres?\s*[:\-]*/i);

  ap = ap.replace(/[^A-ZÁÉÍÓÚÑ\s-]/gi,'').replace(/\s{2,}/g,' ').trim();
  no = no.replace(/[^A-ZÁÉÍÓÚÑ\s-]/gi,'').replace(/\s{2,}/g,' ').trim();

  function findDate(label){
    const r = new RegExp(label+'\\s*[:\\-]?\\s*(\\d{2}[\\/\\-]\\d{2}[\\/\\-]\\d{4})','i');
    const m = text.match(r);
    return m ? m[1].replace(/-/g,'/') : '';
  }
  const fExped = findDate('Fecha de Expedici[oó]n');
  const fReval = findDate('Fecha de Revalidaci[oó]n') || findDate('Revalidaci[oó]n');

  const dniTail = lic ? lic.replace(/^[A-Z]/,'').slice(-8) : '';

  return {
    num: lic || '',
    apellidos: ap || '',
    nombres: no || '',
    fechaEmision: fExped || '',
    fechaVenc: fReval || '',
    dniTail: dniTail || ''
  };
}

/********* NLP Relato (impactos, daños, tardío, etc.) *********/
function analyzeRelato_(text){
  const t = String(text||'').toLowerCase();
  const tags = [];
  const detail = {
    terceros: {present:false, tipos:[]},
    impacto: {frontal:false, trasera:false, lateral_der:false, lateral_izq:false, alcance:false},
    maniobras: {giro:false, adelantamiento:false, retroceso:false, cambio_carril:false, estacionado:false},
    contexto: {semaforo:false, interseccion:false, rotatoria:false, cruce_peatonal:false},
    clima: {lluvia:false, pista_mojada:false},
    riesgo: {velocidad:false, alcohol:false, fuga:false, distraccion:false, sin_soat:false},
    soporte: {policia:false, testigos:false, grua:false, fotos:false},
    hora: {noche:false, madrugada:false, dia:false},
    lugar: [],
    danos: {frontal:false, trasera:false, lateral_der:false, lateral_izq:false, techo:false, vidrios:false},
    late_damage:false,
    segundo_impacto:false,
    objeto_impacto:[]
  };

  // terceros
  const tercerosRx = /(tercer|peaton|peatón|taxi|moto|motocicleta|ciclista|bus|ómnibus|omnibus|camion|camión|micro|coaster|combi|propiedad|poste|pared|muro|reja|port[oó]n)/g;
  const mTerc = t.match(tercerosRx);
  if (mTerc){
    detail.terceros.present = true;
    const uniq = Array.from(new Set(mTerc.map(x=>x.normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'').toLowerCase())));
    detail.terceros.tipos = uniq;
    tags.push('terceros:'+uniq.join(','));
  }

  // impacto
  if (/(frontal|de frente|choque frontal|impacte contra)/.test(t)){ detail.impacto.frontal=true; tags.push('impacto:frontal'); }
  if (/(por atrás|por detras|trasera|alcance|me chocaron atras)/.test(t)){
    detail.impacto.trasera=true;
    detail.impacto.alcance = /alcance/.test(t);
    tags.push('impacto:trasera');
  }
  if (/(lateral derecha|costado derecho|lado derecho)/.test(t)){ detail.impacto.lateral_der=true; tags.push('impacto:lateral_derecha'); }
  if (/(lateral izquierda|costado izquierdo|lado izquierdo)/.test(t)){ detail.impacto.lateral_izq=true; tags.push('impacto:lateral_izquierda'); }

  // maniobras
  if (/(gir[oa]|voltee|doble)/.test(t)){ detail.maniobras.giro=true; tags.push('maniobra:giro'); }
  if (/(adelant|rebas)/.test(t)){ detail.maniobras.adelantamiento=true; tags.push('maniobra:adelantamiento'); }
  if (/(retroced|marcha atr[áa]s)/.test(t)){ detail.maniobras.retroceso=true; tags.push('maniobra:retroceso'); }
  if (/(cambio de carril|me cambi[ée] de carril)/.test(t)){ detail.maniobras.cambio_carril=true; tags.push('maniobra:cambio_carril'); }
  if (/(estacionado|aparcado)/.test(t)){ detail.maniobras.estacionado=true; tags.push('maniobra:estacionado'); }

  // contexto
  if (/(sem[áa]foro)/.test(t)){ detail.contexto.semaforo=true; tags.push('semaforo'); }
  if (/(intersecci[óo]n|cruce)/.test(t)){ detail.contexto.interseccion=true; tags.push('interseccion'); }
  if (/(rotatoria|ovalo)/.test(t)){ detail.contexto.rotatoria=true; tags.push('rotatoria'); }
  if (/(cruce peatonal|zebra)/.test(t)){ detail.contexto.cruce_peatonal=true; tags.push('cruce_peatonal'); }

  // clima
  if (/(lluvia|llov[ií]a)/.test(t)){ detail.clima.lluvia=true; tags.push('clima:lluvia'); }
  if (/(pista mojada|calzada mojada|h[uú]meda)/.test(t)){ detail.clima.pista_mojada=true; tags.push('clima:pista_mojada'); }

  // riesgo
  if (/(exceso de velocidad|muy r[áa]pido|alta velocidad|venia rapido|venía rapido|venia muy rapido|venía muy rápido)/.test(t)){
    detail.riesgo.velocidad=true; tags.push('riesgo:velocidad');
  }
  if (/(alcohol|ebrio|licor|cerveza)/.test(t)){ detail.riesgo.alcohol=true; tags.push('riesgo:alcohol'); }
  if (/(se di[oó] a la fuga|fuga|hu[yí]o)/.test(t)){ detail.riesgo.fuga=true; tags.push('riesgo:fuga'); }
  if (/(celular|whatsapp|distra[ií]d)/.test(t)){ detail.riesgo.distraccion=true; tags.push('riesgo:distraccion'); }
  if (/(sin soat|no ten[ií]a soat)/.test(t)){ detail.riesgo.sin_soat=true; tags.push('riesgo:sin_soat'); }

  // soporte
  if (/(polic[ií]a|pnp)/.test(t)){ detail.soporte.policia=true; tags.push('pnp'); }
  if (/(testig)/.test(t)){ detail.soporte.testigos=true; tags.push('testigos'); }
  if (/(gr[uú]a)/.test(t)){ detail.soporte.grua=true; tags.push('grua'); }
  if (/(foto|imagen|adjunt[ée] pruebas)/.test(t)){ detail.soporte.fotos=true; tags.push('fotos'); }

  // hora
  if (/(noche|pm|de la noche)/.test(t)){ detail.hora.noche=true; tags.push('hora:noche'); }
  if (/(madrugada)/.test(t)){ detail.hora.madrugada=true; tags.push('hora:madrugada'); }
  if (/(ma[ñn]ana|mañana|tarde|am|de la ma[ñn]ana|de la tarde)/.test(t)){ detail.hora.dia=true; tags.push('hora:dia'); }

  // lugar
  const lugarRx = /(en (la )?(av\.?|avenida|jr\.?|jir[oó]n|calle|carretera|panamericana)[^.,;]*)/g;
  const matchLugar = t.match(lugarRx);
  if (matchLugar){
    detail.lugar = Array.from(new Set(matchLugar.map(s=>s.trim())));
    tags.push('lugar:'+detail.lugar.slice(0,2).join('|'));
  }

  // daños por zonas (heurístico sencillo)
  if (/(parachoques delantero|faro delantero|parte delantera|frente del veh[ií]culo|frente de la unidad)/.test(t)){
    detail.danos.frontal = true; tags.push('dano:frontal');
  }
  if (/(parachoques trasero|parte posterior|parte de atr[áa]s|guardabarro trasero)/.test(t)){
    detail.danos.trasera = true; tags.push('dano:trasera');
  }
  if (/(lado derecho|costado derecho|puerta derecha|lateral derecho)/.test(t)){
    detail.danos.lateral_der = true; tags.push('dano:lateral_der');
  }
  if (/(lado izquierdo|costado izquierdo|puerta izquierda|lateral izquierdo)/.test(t)){
    detail.danos.lateral_izq = true; tags.push('dano:lateral_izq');
  }
  if (/(techo|parte superior)/.test(t)){
    detail.danos.techo = true;
  }
  if (/(parabrisas|vidrio|luna)/.test(t)){
    detail.danos.vidrios = true;
  }

  // daño tardío
  const frasesLate = [
    'luego de revisar','luego revisé','luego revise',
    'despu[eé]s del impacto','después del impacto',
    'al d[ií]a siguiente','despu[eé]s me di cuenta',
    'm[aá]s tarde me di cuenta',
    'también me di cuenta de que algo le pasó',
    'despu[eé]s noté'
  ];
  if (frasesLate.some(f=>t.indexOf(f.toLowerCase())!==-1)){
    detail.late_damage = true;
    tags.push('dano:tardio');
  }

  // segundo impacto / objetos
  const objetos = [];
  if (/(impact[ée] contra un poste|choqu[ée] contra un poste)/.test(t)) objetos.push('poste');
  if (/(impact[ée] contra la pared|muro)/.test(t)) objetos.push('pared/muro');
  if (objetos.length){
    detail.segundo_impacto = true;
    detail.objeto_impacto = objetos;
  }

  // severidad (heurística simple)
  let sev = 'BAJA';
  const scoreHeu =
    (detail.impacto.frontal?2:0) +
    (detail.impacto.trasera?2:0) +
    (detail.terceros.present?2:0) +
    (detail.riesgo.velocidad?2:0) +
    (detail.riesgo.alcohol?3:0) +
    (detail.clima.lluvia?1:0) +
    (detail.hora.madrugada?1:0) +
    (detail.danos.frontal?1:0) +
    (detail.danos.trasera?1:0);
  if (scoreHeu>=6) sev='ALTA';
  else if (scoreHeu>=3) sev='MEDIA';

  // map url
  let mapUrl = '';
  if (detail.lugar && detail.lugar.length){
    const q = encodeURIComponent(detail.lugar[0]);
    mapUrl = 'https://www.google.com/maps/search/?api=1&query='+q;
  }

  return {
    tags: Array.from(new Set(tags)),
    detail,
    severity: sev,
    mapUrl
  };
}

/********* Pre-subida base64 *********/
function uploadBase64(payload){
  const root = DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  const folderName = (payload.folderDateKey || Utilities.formatDate(new Date(), tz_(),'yyyyMMdd'))
                    + '_' + (payload.policy || 'SIN_POLIZA');

  let target;
  if (payload.preFolderId) {
    try { target = DriveApp.getFolderById(payload.preFolderId); }
    catch(e) { target = ensureOrCreateChildFolder_(root, folderName); }
  } else {
    target = ensureOrCreateChildFolder_(root, folderName);
  }

  const bytes = Utilities.base64Decode(payload.dataBase64);
  const blob  = Utilities.newBlob(bytes, payload.mime || MimeType.BINARY, payload.name || (payload.label || 'archivo') + '.bin');
  const file = target.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}

  return { ok:true, url:file.getUrl(), name:payload.label || payload.name, folderUrl: target.getUrl(), folderId: target.getId() };
}

/********* Diagnósticos: OCR + Presupuesto + Relato *********/
function buildDiagnostics_(p, linkMap){
  const diag = {
    reported: {
      patente: normalizePlate_(p.Patente || ''),
      doc: String(p.Numero_documento||p.NroDoc||'').replace(/\D/g,''),
      nombreCompleto: String((p.Nombre_asegurado||p.Nomb||'')+' '+(p.Apellido_asegurado||p.Apell||'')).trim().toUpperCase()
    },
    ocr: {
      patente: '',
      patenteMatch: null,
      patenteSource: '',
      license: { num:'', apellidos:'', nombres:'', fechaEmision:'', fechaVenc:'', dniTail:'' }
    },
    presupuesto: { amount:'', currency:'', status:'', zonas:null },
    attachments: { licencia:false, frontal:false, trasera:false, dano:false, opcional:false, presupuesto:false },
    relato: { tags:[], detail:{}, severity:'', mapUrl:'' }
  };

  // Adjuntos
  Object.keys(linkMap||{}).forEach(function(k){
    const arr = linkMap[k]||[];
    const has = arr.length>0;
    if (/licencia/i.test(k)) diag.attachments.licencia = diag.attachments.licencia || has;
    if (/frontal/i.test(k)) diag.attachments.frontal = diag.attachments.frontal || has;
    if (/trasera/i.test(k)) diag.attachments.trasera = diag.attachments.trasera || has;
    if (/dañ|dano/i.test(k)) diag.attachments.dano = diag.attachments.dano || has;
    if (/opcional/i.test(k)) diag.attachments.opcional = diag.attachments.opcional || has;
    if (/presupuesto/i.test(k)) diag.attachments.presupuesto = diag.attachments.presupuesto || has;
  });

  // OCR Patente
  const imgs=[];
  Object.keys(linkMap||{}).forEach(function(k){
    (linkMap[k]||[]).forEach(function(it){
      if (/frontal|trasera|dañ|dano|opcional/i.test((it.name||'')+' '+k)){
        imgs.push({key:k, url:it.url, name:it.name||k});
      }
    });
  });
  for (let i=0;i<imgs.length;i++){
    const id=getIdFromDriveUrl_(imgs[i].url);
    if (!id) continue;
    try{
      const t=ocrFileTextById_(id,'es');
      const best=extractBestPlate_(t);
      if (best){
        diag.ocr.patente = best;
        diag.ocr.patenteMatch = (levenshtein_(normalizePlate_(best), diag.reported.patente) <= 1);
        const label = (imgs[i].name||imgs[i].key).toLowerCase();
        if (/frontal/.test(label)) diag.ocr.patenteSource = 'Frontal';
        else if (/trasera/.test(label)) diag.ocr.patenteSource = 'Trasera';
        else diag.ocr.patenteSource = 'Otro';
        break;
      }
    }catch(err){}
  }

  // OCR Licencia
  const licKey = Object.keys(linkMap||{}).find(k=>/licencia/i.test(k));
  if (licKey && (linkMap[licKey]||[]).length){
    const id = getIdFromDriveUrl_(linkMap[licKey][0].url);
    try{
      const txt=ocrFileTextById_(id,'es');
      const parsed = parsePeruLicenseOCR_(txt);
      diag.ocr.license = parsed;
    }catch(err){}
  }

  // OCR Presupuesto
  const presKey = Object.keys(linkMap||{}).find(k=>/presupuesto/i.test(k));
  if (presKey && (linkMap[presKey]||[]).length){
    const id=getIdFromDriveUrl_(linkMap[presKey][0].url);
    try{
      const txt=ocrFileTextById_(id,'es');
      let m = parseBudgetTotal_(txt);
      if (!m || !m.amount || m.amount <= 0){
        m = moneyParse_(txt);
      }
      if (m.amount>0){
        diag.presupuesto.amount = m.amount;
        diag.presupuesto.currency = m.currency || '$';
        diag.presupuesto.status = 'OK';
      } else {
        diag.presupuesto.status = 'SIN_MONTO';
        diag.presupuesto.currency = '$';
      }
    }catch(err){
      diag.presupuesto.status = 'ERROR_OCR';
      diag.presupuesto.currency = '$';
    }
  } else {
    diag.presupuesto.status = 'SIN_PRESUPUESTO';
    diag.presupuesto.currency = '$';
  }

  // Relato
  const rel = analyzeRelato_(p.Relato_asegurado || p.Relato || '');
  diag.relato.tags = rel.tags;
  diag.relato.detail = rel.detail;
  diag.relato.severity = rel.severity;
  diag.relato.mapUrl = rel.mapUrl;

  return diag;
}

/********* Preview (si lo usas en el front) *********/
function previewDiagnostics(payload, prelinksJson){
  let map = {};
  try{ map = prelinksJson ? JSON.parse(prelinksJson) : {}; }catch(e){}
  return buildDiagnostics_(payload, map);
}

/********* VALIDACIÓN CRÍTICA (BLOQUEO) *********/
function validateOcrAndBlock_(p, linkMap){
  const diag = buildDiagnostics_(p, linkMap);

  // Patente imagen ≠ reportada
  if (diag.ocr.patente && diag.ocr.patenteMatch === false){
    return {
      block:true,
      code:'BLOCK_OCR_PLATE_MISMATCH',
      message:'La patente de la imagen no coincide con la ingresada.',
      diagnostics: diag
    };
  }

  const lic = diag.ocr.license;
  const licNum = lic.num;
  if (!licNum || !/^[A-Z][0-9]{8,9}$/.test(licNum)){
    return {
      block:true,
      code:'BLOCK_OCR_LICENSE',
      message:'No se pudo validar el N° de Licencia (letra + 8/9 dígitos). Adjunta anverso nítido.',
      diagnostics: diag
    };
  }

  const dni = String(p.Numero_documento||p.NroDoc||'').replace(/\D/g,'');
  if (dni && /^\d{8}$/.test(dni)){
    if (lic.dniTail !== dni){
      return {
        block:true,
        code:'BLOCK_OCR_LICENSE',
        message:'El DNI embebido en el N° de Licencia no coincide con el ingresado.',
        diagnostics: diag
      };
    }
  }

  const fullOCR = (lic.nombres+' '+lic.apellidos).replace(/\s+/g,'').toUpperCase();
  const fullRep = String(((p.Nombre_asegurado||p.Nomb||'')+' '+(p.Apellido_asegurado||p.Apell||''))).replace(/\s+/g,'').toUpperCase();
  if (fullOCR && fullRep){
    const dist = levenshtein_(fullOCR, fullRep);
    if (dist > Math.max(3, Math.floor(fullRep.length*0.15))){
      return {
        block:true,
        code:'BLOCK_OCR_LICENSE_NAME',
        message:'El nombre de la Licencia no coincide con el declarado.',
        diagnostics: diag
      };
    }
  }

  const now = new Date();
  if (lic.fechaEmision){
    const em = parseDateSmart_(lic.fechaEmision);
    if (em && em > now){
      return {
        block:true,
        code:'BLOCK_OCR_LICENSE_DATES',
        message:'Fecha de Expedición futura en Licencia.',
        diagnostics: diag
      };
    }
  }
  if (lic.fechaVenc){
    const ve = parseDateSmart_(lic.fechaVenc);
    if (ve && ve < now){
      return {
        block:true,
        code:'BLOCK_OCR_LICENSE_DATES',
        message:'La Licencia aparece vencida.',
        diagnostics: diag
      };
    }
  }

  return { block:false, code:null, message:'OK', diagnostics: diag };
}

/********* HELPERS ANTIFRAUDE *********/
function containsAny_(text, patterns){
  const t = String(text||'').toLowerCase();
  return patterns.some(p => t.indexOf(p.toLowerCase()) !== -1);
}
function indexOfAnyNorm_(hdr, names){
  const normHdr = hdr.map(h=>norm(h));
  for (let i=0; i<names.length; i++){
    const target = norm(names[i]);
    const idx = normHdr.indexOf(target);
    if (idx !== -1) return idx;
  }
  return -1;
}
function buildHistorySnapshot_(ctx){
  const sh = getSheetByGidIdSafe_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return { countDni12m:0, countPat12m:0 };
  }

  const hdr = values[0].map(h=>String(h||'').trim());
  const idxDni = indexOfAnyNorm_(hdr, ['NroDoc','Numero_documento']);
  const idxPat = indexOfAnyNorm_(hdr, ['Patente']);
  const idxFec = indexOfAnyNorm_(hdr, ['FecOcc','Fecha_ocurrencia']);

  const dniActual = String(ctx.p.Numero_documento || ctx.p.NroDoc || '').replace(/\D/g,'');
  const patActual = normalizePlate_(ctx.p.Patente || '');

  if (!dniActual && !patActual) {
    return { countDni12m:0, countPat12m:0 };
  }

  const now = ctx.now || new Date();
  const oneYearAgo = new Date(now.getTime() - 365*24*3600*1000);

  let countDni12m = 0;
  let countPat12m = 0;

  for (let r=1; r<values.length; r++){
    const row = values[r];
    const dF = idxFec>=0 ? parseDateSmart_(row[idxFec]) : null;
    if (!dF || dF < oneYearAgo || dF > now) continue;

    if (idxDni>=0 && dniActual){
      const dniRow = String(row[idxDni]||'').replace(/\D/g,'');
      if (dniRow && dniRow === dniActual) countDni12m++;
    }
    if (idxPat>=0 && patActual){
      const patRow = normalizePlate_(row[idxPat] || '');
      if (patRow && patRow === patActual) countPat12m++;
    }
  }

  return { countDni12m, countPat12m };
}
function decideChannel_(score, flags){
  const criticalFlags = [
    'MULTIZONA_SIN_SEGUNDO_IMPACTO',
    'DANO_TARDIO_OTRA_ZONA',
    'TERCEROS_NO_DECLARADOS',
    'FRECUENCIA_DNI_ALTA',
    'FRECUENCIA_PLACA_ALTA',
    'PATENTE_OCR_NO_COINCIDE'
  ];
  const anyCritical = criticalFlags.some(f => flags[f]);
  if (anyCritical && score >= 40) return 'FRAUDES';
  if (score >= 70) return 'FRAUDES';
  return 'SINIESTROS';
}

/********* MOTOR CENTRAL ANTIFRAUDE *********/
function evaluateRiskRules_(ctx){
  const flags = {};
  const dims = { documental:0, fisico:0, relato:0, presupuesto:0, historico:0, contacto:0 };
  let scoreTotal = 0;
  const chips = [];

  function addFlag(name, dim, weight, chip){
    flags[name] = true;
    scoreTotal += weight;
    dims[dim] = (dims[dim]||0) + weight;
    if (chip) chips.push(chip);
  }

  const p = ctx.p || {};
  const diag = ctx.diag || {};
  const rel = ctx.rel || {};
  const d = rel.detail || {};
  const pres = ctx.pres || {};
  const impacto = d.impacto || {};
  const danos = d.danos || {};
  const lateDamage = !!d.late_damage;
  const severidad = rel.severity || 'BAJA';
  const relText = String(p.Relato_asegurado || p.Relato || '').toLowerCase();
  const tipoSin = String(ctx.tipoSiniestro || p.Tipo_siniestro || '').toLowerCase();
  const parteLower = String(ctx.parteLower || p.Parte_afectada || '').toLowerCase();
  const tercerosCampo = String(ctx.tercerosCampo || p.Dano_a_terceros || p['Daño_a_terceros'] || '').toLowerCase();
  const baseStatus = ctx.baseStatus || '';
  const have = ctx.have || {};
  const dupInfo = ctx.dupInfo || {};
  const now = ctx.now || new Date();
  const dOcc = ctx.fechaOcc || null;
  const lic = (diag.ocr && diag.ocr.license) ? diag.ocr.license : {};
  const reported = diag.reported || {};
  const dniReported = String(reported.doc || p.Numero_documento || p.NroDoc || '').replace(/\D/g,'');
  const patenteReported = normalizePlate_(reported.patente || p.Patente || '');
  const patOCR = diag.ocr ? diag.ocr.patente : '';
  const patOcrMatch = diag.ocr ? diag.ocr.patenteMatch : null;

  /********** 1) DOCUMENTAL / LICENCIA / PÓLIZA **********/
  if (baseStatus && baseStatus !== 'VALIDADO'){
    addFlag('POLIZA_NO_VIGENTE','documental',20,'🔴 Póliza no vigente/base');
  }

  if (lic.fechaVenc){
    const v = parseDateSmart_(lic.fechaVenc);
    if (v && v < now){
      addFlag('LIC_VENCIDA','documental',25,'🔴 Licencia vencida');
    }
  }

  if (dniReported && /^\d{8}$/.test(dniReported) && lic.dniTail && lic.dniTail !== dniReported){
    addFlag('DNI_LIC_NO_MATCH','documental',30,'🔴 DNI de licencia no coincide');
  }

  const fullOCR = (lic.nombres+' '+lic.apellidos).replace(/\s+/g,'').toUpperCase();
  const fullRep = String(reported.nombreCompleto || ((p.Nombre_asegurado||p.Nomb||'')+' '+(p.Apellido_asegurado||p.Apell||''))).replace(/\s+/g,'').toUpperCase();
  if (fullOCR && fullRep){
    const dist = levenshtein_(fullOCR, fullRep);
    if (dist > Math.max(3, Math.floor(fullRep.length*0.15))){
      addFlag('NOMBRE_LIC_NO_MATCH','documental',25,'⚠️ Nombre en licencia difiere');
    }
  }

  if (lic.fechaEmision && dOcc){
    const em = parseDateSmart_(lic.fechaEmision);
    if (em){
      const diffDays = (dOcc - em)/(1000*3600*24);
      if (diffDays >= 0 && diffDays <= 30 && (severidad === 'ALTA' || (pres.amount||0) > 3000)){
        addFlag('LIC_NUEVA_SINIESTRO_ALTO','documental',15,'⚠️ Licencia reciente + siniestro fuerte');
      }
    }
  }

  if (!lic.num){
    addFlag('SIN_LICENCIA_OCR','documental',15,'⚠️ Licencia no legible por OCR');
  }

  /********** 2) FÍSICO / DAÑOS / PARTE **********/
  const impFr = !!impacto.frontal;
  const impTr = !!impacto.trasera;
  const impLd = !!impacto.lateral_der;
  const impLi = !!impacto.lateral_izq;

  const dzFr = !!danos.frontal;
  const dzTr = !!danos.trasera;
  const dzLd = !!danos.lateral_der;
  const dzLi = !!danos.lateral_izq;

  const multiZonaOpuesta =
    (dzFr && dzTr) ||
    (dzFr && (dzLi || dzLd)) ||
    (dzTr && (dzLi || dzLd));

  if (parteLower){
    if (parteLower.indexOf('frontal') > -1 && impTr && !impFr){
      addFlag('PARTE_IMPACTO_INCOHERENTE','fisico',25,'⚠️ Parte frontal pero relato indica impacto trasero');
    }
    if (parteLower.indexOf('trasera') > -1 && impFr && !impTr){
      addFlag('PARTE_IMPACTO_INCOHERENTE','fisico',25,'⚠️ Parte trasera pero relato indica impacto frontal');
    }
    if (parteLower.indexOf('lateral derecha') > -1 && !impLd && (impLi || impFr || impTr)){
      addFlag('PARTE_IMPACTO_INCOHERENTE','fisico',25,'⚠️ Parte lateral derecha, relato apunta a otra zona');
    }
    if (parteLower.indexOf('lateral izquierda') > -1 && !impLi && (impLd || impFr || impTr)){
      addFlag('PARTE_IMPACTO_INCOHERENTE','fisico',25,'⚠️ Parte lateral izquierda, relato apunta a otra zona');
    }
  }

  const sinSegundoImpacto = !d.segundo_impacto && (!d.objeto_impacto || !d.objeto_impacto.length);
  if (multiZonaOpuesta && sinSegundoImpacto){
    addFlag('MULTIZONA_SIN_SEGUNDO_IMPACTO','fisico',30,'🔴 Daños en múltiples zonas opuestas sin segundo impacto declarado');
  }

  if (lateDamage){
    const impactoPrincipalTrasero = impTr && !impFr;
    const impactoPrincipalFrontal = impFr && !impTr;
    if (multiZonaOpuesta ||
        (impactoPrincipalTrasero && dzFr) ||
        (impactoPrincipalFrontal && dzTr)){
      addFlag('DANO_TARDIO_OTRA_ZONA','fisico',25,'🔴 Daño detectado después del impacto en otra zona');
    } else {
      addFlag('DANO_TARDIO_MISMA_ZONA','fisico',15,'⚠️ Daño detectado después del impacto');
    }
  }

  const frasesMin = ['roce leve','golpe leve','apenas me tocó','apenas me toco','apenas un raspon','ligero golpe','golpecito','golpe pequeño','pequeño golpe'];
  if (containsAny_(relText, frasesMin) && (severidad === 'ALTA' || (pres.amount||0) > 3000)){
    addFlag('RELATO_MINIMO_CON_DANO_ALTO','relato',20,'⚠️ Relato minimiza el golpe pero daño/presupuesto alto');
  }

  const frasesEvasivas = ['no recuerdo','no estoy seguro','no estoy segura','no puedo precisar','no puedo indicar','no tengo certeza','no sé exactamente','no se exactamente'];
  if (containsAny_(relText, frasesEvasivas)){
    addFlag('RELATO_EVASIVO','relato',10,'⚠️ Relato con muchas imprecisiones');
  }

  const frasesOrdenIncierto = ['no puedo precisar el orden','no puedo decir en qué momento','no puedo decir en que momento','no sé si fue antes o después','no se si fue antes o despues'];
  if (containsAny_(relText, frasesOrdenIncierto)){
    addFlag('RELATO_ORDEN_INCIERTO','relato',15,'⚠️ Orden de daños poco claro');
  }

  /********** 3) TERCEROS **********/
  if (d.terceros && d.terceros.present){
    const declaraNo = tercerosCampo.indexOf('no') !== -1 || tercerosCampo.indexOf('ninguno') !== -1;
    if (declaraNo){
      addFlag('TERCEROS_NO_DECLARADOS','fisico',25,'🔴 Relato menciona terceros pero cuestionario dice que no');
    }
  }

  const mencionaObjetoFijo = containsAny_(relText,['poste','muro','pared','columna','reja','portón','porton']);
  if (mencionaObjetoFijo && (pres.amount||0) > 3000){
    addFlag('OBJETO_FIJO_PRES_ALTO','fisico',15,'⚠️ Daño contra objeto fijo con presupuesto relativamente alto');
  }

  /********** 4) PRESUPUESTO **********/
  if (tipoSin.indexOf('daño parcial') !== -1 || tipoSin.indexOf('danio parcial') !== -1){
    if ((pres.amount||0) >= 5000){
      addFlag('PRES_OUTLIER_MONTO_ALTO','presupuesto',20,'⚠️ Presupuesto alto para daño parcial');
    }
  }
  if (pres.status === 'SIN_MONTO' || pres.status === 'SIN_PRESUPUESTO'){
    addFlag('PRES_SIN_MONTO','presupuesto',15,'⚠️ Presupuesto sin monto claro');
  }

  /********** 5) HISTÓRICO **********/
  const hist = buildHistorySnapshot_(ctx);
  ctx.hist = hist;
  if (hist.countDni12m >= 3){
    addFlag('FRECUENCIA_DNI_ALTA','historico',25,'⚠️ Alta frecuencia de siniestros por DNI (12 meses)');
  }
  if (hist.countPat12m >= 3){
    addFlag('FRECUENCIA_PLACA_ALTA','historico',25,'⚠️ Alta frecuencia de siniestros por placa (12 meses)');
  }

  /********** 6) CONTACTO / ADJUNTOS **********/
  if (ctx.badEmail || ctx.badPhone){
    addFlag('CONTACTO_DEFICIENTE','contacto',10,'⚠️ Datos de contacto incompletos/incorrectos');
  }

  const totalFotos = (ctx.have.Frontal?1:0) + (ctx.have.Trasera?1:0) + (ctx.have.Dano?1:0) + (ctx.have.Opcional?1:0);
  if (severidad === 'ALTA' && totalFotos <= 1){
    addFlag('EVIDENCIA_ESCASA_SEVERIDAD_ALTA','fisico',15,'⚠️ Pocas fotos para severidad alta');
  }

  if (ctx.faltan && ctx.faltan.length){
    addFlag('ADJUNTOS_INCOMPLETOS','contacto',10,'⚠️ Adjuntos faltantes: '+ctx.faltan.join(', '));
  }

  if (patOCR && patOcrMatch === false){
    addFlag('PATENTE_OCR_NO_COINCIDE','documental',25,'🔴 Patente de imagen no coincide con la declarada');
  }

  /********** 7) SCORE → SEMÁFORO + CANAL **********/
  const cfg = getConfig_();
  const thrR = cfg.thresholds.ROJO || 70;
  const thrA = cfg.thresholds.AMBAR || 40;
  let semaforo = 'VERDE';
  if (scoreTotal >= thrR) semaforo = 'ROJO';
  else if (scoreTotal >= thrA) semaforo = 'AMBAR';

  const canal = decideChannel_(scoreTotal, flags);

  /********** 8) NEXT ACTION **********/
  let nextAction = '';
  if (canal === 'FRAUDES'){
    nextAction = 'Derivar el caso al área de Fraudes para análisis detallado y solicitar aclaración sobre los puntos incoherentes.';
  } else if (flags.MULTIZONA_SIN_SEGUNDO_IMPACTO || flags.DANO_TARDIO_OTRA_ZONA || flags.PARTE_IMPACTO_INCOHERENTE){
    nextAction = 'Revisar en detalle relato y fotos, solicitar explicación de daños en múltiples zonas y confirmar versión del asegurado.';
  } else if (flags.TERCEROS_NO_DECLARADOS){
    nextAction = 'Confirmar presencia de terceros, datos de contacto y validar consistencia con el relato.';
  } else if (ctx.faltan && ctx.faltan.length){
    nextAction = 'Solicitar adjuntos faltantes: '+ctx.faltan.join(', ');
  } else if (ctx.badEmail || ctx.badPhone){
    nextAction = 'Actualizar/confirmar datos de contacto del asegurado.';
  } else if (scoreTotal >= thrA){
    nextAction = 'Priorizar revisión por analista de siniestros.';
  } else {
    nextAction = 'Caso apto para flujo normal de siniestros, sin alertas críticas.';
  }

  if (chips.length === 0){
    chips.push('✅ Sin alertas de fraude relevantes');
  }

  return {
    scoreTotal,
    semaforo,
    flags,
    dims,
    canal,
    resumenChips: chips,
    nextAction,
    hist
  };
}

/********* UI WEBAPP *********/
function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Formulario de Siniestros');
}
function getWebAppUrl(){
  return ScriptApp.getService().getUrl();
}

/********* doPost: núcleo *********/
function doPost(e){
  try{
    const p = e.parameter || {};
    const folderDateKey = Utilities.formatDate(new Date(),tz_(),'yyyyMMdd');
    const root=DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);

    // Prelinks
    const linkMap={};
    if (p.prelinks){
      try{
        const parsed = JSON.parse(p.prelinks);
        Object.keys(parsed).forEach(function(field){
          (parsed[field]||[]).forEach(function(it){
            if (!linkMap[field]) linkMap[field] = [];
            linkMap[field].push({url:it.url, name:it.name||field});
          });
        });
      }catch(err){ Logger.log('Error al parsear prelinks: '+err); }
    }

    // Carpeta objetivo
    let folder;
    if (p.preFolderId){
      try { folder = DriveApp.getFolderById(p.preFolderId); }
      catch(eNotFound){ folder = ensureOrCreateChildFolder_(root, folderDateKey+'_'+(p.Poliza||'SIN_POLIZA')); }
    } else {
      folder = ensureOrCreateChildFolder_(root, folderDateKey+'_'+(p.Poliza||'SIN_POLIZA'));
    }

    // 1) Validación crítica (posible bloqueo)
    const critical = validateOcrAndBlock_(p, linkMap);
    if (critical.block) {
      try { folder.setTrashed(true); } catch(eTrash){}
      return ContentService
        .createTextOutput(JSON.stringify({
          ok:false,
          blocked:true,
          code:critical.code || 'BLOCK',
          message:critical.message || 'Bloqueado',
          diagnostics: critical.diagnostics
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2) Construcción de diagnósticos completos
    const sh=getSheetByGidIdSafe_();
    const headers=ensureHeaders_(sh);

    const tipo=String(p.Tipo_siniestro||p.TipoSin||'');
    const patente=normalizePlate_(p.Patente||'');
    const relato=String(p.Relato_asegurado||p.Relato||'');
    const parte=String(p.Parte_afectada||p.Parte||'').toLowerCase();
    const terceros=String(p.Dano_a_terceros||p['Daño_a_terceros']||p.Ter||'').toLowerCase();
    const email=p.email_asegurado||p.Email||'';
    const celular=p.Celular||p.Fono||'';
    const fechaOcc=p.Fecha_ocurrencia||p.FecOcc||'';
    const validBase=p.Validacion_base||p.ValBase||'';

    const dOcc = parseDateSmart_(fechaOcc);

    // Relato
    const relAnalisis = analyzeRelato_(relato);
    const relTags = relAnalisis.tags.join('; ');
    const relDet = relAnalisis.detail;
    const relSev = relAnalisis.severity;
    const relMapUrl = relAnalisis.mapUrl;

    // Adjuntos
    const have = {
      Licencia: !!(linkMap.Adjuntar_licencia && linkMap.Adjuntar_licencia.length),
      Dano: !!(linkMap.Adjuntar_imagen_dano && linkMap.Adjuntar_imagen_dano.length),
      Presupuesto: !!(linkMap.Adjuntar_presupuesto && linkMap.Adjuntar_presupuesto.length),
      Frontal: !!(linkMap.Adjuntar_imagen_frontal && linkMap.Adjuntar_imagen_frontal.length),
      Trasera: !!(linkMap.Adjuntar_imagen_trasera && linkMap.Adjuntar_imagen_trasera.length),
      Opcional: !!(linkMap.Adjuntar_imagen_opcional && linkMap.Adjuntar_imagen_opcional.length)
    };
    const faltan=[];
    if(tipo.toLowerCase()==='daño parcial' || tipo.toLowerCase()==='danio parcial'){
      if(!have.Dano) faltan.push('Imagen daño');
      if(!have.Presupuesto) faltan.push('Presupuesto');
    }
    if(!have.Licencia) faltan.push('Licencia');

    // Fecha ocurrencia
    let alertaFecha='';
    if(!dOcc) alertaFecha='FECHA_VACIA_O_INVALIDA';
    else{
      const now=new Date(); const diff=(now-dOcc)/(1000*3600*24);
      if(dOcc>now) alertaFecha='FECHA_FUTURA';
      else if(diff>365) alertaFecha='MAYOR_12_MESES';
    }

    // Contacto
    const badEmail=!isValidEmail_(email);
    const badPhone=!isValidPhone_(celular);

    // Terceros relato vs campo
    const relatoSugiereTerceros = !!(relDet.terceros && relDet.terceros.present);
    const alertaRelTerc = (terceros==='no' && relatoSugiereTerceros) ? 'POSIBLE_TERCEROS' : '';

    // Coherencia parte (simple; lo fino lo hace el motor antifraude)
    let alertaParte = '';
    const parteLower = parte;
    const impactoRelFrontal = relDet.impacto && relDet.impacto.frontal;
    const impactoRelTrasera = relDet.impacto && relDet.impacto.trasera;
    const impactoRelLatDer = relDet.impacto && relDet.impacto.lateral_der;
    const impactoRelLatIzq = relDet.impacto && relDet.impacto.lateral_izq;
    if (parteLower.indexOf('frontal')>-1 && impactoRelTrasera) alertaParte='RELATO_APUNTA_TRASERA';
    else if (parteLower.indexOf('trasera')>-1 && impactoRelFrontal) alertaParte='RELATO_APUNTA_FRONTAL';
    else if (parteLower.indexOf('lateral derecha')>-1 && !impactoRelLatDer && (impactoRelLatIzq||impactoRelFrontal||impactoRelTrasera)) alertaParte='RELATO_NO_APUNTA_LATERAL_DERECHA';
    else if (parteLower.indexOf('lateral izquierda')>-1 && !impactoRelLatIzq && (impactoRelLatDer||impactoRelFrontal||impactoRelTrasera)) alertaParte='RELATO_NO_APUNTA_LATERAL_IZQUIERDA';

    // Diagnósticos completos (OCR/Presupuesto/Relato)
    const diag = buildDiagnostics_(p, linkMap);

    // Patente OCR
    const patenteOCR = diag.ocr.patente || '';
    const patenteOcrAlert = patenteOCR
      ? (diag.ocr.patenteMatch===false ? 'NO_COINCIDE' : 'COINCIDE')
      : 'NO_ENCONTRADA';
    const ocrFuente = diag.ocr.patenteSource || '';

    // Presupuesto
    let monto='', moneda='', presEstado='';
    if (diag.presupuesto.status==='OK'){
      monto=diag.presupuesto.amount;
      moneda=diag.presupuesto.currency || '$';
      presEstado='OK';
    }else{
      presEstado=diag.presupuesto.status || 'SIN_MONTO';
    }
    if(String(tipo).toLowerCase()==='daño parcial' && parseFloat(monto)>5000){
      presEstado = 'POSIBLE_OUTLIER';
    }

    // Duplicados
    let dup='', dupRef='';
    const all=sh.getDataRange().getValues();
    if(all.length>1){
      const idxHdr = all[0].map(h=>String(h||'').trim());
      const idxPat=idxHdr.indexOf('Patente');
      const idxFec=idxHdr.indexOf('FecOcc');
      const idxFolder=idxHdr.indexOf('UrlFolder');
      if(plausiblePlate_(patente) && dOcc && idxPat>=0 && idxFec>=0){
        let count=0,last='';
        const links=[];
        for(let r=1;r<all.length;r++){
          const p2=normalizePlate_(all[r][idxPat]||''); if(p2!==patente) continue;
          const d2=parseDateSmart_(all[r][idxFec]); if(!d2) continue;
          const diff=Math.abs((dOcc-d2)/(1000*3600*24));
          if(diff<=DUP_WINDOW_DAYS){
            count++; last=Utilities.formatDate(d2,tz_(),'yyyy-MM-dd');
            if (idxFolder>=0){
              const furl = String(all[r][idxFolder]||'');
              if (furl){
                links.push('HYPERLINK("'+furl+'","'+last+'")');
              }
            }
          }
        }
        if(count>0){
          dup='DUPLICADO_'+count+'_ULT_'+DUP_WINDOW_DAYS+'D (último='+last+')';
          if (links.length) dupRef = '='+links.join(' & " | " & ');
        }
      }
    }

    // Motor antifraude
    const ctx = {
      p,
      diag,
      rel: diag.relato || relAnalisis,
      pres: diag.presupuesto || {},
      baseStatus: validBase,
      fechaOcc: dOcc,
      have,
      dupInfo: { dup: dup || '', dupRef: dupRef || '' },
      now: new Date(),
      badEmail,
      badPhone,
      tipoSiniestro: tipo,
      parteLower,
      tercerosCampo: terceros,
      faltan
    };
    const risk = evaluateRiskRules_(ctx);
    const score = risk.scoreTotal;
    const semaf = risk.semaforo;
    const chips = risk.resumenChips;
    const nextAction = risk.nextAction;
    const canalAsignado = risk.canal;

    // Licencia check simple
    let A_Lic = '';
    const lic = diag.ocr.license || {};
    if (lic.num){
      if (risk.flags.LIC_VENCIDA || risk.flags.DNI_LIC_NO_MATCH || risk.flags.NOMBRE_LIC_NO_MATCH){
        A_Lic = '⚠️';
      } else {
        A_Lic = '✅';
      }
    }

    // A_Pres
    let A_Pres = '';
    if (presEstado === 'OK') A_Pres = E.MONEY;
    else A_Pres = presEstado;

    // A_PatOCR como check
    let A_PatOCR = '';
    if (patenteOCR){
      A_PatOCR = (diag.ocr.patenteMatch===false ? '⚠️' : '✅');
    }

    // A_Vig / alertas
    const A_Vig = (validBase==='VALIDADO')?'':(validBase||'NO_VIGENTE');
    const A_Fec = alertaFecha;
    const A_Contacto = (badEmail||badPhone)
      ? (badEmail?'EMAIL ':'')+(badPhone?'CELULAR ':'')
      : '';
    const A_PatForm = plausiblePlate_(patente)?'':'FORMATO_DUDOSO';
    const A_Adj = faltan.length?'FALTANTES':'';
    const Adj_Falt = faltan.join(', ');
    const A_Ter = risk.flags.TERCEROS_NO_DECLARADOS ? 'TERCEROS_NO_DECLARADOS' : alertaRelTerc;
    const A_Parte = risk.flags.PARTE_IMPACTO_INCOHERENTE ? 'PARTE_IMPACTO_INCOHERENTE' : alertaParte;
    const A_Dup = dup;

    // Prioridad según canal/semaforo
    const prior = (canalAsignado === 'FRAUDES')
      ? 'FRAUDES'
      : (semaf === 'ROJO' ? 'ALTA' : (semaf === 'AMBAR' ? 'MEDIA' : 'BAJA'));

    // Links adjuntos en una celda
    const buildLinksCell=(map)=>{
      const parts=[];
      Object.keys(map).forEach(function(k){
        map[k].forEach(function(it,idx){
          const lab=(it.name||k)+(map[k].length>1?(' '+(idx+1)):'' );
          parts.push('HYPERLINK("'+it.url+'","'+lab+'")');
        });
      });
      return parts.length ? '='+parts.join(' & CHAR(10) & ') : '';
    };

    // Fila según layout corto
    const src = {
      'Poliza': p.Poliza||'',
      'Marca': p.Marca||'',
      'Modelo': p.Modelo||'',
      'Patente': p.Patente||'',
      'Revisado': false,

      'Nomb': p.Nombre_asegurado||p.Nomb||'',
      'Apell': p.Apellido_asegurado||p.Apell||'',
      'TipoDoc': p.Tipo_documento||p.TipoDoc||'',
      'NroDoc': p.Numero_documento||p.NroDoc||'',
      'Email': email,
      'Fono': celular,
      'ValBase': validBase,

      'TipoSin': p.Tipo_siniestro||p.TipoSin||'',
      'FecOcc': fechaOcc,
      'HrOcc': p.Hora_ocurrencia||p.HrOcc||'',
      'Relato': relato,
      'SevRel': relSev,
      'UrlRel': relMapUrl,

      'Parte': p.Parte_afectada||p.Parte||'',
      'Ter': p.Dano_a_terceros||p['Daño_a_terceros']||p.Ter||'',
      'TagsRel': relTags,

      'Adj': buildLinksCell(linkMap),
      'UrlFolder': folder.getUrl(),

      'OCRsrc': ocrFuente,
      'PatOCR': patenteOCR,
      'A_PatOCR': A_PatOCR,

      'MontoPres': monto,
      'A_Pres': A_Pres,

      'Prior': prior,
      'Asig': '',
      'Next': nextAction,

      'A_Lic': A_Lic,

      'JSON': JSON.stringify(diag),
      'DupRef': dupRef || '',
      'Resumen': chips.join('  |  '),

      'A_Vig': A_Vig,
      'A_Fec': A_Fec,
      'A_Contacto': A_Contacto,
      'A_PatForm': A_PatForm,
      'A_Adj': A_Adj,
      'Adj_Falt': Adj_Falt,
      'A_Ter': A_Ter,
      'A_Parte': A_Parte,
      'A_Dup': A_Dup,

      'Score': score,
      'Semaf': semaf
    };

    const row=headers.map(h=> (h in src)?src[h]:'' );
    sh.appendRow(row);

    try{
      const toWrap=['Adj','Resumen','Relato'];
      toWrap.forEach(function(name){
        const c = headers.indexOf(name)+1;
        if(c>0) sh.getRange(1,c,sh.getLastRow(),1).setWrap(true);
      });
    }catch(errWrap){}

    return ContentService
      .createTextOutput(JSON.stringify({
        ok:true,
        code:'OK',
        score:score,
        semaforo:semaf,
        canal:canalAsignado
      }))
      .setMimeType(ContentService.MimeType.JSON);

  }catch(err){
    return ContentService
      .createTextOutput(JSON.stringify({
        ok:false,
        code:'FATAL',
        message:'Error fatal del servidor: '+String(err)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
