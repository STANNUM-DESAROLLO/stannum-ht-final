// ============================================================
// STANNUM · IA SUMMIT TUCUMÁN 2026 — Test de Dominio IA
// Backend Apps Script: guarda en Google Sheet + envía email con resultado
//
// DEPLOY:
//   1. Abrí el Google Sheet → Extensiones → Apps Script.
//   2. Pegá TODO este archivo.
//   3. Guardá. Ejecutá una vez la función setupHeaders() (autoriza permisos).
//   4. Implementar → Nueva implementación → Tipo: Aplicación web.
//      - Ejecutar como: Yo (comercial@stannum.com.ar)
//      - Quién tiene acceso: Cualquier usuario
//   5. Copiá la URL del Web App y pegala en test.html → CONFIG.APPS_SCRIPT_URL
// ============================================================

const SHEET_ID        = "1-8tR8vnWFVHdRsnhsq-q8Mv1NFcPVO6KFdNYEBNgP8Y";
const TEAM_EMAIL      = "comercial@stannum.com.ar";
const DASHBOARD_TOKEN = "stannum2026";          // para el dashboard (doGet)
const WHATSAPP        = "5493814763535";
const EVENT           = "IA Summit Tucumán 2026";
const EVENT_DATE      = "17 de junio de 2026 · Hilton Tucumán · Acreditación 17:30 hs";
const NOTIFY_TEAM     = true;                   // copia interna al equipo por cada lead

const LEVELS = [
  {n:1, name:"Principiante", tag:"Etapa inicial de exploración",
   desc:"Estás en las primeras etapas del uso de IA generativa. La oportunidad es construir una base práctica con casos de uso reales que generen impacto inmediato en tu operación.",
   sol:"TRENNO iA Starter", solD:"Programa de alfabetización ejecutiva con casos de uso aplicados, prompts prácticos y rutinas de adopción para líderes y equipos."},
  {n:2, name:"Intermedio", tag:"Uso regular con potencial de escala",
   desc:"Usás IA de forma regular para tareas específicas. El siguiente paso es sistematizar ese uso, ampliar herramientas y comenzar a medir el impacto concreto en el negocio.",
   sol:"TRENNO iA Starter", solD:"Capacitación aplicada por área con metodología de prompts avanzados, automatización básica y métricas de adopción."},
  {n:3, name:"Avanzado", tag:"Optimización y diseño de sistemas",
   desc:"Optimizás tareas con IA y diseñás prompts personalizados. El siguiente nivel es integrar IA como sistema operativo del equipo con flujos automatizados y gobierno claro.",
   sol:"TRENNO iA Enterprise", solD:"Implementación sistémica con SOPs, automatizaciones conectadas, entrenamiento de equipos y tablero de control de impacto."},
  {n:4, name:"Experto", tag:"IA como ventaja estratégica",
   desc:"Usás IA para automatizar procesos y optimizar estrategias. El foco es el gobierno, la escalabilidad organizacional y el roadmap de IA como ventaja competitiva sostenible.",
   sol:"TRENNO iA Agentic", solD:"Diseño de agentes y workflows autónomos, gobierno de IA y roadmap estratégico para escalar la organización."}
];

const HEADERS = ["Fecha","Nombre","Empresa","Rol","Email","WhatsApp","Equipo","Distribución tiempo",
  "Plataformas","Paga IA","Áreas","Reacción IA","Integración","Funciones","Conceptos","Quiere aprender",
  "Frecuencia","Auto-nivel","Score","Nivel N","Nivel","Solución","ID"];

// ------------------------------------------------------------
function sheet_(){ return SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; }

function setupHeaders(){
  const sh = sheet_();
  sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS])
    .setFontWeight("bold").setBackground("#050606").setFontColor("#32CBBF");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, HEADERS.length);
}

function levelByN_(n){ for (var i=0;i<LEVELS.length;i++){ if(LEVELS[i].n==n) return LEVELS[i]; } return LEVELS[0]; }
function arr_(x){ return Array.isArray(x) ? x.join(", ") : (x||""); }
// Evita que Sheets interprete como fórmula valores que empiezan con + = - @ (ej: teléfonos)
function txt_(x){ x = (x==null?"":String(x)); return /^[\+\=\-@]/.test(x) ? ("'"+x) : x; }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ------------------------------------------------------------
// Recibe el envío del test (POST desde test.html)
function doPost(e){
  try{
    const d = JSON.parse(e.postData.contents);
    appendRow_(d);
    if (d.email) sendResultEmail_(d);
    if (NOTIFY_TEAM) notifyTeam_(d);
    return json_({status:"ok"});
  }catch(err){
    return json_({status:"error", message:String(err)});
  }
}

function appendRow_(d){
  const sh = sheet_();
  if (sh.getLastRow() === 0) setupHeaders();
  sh.appendRow([
    d.date ? new Date(d.date) : new Date(),
    d.name||"", d.company||"", d.role||"", d.email||"", txt_(d.phone),
    d.team||"", d.time||"", arr_(d.plats), d.pago||"", arr_(d.areas),
    d.reaccion||"", d.integracion||"", arr_(d.funciones), arr_(d.conceptos), arr_(d.aprender),
    d.freq||"", d.autonivel||"", d.score||0, d.levelN||"", d.levelName||"", d.solution||"", d.id||""
  ]);
}

// ------------------------------------------------------------
// Email con el resultado al lead
function sendResultEmail_(d){
  const lv = levelByN_(d.levelN);
  const first = (d.name||"").split(/\s+/)[0] || "";
  const waMsg = encodeURIComponent("Hola STANNUM, soy "+(d.name||"")+" ("+(d.company||"")+"). Hice el Test de Dominio IA, nivel "+(d.levelName||"")+" ("+(d.score||0)+"/100). Quiero reservar mi lugar en el "+EVENT+".");
  const wa = "https://wa.me/"+WHATSAPP+"?text="+waMsg;
  const subject = "Tu diagnóstico de Dominio IA — Nivel " + (d.levelName||"");

  const html =
  '<div style="background:#050606;padding:32px 0;font-family:Arial,Helvetica,sans-serif">'+
    '<div style="max-width:600px;margin:0 auto;background:#0b0d0d;border:1px solid rgba(50,203,191,.18);border-radius:16px;overflow:hidden">'+
      '<div style="padding:28px 32px;border-bottom:1px solid rgba(50,203,191,.15)">'+
        '<span style="color:#fff;font-size:18px;font-weight:800;letter-spacing:2px">STAN<span style="color:#32CBBF">NUM</span></span>'+
        '<div style="color:#A0A0A0;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin-top:4px">Diagnóstico Ejecutivo · '+EVENT+'</div>'+
      '</div>'+
      '<div style="padding:36px 32px">'+
        '<p style="color:#A0A0A0;font-size:15px;margin:0 0 8px">Hola '+first+',</p>'+
        '<p style="color:#fff;font-size:16px;line-height:1.6;margin:0 0 28px">Este es tu diagnóstico del <b>Test de Dominio de IA</b>:</p>'+
        '<div style="text-align:center;margin:0 0 28px">'+
          '<div style="color:#32CBBF;font-size:56px;font-weight:800;line-height:1">'+(d.score||0)+'<span style="color:#555;font-size:22px">/100</span></div>'+
          '<div style="color:#fff;font-size:28px;font-weight:800;margin-top:6px">Nivel '+(d.levelName||"")+'</div>'+
          '<div style="color:#32CBBF;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin-top:4px">'+lv.tag+'</div>'+
        '</div>'+
        '<div style="background:rgba(50,203,191,.04);border:1px solid rgba(50,203,191,.15);border-radius:12px;padding:22px;margin-bottom:16px">'+
          '<div style="color:#A0A0A0;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin-bottom:10px">Lectura ejecutiva</div>'+
          '<div style="color:#d6d6d6;font-size:15px;line-height:1.65">'+lv.desc+'</div>'+
        '</div>'+
        '<div style="background:rgba(50,203,191,.04);border:1px solid rgba(50,203,191,.15);border-radius:12px;padding:22px;margin-bottom:28px">'+
          '<div style="color:#A0A0A0;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin-bottom:10px">Solución STANNUM recomendada</div>'+
          '<div style="color:#32CBBF;font-size:20px;font-weight:800;margin-bottom:6px">'+lv.sol+'</div>'+
          '<div style="color:#A0A0A0;font-size:14px;line-height:1.6">'+lv.solD+'</div>'+
        '</div>'+
        '<div style="text-align:center;background:rgba(50,203,191,.06);border-radius:14px;padding:28px 22px">'+
          '<div style="color:#fff;font-size:20px;font-weight:800;margin-bottom:8px">Llevá tu diagnóstico al '+EVENT+'</div>'+
          '<div style="color:#A0A0A0;font-size:14px;line-height:1.6;margin-bottom:20px">'+EVENT_DATE+'. Reservá tu lugar y trabajá tu nivel en vivo con el equipo de STANNUM. Cupos limitados.</div>'+
          '<a href="'+wa+'" style="display:inline-block;background:#32CBBF;color:#000;font-weight:bold;font-size:15px;text-decoration:none;padding:15px 34px;border-radius:999px">Reservar mi lugar por WhatsApp →</a>'+
        '</div>'+
      '</div>'+
      '<div style="padding:20px 32px;border-top:1px solid rgba(255,255,255,.06);color:#666;font-size:11px;text-align:center">'+
        'STANNUM — Alto Rendimiento Argentina · '+TEAM_EMAIL+
      '</div>'+
    '</div>'+
  '</div>';

  GmailApp.sendEmail(d.email, subject, "Tu resultado del Test de Dominio IA de STANNUM. Nivel "+(d.levelName||"")+" — "+(d.score||0)+"/100.", {
    htmlBody: html, name: "STANNUM", replyTo: TEAM_EMAIL
  });
}

// Aviso interno al equipo (lead nuevo)
function notifyTeam_(d){
  const body = "Nuevo lead Test de Dominio IA · "+EVENT+"\n\n"+
    "Nombre: "+(d.name||"")+"\nEmpresa: "+(d.company||"")+"\nRol: "+(d.role||"")+"\n"+
    "Email: "+(d.email||"")+"\nWhatsApp: "+(d.phone||"")+"\n"+
    "Nivel: "+(d.levelName||"")+" ("+(d.score||0)+"/100)\nSolución sugerida: "+(d.solution||"")+"\n";
  MailApp.sendEmail(TEAM_EMAIL, "🔥 Lead Test IA — "+(d.name||"")+" ("+(d.levelName||"")+")", body);
}

// ------------------------------------------------------------
// Dashboard: devuelve todas las respuestas en JSON (protegido por token)
function doGet(e){
  const token = e && e.parameter && e.parameter.token;
  if (token !== DASHBOARD_TOKEN) return json_({status:"error", message:"Unauthorized"});
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last <= 1) return json_({status:"ok", total:0, data:[]});
  const rows = sh.getRange(2,1,last-1,HEADERS.length).getValues();
  const data = rows.map(function(r){
    return {
      fecha:r[0]?new Date(r[0]).toISOString():"", nombre:r[1], empresa:r[2], rol:r[3], email:r[4], whatsapp:r[5],
      equipo:r[6], tiempo:r[7], plataformas:r[8], paga:r[9], areas:r[10], reaccion:r[11], integracion:r[12],
      funciones:r[13], conceptos:r[14], aprender:r[15], frecuencia:r[16], autonivel:r[17],
      score:r[18], nivelN:r[19], nivel:r[20], solucion:r[21], id:r[22]
    };
  });
  return json_({status:"ok", total:data.length, data:data});
}
