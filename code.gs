/******************** UTIL ********************/
// Spreadsheet & sheet helpers
function S(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(n){ return S().getSheetByName(n); }
function ensureSheet_(name, headers){
  let s = sh(name);
  if (!s) s = S().insertSheet(name);
  if (headers && headers.length){
    const lr = s.getLastRow();
    if (lr === 0) s.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return s;
}

/** Normalize a header into a safe key (header-agnostic reader) */
function keyify_(h){
  return String(h||"")
    .trim()
    .toLowerCase()
    .replace(/[\s\/-]+/g,"_")
    .replace(/[^\w]/g,"_")
    .replace(/_+/g,"_")
    .replace(/^_|_$/g,"");
}

/** Map of common header synonyms -> canonical keys we use in code */
const HEADER_ALIASES_ = {
  // universal ids
  "eventid":"event_id", "event":"event_id", "id":"event_id",
  // events meta
  "entrants":"entrants_csv", "players":"entrants_csv",
  "datetime":"date", "start":"date", "start_time":"date",
  // lines
  "american":"american_odds", "odds":"american_odds", "price":"american_odds",
  "param_line":"param","line":"param","handicap":"param","spread":"param",
  "ml":"market", "moneyline":"market",
  // maps
  "map":"map_id", "mapname":"name",
  // users/bets
  "user":"user_id","player":"player_id"
};

/** Canonicalize a key after keyify_ using aliases above */
function canonKey_(k){
  if (!k) return k;
  if (HEADER_ALIASES_[k]) return HEADER_ALIASES_[k];
  return k;
}

/** Load an entire sheet to objects with header-insensitive keys */
function loadTable_(name){
  const s = sh(name);
  if (!s || s.getLastRow() === 0) return {hdr:[], rows:[]};
  const vals = s.getDataRange().getValues();
  if (!vals.length) return {hdr:[], rows:[]};

  // Build header map
  const rawHdr = (vals[0]||[]).map(h=>String(h||""));
  const normHdr = rawHdr.map(h=>canonKey_(keyify_((h))));
  // Deduplicate columns by keeping the first occurrence
  const seen = {};
  const usedIdx = [];
  const finalHdr = [];
  normHdr.forEach((k,i)=>{
    if (!k) return;
    if (seen[k]) return;
    seen[k]=true;
    usedIdx.push(i);
    finalHdr.push(k);
  });

  // Rows
  const rows=[];
  for (let r=1;r<vals.length;r++){
    const ar = vals[r]||[];
    // skip truly empty rows
    if (ar.every(v => v === "")) continue;
    const obj = {};
    for (let j=0;j<usedIdx.length;j++){
      const col = usedIdx[j];
      obj[finalHdr[j]] = ar[col];
    }
    rows.push(obj);
  }
  return {hdr:finalHdr, rows};
}

function parseCSV_(s){
  return String(s||"")
    .split(/[;,|\n]/)     // commas, semicolons, pipes, or line breaks
    .map(x=>x.trim())
    .filter(Boolean);
}

function now_(){ return new Date(); }
function randn_(){ let u=0,v=0; while(u===0)u=Math.random(); while(v===0)v=Math.random(); return Math.sqrt(-2*Math.log(u))*Math.cos(2*Math.PI*v); }
function expWeight_(ageDays, halfLife){ return Math.exp(-Math.LN2 * ageDays / halfLife); }
function applyOverround_(probs, hold){ const sum=Math.max(1e-9,probs.reduce((a,b)=>a+b,0)); const target=1+hold; const scale=target/sum; return probs.map(p=>p*scale); }
function toAmerican_(p){ p=Math.min(Math.max(+p||1e-6,1e-6),1-1e-6); return (p>=0.5)? Math.round(-100*p/(1-p)) : Math.round(100*(1-p)/p); }
function amToDecimal_(am){ const a=Number(am); return a>=0 ? 1 + a/100 : 1 + 100/Math.abs(a); }
function decToAmerican_(d){ if (d>=2) return Math.round((d-1)*100); return Math.round(-100/(d-1)); }

/** Add: signed format helper for spread label text */
function fmtSigned_(x){
  const n = Number(x);
  if (!isFinite(n)) return String(x);
  return (n>0?`+${n.toFixed(1)}`:n.toFixed(1));
}

/******************** LOGGING ********************/
function logError_(where, msg){
  try{
    const s = ensureSheet_("Logs", ["timestamp","where","message"]);
    s.appendRow([new Date(), where, String(msg)]);
  } catch(e){}
}

/******************** ONE-TIME SETUP / MIGRATION ********************/
function setupOnce(){
  ensureSheet_("Settings", ["key","value"]);
  ensureSheet_("Players", ["player_id","name","status"]);
  ensureSheet_("Maps", ["map_id","name","tee","par"]);
  ensureSheet_("Preferences", ["player_id","map_id","affinity"]); // -1,0,1
  ensureSheet_("Scores", ["date","series","event_id","map_id","tee","player_id","score"]);
  ensureSheet_("Events", ["event_id","date","series","mode","map_id","tee","entrants_csv"]);
  ensureSheet_("Model_Output", ["event_id","player_id","mu","sigma","notes"]);
  ensureSheet_("Lines", ["event_id","market","selection","fair_prob","vig_prob","american_odds","param","last_updated"]);
  ensureSheet_("Bets", ["timestamp","user_id","event_id","market","selection","param","odds","stake","potential_payout","status","settled_payout","settled_ts","is_parlay","legs_json"]);
  ensureSheet_("Users", ["user_id","display","coins_balance"]);
  ensureSheet_("Logs", ["timestamp","where","message"]);
  SpreadsheetApp.getUi().alert("Setup complete.");
}

/******************** SETTINGS ********************/
function getSettings_(){
  const table = loadTable_("Settings").rows;
  const m = {}; for (const r of table){
    const k = String(r.key!=null ? r.key : r.setting!=null ? r.setting : r.name!=null ? r.name : "").trim();
    const v = r.value!=null ? r.value : r.val!=null ? r.val : r.setting_value!=null ? r.setting_value : "";
    if (k) m[keyify_(k)] = v;
  }
  return {
    half_life_days: Number(m.half_life_days||60),
    N0: Number(m.n0||m.N0||8),
    theta_player: Number(m.theta_player||0.5),
    theta_map: Number(m.theta_map||0.3),
    hold_multi: Number(m.hold_multi||0.10),
    hold_ou: Number(m.hold_ou||0.045),
    simulations: Number(m.simulations||200),
    min_sd_strokes: Number(m.min_sd_strokes||0.6),
    ou_snap: String(m.ou_snap||"half").toLowerCase(),
    starting_coins: Number(m.starting_coins||500),
    affinity_use: String(m.affinity_use||"on").toLowerCase()==="on",
    affinity_stroke_like: Number(m.affinity_stroke_like||-0.20),
    affinity_stroke_dislike: Number(m.affinity_stroke_dislike||0.15),
    affinity_cap_sd: Number(m.affinity_cap_sd||0.30),
    // spread
    hold_spread: (m.hold_spread!==undefined ? Number(m.hold_spread) : undefined),
    spread_sigma_override: (m.spread_sigma_override!==undefined && m.spread_sigma_override!=="") ? Number(m.spread_sigma_override) : null,
  };
}

/******************** AFFINITY ********************/
function strokeAffinityShift_(aff, sd, st){
  if (!st.affinity_use) return 0;
  let base = 0;
  if (String(aff)==="1") base = st.affinity_stroke_like;
  else if (String(aff)==="-1") base = st.affinity_stroke_dislike;
  const cap = st.affinity_cap_sd * (sd || 1);
  return Math.max(-cap, Math.min(cap, base));
}

/******************** NORMAL CDF + INV ********************/
function erf_(x){
  const a1=0.254829592,a2=-0.284496736,a3=1.421413741,a4=-1.453152027,a5=1.061405429,p=0.3275911;
  const sign = x<0?-1:1; x=Math.abs(x);
  const t = 1/(1+p*x);
  const y = 1-((((a5*t+a4)*t+a3)*t+a2)*t+a1)*t*Math.exp(-x*x);
  return sign*y;
}
function normCdf_(z){ return 0.5*(1+erf_(z/Math.SQRT2)); }
function normInv_(p){
  if (!(p>0 && p<1)) throw new Error("normInv_ p out of (0,1)");
  const a=[-39.69683028665376,220.9460984245205,-275.9285104469687,138.3577518672690,-30.66479806614716,2.506628277459239];
  const b=[-54.47609879822406,161.5858368580409,-155.6989798598866,66.80131188771972,-13.28068155288572];
  const c=[-0.007784894002430293,-0.3223964580411365,-2.400758277161838,-2.549732539343734,4.374664141464968,2.938163982698783];
  const d=[0.007784695709041462,0.3224671290700398,2.445134137142996,3.754408661907416];
  const plow=0.02425, phigh=1-plow;
  let q,r;
  if (p<plow){
    q=Math.sqrt(-2*Math.log(p));
    return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5])/((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
  }
  if (p>phigh){
    q=Math.sqrt(-2*Math.log(1-p));
    return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5])/((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1);
  }
  q=p-0.5; r=q*q;
  return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q/(((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+1));
}

/******************** STATS (RECENCY + EB) ********************/
/** Uses EVERY row in Scores (no series filter) */
function buildStats_(st){
  const today = new Date();
  const prefs      = loadTable_("Preferences").rows;
  const allScores  = loadTable_("Scores").rows;   // <-- no filtering by series

  // Affinity lookup: { [player_id]: { [map_id]: -1|0|1 } }
  const affinity = {};
  for (const r of prefs){
    const p = String(r.player_id||"").trim();
    const m = String(r.map_id||"").trim();
    if (!p || !m) continue;
    (affinity[p]||(affinity[p]={}))[m] = Number(r.affinity||0);
  }

  // Use ALL scores
  const scores = allScores;

  // Weighted aggregates
  const W = {}, P = {}, M = {};
  let gSum=0, gW=0;

  for (const r of scores){
    const p = String(r.player_id||"").trim();
    const m = String(r.map_id||"").trim();
    const score = Number(r.score);
    if (!p || !m || !isFinite(score)) continue;

    const d = r.date ? new Date(r.date) : today;
    const age = Math.max(0, (today - d)/86400000);
    const w = expWeight_(age, st.half_life_days);

    const key = p+"|"+m;
    if (!W[key]) W[key]={wSum:0,xwSum:0,x2wSum:0,n:0};
    W[key].wSum+=w; W[key].xwSum+=w*score; W[key].x2wSum+=w*score*score; W[key].n+=w;

    if (!P[p]) P[p]={wSum:0,xwSum:0,x2wSum:0,n:0};
    P[p].wSum+=w; P[p].xwSum+=w*score; P[p].x2wSum+=w*score*score; P[p].n+=w;

    if (!M[m]) M[m]={wSum:0,xwSum:0,x2wSum:0,n:0};
    M[m].wSum+=w; M[m].xwSum+=w*score; M[m].x2wSum+=w*score*score; M[m].n+=w;

    gSum+=w*score; gW+=w;
  }

  const gMean = gW>0 ? gSum/gW : 0;

  function meanVar_(obj){
    const out={};
    for (const k in obj){
      const o=obj[k];
      const mu = o.wSum>0 ? o.xwSum/o.wSum : gMean;
      const varW = o.wSum>0 ? Math.max(1e-6, o.x2wSum/o.wSum - mu*mu) : Math.pow(st.min_sd_strokes,2);
      out[k]={mu, sd: Math.max(st.min_sd_strokes, Math.sqrt(varW)), n:o.n};
    }
    return out;
  }

  const PM = meanVar_(W);  // player+map
  const PP = meanVar_(P);  // player overall
  const MM = meanVar_(M);  // map overall

  function ebStroke_(p,m){
    const key = p+"|"+m;
    const base = PM[key] || {mu:gMean, sd:st.min_sd_strokes, n:0};
    const np = (PP[p]&&PP[p].mu!=null)?PP[p].mu:gMean;
    const nm = (MM[m]&&MM[m].mu!=null)?MM[m].mu:gMean;

    const alpha = base.n / (base.n + Number(st.N0));
    const backstop = st.theta_player*np + st.theta_map*nm + (1 - st.theta_player - st.theta_map)*gMean;

    const mu = alpha*base.mu + (1-alpha)*backstop;

    const sdPlayer = (PP[p]?.sd)||st.min_sd_strokes;
    const sdMap    = (MM[m]?.sd)||st.min_sd_strokes;
    const sdBack   = st.theta_player*sdPlayer + st.theta_map*sdMap + (1-st.theta_player-st.theta_map)*((sdPlayer+sdMap)/2);
    const sd = Math.max(st.min_sd_strokes, alpha*base.sd + (1-alpha)*sdBack);

    return {mu, sd};
  }

  return {ebStroke_, affinity};
}


/******************** TOUR PRICING (ML + O/U + SPREAD) ********************/
function snapToHalf_(x){
  let y = Math.round(x*2)/2;
  if (Math.abs(y - Math.round(y)) < 1e-9) y += 0.5*(y>=x?1:-1);
  return y;
}
function pricePlayerOU_(muAdj, sd, st){
  const L = (String(st.ou_snap||"half")==="half") ? snapToHalf_(muAdj) : muAdj;
  const z = (L - muAdj)/sd;
  const pU = normCdf_(z), pO = 1 - pU;
  const vig = applyOverround_([pU, pO], Number(st.hold_ou||0.045));
  return {line:L, pUnderFair:pU, pOverFair:pO, pUnderVig:vig[0], pOverVig:vig[1]};
}
function simulateTourML_(entrants, map_id, ebStroke, affinity, st){
  const sims = Math.max(1, Number(st.simulations||200));
  const wins = Object.fromEntries(entrants.map(p=>[p,0]));
  for (let s=0;s<sims;s++){
    let best=null, bestP=null;
    for (const p of entrants){
      const est = ebStroke(p, map_id);
      const aff = ((affinity[p]||{})[map_id]||0);
      const muAdj = est.mu + strokeAffinityShift_(aff, est.sd, st);
      const draw = muAdj + est.sd*randn_();
      if (best===null || draw < best){ best=draw; bestP=p; }
    }
    wins[bestP] += 1;
  }
  return entrants.map(p => (wins[p]||0)/sims);
}
function makeSpreadsFromWinProb_(pA, sigma, spreadsList, st){
  const eps = 1e-6;
  pA = Math.min(1-eps, Math.max(eps, Number(pA)||0));
  sigma = Math.max(1e-6, Number(sigma)||1);
  const mu = -sigma * normInv_(pA);
  const fair = snapToHalf_(mu);

  const want = [];
  const seen = {};
  function add(x){ const k=String(Math.round(x*10)/10); if(!seen[k]){ seen[k]=1; want.push(x); } }
  add(fair);
  (spreadsList||[]).forEach(add);

  const hold = Number(st.hold_spread ?? st.hold_ou ?? 0.045);
  const rows = [];

  for (const s of want){
    const qAcover = normCdf_((s - mu)/sigma);
    const qBcover = 1 - qAcover;
    const vig = applyOverround_([qAcover, qBcover], hold);
    rows.push({
      spread: s,
      A_cover_fair: qAcover,
      B_cover_fair: qBcover,
      A_cover_vig:  vig[0],
      B_cover_vig:  vig[1],
      A_odds: toAmerican_(vig[0]),
      B_odds: toAmerican_(vig[1])
    });
  }
  return { fairSpread: fair, mu, sigma, rows };
}

/******************** PRICING ENTRYPOINTS ********************/
function computeAndPriceAllEventsCore_({from='menu'} = {}) {
  const st = getSettings_();
  const {ebStroke_, affinity} = buildStats_(st);

  const eventsTbl = loadTable_("Events");
  const linesSh = ensureSheet_("Lines", ["event_id","market","selection","fair_prob","vig_prob","american_odds","param","last_updated"]);

  const out = [];
  let processed = 0;
  const start = Date.now();
  const norm = v => String(v==null?"":v).trim();

  for (const e of eventsTbl.rows){
    if ((Date.now()-start) > 5.5*60*1000){ break; }

    const event_id = norm(_g(e,"event_id","id"));
    const seriesLC = (norm(_g(e,"series")) || "tour").toLowerCase();
    const modeLC   = (norm(_g(e,"mode"))   || "stroke").toLowerCase();
    const map_id   = norm(_g(e,"map_id","map"));
    const entrants = parseCSV_( _g(e,"entrants_csv","entrants","players") );

    if (!event_id || seriesLC!=="tour" || modeLC!=="stroke" || !map_id || entrants.length<2) continue;

    // ML
    const probsML = simulateTourML_(entrants, map_id, ebStroke_, affinity, st);
    const vigML   = applyOverround_(probsML, Number(st.hold_multi||0.10));
    for (let i=0;i<entrants.length;i++){
      out.push([event_id,"ML",entrants[i], probsML[i], vigML[i], toAmerican_(vigML[i]), "", new Date()]);
    }

    // Player O/U
    for (const p of entrants){
      const est = ebStroke_(p, map_id);
      const aff = ((affinity[p]||{})[map_id]||0);
      const muAdj = est.mu + strokeAffinityShift_(aff, est.sd, st);
      const sd = est.sd;
      const ou = pricePlayerOU_(muAdj, sd, st);
      out.push([event_id,"OU",`${p} UNDER`, ou.pUnderFair, ou.pUnderVig, toAmerican_(ou.pUnderVig), String(ou.line), new Date()]);
      out.push([event_id,"OU",`${p} OVER`,  ou.pOverFair,  ou.pOverVig,  toAmerican_(ou.pOverVig),  String(ou.line), new Date()]);
    }

    // H2H spreads (ensure selection has explicit sign & param is numeric)
    if (entrants.length === 2){
      const [A, B] = entrants;
      const pA = Number(probsML[0]);

      const estA = ebStroke_(A, map_id);
      const estB = ebStroke_(B, map_id);
      const sigmaLearned = Math.sqrt(estA.sd*estA.sd + estB.sd*estB.sd);
      const sigma = (st.spread_sigma_override && isFinite(st.spread_sigma_override) && st.spread_sigma_override>0)
                    ? st.spread_sigma_override : sigmaLearned;

      const pack = makeSpreadsFromWinProb_(pA, sigma, [], st);
      const s = Number(pack.fairSpread);
      const row = pack.rows.find(r => Math.abs(r.spread - s) < 1e-9) || pack.rows[0];

      out.push([
        event_id,"SPREAD",`${A} ${fmtSigned_(s)}`,
        row.A_cover_fair,row.A_cover_vig,row.A_odds,
        s,new Date()
      ]);
      out.push([
        event_id,"SPREAD",`${B} ${fmtSigned_(-s)}`,
        row.B_cover_fair,row.B_cover_vig,row.B_odds,
        -s,new Date()
      ]);
    }

    processed++;
  }

  if (out.length){
    const lr = linesSh.getLastRow(); if (lr>1) linesSh.getRange(2,1,lr-1,8).clearContent();
    linesSh.getRange(2,1,out.length,8).setValues(out);
  } else {
    logError_("computeAndPriceAllEventsCore_", "No Tour events priced — check Events sheet (series='tour', mode='stroke', 2+ entrants).");
  }

  return {events_scanned: eventsTbl.rows.length, events_priced: processed, lines_written: out.length, context: from};
}
function computeAndPriceAllEvents(){
  try{
    const res = computeAndPriceAllEventsCore_({from:'menu'});
    SpreadsheetApp.getActive().toast(`Priced ${res.events_priced} event(s).`, "Sportsbook", 5);
  } catch(err){
    logError_("computeAndPriceAllEvents(menu)", err && err.stack ? err.stack : err);
    SpreadsheetApp.getUi().alert("Error: "+err);
  }
}
function repriceNow(){
  try{
    const res = computeAndPriceAllEventsCore_({from:'web'});
    return { ok:true, ...res };
  } catch (e){
    logError_("repriceNow", e && e.stack ? e.stack : e);
    return { ok:false, error:String(e) };
  }
}

/******************** DATA API FOR UI ********************/
function _g(obj /*, k1, k2, ...*/){
  for (let i=1;i<arguments.length;i++){
    const k = arguments[i];
    if (obj[k]!==undefined && obj[k]!==null && obj[k]!=="" ) return obj[k];
  }
  return "";
}

function getOddsBoard(){
  try{
    const norm = v => String(v==null?"":v).trim();
    const maps = loadTable_("Maps").rows.reduce((a,r)=>{
      const id = norm(_g(r,"map_id","id","map"));
      if (id) a[id]={ name: norm(_g(r,"name","map_name")), tee: norm(_g(r,"tee")) };
      return a;
    },{});
    const eventsTab = loadTable_("Events").rows;
    const findEvent = (id)=> eventsTab.find(e=>norm(_g(e,"event_id","id"))===id);

    const lines = loadTable_("Lines").rows;
    const events = {};

    for (const L of lines){
      const id = norm(_g(L,"event_id","id","event"));
      const market = norm(_g(L,"market","type")).toUpperCase();
      const selection = norm(_g(L,"selection","team","player","pick"));
      const param = norm(_g(L,"param","line","handicap","spread"));
      const amOdds = Number(_g(L,"american_odds","odds","price"));
      const vigProb = _g(L,"vig_prob","prob","prob_vig");
      const fairProb = _g(L,"fair_prob","prob_fair");
      if (!id || !market || !selection) continue;

      if (!events[id]){
        const evMatch = findEvent(id);
        let mapName="", label=`Event ${id}`, series="Tour", mode="stroke", map_id="", tee="";
        let dateVal = evMatch ? _g(evMatch,"date","start","datetime") : "";
        if (evMatch){
          map_id = norm(_g(evMatch,"map_id","map"));
          const map = maps[map_id] || {};
          mapName = norm(map.name) || map_id || "";
          tee = norm(map.tee) || norm(_g(evMatch,"tee")) || "";
          if (mapName) label = tee ? `${mapName} (${tee})` : mapName;
          series = norm(_g(evMatch,"series"))||series;
          mode   = norm(_g(evMatch,"mode"))||mode;
        }
        if (dateVal && Object.prototype.toString.call(dateVal)==="[object Date]") {
          dateVal = dateVal.toISOString();
        }
        events[id] = {
          event_id:id, date:dateVal,
          series, mode, map_id, map_name:mapName, tee,
          map_label:label, markets:[]
        };
      }

      events[id].markets.push({
        market,
        selection,
        param,
        american_odds: isFinite(amOdds) ? amOdds : 0,
        vig_prob: (vigProb===""||vigProb==null)?null:Number(vigProb),
        fair_prob: (fairProb===""||fairProb==null)?null:Number(fairProb)
      });
    }

    const out = Object.values(events).map(e=>{
      e.markets.sort((a,b)=>{
        if (a.market!==b.market){
          if (a.market==="ML") return -1;
          if (b.market==="ML") return 1;
          if (a.market==="SPREAD" && b.market==="OU") return -1;
          if (a.market==="OU" && b.market==="SPREAD") return 1;
          return a.market<b.market?-1:1;
        }
        if (a.selection!==b.selection) return a.selection<b.selection?-1:1;
        return Math.abs(a.american_odds)-Math.abs(b.american_odds);
      });
      return e;
    });
    return out;
  } catch (e){
    logError_("getOddsBoard", e && e.stack ? e.stack : e);
    return [];
  }
}

/** Optional small debug endpoint the UI uses when nothing renders */
function getOddsBoardDebug(){
  const L = loadTable_("Lines").rows;
  const c = L.reduce((m,r)=>{
    const mk = String(r.market||'').toUpperCase().trim();
    m[mk]=(m[mk]||0)+1; return m;
  },{});
  return {lines_count:L.length, by_market:c, sample:L.slice(0,5)};
}

/******************** BETS ********************/
function getWallet(user_id){
  const st = getSettings_();
  const usersSh = ensureSheet_("Users", ["user_id","display","coins_balance"]);
  const vals = usersSh.getDataRange().getValues();
  let row = -1;
  for (let i=1;i<vals.length;i++){
    if (String(vals[i][0])===String(user_id)){ row=i; break; }
  }
  if (row===-1){
    usersSh.appendRow([user_id, user_id, st.starting_coins]);
    return {user_id, display:user_id, coins_balance: st.starting_coins};
  }
  return {user_id, display: vals[row][1], coins_balance: Number(vals[row][2]||0)};
}

/** Robust: trims & lowercases both sides so stray spaces/casing don't hide rows */
function listMyBets(user_id){
  const norm = s => String(s||"").trim().toLowerCase();
  const me = norm(user_id);
  const bets = loadTable_("Bets").rows.filter(b=> norm(b.user_id) === me );
  bets.sort((a,b)=> new Date(b.timestamp) - new Date(a.timestamp));
  return bets;
}

function placeBet(payload){
  try{
    const usersSh = ensureSheet_("Users", ["user_id","display","coins_balance"]);
    const betsSh  = ensureSheet_("Bets", ["timestamp","user_id","event_id","market","selection","param","odds","stake","potential_payout","status","settled_payout","settled_ts","is_parlay","legs_json"]);

    const vals = usersSh.getDataRange().getValues();
    let row = -1; for (let i=1;i<vals.length;i++){ if (String(vals[i][0])===String(payload.user_id)){ row=i; break; } }
    if (row===-1) throw new Error("User not found; reload app.");

    const balance = Number(vals[row][2]||0);
    const stake = Number(payload.stake||0);
    if (!(stake>0)) throw new Error("Stake must be > 0.");
    if (stake > balance) throw new Error("Insufficient coins.");

    let market="SINGLE", event_id="", selection="", param="", amOdds=0, legsJson="", isParlay=false, potential=0;

    if (payload.type==='parlay'){
      isParlay = true;
      const legs = payload.legs||[];
      if (legs.length<2) throw new Error("Parlay needs 2+ legs.");
      const seen={};
      for (const lg of legs){ if (seen[lg.event_id]) throw new Error("Parlay cannot have multiple legs from the same event."); seen[lg.event_id]=true; }
      const dec = legs.map(l=>amToDecimal_(l.odds)).reduce((a,b)=>a*b,1);
      amOdds = decToAmerican_(dec);
      market = "PARLAY";
      selection = legs.map(l=>`${l.event_id}:${l.market}:${l.selection}${l.param?`@${l.param}`:''}`).join(" | ");
      legsJson = JSON.stringify(legs);
      potential = stake * dec;
    } else {
      const s = payload.selection;
      if (!s) throw new Error("Missing selection.");
      market = s.market; event_id = s.event_id; selection = s.selection; param = s.param||""; amOdds = Number(s.odds);
      const dec = amToDecimal_(amOdds);
      potential = stake * dec;
    }

    // status starts "open" (shows as PENDING in UI)
    betsSh.appendRow([new Date(), payload.user_id, event_id, market, selection, param, amOdds, stake, potential, "open", "", "", isParlay, legsJson]);

    usersSh.getRange(row+1,3).setValue(balance - stake);
    return {ok:true, new_balance: balance - stake};
  } catch(e){
    logError_("placeBet", e && e.stack ? e.stack : e);
    throw e;
  }
}

/******************** SETTLEMENT ********************/
function settleEventPrompt(){
  const ui = SpreadsheetApp.getUi();
  const id = ui.prompt("Settle which event_id?", "Enter event_id exactly (e.g., E123).", ui.ButtonSet.OK_CANCEL).getResponseText();
  if (id) { const res = settleEvent(String(id).trim()); ui.alert(`Settled ${res.settled} bet(s).`); }
}
function settleAllWithScores(){
  const evs = loadTable_("Events").rows;
  let total=0;
  for (const e of evs){
    const id = String(e.event_id||e.id||"").trim(); if (!id) continue;
    const res = settleEvent(id);
    total += res.settled;
  }
  SpreadsheetApp.getUi().alert(`Settled ${total} bet(s) across all events with scores.`);
}
function settleEvent(event_id){
  const scores = loadTable_("Scores").rows.filter(r=> (String(r.series||"")==="Tour" || String(r.series||"").toLowerCase()==="tour") && String(r.event_id||r.id||"")===String(event_id));
  if (!scores.length) return {settled:0};

  const playerScore = {};
  for (const r of scores){
    const pid = String(r.player_id||"").trim();
    const sc = Number(r.score);
    if (!pid || !isFinite(sc)) continue;
    playerScore[pid] = sc;
  }
  let minScore = Infinity;
  for (const pid in playerScore){ if (playerScore[pid] < minScore) minScore = playerScore[pid]; }
  const winners = new Set(Object.keys(playerScore).filter(p=>playerScore[p]===minScore));

  const betsSh = ensureSheet_("Bets", ["timestamp","user_id","event_id","market","selection","param","odds","stake","potential_payout","status","settled_payout","settled_ts","is_parlay","legs_json"]);
  const usersSh = ensureSheet_("Users", ["user_id","display","coins_balance"]);
  const bVals = betsSh.getDataRange().getValues();
  const hdr = bVals[0].map(String);
  const idx = Object.fromEntries(hdr.map((h,i)=>[h,i]));
  let settled=0;

  for (let i=1;i<bVals.length;i++){
    const row = bVals[i];
    const status = String(row[idx.status]||"");
    if (status!=="open") continue;

    const isParlay = String(row[idx.is_parlay]||"") === "TRUE" || row[idx.is_parlay]===true;
    if (isParlay){
      const legs = JSON.parse(String(row[idx.legs_json]||"[]"));
      let canSettle = true;
      const legResults = [];
      for (const lg of legs){
        const sc = loadTable_("Scores").rows.filter(r=> (String(r.series||"")==="Tour" || String(r.series||"").toLowerCase()==="tour") && String(r.event_id||"")===String(lg.event_id));
        if (!sc.length){ canSettle=false; break; }
        legResults.push(gradeSingleLeg_(lg, sc));
      }
      if (!canSettle) continue;

      const anyLose = legResults.some(r=>r.result==="lost");
      const anyPush = legResults.some(r=>r.result==="push");
      let result="won";
      if (anyLose) result="lost";
      else if (anyPush) result="push";

      const stake = Number(row[idx.stake]||0);
      let payout = 0;
      if (result==="won"){
        const dec = legResults.filter(r=>r.result==="won").map(r=>amToDecimal_(r.odds)).reduce((a,b)=>a*b,1);
        payout = stake * dec;
      } else if (result==="push"){
        const wonLegs = legResults.filter(r=>r.result==="won");
        payout = wonLegs.length===0 ? stake : stake * wonLegs.map(r=>amToDecimal_(r.odds)).reduce((a,b)=>a*b,1);
      }

      row[idx.status] = result;
      row[idx.settled_payout] = payout;
      row[idx.settled_ts] = new Date();
      betsSh.getRange(i+1,1,1,hdr.length).setValues([row]);
      creditUser_(String(row[idx.user_id]), payout);
      settled++;
      continue;
    }

    const betEvent = String(row[idx.event_id]||"").trim();
    if (betEvent!==String(event_id)) continue;
    const graded = gradeSingleRow_(row, winners, playerScore, idx);
    if (!graded) continue;

    betsSh.getRange(i+1,1,1,hdr.length).setValues([graded.row]);
    creditUser_(String(row[idx.user_id]), graded.payout);
    settled++;
  }
  return {settled};
}
function creditUser_(user_id, amount){
  if (!(amount>0)) return;
  const usersSh = ensureSheet_("Users", ["user_id","display","coins_balance"]);
  const vals = usersSh.getDataRange().getValues();
  for (let i=1;i<vals.length;i++){
    if (String(vals[i][0])===String(user_id)){
      const bal = Number(vals[i][2]||0);
      usersSh.getRange(i+1,3).setValue(bal + amount);
      return;
    }
  }
}
function gradeSingleRow_(row, winners, playerScore, idx){
  const market = String(row[idx.market]||"");
  const selection = String(row[idx.selection]||"");
  const param = row[idx.param];
  const odds = Number(row[idx.odds]||0);
  const stake = Number(row[idx.stake]||0);

  let result="lost", payout=0;

  if (market==="ML"){
    if (winners.has(selection)) { result="won"; payout = stake * amToDecimal_(odds); }
    else if (winners.size===0){ result="push"; payout=stake; }
  } else if (market==="OU"){
    const parts = selection.split(" ");
    const pid = parts.slice(0, -1).join(" ") || parts[0];
    const side = (parts[parts.length-1]||"").toUpperCase();
    const line = Number(param);
    const sc = playerScore[pid];
    if (typeof sc === "number"){
      if (side==="UNDER"){ if (sc < line){ result="won"; payout = stake * amToDecimal_(odds); } else if (sc === line){ result="push"; payout = stake; } }
      else { if (sc > line){ result="won"; payout = stake * amToDecimal_(odds); } else if (sc === line){ result="push"; payout = stake; } }
    } else return null;
  } else if (market==="SPREAD"){
    const sel = String(selection);
    const m = sel.match(/^(.*)\s+([+-]?\d+(?:\.\d+)?)$/);
    if (!m) return null;
    const pid = String(m[1]).trim();
    const hSel = Number(m[2]);
    const sc = playerScore[pid];
    if (typeof sc !== "number") return null;

    const ids = Object.keys(playerScore).filter(k=>k!==pid);
    if (ids.length !== 1) { result="push"; payout=stake; }
    else {
      const scOpp = playerScore[ids[0]];
      if (sc + hSel < scOpp){ result="won"; payout = stake * amToDecimal_(odds); }
      else if (sc + hSel === scOpp){ result="push"; payout = stake; }
      else { result="lost"; }
    }
  } else return null;

  row[idx.status] = result;
  row[idx.settled_payout] = payout;
  row[idx.settled_ts] = new Date();
  return {row, payout};
}
function gradeSingleLeg_(leg, scoresRows){
  const ps = {}; for (const r of scoresRows){ ps[String(r.player_id)]=Number(r.score); }
  let winners = new Set();
  let min = Infinity; for (const k in ps){ if (ps[k]<min) min=ps[k]; }
  for (const k in ps){ if (ps[k]===min) winners.add(k); }

  if (leg.market==="ML"){
    let result="lost";
    if (winners.has(leg.selection)) result="won";
    else if (winners.size===0) result="push";
    return {result, odds:Number(leg.odds)};
  } else if (leg.market==="OU"){
    const parts = String(leg.selection).split(" ");
    const pid = parts.slice(0, -1).join(" ") || parts[0];
    const side=(parts[parts.length-1]||"").toUpperCase();
    const line=Number(leg.param), sc=ps[pid];
    if (typeof sc !== "number") return {result:"push", odds:Number(leg.odds)};
    if (side==="UNDER"){
      if (sc<line) return {result:"won", odds:Number(leg.odds)};
      if (sc===line) return {result:"push", odds:Number(leg.odds)};
      return {result:"lost", odds:Number(leg.odds)};
    } else {
      if (sc>line) return {result:"won", odds:Number(leg.odds)};
      if (sc===line) return {result:"push", odds:Number(leg.odds)};
      return {result:"lost", odds:Number(leg.odds)};
    }
  } else if (leg.market==="SPREAD"){
    const sel = String(leg.selection);
    const mm = sel.match(/^(.*)\s+([+-]?\d+(?:\.\d+)?)$/);
    if (!mm) return {result:"push", odds:Number(leg.odds)};
    const pid = String(mm[1]).trim();
    const hSel = Number(mm[2]);
    const sc = ps[pid];
    if (typeof sc !== "number") return {result:"push", odds:Number(leg.odds)};
    const ids = Object.keys(ps).filter(k=>k!==pid);
    if (ids.length !== 1) return {result:"push", odds:Number(leg.odds)};
    const scOpp = ps[ids[0]];
    if (sc + hSel < scOpp) return {result:"won", odds:Number(leg.odds)};
    if (sc + hSel === scOpp) return {result:"push", odds:Number(leg.odds)};
    return {result:"lost", odds:Number(leg.odds)};
  }
  return {result:"push", odds:Number(leg.odds)};
}

/******************** WEB APP ********************/
function doGet(){
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Tour Sportsbook")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/******************** MENU ********************/
function onOpen(){
  SpreadsheetApp.getUi().createMenu("Sportsbook")
    .addItem("Compute & Price All Events","computeAndPriceAllEvents")
    .addItem("Settle Event (prompt)","settleEventPrompt")
    .addItem("Settle All With Scores","settleAllWithScores")
    .addItem("Setup / Migrate Tabs","setupOnce")
    .addToUi();
}
function testBoard(){
  const board = getOddsBoard();
  Logger.log(JSON.stringify({count: board.length, sample: board[0]}));
}

/** -------- DIAGNOSTICS -------- **/
function debugEbStroke(player_id, map_id){
  const st = getSettings_();
  const { ebStroke_ } = buildStats_(st);
  const est = ebStroke_(String(player_id), String(map_id));
  Logger.log(JSON.stringify({ player: player_id, map: map_id, mu: est.mu, sd: est.sd }, null, 2));
}
function debugEvent(event_id){
  const st = getSettings_();
  const { ebStroke_, affinity } = buildStats_(st);

  // find event & entrants
  const events = loadTable_("Events").rows;
  const ev = events.find(e => String(e.event_id||e.id||"").trim() === String(event_id));
  if (!ev){ Logger.log("Event not found: "+event_id); return; }

  const map_id   = String(ev.map_id||ev.map||"").trim();
  const entrants = String(ev.entrants_csv||ev.entrants||"").split(/[,\n;]/).map(s=>s.trim()).filter(Boolean);
  if (!map_id || entrants.length < 2){ Logger.log("Need map and 2+ entrants"); return; }

  // show the EB parameters for each entrant
  const rows = entrants.map(p=>{
    const est = ebStroke_(p, map_id);
    const aff = ( (affinity[p]||{})[map_id] || 0 );
    const muAdj = est.mu + strokeAffinityShift_(aff, est.sd, st);
    return { player:p, mu:est.mu, sd:est.sd, affinity:aff, muAdj };
  });

  // simulate ML exactly like pricing does
  const probs = simulateTourML_(entrants, map_id, ebStroke_, affinity, st);
  const ml = entrants.map((p,i)=>({ player:p, fair_prob: probs[i], american: toAmerican_(probs[i]) }));

  Logger.log(JSON.stringify({
    event_id, map_id,
    entrants: rows,
    ml_fair: ml
  }, null, 2));
}

// (Optional) comment out any direct debug calls to avoid auto-execution
// debugEbStroke("Wick", "UPSIDE_TOWN_EASY");
// debugEbStroke("Liberator", "UPSIDE_TOWN_EASY");
// debugEvent("tour");
