[bordsplanering_1.html](https://github.com/user-attachments/files/29959878/bordsplanering_1.html)
<!doctype html>
<html lang="sv">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Bordsplanering — planlösning &amp; bordsplacering</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --paper:#efe9dc;
    --paper-2:#e4ddcc;
    --ink:#1c2b3a;
    --ink-soft:#4b5c6b;
    --line:#cdc3aa;
    --accent:#c98a3d;
    --accent-2:#8a5a22;
    --blue-deep:#132133;
    --blue-mid:#1c3049;
    --blue-grid-minor:#22394f;
    --blue-grid-major:#2c4763;
    --wall:#f2ead9;
    --seat-empty:#31506e;
    --seat-filled:#c98a3d;
    --danger:#b3543c;
  }
  *{box-sizing:border-box;}
  html,body{height:100%;margin:0;}
  body{
    font-family:'Space Grotesk',system-ui,sans-serif;
    color:var(--ink);
    background:var(--paper);
    display:flex;
    overflow:hidden;
  }
  .app{display:flex;width:100%;height:100vh;}

  .sidebar{
    width:340px;
    min-width:340px;
    background:var(--paper);
    border-right:1px solid var(--line);
    display:flex;
    flex-direction:column;
    overflow-y:auto;
  }
  .brand{padding:18px 18px 12px;border-bottom:1px solid var(--line);}
  .brand h1{font-size:16px;letter-spacing:.04em;text-transform:uppercase;margin:0 0 2px;font-weight:700;}
  .brand .save-status{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--ink-soft);}
  .section{padding:14px 18px;border-bottom:1px solid var(--line);}
  .section h2{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--ink-soft);margin:0 0 10px;font-weight:600;}
  .btn-row{display:flex;flex-wrap:wrap;gap:6px;}
  button{
    font-family:'Space Grotesk',sans-serif;font-size:12.5px;border:1px solid var(--ink);
    background:var(--paper);color:var(--ink);padding:7px 10px;border-radius:3px;cursor:pointer;
    display:inline-flex;align-items:center;gap:6px;transition:background .12s, color .12s;
  }
  button:hover{background:var(--ink);color:var(--paper);}
  button.primary{background:var(--accent);border-color:var(--accent-2);color:#fff;}
  button.primary:hover{background:var(--accent-2);}
  button.ghost{border-color:var(--line);color:var(--ink-soft);}
  button.ghost:hover{background:var(--ink-soft);color:#fff;}
  button.danger{border-color:var(--danger);color:var(--danger);}
  button.danger:hover{background:var(--danger);color:#fff;}
  button:disabled{opacity:.35;cursor:not-allowed;}
  button:disabled:hover{background:var(--paper);color:var(--ink);}
  button.full{width:100%;justify-content:center;}

  .hint{font-size:11.5px;color:var(--ink-soft);margin-top:8px;line-height:1.4;}
  label.check{display:flex;align-items:center;gap:6px;font-size:12.5px;color:var(--ink);margin-top:4px;}

  .field{margin-bottom:10px;}
  .field label{display:block;font-size:11px;color:var(--ink-soft);margin-bottom:4px;text-transform:uppercase;letter-spacing:.04em;}
  .field input[type=text], .field input[type=number], .field select{
    width:100%;font-family:'IBM Plex Mono',monospace;font-size:13px;padding:6px 8px;
    border:1px solid var(--line);border-radius:3px;background:#fff;color:var(--ink);
  }
  .field select{font-family:'Space Grotesk',sans-serif;}
  .field-row{display:flex;gap:8px;}
  .field-row .field{flex:1;}
  .recommend{font-size:11.5px;color:var(--ink-soft);margin-top:-4px;margin-bottom:10px;}
  .recommend a{color:var(--accent-2);cursor:pointer;text-decoration:underline;}
  .empty-note{font-size:12px;color:var(--ink-soft);font-style:italic;}

  .guest-add{display:flex;gap:6px;margin-bottom:10px;}
  .guest-add input{flex:1;font-family:'Space Grotesk',sans-serif;font-size:13px;padding:7px 8px;border:1px solid var(--line);border-radius:3px;}
  .guest-list{list-style:none;margin:0;padding:0;max-height:160px;overflow-y:auto;}
  .guest-list li{display:flex;align-items:center;justify-content:space-between;padding:6px 8px;border-radius:3px;font-size:13px;cursor:pointer;border:1px solid transparent;}
  .guest-list li:hover{background:var(--paper-2);}
  .guest-list li.armed{background:var(--accent);color:#fff;border-color:var(--accent-2);}
  .guest-list li .rm{font-family:'IBM Plex Mono',monospace;color:var(--ink-soft);padding:0 4px;}
  .guest-list li.armed .rm{color:#fff;}
  .guest-list li .rm:hover{color:var(--danger);}
  .armed-banner{font-size:12px;background:var(--accent);color:#fff;padding:8px 10px;border-radius:3px;margin-bottom:10px;line-height:1.4;}
  .progress{font-family:'IBM Plex Mono',monospace;font-size:11.5px;color:var(--ink-soft);margin-bottom:8px;}

  .overview-item{padding:8px 0;border-bottom:1px dashed var(--line);font-size:12.5px;}
  .overview-item:last-child{border-bottom:none;}
  .overview-item .t-name{font-weight:600;}
  .overview-item .t-count{font-family:'IBM Plex Mono',monospace;color:var(--ink-soft);font-size:11.5px;}
  .overview-item .t-guests{color:var(--ink-soft);margin-top:2px;}

  main.canvas-wrap{flex:1;display:flex;flex-direction:column;background:var(--blue-deep);min-width:0;}
  .toolbar{
    background:var(--blue-mid);border-bottom:1px solid #0e1826;padding:10px 16px;
    display:flex;align-items:center;gap:10px;flex-wrap:wrap;color:var(--paper);
  }
  .toolbar .grp{display:flex;align-items:center;gap:6px;padding-right:10px;border-right:1px solid #2c4763;}
  .toolbar .grp:last-child{border-right:none;}
  .toolbar button{background:transparent;border:1px solid #3a577a;color:var(--paper);}
  .toolbar button:hover{background:#2c4763;color:#fff;}
  .toolbar .mode-hint{font-family:'IBM Plex Mono',monospace;font-size:11.5px;color:#9fb4c8;}
  .scroll-area{flex:1;overflow:auto;display:flex;align-items:flex-start;justify-content:flex-start;padding:24px;}
  svg#plan{display:block;touch-action:none;}

  .table-group{cursor:grab;}
  .table-group.selected .table-shape{stroke:var(--accent);stroke-width:4;}
  .table-shape{fill:var(--wall);stroke:var(--blue-deep);stroke-width:2.5;}
  .table-label{font-family:'Space Grotesk',sans-serif;font-weight:600;fill:var(--blue-deep);text-anchor:middle;dominant-baseline:middle;pointer-events:none;}
  .seat{cursor:pointer;}
  .seat circle{fill:var(--seat-empty);stroke:var(--blue-deep);stroke-width:1.5;}
  .seat.filled circle{fill:var(--seat-filled);stroke:var(--accent-2);}
  .seat text{font-family:'IBM Plex Mono',monospace;font-size:11px;fill:#fff;text-anchor:middle;dominant-baseline:middle;pointer-events:none;font-weight:600;}
  .room-outline{fill:none;stroke:var(--wall);stroke-width:10;stroke-linejoin:round;}
  .table-delete circle{fill:var(--danger);}
  .table-delete text{fill:#fff;font-family:'IBM Plex Mono',monospace;font-size:14px;text-anchor:middle;dominant-baseline:middle;}
  .table-delete{cursor:pointer;}

  #printSummary{display:none;}
  @media print{
    .sidebar, .toolbar{display:none !important;}
    .scroll-area{padding:0;overflow:visible;}
    body{background:#fff;}
    main.canvas-wrap{background:#fff;}
    #printSummary{display:block;padding:20px;font-family:'Space Grotesk',sans-serif;}
    #printSummary h2{font-size:16px;margin-bottom:10px;}
    #printSummary .p-item{margin-bottom:6px;font-size:13px;}
  }
</style>
</head>
<body>
<div class="app">

  <aside class="sidebar">
    <div class="brand">
      <h1>Bordsplanering</h1>
      <div class="save-status" id="saveStatus">Laddar…</div>
    </div>

    <div class="section">
      <h2>Rum</h2>
      <div class="field-row">
        <div class="field"><label>Bredd (cm)</label><input type="number" id="roomWidth" min="200" max="5000" step="10"></div>
        <div class="field"><label>Längd (cm)</label><input type="number" id="roomLength" min="200" max="5000" step="10"></div>
      </div>
      <div class="hint">Ritytan skalas automatiskt efter rummets mått. ▭ ruta i bakgrunden = 50 cm.</div>
    </div>

    <div class="section">
      <h2>Automatisk bordsplacering</h2>
      <div class="field"><label>Antal gäster</label><input type="number" id="guestCount" min="1" max="1000" step="1"></div>
      <div class="field"><label>Bordstyp, preferens</label>
        <select id="tablePref">
          <option value="both">Både runda och rektangulära</option>
          <option value="round">Endast runda bord</option>
          <option value="rect">Endast rektangulära bord</option>
        </select>
      </div>
      <div class="field-row">
        <div class="field"><label>Diameter runt bord (cm)</label><input type="number" id="acRoundD" min="60" max="300" step="10"></div>
      </div>
      <div class="field-row">
        <div class="field"><label>Bredd rekt. bord (cm)</label><input type="number" id="acRectW" min="60" max="400" step="10"></div>
        <div class="field"><label>Djup rekt. bord (cm)</label><input type="number" id="acRectH" min="40" max="200" step="10"></div>
      </div>
      <label class="check"><input type="checkbox" id="includeHead" checked> Hedersbord/toppbord (bänkplacering, endast en sida — t.ex. för brudpar)</label>
      <div class="field" style="margin-top:8px;"><label>Platser vid hedersbord</label><input type="number" id="headSeats" min="2" max="40" step="1"></div>
      <div class="hint">Hedersbord delas automatiskt upp i flera bord i rad om det blir för brett (max ~8 platser per bord), med bänkplacering på framsidan så gäster kan komma fram och prata.</div>
      <button class="primary full" id="generateLayout" style="margin-top:10px;">✨ Föreslå bordsplacering</button>
    </div>

    <div class="section">
      <h2>Lägg till bord manuellt</h2>
      <div class="btn-row">
        <button id="addRound">◯ Runt bord</button>
        <button id="addRect">▭ Rektangulärt bord</button>
        <button id="addBench">▭ Bänkbord</button>
      </div>
    </div>

    <div class="section" id="propsSection" style="display:none;">
      <h2>Bordets egenskaper</h2>
      <div class="field">
        <label>Namn</label>
        <input type="text" id="propName">
      </div>
      <div id="propRound">
        <div class="field"><label>Diameter (cm)</label><input type="number" id="propDiameter" min="40" max="400" step="5"></div>
      </div>
      <div id="propRect" style="display:none;">
        <div class="field-row">
          <div class="field"><label>Bredd (cm)</label><input type="number" id="propW" min="40" max="600" step="5"></div>
          <div class="field"><label>Djup (cm)</label><input type="number" id="propH" min="40" max="500" step="5"></div>
        </div>
        <div class="field"><label>Rotation (grader)</label><input type="number" id="propRot" min="0" max="359" step="15"></div>
        <label class="check"><input type="checkbox" id="propBench"> Bänkplacering (endast framsidan)</label>
      </div>
      <div class="field" style="margin-top:10px;"><label>Antal platser</label><input type="number" id="propSeats" min="1" max="40" step="1"></div>
      <div class="recommend" id="propRecommend"></div>
      <button class="danger full" id="deleteTable">Ta bort bordet</button>
    </div>

    <div class="section">
      <h2>Gästlista</h2>
      <div class="hint" style="margin-top:0;">Importera från Excel (.xlsx) eller CSV. Kolumner för förnamn/efternamn känns av automatiskt — annars används första kolumnen.</div>
      <input type="file" id="guestFile" accept=".csv,.xlsx,.xls" style="margin-top:8px;font-size:12px;width:100%;">
      <div class="hint" id="importStatus"></div>
      <div class="guest-add" style="margin-top:12px;">
        <input type="text" id="guestInput" placeholder="Namn på gäst…">
        <button id="addGuest">Lägg till</button>
      </div>
      <div class="progress" id="guestProgress"></div>
      <div id="armedBanner"></div>
      <ul class="guest-list" id="guestList"></ul>
      <div class="hint">Klicka på en gäst, klicka sedan på en ledig plats vid ett bord för att tilldela. Klicka på en tilldelad plats för att ta bort gästen därifrån.</div>
    </div>

    <div class="section">
      <h2>Bordsöversikt</h2>
      <div id="overview"></div>
    </div>

    <div class="section">
      <h2>Gästlista i bokstavsordning</h2>
      <div class="hint" style="margin-top:0;">Sorterad på efternamn, med vilket bord respektive gäst sitter vid.</div>
      <div id="alphaList" style="margin-top:8px;max-height:220px;overflow-y:auto;"></div>
      <div class="field-row" style="margin-top:10px;">
        <div class="field"><label>Linjefärg</label><input type="color" id="exportColor2" value="#1c2b3a" style="height:34px;padding:2px;"></div>
      </div>
      <div class="hint">Bakgrunden är alltid transparent.</div>
      <button class="full" id="exportNameList" style="margin-top:8px;">⬇ Exportera gästlista (SVG)</button>
    </div>

    <div class="section">
      <h2>Exportera som vektorfil (SVG)</h2>
      <div class="field-row">
        <div class="field"><label>Linjefärg</label><input type="color" id="exportColor" value="#1c2b3a" style="height:34px;padding:2px;"></div>
      </div>
      <div class="hint">Bakgrunden är alltid transparent, så filen kan läggas ovanpå valfri färg eller material.</div>
      <button class="full" id="exportSvg" style="margin-top:8px;">⬇ Exportera SVG</button>
    </div>

    <div class="section" style="border-bottom:none;">
      <h2>Åtgärder</h2>
      <div class="btn-row">
        <button id="saveNow" class="ghost">💾 Spara nu</button>
        <button id="printBtn" class="ghost">🖨 Skriv ut</button>
        <button id="clearAll" class="danger">Rensa allt</button>
      </div>
    </div>
  </aside>

  <main class="canvas-wrap">
    <div class="toolbar">
      <div class="grp">
        <button id="zoomOut">−</button>
        <span class="mode-hint" id="zoomLabel">100%</span>
        <button id="zoomIn">+</button>
        <button id="zoomReset" class="ghost">Återställ vy</button>
      </div>
      <div class="grp">
        <label class="check" style="margin:0;color:#c9d6e2;"><input type="checkbox" id="snapToggle" checked> Snäppa 10 cm</label>
      </div>
      <div class="grp" style="border-right:none;">
        <span class="mode-hint">Dra bord för att placera dem. Klicka för att välja.</span>
      </div>
    </div>
    <div class="scroll-area">
      <svg id="plan"></svg>
    </div>
  </main>

</div>

<div id="printSummary"></div>

<script>
(function(){
  "use strict";

  const MARGIN = 150;
  const BASE_PX_PER_CM = 0.5;
  const SEAT_SPACING = 65;
  const CHAIR_OFFSET = 28;
  const CHAIR_R = 14;
  const MAX_PER_HEAD_TABLE = 8;

  let uidCounter = 1;
  function uid(){ return 'id' + (uidCounter++) + '_' + Math.random().toString(36).slice(2,7); }

  const state = {
    room: { width: 1000, length: 700 },
    tables: [],
    guests: [],
    selectedTableId: null,
    armedGuestId: null,
    snap: true,
    zoom: 1,
  };

  let dragging = null;

  function snapVal(v){ return state.snap ? Math.round(v/10)*10 : Math.round(v); }
  function clamp(v,a,b){ return Math.max(a, Math.min(b, v)); }

  function tableFootprint(t){
    if (t.type === 'round') return {w: t.diameter, h: t.diameter};
    return {w: t.w, h: t.h};
  }

  function recommendedSeats(t){
    if (t.type === 'round'){
      return Math.max(2, Math.round((Math.PI * t.diameter) / SEAT_SPACING));
    }
    if (t.bench){
      return Math.max(1, Math.floor(t.w / 60));
    }
    return Math.max(2, Math.round((2*(t.w+t.h)) / SEAT_SPACING));
  }

  function getSeatCount(t){
    return (t.seatsOverride != null) ? t.seatsOverride : recommendedSeats(t);
  }

  function ensureSeatsLength(t){
    const n = getSeatCount(t);
    if (t.seats.length > n) t.seats.length = n;
    while (t.seats.length < n) t.seats.push(null);
  }

  function rectPerimeterPoint(w,h,d){
    const per = 2*(w+h);
    let dd = ((d % per) + per) % per;
    if (dd <= w) return {x:-w/2+dd, y:-h/2, nx:0, ny:-1};
    dd -= w;
    if (dd <= h) return {x:w/2, y:-h/2+dd, nx:1, ny:0};
    dd -= h;
    if (dd <= w) return {x:w/2-dd, y:h/2, nx:0, ny:1};
    dd -= w;
    return {x:-w/2, y:h/2-dd, nx:-1, ny:0};
  }

  function getSeatLocalPositions(t){
    const n = getSeatCount(t);
    const positions = [];
    if (t.type === 'round'){
      const r = t.diameter/2 + CHAIR_OFFSET;
      for (let i=0;i<n;i++){
        const angle = (i/n)*Math.PI*2 - Math.PI/2;
        positions.push({x: Math.cos(angle)*r, y: Math.sin(angle)*r});
      }
    } else if (t.bench){
      const spacing = t.w / n;
      for (let i=0;i<n;i++){
        positions.push({x: -t.w/2 + spacing*(i+0.5), y: -t.h/2 - CHAIR_OFFSET});
      }
    } else {
      const perim = 2*(t.w+t.h);
      for (let i=0;i<n;i++){
        const d = (i+0.5) * (perim/n);
        const p = rectPerimeterPoint(t.w, t.h, d);
        positions.push({x: p.x + p.nx*CHAIR_OFFSET, y: p.y + p.ny*CHAIR_OFFSET});
      }
    }
    return positions;
  }

  function tableName(t, idx){
    return (t.name && t.name.trim()) ? t.name.trim() : ('Bord ' + (idx+1));
  }

  function unassignedGuests(){
    return state.guests.filter(g => !state.tables.some(t => t.seats.includes(g.id)));
  }
  function findTable(id){ return state.tables.find(t=>t.id===id); }
  function findGuest(id){ return state.guests.find(g=>g.id===id); }

  function splitName(full){
    const parts = String(full||'').trim().split(/\s+/).filter(Boolean);
    if (parts.length <= 1) return {first:'', last: parts[0]||''};
    return {first: parts.slice(0,-1).join(' '), last: parts[parts.length-1]};
  }
  function guestTableInfo(gid){
    for (let idx=0; idx<state.tables.length; idx++){
      if (state.tables[idx].seats.includes(gid)) return tableName(state.tables[idx], idx);
    }
    return null;
  }
  function sortedGuestsByLastName(){
    return [...state.guests].sort((a,b) => splitName(a.name).last.toLowerCase().localeCompare(splitName(b.name).last.toLowerCase(), 'sv'));
  }

  function makeTable(type, opts){
    const t = Object.assign({
      id: uid(), type, name: '', x: 300, y: 300,
      diameter: 120, w: 180, h: 75, bench:false, rotation: 0,
      seatsOverride: null, seats: []
    }, opts||{});
    ensureSeatsLength(t);
    return t;
  }

  function addTable(type, bench){
    const idx = state.tables.length;
    const t = makeTable(type, {
      x: MARGIN + 150 + (idx%5)*160,
      y: MARGIN + 150 + Math.floor(idx/5)*160,
      bench: !!bench,
      w: bench ? 240 : 180,
      h: bench ? 70 : 75
    });
    state.tables.push(t);
    state.selectedTableId = t.id;
    commit();
  }

  function deleteTable(id){
    state.tables = state.tables.filter(t=>t.id!==id);
    if (state.selectedTableId === id) state.selectedTableId = null;
    commit();
  }

  function assignGuestToSeat(tableId, seatIndex, guestId){
    state.tables.forEach(t => { t.seats = t.seats.map(g => g===guestId ? null : g); });
    findTable(tableId).seats[seatIndex] = guestId;
  }
  function unassignSeat(tableId, seatIndex){
    findTable(tableId).seats[seatIndex] = null;
  }
  function addGuest(name){
    name = name.trim();
    if (!name) return;
    state.guests.push({id: uid(), name});
    commit();
  }
  function removeGuest(id){
    state.guests = state.guests.filter(g=>g.id!==id);
    state.tables.forEach(t => { t.seats = t.seats.map(g => g===id ? null : g); });
    if (state.armedGuestId === id) state.armedGuestId = null;
    commit();
  }

  function autoLayout(){
    if (state.tables.length && !confirm('Detta ersätter nuvarande bordsplacering. Fortsätta?')) return;

    const guestCount = clamp(parseInt(document.getElementById('guestCount').value||1,10), 1, 2000);
    const pref = document.getElementById('tablePref').value;
    const roundD = clamp(parseInt(document.getElementById('acRoundD').value||150,10), 60, 300);
    const rectW = clamp(parseInt(document.getElementById('acRectW').value||180,10), 60, 400);
    const rectH = clamp(parseInt(document.getElementById('acRectH').value||75,10), 40, 200);
    const includeHead = document.getElementById('includeHead').checked;
    const headSeatsWanted = clamp(parseInt(document.getElementById('headSeats').value||8,10), 2, 40);

    let remaining = guestCount;
    const newTables = [];

    if (includeHead && remaining > 0){
      const headTotal = Math.min(headSeatsWanted, remaining);
      const headCount = Math.max(1, Math.ceil(headTotal / MAX_PER_HEAD_TABLE));
      const base = Math.floor(headTotal / headCount);
      let extra = headTotal - base*headCount;
      for (let i=0;i<headCount;i++){
        const seats = base + (i < extra ? 1 : 0);
        if (seats <= 0) continue;
        const w = Math.max(120, seats*60);
        newTables.push(makeTable('rect', {
          w, h:75, bench:true, seatsOverride: seats,
          name: 'Hedersbord' + (headCount>1 ? ' '+(i+1) : '')
        }));
      }
      remaining -= headTotal;
    }

    const roundCap = recommendedSeats({type:'round', diameter: roundD});
    const rectCap = recommendedSeats({type:'rect', w: rectW, h: rectH, bench:false});
    let toggle = 0;
    let guard = 0;
    while (remaining > 0 && guard < 500){
      guard++;
      let useType = pref === 'round' ? 'round' : pref === 'rect' ? 'rect' : (toggle%2===0 ? 'round' : 'rect');
      toggle++;
      const cap = useType === 'round' ? roundCap : rectCap;
      const seatsForThis = Math.min(cap, remaining);
      if (useType === 'round'){
        newTables.push(makeTable('round', {
          diameter: roundD, seatsOverride: seatsForThis < cap ? seatsForThis : null
        }));
      } else {
        newTables.push(makeTable('rect', {
          w: rectW, h: rectH, bench:false, seatsOverride: seatsForThis < cap ? seatsForThis : null
        }));
      }
      remaining -= seatsForThis;
    }

    arrangeTablesInRoom(newTables);
    state.tables = newTables;
    state.selectedTableId = null;
    commit();
  }

  function arrangeTablesInRoom(tables){
    const pad = 90;
    const heads = tables.filter(t => t.bench);
    const others = tables.filter(t => !t.bench);

    if (heads.length){
      const totalW = heads.reduce((s,t)=>s+t.w,0) + pad*(heads.length-1);
      let cursor = MARGIN + Math.max(20, (state.room.width - totalW)/2);
      heads.forEach(t => {
        t.x = cursor + t.w/2;
        t.y = MARGIN + 90;
        cursor += t.w + pad;
      });
    }

    let startY = heads.length ? MARGIN + 90 + 190 : MARGIN + 130;
    let cx = MARGIN + 130, cy = startY, rowMaxH = 0;
    const rightLimit = MARGIN + state.room.width - 60;
    others.forEach(t => {
      const fp = tableFootprint(t);
      if (cx + fp.w/2 > rightLimit && cx > MARGIN + 130){
        cx = MARGIN + 130; cy += rowMaxH + pad; rowMaxH = 0;
      }
      t.x = cx + fp.w/2;
      t.y = cy + fp.h/2;
      cx += fp.w + pad;
      rowMaxH = Math.max(rowMaxH, fp.h);
    });
  }

  function canvasDims(){
    let maxX = state.room.width + MARGIN;
    let maxY = state.room.length + MARGIN;
    state.tables.forEach(t => {
      const fp = tableFootprint(t);
      maxX = Math.max(maxX, t.x + fp.w/2 + CHAIR_OFFSET + 40 - MARGIN);
      maxY = Math.max(maxY, t.y + fp.h/2 + CHAIR_OFFSET + 40 - MARGIN);
    });
    return { w: maxX + MARGIN, h: maxY + MARGIN };
  }

  const svg = document.getElementById('plan');

  function applyZoom(dims){
    const w = dims.w * BASE_PX_PER_CM * state.zoom;
    const h = dims.h * BASE_PX_PER_CM * state.zoom;
    svg.style.width = w + 'px';
    svg.style.height = h + 'px';
    document.getElementById('zoomLabel').textContent = Math.round(state.zoom*100) + '%';
  }

  function toSvgPoint(evt){
    const pt = svg.createSVGPoint();
    pt.x = evt.clientX; pt.y = evt.clientY;
    const ctm = svg.getScreenCTM().inverse();
    const p = pt.matrixTransform(ctm);
    return {x:p.x, y:p.y};
  }

  function esc(str){
    return String(str).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  function render(){
    const dims = canvasDims();
    svg.setAttribute('viewBox', `0 0 ${dims.w} ${dims.h}`);
    let parts = [];

    parts.push(`<defs>
      <pattern id="gminor" width="50" height="50" patternUnits="userSpaceOnUse">
        <path d="M 50 0 L 0 0 0 50" fill="none" stroke="var(--blue-grid-minor)" stroke-width="1"/>
      </pattern>
      <pattern id="gmajor" width="200" height="200" patternUnits="userSpaceOnUse">
        <rect width="200" height="200" fill="url(#gminor)"/>
        <path d="M 200 0 L 0 0 0 200" fill="none" stroke="var(--blue-grid-major)" stroke-width="1.5"/>
      </pattern>
    </defs>`);
    parts.push(`<rect x="0" y="0" width="${dims.w}" height="${dims.h}" fill="url(#gmajor)"/>`);
    parts.push(`<rect class="room-outline" x="${MARGIN}" y="${MARGIN}" width="${state.room.width}" height="${state.room.length}"/>`);
    parts.push(`<text x="${MARGIN}" y="${MARGIN-16}" font-family="IBM Plex Mono, monospace" font-size="12" fill="#9fb4c8">${(state.room.width/100).toFixed(1)} × ${(state.room.length/100).toFixed(1)} m</text>`);

    state.tables.forEach((t, idx) => {
      ensureSeatsLength(t);
      const selected = t.id === state.selectedTableId;
      let shape = '';
      if (t.type === 'round'){
        shape = `<circle class="table-shape" cx="0" cy="0" r="${t.diameter/2}"/>`;
      } else {
        shape = `<rect class="table-shape" x="${-t.w/2}" y="${-t.h/2}" width="${t.w}" height="${t.h}" rx="6"/>`;
      }
      const label = `<text class="table-label" x="0" y="4" font-size="16">${esc(tableName(t,idx))}</text>`;
      const seatPositions = getSeatLocalPositions(t);
      let seats = '';
      seatPositions.forEach((p, si) => {
        const guestId = t.seats[si];
        const guest = guestId ? findGuest(guestId) : null;
        const filled = !!guest;
        const initials = guest ? guest.name.trim().slice(0,2).toUpperCase() : '';
        seats += `<g class="seat ${filled?'filled':''}" data-table="${t.id}" data-seat="${si}">
          <circle cx="${p.x}" cy="${p.y}" r="${CHAIR_R}"/>
          ${filled ? `<text x="${p.x}" y="${p.y+1}">${esc(initials)}</text>` : ''}
        </g>`;
      });
      const rot = t.type === 'rect' ? (t.rotation||0) : 0;
      let del = '';
      if (selected){
        const dx = (t.type==='round'? t.diameter/2 : t.w/2) + 12;
        const dy = -(t.type==='round'? t.diameter/2 : t.h/2) - 12;
        del = `<g class="table-delete" data-table="${t.id}" transform="translate(${dx},${dy})">
          <circle r="12"/><text x="0" y="1">×</text>
        </g>`;
      }
      parts.push(`<g class="table-group ${selected?'selected':''}" data-table="${t.id}" transform="translate(${t.x},${t.y}) rotate(${rot})">
        ${shape}${seats}
        <g transform="rotate(${-rot})">${label}</g>
        ${del}
      </g>`);
    });

    svg.innerHTML = parts.join('');
    applyZoom(dims);
    renderSidebar();
  }

  function renderSidebar(){
    document.getElementById('roomWidth').value = state.room.width;
    document.getElementById('roomLength').value = state.room.length;

    const propsSection = document.getElementById('propsSection');
    const table = state.selectedTableId ? findTable(state.selectedTableId) : null;
    if (table){
      propsSection.style.display = '';
      const idx = state.tables.indexOf(table);
      document.getElementById('propName').value = table.name || '';
      document.getElementById('propName').placeholder = tableName(table, idx);
      const isRound = table.type === 'round';
      document.getElementById('propRound').style.display = isRound ? '' : 'none';
      document.getElementById('propRect').style.display = isRound ? 'none' : '';
      if (isRound) document.getElementById('propDiameter').value = table.diameter;
      else {
        document.getElementById('propW').value = table.w;
        document.getElementById('propH').value = table.h;
        document.getElementById('propRot').value = table.rotation || 0;
        document.getElementById('propBench').checked = !!table.bench;
      }
      document.getElementById('propSeats').value = getSeatCount(table);
      const rec = recommendedSeats(table);
      document.getElementById('propRecommend').innerHTML =
        `Rekommenderat antal platser: <strong>${rec}</strong>` +
        (table.seatsOverride != null ? ` — <a id="resetSeats">återställ till rekommenderat</a>` : '');
    } else {
      propsSection.style.display = 'none';
    }

    const unassigned = unassignedGuests();
    const listEl = document.getElementById('guestList');
    if (unassigned.length === 0){
      listEl.innerHTML = state.guests.length
        ? `<li class="empty-note" style="cursor:default;">Alla gäster är placerade.</li>`
        : `<li class="empty-note" style="cursor:default;">Inga gäster tillagda än.</li>`;
    } else {
      listEl.innerHTML = unassigned.map(g => `
        <li data-guest="${g.id}" class="${state.armedGuestId===g.id?'armed':''}">
          <span>${esc(g.name)}</span><span class="rm" data-rm="${g.id}">✕</span>
        </li>`).join('');
    }
    const totalPlaced = state.guests.length - unassigned.length;
    document.getElementById('guestProgress').textContent =
      state.guests.length ? `${totalPlaced} av ${state.guests.length} gäster placerade` : '';

    const armedGuest = state.armedGuestId ? findGuest(state.armedGuestId) : null;
    document.getElementById('armedBanner').innerHTML = armedGuest
      ? `<div class="armed-banner">${esc(armedGuest.name)} vald — klicka på en ledig plats. Klicka på namnet igen för att avbryta.</div>`
      : '';

    const overviewEl = document.getElementById('overview');
    if (state.tables.length === 0){
      overviewEl.innerHTML = `<div class="empty-note">Inga bord tillagda än. Ange antal gäster ovan och tryck på "Föreslå bordsplacering", eller lägg till bord manuellt.</div>`;
    } else {
      const totalSeats = state.tables.reduce((s,t)=>s+getSeatCount(t),0);
      overviewEl.innerHTML = `<div class="progress" style="margin-bottom:8px;">${state.tables.length} bord — totalt ${totalSeats} platser</div>` +
        state.tables.map((t, idx) => {
          const n = getSeatCount(t);
          const names = t.seats.filter(Boolean).map(gid => { const g=findGuest(gid); return g?g.name:''; }).filter(Boolean);
          return `<div class="overview-item">
            <div><span class="t-name">${esc(tableName(t,idx))}</span> <span class="t-count">— ${names.length}/${n} platser</span></div>
            ${names.length ? `<div class="t-guests">${esc(names.join(', '))}</div>` : ''}
          </div>`;
        }).join('');
    }

    const alphaEl = document.getElementById('alphaList');
    const sorted = sortedGuestsByLastName();
    if (!sorted.length){
      alphaEl.innerHTML = `<div class="empty-note">Inga gäster tillagda än.</div>`;
    } else {
      alphaEl.innerHTML = sorted.map(g => {
        const tn = guestTableInfo(g.id);
        return `<div class="overview-item"><span class="t-name">${esc(g.name)}</span> <span class="t-count">— ${tn?esc(tn):'Ej placerad'}</span></div>`;
      }).join('');
    }

    document.getElementById('snapToggle').checked = state.snap;
  }

  const statusEl = document.getElementById('saveStatus');
  let saveTimer = null;

  function exportableState(){
    return { room: state.room, tables: state.tables, guests: state.guests };
  }
  async function doSave(){
    try{
      statusEl.textContent = 'Sparar…';
      await window.storage.set('floorplan-v2', JSON.stringify(exportableState()));
      statusEl.textContent = 'Sparat ' + new Date().toLocaleTimeString('sv-SE', {hour:'2-digit', minute:'2-digit'});
    }catch(e){
      statusEl.textContent = 'Kunde inte spara (endast lokalt i sessionen)';
    }
  }
  function scheduleSave(){
    clearTimeout(saveTimer);
    statusEl.textContent = 'Väntar…';
    saveTimer = setTimeout(doSave, 800);
  }
  function commit(){ render(); scheduleSave(); }

  async function loadState(){
    try{
      const res = await window.storage.get('floorplan-v2');
      if (res && res.value){
        const data = JSON.parse(res.value);
        if (data.room) state.room = data.room;
        state.tables = data.tables || [];
        state.guests = data.guests || [];
        state.tables.forEach(t => { if(t.bench===undefined) t.bench=false; ensureSeatsLength(t); });
        statusEl.textContent = 'Tidigare planlösning inläst';
      } else {
        statusEl.textContent = 'Ny planlösning';
      }
    }catch(e){
      statusEl.textContent = 'Ny planlösning';
    }
    render();
  }

  svg.addEventListener('pointerdown', (e) => {
    const pt = toSvgPoint(e);
    const target = e.target;

    const delEl = target.closest('.table-delete');
    if (delEl){ deleteTable(delEl.dataset.table); return; }

    const seatEl = target.closest('.seat');
    if (seatEl && seatEl.dataset.table){
      const tId = seatEl.dataset.table, sIdx = parseInt(seatEl.dataset.seat,10);
      const t = findTable(tId);
      state.selectedTableId = tId;
      if (state.armedGuestId){
        assignGuestToSeat(tId, sIdx, state.armedGuestId);
        state.armedGuestId = null;
      } else if (t.seats[sIdx]){
        unassignSeat(tId, sIdx);
      }
      commit();
      return;
    }

    const tableEl = target.closest('.table-group');
    if (tableEl){
      const tId = tableEl.dataset.table;
      const t = findTable(tId);
      state.selectedTableId = tId;
      dragging = {tableId: tId, startPt: pt, orig: {x:t.x, y:t.y}, moved:false};
      render();
      return;
    }

    state.selectedTableId = null;
    render();
  });

  svg.addEventListener('pointermove', (e) => {
    if (!dragging) return;
    const pt = toSvgPoint(e);
    const dx = pt.x - dragging.startPt.x, dy = pt.y - dragging.startPt.y;
    if (Math.hypot(dx,dy) > 2) dragging.moved = true;
    const t = findTable(dragging.tableId);
    if (!t) return;
    t.x = snapVal(Math.max(0, dragging.orig.x + dx));
    t.y = snapVal(Math.max(0, dragging.orig.y + dy));
    render();
  });
  window.addEventListener('pointerup', () => {
    if (dragging && dragging.moved) scheduleSave();
    dragging = null;
  });

  document.getElementById('roomWidth').addEventListener('input', (e) => {
    state.room.width = clamp(parseInt(e.target.value||200,10), 200, 5000); commit();
  });
  document.getElementById('roomLength').addEventListener('input', (e) => {
    state.room.length = clamp(parseInt(e.target.value||200,10), 200, 5000); commit();
  });

  document.getElementById('snapToggle').addEventListener('change', (e) => { state.snap = e.target.checked; });
  document.getElementById('zoomIn').addEventListener('click', () => { state.zoom = clamp(state.zoom*1.2, 0.3, 3); applyZoom(canvasDims()); });
  document.getElementById('zoomOut').addEventListener('click', () => { state.zoom = clamp(state.zoom/1.2, 0.3, 3); applyZoom(canvasDims()); });
  document.getElementById('zoomReset').addEventListener('click', () => { state.zoom = 1; applyZoom(canvasDims()); });

  document.getElementById('addRound').addEventListener('click', () => addTable('round', false));
  document.getElementById('addRect').addEventListener('click', () => addTable('rect', false));
  document.getElementById('addBench').addEventListener('click', () => addTable('rect', true));

  document.getElementById('generateLayout').addEventListener('click', autoLayout);

  document.getElementById('propName').addEventListener('input', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return; t.name = e.target.value; commit();
  });
  document.getElementById('propDiameter').addEventListener('input', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return;
    t.diameter = clamp(parseInt(e.target.value||40,10), 40, 400); ensureSeatsLength(t); commit();
  });
  document.getElementById('propW').addEventListener('input', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return;
    t.w = clamp(parseInt(e.target.value||40,10), 40, 600); ensureSeatsLength(t); commit();
  });
  document.getElementById('propH').addEventListener('input', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return;
    t.h = clamp(parseInt(e.target.value||40,10), 40, 500); ensureSeatsLength(t); commit();
  });
  document.getElementById('propRot').addEventListener('input', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return;
    t.rotation = ((parseInt(e.target.value||0,10) % 360) + 360) % 360; commit();
  });
  document.getElementById('propBench').addEventListener('change', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return;
    t.bench = e.target.checked; ensureSeatsLength(t); commit();
  });
  document.getElementById('propSeats').addEventListener('input', (e) => {
    const t = findTable(state.selectedTableId); if(!t) return;
    t.seatsOverride = clamp(parseInt(e.target.value||1,10), 1, 40); ensureSeatsLength(t); commit();
  });
  document.addEventListener('click', (e) => {
    if (e.target && e.target.id === 'resetSeats'){
      const t = findTable(state.selectedTableId); if(!t) return;
      t.seatsOverride = null; ensureSeatsLength(t); commit();
    }
  });
  document.getElementById('deleteTable').addEventListener('click', () => {
    if (state.selectedTableId) deleteTable(state.selectedTableId);
  });

  document.getElementById('addGuest').addEventListener('click', () => {
    const input = document.getElementById('guestInput');
    addGuest(input.value); input.value=''; input.focus();
  });
  document.getElementById('guestInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter'){ addGuest(e.target.value); e.target.value=''; }
  });
  document.getElementById('guestList').addEventListener('click', (e) => {
    const rm = e.target.closest('[data-rm]');
    if (rm){ removeGuest(rm.dataset.rm); return; }
    const li = e.target.closest('[data-guest]');
    if (li){
      const gId = li.dataset.guest;
      state.armedGuestId = (state.armedGuestId === gId) ? null : gId;
      renderSidebar();
    }
  });

  document.getElementById('saveNow').addEventListener('click', () => { clearTimeout(saveTimer); doSave(); });
  document.getElementById('clearAll').addEventListener('click', () => {
    if (!confirm('Rensa hela planlösningen, alla bord och gäster? Detta går inte att ångra.')) return;
    state.tables = []; state.guests = [];
    state.selectedTableId = null; state.armedGuestId = null;
    commit();
  });
  document.getElementById('printBtn').addEventListener('click', () => window.print());
  window.addEventListener('beforeprint', () => {
    const el = document.getElementById('printSummary');
    let html = '<h2>Bordsöversikt</h2>';
    state.tables.forEach((t, idx) => {
      const n = getSeatCount(t);
      const names = t.seats.filter(Boolean).map(gid => { const g=findGuest(gid); return g?g.name:''; }).filter(Boolean);
      html += `<div class="p-item"><strong>${esc(tableName(t,idx))}</strong> (${names.length}/${n} platser): ${esc(names.join(', ') || '—')}</div>`;
    });
    el.innerHTML = html;
  });

  function downloadBlob(content, filename, mime){
    const blob = new Blob([content], {type: mime});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function extractNames(rows){
    if (!rows || !rows.length) return [];
    const first = rows[0].map(c => String(c==null?'':c).toLowerCase().trim());
    const headerKeywords = ['namn','förnamn','fornamn','efternamn','name','first','last','gäst','gast','guest'];
    const looksHeader = first.some(c => headerKeywords.some(k => c.includes(k)));
    let firstCol=-1, lastCol=-1, nameCol=-1, startIdx=0;
    if (looksHeader){
      first.forEach((c,i) => {
        if (/förnamn|fornamn|first/.test(c)) firstCol=i;
        else if (/efternamn|last/.test(c)) lastCol=i;
        else if (nameCol===-1 && /namn|name|gäst|gast|guest/.test(c)) nameCol=i;
      });
      startIdx = 1;
    }
    const out = [];
    for (let i=startIdx;i<rows.length;i++){
      const row = rows[i];
      if (!row || !row.length) continue;
      let name = '';
      if (firstCol>-1 || lastCol>-1){
        const f = firstCol>-1 ? String(row[firstCol]==null?'':row[firstCol]).trim() : '';
        const l = lastCol>-1 ? String(row[lastCol]==null?'':row[lastCol]).trim() : '';
        name = [f,l].filter(Boolean).join(' ');
      } else if (nameCol>-1){
        name = String(row[nameCol]==null?'':row[nameCol]).trim();
      } else {
        const c0 = String(row[0]==null?'':row[0]).trim();
        const c1 = String(row[1]==null?'':row[1]).trim();
        name = c1 ? `${c0} ${c1}` : c0;
      }
      if (name) out.push(name);
    }
    return out;
  }

  document.getElementById('guestFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const statusEl2 = document.getElementById('importStatus');
    statusEl2.textContent = 'Läser fil…';
    try{
      const buf = await file.arrayBuffer();
      let rows;
      if (/\.csv$/i.test(file.name)){
        const text = new TextDecoder('utf-8').decode(buf);
        rows = text.split(/\r?\n/).filter(r => r.trim().length)
          .map(r => r.split(/[,;\t]/).map(c => c.trim().replace(/^"|"$/g,'')));
      } else {
        const wb = XLSX.read(buf, {type:'array'});
        const sheet = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(sheet, {header:1, raw:false, defval:''});
      }
      const names = extractNames(rows);
      const existing = new Set(state.guests.map(g => g.name.trim().toLowerCase()));
      let added=0, skipped=0;
      names.forEach(n => {
        const key = n.trim().toLowerCase();
        if (!n.trim() || existing.has(key)){ skipped++; return; }
        existing.add(key);
        state.guests.push({id: uid(), name: n.trim()});
        added++;
      });
      commit();
      statusEl2.textContent = `${added} gäster importerade` + (skipped ? `, ${skipped} hoppades över (tomma/dubbletter).` : '.');
    }catch(err){
      statusEl2.textContent = 'Kunde inte läsa filen. Kontrollera att det är en giltig Excel- eller CSV-fil.';
    }
    e.target.value = '';
  });

  document.getElementById('exportSvg').addEventListener('click', () => {
    const color = document.getElementById('exportColor').value;
    const dims = canvasDims();
    let parts = [];
    parts.push(`<rect x="${MARGIN}" y="${MARGIN}" width="${state.room.width}" height="${state.room.length}" fill="none" stroke="${color}" stroke-width="8" rx="2"/>`);
    state.tables.forEach((t, idx) => {
      ensureSeatsLength(t);
      let shape = '';
      if (t.type === 'round') shape = `<circle cx="0" cy="0" r="${t.diameter/2}" fill="none" stroke="${color}" stroke-width="3"/>`;
      else shape = `<rect x="${-t.w/2}" y="${-t.h/2}" width="${t.w}" height="${t.h}" rx="6" fill="none" stroke="${color}" stroke-width="3"/>`;
      const label = `<text x="0" y="4" font-size="16" text-anchor="middle" font-family="Space Grotesk, sans-serif" font-weight="600" fill="${color}">${esc(tableName(t,idx))}</text>`;
      const seatPositions = getSeatLocalPositions(t);
      let seats = '';
      seatPositions.forEach((p, si) => {
        const guestId = t.seats[si];
        const guest = guestId ? findGuest(guestId) : null;
        if (guest){
          const initials = guest.name.trim().slice(0,2).toUpperCase();
          seats += `<circle cx="${p.x}" cy="${p.y}" r="${CHAIR_R}" fill="${color}"/>
            <text x="${p.x}" y="${p.y+1}" font-size="11" text-anchor="middle" dominant-baseline="middle" font-family="IBM Plex Mono, monospace" fill="#ffffff">${esc(initials)}</text>`;
        } else {
          seats += `<circle cx="${p.x}" cy="${p.y}" r="${CHAIR_R}" fill="none" stroke="${color}" stroke-width="2"/>`;
        }
      });
      const rot = t.type === 'rect' ? (t.rotation||0) : 0;
      parts.push(`<g transform="translate(${t.x},${t.y}) rotate(${rot})">${shape}${seats}<g transform="rotate(${-rot})">${label}</g></g>`);
    });
    const svgStr = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${dims.w} ${dims.h}" width="${dims.w}" height="${dims.h}">
${parts.join('\n')}
</svg>`;
    downloadBlob(svgStr, 'bordsplacering.svg', 'image/svg+xml');
  });

  document.getElementById('exportNameList').addEventListener('click', () => {
    const color = document.getElementById('exportColor2').value;
    const sorted = sortedGuestsByLastName();
    const lineH = 30, width = 700, top = 74;
    const height = top + Math.max(1, sorted.length) * lineH + 30;
    let body = `<text x="30" y="42" font-family="Space Grotesk, sans-serif" font-size="24" font-weight="700" fill="${color}">Gästlista — bordsplacering</text>`;
    if (!sorted.length){
      body += `<text x="30" y="${top+24}" font-family="Space Grotesk, sans-serif" font-size="14" fill="${color}">Inga gäster tillagda än.</text>`;
    }
    sorted.forEach((g, i) => {
      const y = top + i*lineH + 18;
      const tn = guestTableInfo(g.id) || 'Ej placerad';
      const sn = splitName(g.name);
      const display = sn.first ? `${sn.last}, ${sn.first}` : sn.last;
      body += `<text x="30" y="${y}" font-family="IBM Plex Mono, monospace" font-size="14" fill="${color}">${esc(display)}</text>`;
      body += `<text x="${width-30}" y="${y}" font-family="IBM Plex Mono, monospace" font-size="14" text-anchor="end" fill="${color}">${esc(tn)}</text>`;
    });
    const svgStr = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="${width}" height="${height}">
${body}
</svg>`;
    downloadBlob(svgStr, 'gastlista.svg', 'image/svg+xml');
  });

  document.getElementById('guestCount').value = 80;
  document.getElementById('acRoundD').value = 150;
  document.getElementById('acRectW').value = 180;
  document.getElementById('acRectH').value = 75;
  document.getElementById('headSeats').value = 10;

  loadState();
})();
</script>
</body>
</html>
