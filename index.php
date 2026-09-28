<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Kovil Kanakku — Temple Trust Accounts</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Yatra+One&family=Mukta:wght@400;500;600;700&family=Cormorant+Garamond:ital,wght@0,500;0,600;1,500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --maroon-950:#2b0a0f;
    --maroon-900:#3d0f16;
    --maroon-800:#5a1520;
    --maroon-700:#7a1e2a;
    --gold-500:#c9a227;
    --gold-400:#dcb94a;
    --gold-300:#ecd28a;
    --cream-100:#fbf4e2;
    --cream-200:#f3e6c4;
    --cream-50:#fffdf7;
    --ink:#2c1a10;
    --ink-soft:#5b4635;
    --green-ok:#1f6b3a;
    --red-bad:#a12626;
    --shadow: 0 10px 30px rgba(43,10,15,.25);
    --radius: 14px;
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;height:100%;}
  body{
    font-family:'Mukta',sans-serif;
    background:
      radial-gradient(circle at 15% 10%, rgba(201,162,39,.10), transparent 40%),
      radial-gradient(circle at 90% 85%, rgba(122,30,42,.12), transparent 45%),
      var(--cream-100);
    color:var(--ink);
    min-height:100vh;
  }
  .app{display:flex; min-height:100vh;}

  /* ---------- Sidebar ---------- */
  .sidebar{
    width:250px; flex-shrink:0;
    background: linear-gradient(180deg, var(--maroon-950), var(--maroon-800) 60%, var(--maroon-900));
    color:var(--cream-100);
    display:flex; flex-direction:column;
    position:relative;
    box-shadow: 4px 0 20px rgba(0,0,0,.25);
  }
  .sidebar::after{
    content:"";
    position:absolute; top:0; right:0; bottom:0; width:4px;
    background:linear-gradient(180deg, var(--gold-300), var(--gold-500), var(--gold-300));
  }
  .brand{
    padding:26px 22px 18px;
    border-bottom:1px solid rgba(220,185,74,.25);
  }
  .brand .mark{
    font-family:'Yatra One',cursive;
    font-size:28px;
    color:var(--gold-300);
    letter-spacing:.5px;
    display:flex; align-items:center; gap:10px;
  }
  .brand .sub{
    font-family:'Cormorant Garamond',serif;
    font-style:italic;
    color:var(--cream-200);
    opacity:.8;
    font-size:13.5px;
    margin-top:4px;
  }
  .brand-logo{
    width:50px; height:50px; margin:0 auto 8px; border-radius:8px;
    background:rgba(243,230,196,.15); display:flex; align-items:center; justify-content:center;
    overflow:hidden;
  }
  .brand-logo img{
    width:100%; height:100%; object-fit:cover;
  }
  .lang-toggle{
    display:flex; gap:4px; padding:8px 12px; margin:8px 0;
    background:rgba(220,185,74,.15); border-radius:8px;
  }
  .lang-toggle button{
    all:unset; cursor:pointer; flex:1; padding:6px 8px;
    font-size:11px; font-weight:600; border-radius:6px;
    color:var(--cream-200); transition:.15s; border:1px solid transparent;
  }
  .lang-toggle button.active{
    background:var(--gold-400); color:var(--maroon-950);
    border:1px solid var(--gold-500);
  }
  nav{ padding:16px 12px; display:flex; flex-direction:column; gap:4px; flex:1;}
  nav button{
    all:unset; cursor:pointer;
    padding:12px 14px;
    border-radius:10px;
    font-size:14.5px;
    font-weight:600;
    color:var(--cream-200);
    display:flex; align-items:center; gap:10px;
    transition:.18s;
  }
  nav button svg{width:18px;height:18px; opacity:.85; flex-shrink:0;}
  nav button:hover{background:rgba(236,210,138,.10); color:var(--gold-300);}
  nav button.active{
    background:linear-gradient(90deg, rgba(201,162,39,.28), rgba(201,162,39,.05));
    color:var(--gold-300);
    box-shadow: inset 3px 0 0 var(--gold-400);
  }
  .sidebar-foot{
    padding:16px 22px 22px;
    font-size:11.5px;
    color:rgba(243,230,196,.55);
    border-top:1px solid rgba(220,185,74,.15);
  }

  /* ---------- Main ---------- */
  main{ flex:1; padding:34px 40px; max-width:1180px; }
  .page-head{ display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:26px; flex-wrap:wrap; gap:14px;}
  .page-head h1{
    font-family:'Yatra One',cursive; font-weight:400;
    font-size:32px; color:var(--maroon-800); margin:0 0 4px;
  }
  .page-head p{ margin:0; color:var(--ink-soft); font-size:14px; }
  .btn{
    all:unset; cursor:pointer; font-family:'Mukta',sans-serif;
    background:linear-gradient(180deg, var(--gold-400), var(--gold-500));
    color:var(--maroon-950); font-weight:700; font-size:14px;
    padding:11px 20px; border-radius:10px;
    box-shadow:0 4px 14px rgba(201,162,39,.35);
    transition:.15s; border:1px solid rgba(43,10,15,.15);
  }
  .btn:hover{ transform:translateY(-1px); box-shadow:0 6px 18px rgba(201,162,39,.5); }
  .btn.secondary{
    background:transparent; color:var(--maroon-800);
    border:1.5px solid var(--maroon-700); box-shadow:none;
  }
  .btn.danger{ background:linear-gradient(180deg,#c94848,#a12626); color:#fff; box-shadow:0 4px 14px rgba(161,38,38,.3);}
  .btn.small{ padding:7px 13px; font-size:12.5px; border-radius:8px;}

  /* cards / stat grid */
  .grid-stats{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:18px; margin-bottom:30px;}
  .stat-card{
    background:var(--cream-50); border-radius:var(--radius);
    padding:20px 22px; box-shadow:var(--shadow); position:relative; overflow:hidden;
    border:1px solid var(--cream-200);
  }
  .stat-card::before{
    content:""; position:absolute; top:-30px; right:-30px; width:90px; height:90px;
    border-radius:50%; background:radial-gradient(circle, rgba(201,162,39,.18), transparent 70%);
  }
  .stat-card .label{ font-size:12.5px; text-transform:uppercase; letter-spacing:.08em; color:var(--ink-soft); font-weight:600;}
  .stat-card .value{ font-family:'Cormorant Garamond',serif; font-size:32px; font-weight:600; margin-top:6px; color:var(--maroon-800);}
  .stat-card.income .value{ color:var(--green-ok); }
  .stat-card.expense .value{ color:var(--red-bad); }

  .panel{
    background:var(--cream-50); border-radius:var(--radius); box-shadow:var(--shadow);
    border:1px solid var(--cream-200); padding:22px 24px; margin-bottom:24px;
  }
  .panel h2{ font-family:'Cormorant Garamond',serif; font-size:21px; color:var(--maroon-800); margin:0 0 14px; font-weight:600;}

  table{ width:100%; border-collapse:collapse; font-size:13.8px;}
  thead th{
    text-align:left; padding:10px 10px; color:var(--ink-soft);
    text-transform:uppercase; font-size:11px; letter-spacing:.06em;
    border-bottom:2px solid var(--gold-300);
  }
  tbody td{ padding:11px 10px; border-bottom:1px solid var(--cream-200); vertical-align:middle;}
  tbody tr:hover{ background:rgba(201,162,39,.06); }
  .pill{ display:inline-block; padding:3px 10px; border-radius:20px; font-size:11.5px; font-weight:700;}
  .pill.income{ background:#e2f3e6; color:var(--green-ok);}
  .pill.expense{ background:#fbe4e4; color:var(--red-bad);}
  .row-actions{ display:flex; gap:6px; }
  .icon-btn{
    all:unset; cursor:pointer; width:28px; height:28px; display:flex; align-items:center; justify-content:center;
    border-radius:7px; color:var(--maroon-700);
  }
  .icon-btn:hover{ background:var(--cream-200); }
  .icon-btn svg{ width:15px; height:15px; }

  .empty{ text-align:center; padding:40px 10px; color:var(--ink-soft); }
  .empty svg{ width:46px; height:46px; opacity:.4; margin-bottom:10px;}

  /* filters/toolbar */
  .toolbar{ display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:16px;}
  .toolbar input, .toolbar select{
    font-family:'Mukta'; padding:9px 12px; border-radius:8px; border:1.5px solid var(--cream-200);
    background:var(--cream-50); font-size:13.5px; color:var(--ink);
  }
  .toolbar input:focus, .toolbar select:focus{ outline:none; border-color:var(--gold-400); }
  .spacer{ flex:1; }

  /* modal */
  .overlay{
    position:fixed; inset:0; background:rgba(43,10,15,.55); backdrop-filter:blur(2px);
    display:none; align-items:flex-start; justify-content:center; padding:5vh 20px; z-index:100; overflow-y:auto;
  }
  .overlay.show{ display:flex; }
  .modal{
    background:var(--cream-50); border-radius:16px; padding:28px 30px; width:100%; max-width:520px;
    box-shadow:0 20px 60px rgba(0,0,0,.4); border:1px solid var(--gold-300); animation:pop .18s ease;
  }
  @keyframes pop{ from{ transform:translateY(-8px); opacity:0;} to{ transform:translateY(0); opacity:1;} }
  .modal h3{ font-family:'Yatra One',cursive; font-weight:400; color:var(--maroon-800); font-size:22px; margin:0 0 18px;}
  .field{ margin-bottom:14px; }
  .field label{ display:block; font-size:12.5px; font-weight:700; color:var(--ink-soft); margin-bottom:5px; text-transform:uppercase; letter-spacing:.04em;}
  .field input, .field select, .field textarea{
    width:100%; padding:10px 12px; border-radius:9px; border:1.5px solid var(--cream-200);
    font-family:'Mukta'; font-size:14px; background:#fff; color:var(--ink);
  }
  .field input:focus, .field select:focus, .field textarea:focus{ outline:none; border-color:var(--gold-400); }
  .field-row{ display:grid; grid-template-columns:1fr 1fr; gap:12px; }
  .modal-actions{ display:flex; justify-content:flex-end; gap:10px; margin-top:22px;}

  /* receipt */
  .receipt{
    max-width:480px; margin:0 auto; background:var(--cream-50); border:2px solid var(--gold-400);
    border-radius:12px; padding:30px 34px; position:relative;
  }
  .receipt::before, .receipt::after{
    content:""; position:absolute; left:16px; right:16px; height:1px; background:repeating-linear-gradient(90deg, var(--gold-400) 0 6px, transparent 6px 12px);
  }
  .receipt::before{ top:8px; } .receipt::after{ bottom:8px; }
  .receipt-head{ text-align:center; margin-bottom:18px; }
  .receipt-headbar{ display:flex; align-items:center; gap:10px; margin-bottom:18px; }
  .receipt-headbar .logo-slot{ width:70px; flex-shrink:0; display:flex; align-items:center; justify-content:flex-start; }
  .receipt-headbar .logo-slot img{ max-width:70px; max-height:60px; object-fit:contain; }
  .receipt-headbar .temple-info{ flex:1; text-align:center; }
  .receipt-head .temple{ font-family:'Yatra One',cursive; font-size:24px; color:var(--maroon-800); }
  .receipt-head .addr{ font-size:12px; color:var(--ink-soft); margin-top:2px;}
  .receipt-headbar .temple{ font-family:'Yatra One',cursive; font-size:24px; color:var(--maroon-800); }
  .receipt-headbar .addr{ font-size:12px; color:var(--ink-soft); margin-top:2px;}
  .receipt-title{ text-align:center; font-weight:700; letter-spacing:.1em; text-transform:uppercase; font-size:12.5px; color:var(--gold-500); margin:14px 0; border-top:1px dashed var(--gold-300); border-bottom:1px dashed var(--gold-300); padding:6px 0;}
  .receipt-grid{ display:grid; grid-template-columns:auto 1fr; gap:7px 14px; font-size:13.5px; margin-bottom:16px;}
  .receipt-grid .k{ color:var(--ink-soft); font-weight:600;}
  .receipt-amount{ text-align:center; font-family:'Cormorant Garamond',serif; font-size:30px; color:var(--green-ok); font-weight:700; margin:14px 0;}
  .receipt-foot{ text-align:center; font-size:11.5px; color:var(--ink-soft); margin-top:20px; font-style:italic; }

  select#txnType.income-mode{}

  /* report bars */
  .bar-row{ display:grid; grid-template-columns:140px 1fr 90px; align-items:center; gap:10px; margin-bottom:10px; font-size:13px;}
  .bar-track{ background:var(--cream-200); border-radius:6px; height:14px; overflow:hidden; }
  .bar-fill{ height:100%; border-radius:6px; background:linear-gradient(90deg, var(--gold-400), var(--maroon-700)); }

  .toast{
    position:fixed; bottom:24px; left:50%; transform:translateX(-50%) translateY(20px);
    background:var(--maroon-900); color:var(--gold-300); padding:12px 22px; border-radius:10px;
    font-size:13.5px; font-weight:600; box-shadow:0 8px 24px rgba(0,0,0,.35);
    opacity:0; pointer-events:none; transition:.25s; z-index:200; border:1px solid var(--gold-400);
  }
  .toast.show{ opacity:1; transform:translateX(-50%) translateY(0); }

  @media (max-width:820px){
    .app{ flex-direction:column; }
    .sidebar{ width:100%; }
    nav{ flex-direction:row; overflow-x:auto; }
    main{ padding:22px 18px; }
  }

  @media print{
    body *{ visibility:hidden; }
    #printArea, #printArea *{ visibility:visible; }
    #printArea{ position:absolute; top:0; left:0; width:100%; }
    .print-hide{ display:none !important; }
  }
  .print-hide{ display:block; }

</style>
</head>
<body>
<div class="app">

  <aside class="sidebar">
    <div class="brand">
      <div class="brand-logo" id="sidebarLogo"></div>
      <div class="mark" id="brandName">Kovil Kanakku</div>
      <div class="sub" id="brandSub">Temple Trust Account Register</div>
    </div>
    <div class="lang-toggle">
      <button onclick="setAppLang('en')" id="langBtn-en" class="active">EN</button>
      <button onclick="setAppLang('ta')" id="langBtn-ta">TA</button>
    </div>
    <nav id="nav"></nav>
    <div class="sidebar-foot">Shared trust ledger — visible to all trustees &amp; staff using this app.</div>
  </aside>

  <main id="main"></main>
</div>

<div class="overlay" id="overlay"><div class="modal" id="modalBody"></div></div>
<div class="toast" id="toast"></div>

<script>
/* ============ STATE ============ */
let state = {
  settings: { templeName:'Sri Temple Trust', address:'', regNo:'', phone:'', logo:null },
  accounts: [],
  donors: [],
  members: [],
  transactions: []
};
let currentView = 'dashboard';
let loaded = false;

const fmt = n => '₹' + Number(n||0).toLocaleString('en-IN', {maximumFractionDigits:2});
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2,7);
const todayStr = () => new Date().toISOString().slice(0,10);

async function loadState(){
  try{
    const res = await fetch('load_state.php');
    const json = await res.json();
    if(json.success && json.state){
      state = Object.assign(state, json.state);
    } else if(!json.success){
      showToast('⚠ ' + (json.error || 'Could not load data from server'));
    }
  }catch(e){
    showToast('⚠ Could not reach server — check XAMPP (Apache + MySQL running?)');
    console.error(e);
  }
  loaded = true;
}
async function saveState(){
  try{
    const res = await fetch('save_state.php', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(state)
    });
    const json = await res.json();
    if(!json.success){ showToast('⚠ Could not save — please retry'); }
  }catch(e){
    showToast('⚠ Could not save — please retry');
    console.error(e);
  }
}
function showToast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(()=>t.classList.remove('show'), 2200);
}

/* ============ NAV ============ */
document.getElementById('nav').addEventListener('click', e=>{
  const btn = e.target.closest('button[data-view]');
  if(!btn) return;
  currentView = btn.dataset.view;
  render();
});

function setActiveNav(){
  document.querySelectorAll('#nav button').forEach(b=>{
    b.classList.toggle('active', b.dataset.view===currentView);
  });
}

function renderNav(){
  const navItems = [
    {view:'dashboard', label:t('dashboard'), icon:'📊'},
    {view:'transactions', label:t('transactions'), icon:'📋'},
    {view:'accounts', label:t('accounts'), icon:'🏦'},
    {view:'donors', label:t('donors'), icon:'🙏'},
    {view:'members', label:t('members'), icon:'👥'},
    {view:'receipts', label:t('receipts'), icon:'📄'},
    {view:'reports', label:t('reports'), icon:'📊'},
    {view:'settings', label:t('settings'), icon:'⚙️'},
  ];
  const nav = document.getElementById('nav');
  nav.innerHTML = navItems.map(item => 
    `<button data-view="${item.view}" onclick="currentView='${item.view}';render()">${item.icon}${item.label}</button>`
  ).join('');
  setActiveNav();
}

/* ============ MODAL HELPERS ============ */
function openModal(html){
  document.getElementById('modalBody').innerHTML = html;
  document.getElementById('overlay').classList.add('show');
}
function closeModal(){ document.getElementById('overlay').classList.remove('show'); }
document.getElementById('overlay').addEventListener('click', e=>{
  if(e.target.id==='overlay') closeModal();
});

/* ============ RENDER ROOT ============ */
function render(){
  renderNav();
  updateBrand();
  document.getElementById('langBtn-en').classList.toggle('active', appLang==='en');
  document.getElementById('langBtn-ta').classList.toggle('active', appLang==='ta');
  const main = document.getElementById('main');
  if(currentView==='dashboard') return renderDashboard(main);
  if(currentView==='transactions') return renderTransactions(main);
  if(currentView==='accounts') return renderAccounts(main);
  if(currentView==='donors') return renderDonors(main);
  if(currentView==='members') return renderMembers(main);
  if(currentView==='receipts') return renderReceipts(main);
  if(currentView==='reports') return renderReports(main);
  if(currentView==='settings') return renderSettings(main);
}

/* ============ DASHBOARD ============ */
function renderDashboard(main){
  const income = state.transactions.filter(t=>t.type==='income').reduce((s,t)=>s+Number(t.amount),0);
  const expense = state.transactions.filter(t=>t.type==='expense').reduce((s,t)=>s+Number(t.amount),0);
  const balance = income - expense;
  const recent = [...state.transactions].sort((a,b)=>b.date.localeCompare(a.date)||b.id.localeCompare(a.id)).slice(0,7);

  main.innerHTML = `
    <div class="page-head">
      <div><h1>${t('dashboard')}</h1><p>${escapeHtml(state.settings.templeName || 'Temple Trust')} — ${t('dashboardSub')}</p></div>
    </div>
    <div class="grid-stats">
      <div class="stat-card income"><div class="label">Total Income</div><div class="value">${fmt(income)}</div></div>
      <div class="stat-card expense"><div class="label">Total Expenses</div><div class="value">${fmt(expense)}</div></div>
      <div class="stat-card"><div class="label">Net Balance</div><div class="value">${fmt(balance)}</div></div>
      <div class="stat-card"><div class="label">Fund Accounts</div><div class="value">${state.accounts.length}</div></div>
      <div class="stat-card"><div class="label">Registered Donors</div><div class="value">${state.donors.length}</div></div>
    </div>

    <div class="panel">
      <h2>Recent Activity</h2>
      ${recent.length ? `
      <table>
        <thead><tr><th>Date</th><th>Type</th><th>Category</th><th>Account</th><th>Donor</th><th style="text-align:right">Amount</th></tr></thead>
        <tbody>
          ${recent.map(t=>`
            <tr>
              <td>${fmtDate(t.date)}</td>
              <td><span class="pill ${t.type}">${t.type==='income'?'Income':'Expense'}</span></td>
              <td>${escapeHtml(t.category||'-')}</td>
              <td>${escapeHtml(accountName(t.accountId))}</td>
              <td>${escapeHtml(donorName(t.donorId) || memberName(t.memberId) || '-')}</td>
              <td style="text-align:right;font-weight:700;color:${t.type==='income'?'var(--green-ok)':'var(--red-bad)'}">${t.type==='income'?'+':'-'}${fmt(t.amount)}</td>
            </tr>`).join('')}
        </tbody>
      </table>` : emptyState('No transactions recorded yet.')}
    </div>

    <div class="panel">
      <h2>Balance by Fund Account</h2>
      ${state.accounts.length ? state.accounts.map(a=>{
        const bal = accountBalance(a.id);
        const max = Math.max(1,...state.accounts.map(x=>Math.abs(accountBalance(x.id))));
        const pct = Math.min(100, Math.abs(bal)/max*100);
        return `<div class="bar-row">
          <div>${escapeHtml(a.name)}</div>
          <div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div>
          <div style="text-align:right;font-weight:700;">${fmt(bal)}</div>
        </div>`;
      }).join('') : emptyState('No fund accounts yet. Add one under "Fund Accounts".')}
    </div>
  `;
}

function accountBalance(accId){
  const acc = state.accounts.find(a=>a.id===accId);
  const opening = acc ? Number(acc.openingBalance||0) : 0;
  const inSum = state.transactions.filter(t=>t.accountId===accId && t.type==='income').reduce((s,t)=>s+Number(t.amount),0);
  const outSum = state.transactions.filter(t=>t.accountId===accId && t.type==='expense').reduce((s,t)=>s+Number(t.amount),0);
  return opening + inSum - outSum;
}
function accountName(id){ const a = state.accounts.find(x=>x.id===id); return a? a.name : '—'; }
function donorName(id){ const d = state.donors.find(x=>x.id===id); return d? d.name : ''; }
function memberName(id){ const m = state.members.find(x=>x.id===id); return m? m.name : ''; }
function fmtDate(d){ if(!d) return '-'; const dt = new Date(d+'T00:00:00'); return dt.toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'}); }
function escapeHtml(s){ return String(s??'').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function emptyState(msg){
  return `<div class="empty"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="12" cy="12" r="9"/><path d="M9 10h.01M15 10h.01M8 15c1.2 1 2.6 1.5 4 1.5s2.8-.5 4-1.5"/></svg><div>${msg}</div></div>`;
}

/* ============ ACCOUNTS ============ */
function renderAccounts(main){
  main.innerHTML = `
    <div class="page-head">
      <div><h1>${t('accounts')}</h1><p>${t('accountsSub')}</p></div>
      <button class="btn" id="addAccBtn">${t('newFundAccount')}</button>
    </div>
    <div class="panel">
      ${state.accounts.length ? `
      <table>
        <thead><tr><th>Name</th><th>Type</th><th>Opening Balance</th><th>Current Balance</th><th></th></tr></thead>
        <tbody>
          ${state.accounts.map(a=>`
            <tr>
              <td style="font-weight:700;">${escapeHtml(a.name)}</td>
              <td>${escapeHtml(a.type||'-')}</td>
              <td>${fmt(a.openingBalance)}</td>
              <td style="font-weight:700;color:var(--maroon-800)">${fmt(accountBalance(a.id))}</td>
              <td class="row-actions">
                <button class="icon-btn" data-edit-acc="${a.id}" title="Edit">${editIcon()}</button>
                <button class="icon-btn" data-del-acc="${a.id}" title="Delete">${trashIcon()}</button>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>` : emptyState('No fund accounts yet — click "New Fund Account" to create one.')}
    </div>
  `;
  document.getElementById('addAccBtn').onclick = ()=>openAccountModal();
  main.querySelectorAll('[data-edit-acc]').forEach(b=>b.onclick=()=>openAccountModal(b.dataset.editAcc));
  main.querySelectorAll('[data-del-acc]').forEach(b=>b.onclick=()=>deleteAccount(b.dataset.delAcc));
}
function openAccountModal(id){
  const acc = id ? state.accounts.find(a=>a.id===id) : null;
  openModal(`
    <h3>${acc?'Edit':'New'} Fund Account</h3>
    <div class="field"><label>Account Name</label><input id="fAccName" value="${acc?escapeHtml(acc.name):''}" placeholder="e.g. General Fund, Hundi Collection"></div>
    <div class="field"><label>Type</label>
      <select id="fAccType">
        ${['General','Hundi / Donation Box','Annadhanam','Renovation / Construction','Festival Fund','Education / Trust Charity','Other'].map(t=>`<option ${acc&&acc.type===t?'selected':''}>${t}</option>`).join('')}
      </select>
    </div>
    <div class="field"><label>Opening Balance (₹)</label><input id="fAccOpen" type="number" step="0.01" value="${acc?acc.openingBalance:0}"></div>
    <div class="modal-actions">
      <button class="btn secondary" id="cancelBtn">Cancel</button>
      <button class="btn" id="saveAccBtn">Save Account</button>
    </div>
  `);
  document.getElementById('cancelBtn').onclick = closeModal;
  document.getElementById('saveAccBtn').onclick = async ()=>{
    const name = document.getElementById('fAccName').value.trim();
    if(!name){ showToast('Please enter an account name'); return; }
    const type = document.getElementById('fAccType').value;
    const openingBalance = Number(document.getElementById('fAccOpen').value||0);
    if(acc){ acc.name=name; acc.type=type; acc.openingBalance=openingBalance; }
    else{ state.accounts.push({id:uid(), name, type, openingBalance}); }
    await saveState();
    closeModal(); showToast('Fund account saved'); render();
  };
}
async function deleteAccount(id){
  const used = state.transactions.some(t=>t.accountId===id);
  if(used && !confirm('This account has transactions linked to it. Delete anyway?')) return;
  if(!used && !confirm('Delete this fund account?')) return;
  state.accounts = state.accounts.filter(a=>a.id!==id);
  await saveState(); showToast('Account deleted'); render();
}

/* ============ DONORS ============ */
function renderDonors(main){
  main.innerHTML = `
    <div class="page-head">
      <div><h1>${t('donors')}</h1><p>${t('donorList')}</p></div>
      <button class="btn" id="addDonorBtn">+ ${t('donors')}</button>
    </div>
    <div class="toolbar"><input id="donorSearch" placeholder="Search donors by name or phone..." style="min-width:260px"></div>
    <div class="panel" id="donorPanel"></div>
  `;
  const draw = (q='') => {
    const list = state.donors.filter(d => !q || (d.name+d.phone).toLowerCase().includes(q.toLowerCase()));
    document.getElementById('donorPanel').innerHTML = list.length ? `
      <table>
        <thead><tr><th>Name</th><th>Phone</th><th>PAN</th><th style="text-align:right">Total Donated</th><th></th></tr></thead>
        <tbody>
          ${list.map(d=>{
            const total = state.transactions.filter(t=>t.donorId===d.id && t.type==='income').reduce((s,t)=>s+Number(t.amount),0);
            return `<tr>
              <td style="font-weight:700;">${escapeHtml(d.name)}</td>
              <td>${escapeHtml(d.phone||'-')}</td>
              <td>${escapeHtml(d.pan||'-')}</td>
              <td style="text-align:right;font-weight:700;color:var(--green-ok)">${fmt(total)}</td>
              <td class="row-actions">
                <button class="icon-btn" data-edit-donor="${d.id}">${editIcon()}</button>
                <button class="icon-btn" data-del-donor="${d.id}">${trashIcon()}</button>
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>` : emptyState('No donors found.');
    document.getElementById('donorPanel').querySelectorAll('[data-edit-donor]').forEach(b=>b.onclick=()=>openDonorModal(b.dataset.editDonor));
    document.getElementById('donorPanel').querySelectorAll('[data-del-donor]').forEach(b=>b.onclick=()=>deleteDonor(b.dataset.delDonor));
  };
  draw();
  document.getElementById('addDonorBtn').onclick = ()=>openDonorModal();
  document.getElementById('donorSearch').oninput = e=>draw(e.target.value);
}
function openDonorModal(id){
  const d = id ? state.donors.find(x=>x.id===id) : null;
  openModal(`
    <h3>${d?'Edit':'New'} Donor</h3>
    <div class="field"><label>Full Name</label><input id="fDName" value="${d?escapeHtml(d.name):''}"></div>
    <div class="field-row">
      <div class="field"><label>Phone</label><input id="fDPhone" value="${d?escapeHtml(d.phone||''):''}"></div>
      <div class="field"><label>PAN (optional)</label><input id="fDPan" value="${d?escapeHtml(d.pan||''):''}"></div>
    </div>
    <div class="field"><label>Email</label><input id="fDEmail" value="${d?escapeHtml(d.email||''):''}"></div>
    <div class="field"><label>Address</label><textarea id="fDAddr" rows="2">${d?escapeHtml(d.address||''):''}</textarea></div>
    <div class="modal-actions">
      <button class="btn secondary" id="cancelBtn">Cancel</button>
      <button class="btn" id="saveDonorBtn">Save Donor</button>
    </div>
  `);
  document.getElementById('cancelBtn').onclick = closeModal;
  document.getElementById('saveDonorBtn').onclick = async ()=>{
    const name = document.getElementById('fDName').value.trim();
    if(!name){ showToast('Please enter a name'); return; }
    const payload = {
      name,
      phone: document.getElementById('fDPhone').value.trim(),
      pan: document.getElementById('fDPan').value.trim(),
      email: document.getElementById('fDEmail').value.trim(),
      address: document.getElementById('fDAddr').value.trim(),
    };
    if(d) Object.assign(d, payload);
    else state.donors.push({id:uid(), ...payload});
    await saveState(); closeModal(); showToast('Donor saved'); render();
  };
}
async function deleteDonor(id){
  if(!confirm('Delete this donor record?')) return;
  state.donors = state.donors.filter(d=>d.id!==id);
  await saveState(); showToast('Donor deleted'); render();
}

/* ============ MEMBERS ============ */
function renderMembers(main){
  main.innerHTML = `
    <div class="page-head">
      <div><h1>${t('members')}</h1><p>${t('memberList')}</p></div>
      <button class="btn" id="addMemberBtn">+ ${t('members')}</button>
    </div>
    <div class="toolbar"><input id="memberSearch" placeholder="Search members by name or phone..." style="min-width:260px"></div>
    <div class="panel" id="memberPanel"></div>
  `;
  const draw = (q='') => {
    const list = state.members.filter(m => !q || (m.name+m.phone).toLowerCase().includes(q.toLowerCase()));
    document.getElementById('memberPanel').innerHTML = list.length ? `
      <table>
        <thead><tr><th>Name</th><th>Phone</th><th>Role</th><th></th></tr></thead>
        <tbody>
          ${list.map(m=>`<tr>
            <td style="font-weight:700;">${escapeHtml(m.name)}</td>
            <td>${escapeHtml(m.phone||'-')}</td>
            <td>${escapeHtml(m.role||'Member')}</td>
            <td class="row-actions">
              <button class="icon-btn" data-edit-member="${m.id}">${editIcon()}</button>
              <button class="icon-btn" data-del-member="${m.id}">${trashIcon()}</button>
            </td>
          </tr>`).join('')}
        </tbody>
      </table>` : emptyState('No members found.');
    document.getElementById('memberPanel').querySelectorAll('[data-edit-member]').forEach(b=>b.onclick=()=>openMemberModal(b.dataset.editMember));
    document.getElementById('memberPanel').querySelectorAll('[data-del-member]').forEach(b=>b.onclick=()=>deleteMember(b.dataset.delMember));
  };
  draw();
  document.getElementById('addMemberBtn').onclick = ()=>openMemberModal();
  document.getElementById('memberSearch').oninput = e=>draw(e.target.value);
}

function openMemberModal(id){
  const m = id ? state.members.find(x=>x.id===id) : null;
  openModal(`
    <h3>${m?'Edit':'New'} Member</h3>
    <div class="field"><label>Full Name</label><input id="fMName" value="${m?escapeHtml(m.name):''}"></div>
    <div class="field-row">
      <div class="field"><label>Phone</label><input id="fMPhone" value="${m?escapeHtml(m.phone||''):''}"></div>
      <div class="field"><label>Role</label><input id="fMRole" value="${m?escapeHtml(m.role||''):''}"></div>
    </div>
    <div class="field"><label>Email</label><input id="fMEmail" value="${m?escapeHtml(m.email||''):''}"></div>
    <div class="field"><label>Address</label><textarea id="fMAddr" rows="2">${m?escapeHtml(m.address||''):''}</textarea></div>
    <div class="modal-actions">
      <button class="btn secondary" id="cancelBtn">Cancel</button>
      <button class="btn" id="saveMemberBtn">Save Member</button>
    </div>
  `);
  document.getElementById('cancelBtn').onclick = closeModal;
  document.getElementById('saveMemberBtn').onclick = async ()=>{
    const name = document.getElementById('fMName').value.trim();
    if(!name){ showToast('Please enter a name'); return; }
    const payload = {
      name,
      phone: document.getElementById('fMPhone').value.trim(),
      email: document.getElementById('fMEmail').value.trim(),
      role: document.getElementById('fMRole').value.trim(),
      address: document.getElementById('fMAddr').value.trim(),
    };
    if(m) Object.assign(m, payload);
    else state.members.push({id:uid(), ...payload});
    await saveState(); closeModal(); showToast('Member saved'); render();
  };
}

async function deleteMember(id){
  if(!confirm('Delete this member record?')) return;
  state.members = state.members.filter(m=>m.id!==id);
  await saveState(); showToast('Member deleted'); render();
}

/* ============ TRANSACTIONS ============ */
function renderTransactions(main){
  main.innerHTML = `
    <div class="page-head">
      <div><h1>${t('transactions')}</h1><p>${t('transactionsSub')}</p></div>
      <div style="display:flex;gap:10px;">
        <button class="btn secondary" id="addExpBtn">${t('addExpense')}</button>
        <button class="btn" id="addIncBtn">${t('addIncome')}</button>
      </div>
    </div>
    <div class="toolbar">
      <select id="filterType"><option value="">All Types</option><option value="income">Income</option><option value="expense">Expense</option></select>
      <select id="filterAcc"><option value="">All Accounts</option>${state.accounts.map(a=>`<option value="${a.id}">${escapeHtml(a.name)}</option>`).join('')}</select>
      <input id="filterSearch" placeholder="Search category / notes...">
      <div class="spacer"></div>
    </div>
    <div class="panel" id="txnPanel"></div>
  `;
  const draw = ()=>{
    const ty = document.getElementById('filterType').value;
    const ac = document.getElementById('filterAcc').value;
    const q = document.getElementById('filterSearch').value.toLowerCase();
    let list = [...state.transactions].sort((a,b)=>b.date.localeCompare(a.date)||b.id.localeCompare(a.id));
    if(ty) list = list.filter(t=>t.type===ty);
    if(ac) list = list.filter(t=>t.accountId===ac);
    if(q) list = list.filter(t=>(t.category+t.notes).toLowerCase().includes(q));
    document.getElementById('txnPanel').innerHTML = list.length ? `
      <table>
        <thead><tr><th>Date</th><th>Type</th><th>Category</th><th>Account</th><th>Donor/Payee</th><th>Mode</th><th style="text-align:right">Amount</th><th></th></tr></thead>
        <tbody>
          ${list.map(t=>`
            <tr>
              <td>${fmtDate(t.date)}</td>
              <td><span class="pill ${t.type}">${t.type==='income'?'Income':'Expense'}</span></td>
              <td>${escapeHtml(t.category||'-')}</td>
              <td>${escapeHtml(accountName(t.accountId))}</td>
              <td>${escapeHtml(t.type==='income' ? (donorName(t.donorId)||memberName(t.memberId)||t.payee||'-') : (t.payee||'-'))}</td>
              <td>${escapeHtml(t.mode||'-')}</td>
              <td style="text-align:right;font-weight:700;color:${t.type==='income'?'var(--green-ok)':'var(--red-bad)'}">${t.type==='income'?'+':'-'}${fmt(t.amount)}</td>
              <td class="row-actions">
                ${t.type==='income'?`<button class="icon-btn" data-receipt="${t.id}" title="Receipt">${receiptIcon()}</button>`:''}
                <button class="icon-btn" data-del-txn="${t.id}" title="Delete">${trashIcon()}</button>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>` : emptyState('No transactions match your filters.');
    document.getElementById('txnPanel').querySelectorAll('[data-del-txn]').forEach(b=>b.onclick=()=>deleteTxn(b.dataset.delTxn));
    document.getElementById('txnPanel').querySelectorAll('[data-receipt]').forEach(b=>b.onclick=()=>{ currentView='receipts'; render(); setTimeout(()=>showReceipt(b.dataset.receipt),0); });
  };
  draw();
  ['filterType','filterAcc','filterSearch'].forEach(id=>document.getElementById(id).addEventListener('input', draw));
  document.getElementById('addIncBtn').onclick = ()=>openTxnModal('income');
  document.getElementById('addExpBtn').onclick = ()=>openTxnModal('expense');
}
async function deleteTxn(id){
  if(!confirm('Delete this transaction? This cannot be undone.')) return;
  state.transactions = state.transactions.filter(t=>t.id!==id);
  await saveState(); showToast('Transaction deleted'); render();
}
function openTxnModal(type){
  if(state.accounts.length===0){ showToast('Create a fund account first'); return; }
  const incomeCats = ['Hundi Collection','General Donation','Annadhanam Sponsorship','Festival Sponsorship','Pooja Booking','Renovation Donation','Other'];
  const expenseCats = ['Priest Honorarium','Flowers & Pooja Items','Electricity','Maintenance & Repairs','Annadhanam Expense','Festival Expense','Staff Salary','Other'];
  const cats = type==='income'?incomeCats:expenseCats;
  openModal(`
    <h3>${type==='income'?'Record Donation / Income':'Record Expense'}</h3>
    <div class="field-row">
      <div class="field"><label>Date</label><input id="fDate" type="date" value="${todayStr()}"></div>
      <div class="field"><label>Amount (₹)</label><input id="fAmount" type="number" step="0.01" placeholder="0.00"></div>
    </div>
    <div class="field"><label>Fund Account</label>
      <select id="fAccount">${state.accounts.map(a=>`<option value="${a.id}">${escapeHtml(a.name)}</option>`).join('')}</select>
    </div>
    <div class="field"><label>Category</label>
      <select id="fCategory">${cats.map(c=>`<option>${c}</option>`).join('')}</select>
    </div>
    ${type==='income' ? `
    <div class="field-row">
      <div class="field"><label>Donor (optional)</label>
        <select id="fDonor"><option value="">— Anonymous / Walk-in —</option>${state.donors.map(d=>`<option value="${d.id}">${escapeHtml(d.name)}</option>`).join('')}</select>
      </div>
      <div class="field"><label>Member (optional)</label>
        <select id="fMember"><option value="">— None —</option>${state.members.map(m=>`<option value="${m.id}">${escapeHtml(m.name)}</option>`).join('')}</select>
      </div>
    </div>` : `
    <div class="field"><label>Paid To</label><input id="fPayee" placeholder="e.g. vendor / staff name"></div>
    `}
    <div class="field-row">
      <div class="field"><label>Payment Mode</label>
        <select id="fMode"><option>Cash</option><option>UPI</option><option>Bank Transfer</option><option>Cheque</option><option>Card</option></select>
      </div>
      <div class="field"><label>Reference No. (optional)</label><input id="fRef"></div>
    </div>
    <div class="field"><label>Notes</label><textarea id="fNotes" rows="2"></textarea></div>
    <div class="modal-actions">
      <button class="btn secondary" id="cancelBtn">Cancel</button>
      <button class="btn" id="saveTxnBtn">Save ${type==='income'?'Donation':'Expense'}</button>
    </div>
  `);
  document.getElementById('cancelBtn').onclick = closeModal;
  document.getElementById('saveTxnBtn').onclick = async ()=>{
    const amount = Number(document.getElementById('fAmount').value);
    if(!amount || amount<=0){ showToast('Please enter a valid amount'); return; }
    const txn = {
      id: uid(),
      type,
      date: document.getElementById('fDate').value || todayStr(),
      amount,
      accountId: document.getElementById('fAccount').value,
      category: document.getElementById('fCategory').value,
      mode: document.getElementById('fMode').value,
      ref: document.getElementById('fRef').value.trim(),
      notes: document.getElementById('fNotes').value.trim(),
      donorId: type==='income' ? (document.getElementById('fDonor').value || null) : null,
      memberId: type==='income' ? (document.getElementById('fMember').value || null) : null,
      payee: type==='expense' ? document.getElementById('fPayee').value.trim() : '',
      receiptNo: type==='income' ? 'RCPT-' + (state.transactions.filter(t=>t.type==='income').length+1).toString().padStart(4,'0') : null
    };
    state.transactions.push(txn);
    await saveState(); closeModal();
    showToast(type==='income' ? 'Donation recorded' : 'Expense recorded');
    render();
  };
}

/* ============ RECEIPTS ============ */
function renderReceipts(main){
  const incomeTxns = [...state.transactions].filter(t=>t.type==='income').sort((a,b)=>b.date.localeCompare(a.date));
  main.innerHTML = `
    <div class="page-head"><div><h1>${t('receipts')}</h1><p>${t('receiptsSub')}</p></div></div>
    <div class="panel">
      ${incomeTxns.length ? `
      <table>
        <thead><tr><th>Receipt No.</th><th>Date</th><th>Donor</th><th>Category</th><th style="text-align:right">Amount</th><th></th></tr></thead>
        <tbody>
          ${incomeTxns.map(t=>`
            <tr>
              <td style="font-weight:700;">${t.receiptNo||'-'}</td>
              <td>${fmtDate(t.date)}</td>
              <td>${escapeHtml(donorName(t.donorId) || memberName(t.memberId) || 'Anonymous')}</td>
              <td>${escapeHtml(t.category)}</td>
              <td style="text-align:right;font-weight:700;color:var(--green-ok)">${fmt(t.amount)}</td>
              <td><button class="btn small secondary" data-view-receipt="${t.id}">View / Print</button></td>
            </tr>`).join('')}
        </tbody>
      </table>` : emptyState('No donations recorded yet — receipts appear here once you add income transactions.')}
    </div>
    <div id="receiptView"></div>
  `;
  main.querySelectorAll('[data-view-receipt]').forEach(b=>b.onclick=()=>showReceipt(b.dataset.viewReceipt));
}
function showReceipt(txnId, lang){
  const txn = state.transactions.find(x=>x.id===txnId);
  if(!txn) return;
  lang = lang || appLang;
  const s = state.settings;
  const html = `
    <div class="panel" style="margin-top:20px;">
      <div id="printArea">
        <div class="receipt">
          <div class="receipt-headbar">
            <div class="logo-slot">${s.logo ? `<img src="${s.logo}">` : ''}</div>
            <div class="temple-info">
              <div class="temple">${escapeHtml(s.templeName||'Temple Trust')}</div>
              <div class="addr">${escapeHtml(s.address||'')}</div>
              <div class="addr">${s.regNo?(tr('regNoLabel',lang)+' '+escapeHtml(s.regNo)):''} ${s.phone?(' • '+tr('phLabel',lang)+' '+escapeHtml(s.phone)):''}</div>
            </div>
            <div class="logo-slot" style="justify-content:flex-end;visibility:hidden;">${s.logo ? `<img src="${s.logo}">` : ''}</div>
          </div>
          <div class="receipt-title">${tr('receiptDocTitle',lang)}</div>
          <div class="receipt-grid">
            <div class="k">${tr('receiptNoLabel',lang)}</div><div>${txn.receiptNo||'-'}</div>
            <div class="k">${tr('dateLabel',lang)}</div><div>${fmtDate(txn.date)}</div>
            <div class="k">${tr('receivedFromLabel',lang)}</div><div>${escapeHtml(donorName(txn.donorId) || memberName(txn.memberId) || tr('anonymousWalkin',lang))}</div>
            <div class="k">${tr('purposeLabel',lang)}</div><div>${escapeHtml(txn.category)}</div>
            <div class="k">${tr('fundAccountLabel',lang)}</div><div>${escapeHtml(accountName(txn.accountId))}</div>
            <div class="k">${tr('paymentModeLabel',lang)}</div><div>${escapeHtml(txn.mode||'-')}</div>
          </div>
          <div class="receipt-amount">${fmt(txn.amount)}</div>
          <div class="receipt-foot">${tr('thankYouMsg',lang)}<br>${tr('systemGenerated',lang)}</div>
        </div>
      </div>
      <div class="modal-actions print-hide" style="justify-content:center;gap:16px;margin-top:20px;flex-wrap:wrap;">
        <div style="display:flex;gap:8px;align-items:center;">
          <label style="font-weight:600;color:var(--maroon-800);">${t('printLanguage')}</label>
          <select id="receiptLangSelect" style="padding:8px 12px;border:1px solid var(--gold-300);border-radius:6px;background:white;color:var(--maroon-800);">
            <option value="en" ${lang==='en'?'selected':''}>English</option>
            <option value="ta" ${lang==='ta'?'selected':''}>Tamil (தமிழ்)</option>
          </select>
        </div>
        <button class="btn secondary" onclick="document.getElementById('receiptView').innerHTML=''">${t('close')}</button>
        <button class="btn" onclick="window.print()">🖨 ${t('printReceipt')}</button>
      </div>
    </div>
  `;
  document.getElementById('receiptView').innerHTML = html;
  document.getElementById('receiptView').scrollIntoView({behavior:'smooth'});
  document.getElementById('receiptLangSelect').onchange = (e)=> showReceipt(txnId, e.target.value);
}

/* ============ REPORTS ============ */
function renderReports(main){
  main.innerHTML = `
    <div class="page-head">
      <div><h1>${t('reports')}</h1><p>${t('reportsSub')}</p></div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;">
        <button class="btn secondary small" id="repExportCsv">⬇ Export CSV</button>
        <button class="btn small" id="repDownloadPdf">⬇ Download Report (PDF)</button>
      </div>
    </div>
    <div class="toolbar">
      <label style="font-size:13px;color:var(--ink-soft);font-weight:600;">From <input id="repFrom" type="date"></label>
      <label style="font-size:13px;color:var(--ink-soft);font-weight:600;">To <input id="repTo" type="date"></label>
      <button class="btn small secondary" id="repRun">Apply</button>
    </div>
    <div id="repOut"></div>
    <div id="repPrintPreview"></div>
  `;
  let lastReport = null;
  const run = ()=>{
    const from = document.getElementById('repFrom').value;
    const to = document.getElementById('repTo').value;
    let list = state.transactions;
    if(from) list = list.filter(t=>t.date>=from);
    if(to) list = list.filter(t=>t.date<=to);
    const income = list.filter(t=>t.type==='income').reduce((s,t)=>s+Number(t.amount),0);
    const expense = list.filter(t=>t.type==='expense').reduce((s,t)=>s+Number(t.amount),0);

    const byCat = {};
    list.forEach(t=>{
      const k = t.type+'|'+(t.category||'Other');
      byCat[k] = (byCat[k]||0) + Number(t.amount);
    });
    const incomeCats = Object.entries(byCat).filter(([k])=>k.startsWith('income|')).sort((a,b)=>b[1]-a[1]);
    const expenseCats = Object.entries(byCat).filter(([k])=>k.startsWith('expense|')).sort((a,b)=>b[1]-a[1]);
    const maxCat = Math.max(1, ...Object.values(byCat));

    const barsHtml = (arr) => arr.length ? arr.map(([k,v])=>{
      const label = k.split('|')[1];
      const pct = v/maxCat*100;
      return `<div class="bar-row"><div>${escapeHtml(label)}</div><div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div><div style="text-align:right;font-weight:700;">${fmt(v)}</div></div>`;
    }).join('') : emptyState('No data in this period.');

    document.getElementById('repOut').innerHTML = `
      <div class="grid-stats">
        <div class="stat-card income"><div class="label">Income (period)</div><div class="value">${fmt(income)}</div></div>
        <div class="stat-card expense"><div class="label">Expenses (period)</div><div class="value">${fmt(expense)}</div></div>
        <div class="stat-card"><div class="label">Net (period)</div><div class="value">${fmt(income-expense)}</div></div>
      </div>
      <div class="panel"><h2>Income by Category</h2>${barsHtml(incomeCats)}</div>
      <div class="panel"><h2>Expenses by Category</h2>${barsHtml(expenseCats)}</div>
    `;

    lastReport = { from, to, list, income, expense, incomeCats, expenseCats };
    document.getElementById('repPrintPreview').innerHTML = '';
  };
  document.getElementById('repRun').onclick = run;
  document.getElementById('repExportCsv').onclick = ()=> exportReportCsv(lastReport);
  document.getElementById('repDownloadPdf').onclick = ()=> downloadReportPdf(lastReport);
  run();
}

function reportPeriodLabel(rep, lang){
  lang = lang || appLang;
  if(rep.from && rep.to) return `${fmtDate(rep.from)} — ${fmtDate(rep.to)}`;
  if(rep.from) return `${tr('fromLabel',lang)} ${fmtDate(rep.from)}`;
  if(rep.to) return `${tr('uptoLabel',lang)} ${fmtDate(rep.to)}`;
  return tr('allTime',lang);
}

/* ---- CSV export (raw transaction data for the selected period) ---- */
function exportReportCsv(rep){
  if(!rep){ showToast('Run the report first'); return; }
  if(!rep.list.length){ showToast('No transactions in this period'); return; }
  const headers = ['Date','Type','Category','Account','Donor/Payee','Payment Mode','Reference','Amount','Notes'];
  const rows = [...rep.list]
    .sort((a,b)=>a.date.localeCompare(b.date))
    .map(t=>[
      t.date,
      t.type==='income'?'Income':'Expense',
      t.category||'',
      accountName(t.accountId),
      t.type==='income' ? (donorName(t.donorId)||memberName(t.memberId)||t.payee||'Anonymous') : (t.payee||''),
      t.mode||'',
      t.ref||'',
      t.amount,
      t.notes||''
    ]);
  const csvEscape = v => {
    const s = String(v ?? '');
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
  };
  const lines = [headers, ...rows].map(r=>r.map(csvEscape).join(','));
  lines.push('');
  lines.push(['','','','','','','Total Income','',rep.income].map(csvEscape).join(','));
  lines.push(['','','','','','','Total Expenses','',rep.expense].map(csvEscape).join(','));
  lines.push(['','','','','','','Net','',rep.income-rep.expense].map(csvEscape).join(','));
  const csv = lines.join('\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const stamp = todayStr();
  a.href = url;
  a.download = `trust-report_${rep.from||'all'}_to_${rep.to||stamp}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  showToast('CSV downloaded');
}

/* ---- Full printable report (use browser's Save as PDF) ---- */
function downloadReportPdf(rep, lang){
  if(!rep){ showToast('Run the report first'); return; }
  lang = lang || appLang;
  const s = state.settings;
  const rowsHtml = (arr) => arr.length ? `
    <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:6px;">
      <thead><tr><th style="text-align:left;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colCategory',lang)}</th><th style="text-align:right;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colAmount',lang)}</th></tr></thead>
      <tbody>${arr.map(([k,v])=>`<tr><td style="padding:6px 4px;border-bottom:1px solid var(--cream-200);">${escapeHtml(k.split('|')[1])}</td><td style="padding:6px 4px;border-bottom:1px solid var(--cream-200);text-align:right;font-weight:700;">${fmt(v)}</td></tr>`).join('')}</tbody>
    </table>` : `<p style="color:var(--ink-soft);font-size:13px;">${tr('noDataPeriod',lang)}</p>`;

  const txnRowsHtml = [...rep.list].sort((a,b)=>a.date.localeCompare(b.date)).map(t=>`
    <tr>
      <td style="padding:5px 4px;border-bottom:1px solid var(--cream-200);">${fmtDate(t.date)}</td>
      <td style="padding:5px 4px;border-bottom:1px solid var(--cream-200);">${t.type==='income'?tr('typeIncome',lang):tr('typeExpense',lang)}</td>
      <td style="padding:5px 4px;border-bottom:1px solid var(--cream-200);">${escapeHtml(t.category||'-')}</td>
      <td style="padding:5px 4px;border-bottom:1px solid var(--cream-200);">${escapeHtml(accountName(t.accountId))}</td>
      <td style="padding:5px 4px;border-bottom:1px solid var(--cream-200);">${escapeHtml(t.type==='income' ? (donorName(t.donorId)||memberName(t.memberId)||t.payee||'Anonymous') : (t.payee||'-'))}</td>
      <td style="padding:5px 4px;border-bottom:1px solid var(--cream-200);text-align:right;font-weight:700;color:${t.type==='income'?'var(--green-ok)':'var(--red-bad)'}">${t.type==='income'?'+':'-'}${fmt(t.amount)}</td>
    </tr>`).join('');

  const html = `
    <div class="panel" style="margin-top:20px;">
      <div id="printArea">
        <div style="max-width:760px;margin:0 auto;">
          <div style="margin-bottom:16px;">
            <div style="display:flex;align-items:center;gap:14px;">
              ${s.logo ? `<div style="width:80px;flex-shrink:0;"><img src="${s.logo}" style="max-width:80px;max-height:70px;object-fit:contain;"></div>` : ''}
              <div style="flex:1;text-align:center;">
                <div style="font-family:'Yatra One',cursive;font-size:24px;color:var(--maroon-800);">${escapeHtml(s.templeName||'Temple Trust')}</div>
                <div style="font-size:12px;color:var(--ink-soft);">${escapeHtml(s.address||'')}</div>
                <div style="font-size:12px;color:var(--ink-soft);">${s.regNo?(tr('regNoLabel',lang)+' '+escapeHtml(s.regNo)):''} ${s.phone?(' • '+tr('phLabel',lang)+' '+escapeHtml(s.phone)):''}</div>
              </div>
              ${s.logo ? `<div style="width:80px;flex-shrink:0;visibility:hidden;"></div>` : ''}
            </div>
            <div style="text-align:center;font-weight:700;letter-spacing:.1em;text-transform:uppercase;font-size:12.5px;color:var(--gold-500);margin:14px 0;border-top:1px dashed var(--gold-300);border-bottom:1px dashed var(--gold-300);padding:6px 0;">${tr('reportDocTitle',lang)}</div>
            <div style="text-align:center;font-size:13px;color:var(--ink-soft);">${tr('periodLabel',lang)} ${reportPeriodLabel(rep,lang)}</div>
          </div>

          <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px;">
            <div style="border:1px solid var(--cream-200);border-radius:10px;padding:12px;text-align:center;">
              <div style="font-size:11px;text-transform:uppercase;color:var(--ink-soft);">${tr('incomeLabel',lang)}</div>
              <div style="font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:700;color:var(--green-ok);">${fmt(rep.income)}</div>
            </div>
            <div style="border:1px solid var(--cream-200);border-radius:10px;padding:12px;text-align:center;">
              <div style="font-size:11px;text-transform:uppercase;color:var(--ink-soft);">${tr('expensesLabel',lang)}</div>
              <div style="font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:700;color:var(--red-bad);">${fmt(rep.expense)}</div>
            </div>
            <div style="border:1px solid var(--cream-200);border-radius:10px;padding:12px;text-align:center;">
              <div style="font-size:11px;text-transform:uppercase;color:var(--ink-soft);">${tr('netLabel',lang)}</div>
              <div style="font-family:'Cormorant Garamond',serif;font-size:22px;font-weight:700;color:var(--maroon-800);">${fmt(rep.income-rep.expense)}</div>
            </div>
          </div>

          <h3 style="font-family:'Cormorant Garamond',serif;color:var(--maroon-800);font-size:17px;margin:16px 0 8px;">${tr('incomeByCategory',lang)}</h3>
          ${rowsHtml(rep.incomeCats)}

          <h3 style="font-family:'Cormorant Garamond',serif;color:var(--maroon-800);font-size:17px;margin:16px 0 8px;">${tr('expensesByCategory',lang)}</h3>
          ${rowsHtml(rep.expenseCats)}

          <h3 style="font-family:'Cormorant Garamond',serif;color:var(--maroon-800);font-size:17px;margin:16px 0 8px;">${tr('transactionDetail',lang)}</h3>
          ${rep.list.length ? `
          <table style="width:100%;border-collapse:collapse;font-size:12.5px;">
            <thead><tr>
              <th style="text-align:left;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colDate',lang)}</th>
              <th style="text-align:left;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colType',lang)}</th>
              <th style="text-align:left;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colCategory',lang)}</th>
              <th style="text-align:left;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colAccount',lang)}</th>
              <th style="text-align:left;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colDonorPayee',lang)}</th>
              <th style="text-align:right;padding:6px 4px;border-bottom:2px solid var(--gold-300);">${tr('colAmount',lang)}</th>
            </tr></thead>
            <tbody>${txnRowsHtml}</tbody>
          </table>` : `<p style="color:var(--ink-soft);font-size:13px;">${tr('noTxnPeriod',lang)}</p>`}

          <div style="text-align:center;font-size:11px;color:var(--ink-soft);margin-top:24px;font-style:italic;">${tr('generatedOn',lang)} ${fmtDate(todayStr())} — ${tr('systemGeneratedReport',lang)}</div>
        </div>
      </div>
      <div class="modal-actions print-hide" style="justify-content:center;gap:16px;margin-top:20px;flex-wrap:wrap;">
        <div style="display:flex;gap:8px;align-items:center;">
          <label style="font-weight:600;color:var(--maroon-800);">${t('printLanguage')}</label>
          <select id="printLangSelect" style="padding:8px 12px;border:1px solid var(--gold-300);border-radius:6px;background:white;color:var(--maroon-800);">
            <option value="en" ${lang==='en'?'selected':''}>English</option>
            <option value="ta" ${lang==='ta'?'selected':''}>Tamil (தமிழ்)</option>
          </select>
        </div>
        <button class="btn secondary" onclick="document.getElementById('repPrintPreview').innerHTML=''">${t('close')}</button>
        <button class="btn" onclick="window.print()">🖨 ${t('printReport')}</button>
      </div>
    </div>
  `;
  document.getElementById('repPrintPreview').innerHTML = html;
  document.getElementById('repPrintPreview').scrollIntoView({behavior:'smooth'});
  document.getElementById('printLangSelect').onchange = (e)=> downloadReportPdf(rep, e.target.value);
}

/* ============ SETTINGS ============ */
function updateBrand(){
  const s = state.settings;
  document.getElementById('brandName').textContent = s.templeName ? s.templeName : 'Kovil Kanakku';
  if(s.logo){
    document.getElementById('sidebarLogo').innerHTML = `<img src="${s.logo}" alt="Logo">`;
  } else {
    document.getElementById('sidebarLogo').textContent = '';
  }
}

function renderSettings(main){
  const s = state.settings;
  main.innerHTML = `
    <div class="page-head"><div><h1>${t('settings')}</h1><p>${t('settingsSub')}</p></div></div>
    <div class="panel" style="max-width:520px;">
      <h3 style="font-size:16px;color:var(--maroon-800);margin-top:0;">⚙️ Temple Information</h3>
      <div class="field"><label>Temple / Trust Name</label><input id="sName" value="${escapeHtml(s.templeName||'')}"></div>
      <div class="field"><label>Address</label><textarea id="sAddr" rows="2">${escapeHtml(s.address||'')}</textarea></div>
      <div class="field-row">
        <div class="field"><label>Trust Registration No.</label><input id="sReg" value="${escapeHtml(s.regNo||'')}"></div>
        <div class="field"><label>Phone</label><input id="sPhone" value="${escapeHtml(s.phone||'')}"></div>
      </div>
      <button class="btn" id="saveSettingsBtn">Save Settings</button>
    </div>

    <div class="panel" style="max-width:520px;">
      <h3 style="font-size:16px;color:var(--maroon-800);margin-top:0;">🖼️ Temple Logo</h3>
      <div class="field">
        <label>Upload Logo/Image</label>
        <div style="display:flex;gap:10px;margin-top:10px;">
          <input type="file" id="logoFile" accept="image/*" style="flex:1;padding:8px;border:1px solid var(--cream-200);border-radius:8px;">
          <button class="btn small" onclick="uploadLogo()">📤 Upload</button>
          ${s.logo ? `<button class="btn small danger" onclick="clearLogo()">❌ Clear</button>` : ''}
        </div>
      </div>
      <div style="margin-top:12px;padding:12px;background:rgba(201,162,39,.1);border-radius:8px;text-align:center;min-height:80px;">
        ${s.logo ? `<img src="${s.logo}" style="max-width:100px;max-height:70px;border-radius:6px;">` : '<span style="color:var(--ink-soft);">No logo uploaded</span>'}
      </div>
    </div>

    <div class="panel" style="max-width:520px;">
      <h3 style="font-size:16px;color:var(--maroon-800);margin-top:0;">📥 Bulk Import</h3>
      <input type="file" id="importFile" accept=".xlsx,.xls" style="display:block;width:100%;padding:8px;border:1px solid var(--cream-200);border-radius:8px;margin-bottom:8px;font-size:12px;">
      <button class="btn small" onclick="bulkImport()" style="width:100%;margin-bottom:8px;">Import Data</button>
      <button class="btn small secondary" onclick="downloadImportTemplate()" style="width:100%;">⬇ Download Template</button>
    </div>

    <div class="panel" style="max-width:520px;">
      <h3 style="font-size:16px;color:var(--maroon-800);margin-top:0;">📤 Bulk Export</h3>
      <button class="btn small secondary" onclick="bulkExport()" style="width:100%;">Export All Data</button>
      <p style="font-size:11px;color:var(--ink-soft);margin-top:8px;">Exports as an Excel (.xlsx) file with all records</p>
    </div>
  `;
  document.getElementById('saveSettingsBtn').onclick = async ()=>{
    state.settings = {
      templeName: document.getElementById('sName').value.trim(),
      address: document.getElementById('sAddr').value.trim(),
      regNo: document.getElementById('sReg').value.trim(),
      phone: document.getElementById('sPhone').value.trim(),
      logo: state.settings.logo
    };
    await saveState(); 
    updateBrand();
    showToast('Settings saved');
  };
}

function downloadImportTemplate(){
  if(typeof XLSX === 'undefined'){ showToast('Excel library failed to load — check your connection'); return; }
  const wb = XLSX.utils.book_new();

  const settingsSheet = XLSX.utils.json_to_sheet([
    { Field:'Temple Name', Value:'' },
    { Field:'Address', Value:'' },
    { Field:'Trust Registration No.', Value:'' },
    { Field:'Phone', Value:'' },
    { Field:'Logo (Base64)', Value:'' },
  ]);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  const accountsSheet = XLSX.utils.json_to_sheet([
    { id:'acc1', name:'General Fund', type:'General', openingBalance:0 }
  ]);
  XLSX.utils.book_append_sheet(wb, accountsSheet, 'Accounts');

  const donorsSheet = XLSX.utils.json_to_sheet([
    { id:'don1', name:'Sample Donor', phone:'9876543210', pan:'', email:'', address:'' }
  ]);
  XLSX.utils.book_append_sheet(wb, donorsSheet, 'Donors');

  const membersSheet = XLSX.utils.json_to_sheet([
    { id:'mem1', name:'Sample Member', phone:'9876543210', role:'', email:'', address:'' }
  ]);
  XLSX.utils.book_append_sheet(wb, membersSheet, 'Members');

  const transactionsSheet = XLSX.utils.json_to_sheet([
    { id:'txn1', type:'income', date:'2026-01-01', amount:1000, accountId:'acc1',
      accountName:'General Fund', category:'General Donation', mode:'Cash',
      ref:'', notes:'', donorId:'don1', donorName:'Sample Donor',
      memberId:'', memberName:'', payee:'', receiptNo:'RCPT-0001' }
  ]);
  XLSX.utils.book_append_sheet(wb, transactionsSheet, 'Transactions');

  const receiptsSheet = XLSX.utils.json_to_sheet([
    { receiptNo:'RCPT-0001', date:'2026-01-01', donorName:'Sample Donor', category:'General Donation', fundAccount:'General Fund', mode:'Cash', amount:1000 }
  ]);
  XLSX.utils.book_append_sheet(wb, receiptsSheet, 'Receipts');

  XLSX.writeFile(wb, 'temple-import-template.xlsx');
  showToast('Template downloaded ✓');
}

function bulkExport(){
  if(typeof XLSX === 'undefined'){ showToast('Excel library failed to load — check your connection'); return; }
  const wb = XLSX.utils.book_new();

  const settingsSheet = XLSX.utils.json_to_sheet([
    { Field:'Temple Name', Value: state.settings.templeName||'' },
    { Field:'Address', Value: state.settings.address||'' },
    { Field:'Trust Registration No.', Value: state.settings.regNo||'' },
    { Field:'Phone', Value: state.settings.phone||'' },
    { Field:'Logo (Base64)', Value: state.settings.logo||'' },
  ]);
  XLSX.utils.book_append_sheet(wb, settingsSheet, 'Settings');

  const accountsSheet = XLSX.utils.json_to_sheet(state.accounts.map(a=>({
    id:a.id, name:a.name, type:a.type||'', openingBalance:a.openingBalance||0
  })));
  XLSX.utils.book_append_sheet(wb, accountsSheet, 'Accounts');

  const donorsSheet = XLSX.utils.json_to_sheet(state.donors.map(d=>({
    id:d.id, name:d.name, phone:d.phone||'', pan:d.pan||'', email:d.email||'', address:d.address||''
  })));
  XLSX.utils.book_append_sheet(wb, donorsSheet, 'Donors');

  const membersSheet = XLSX.utils.json_to_sheet(state.members.map(m=>({
    id:m.id, name:m.name, phone:m.phone||'', role:m.role||'', email:m.email||'', address:m.address||''
  })));
  XLSX.utils.book_append_sheet(wb, membersSheet, 'Members');

  const transactionsSheet = XLSX.utils.json_to_sheet(state.transactions.map(t=>({
    id:t.id, type:t.type, date:t.date, amount:t.amount, accountId:t.accountId,
    accountName: accountName(t.accountId), category:t.category||'', mode:t.mode||'',
    ref:t.ref||'', notes:t.notes||'', donorId:t.donorId||'', donorName: donorName(t.donorId),
    memberId:t.memberId||'', memberName: memberName(t.memberId), payee:t.payee||'',
    receiptNo:t.receiptNo||''
  })));
  XLSX.utils.book_append_sheet(wb, transactionsSheet, 'Transactions');

  const receiptsSheet = XLSX.utils.json_to_sheet(
    state.transactions.filter(t=>t.type==='income').map(t=>({
      receiptNo:t.receiptNo||'', date:t.date, donorName: donorName(t.donorId)||memberName(t.memberId)||'Anonymous',
      category:t.category||'', fundAccount: accountName(t.accountId), mode:t.mode||'', amount:t.amount
    }))
  );
  XLSX.utils.book_append_sheet(wb, receiptsSheet, 'Receipts');

  const stamp = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `temple-backup_${(state.settings.templeName||'trust').replace(/\s+/g,'_')}_${stamp}.xlsx`);
  showToast('Data exported ✓');
}

function bulkImport(){
  if(typeof XLSX === 'undefined'){ showToast('Excel library failed to load — check your connection'); return; }
  const file = document.getElementById('importFile').files[0];
  if(!file){ showToast('Select an Excel (.xlsx) file'); return; }
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, {type:'array'});
      const sheet = (name) => wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name]) : [];

      const settingsRows = sheet('Settings');
      const getField = (label) => { const r = settingsRows.find(row=>row.Field===label); return r ? String(r.Value||'') : ''; };
      const importedLogo = getField('Logo (Base64)');
      const newSettings = settingsRows.length ? {
        templeName: getField('Temple Name'),
        address: getField('Address'),
        regNo: getField('Trust Registration No.'),
        phone: getField('Phone'),
        logo: importedLogo ? importedLogo : state.settings.logo
      } : state.settings;

      const accounts = sheet('Accounts').map(r=>({ id:String(r.id), name:r.name||'', type:r.type||'', openingBalance:Number(r.openingBalance)||0 }));
      const donors = sheet('Donors').map(r=>({ id:String(r.id), name:r.name||'', phone:r.phone?String(r.phone):'', pan:r.pan||'', email:r.email||'', address:r.address||'' }));
      const members = sheet('Members').map(r=>({ id:String(r.id), name:r.name||'', phone:r.phone?String(r.phone):'', role:r.role||'', email:r.email||'', address:r.address||'' }));
      const transactions = sheet('Transactions').map(r=>({
        id:String(r.id), type:r.type, date:String(r.date), amount:Number(r.amount)||0,
        accountId:r.accountId?String(r.accountId):null, category:r.category||'', mode:r.mode||'',
        ref:r.ref?String(r.ref):'', notes:r.notes||'',
        donorId:r.donorId?String(r.donorId):null, memberId:r.memberId?String(r.memberId):null,
        payee:r.payee||'', receiptNo:r.receiptNo?String(r.receiptNo):null
      }));

      if(!accounts.length && !donors.length && !members.length && !transactions.length){
        showToast('Invalid or empty Excel file'); return;
      }

      if(confirm('This will replace all current data. Continue?')){
        state.settings = newSettings;
        state.accounts = accounts;
        state.donors = donors;
        state.members = members;
        state.transactions = transactions;

        saveState();
        updateBrand();
        document.getElementById('importFile').value = '';
        showToast('Data imported ✓');
        render();
      }
    } catch(err) {
      showToast('Error reading file: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function uploadLogo(){
  const file = document.getElementById('logoFile').files[0];
  if(!file){ showToast('Select an image'); return; }
  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      const maxDim = 160;
      let w = img.width, h = img.height;
      if(w > h){ if(w > maxDim){ h = Math.round(h * maxDim / w); w = maxDim; } }
      else { if(h > maxDim){ w = Math.round(w * maxDim / h); h = maxDim; } }
      const canvas = document.createElement('canvas');
      canvas.width = w; canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, w, h);
      state.settings.logo = canvas.toDataURL('image/png');
      saveState();
      updateBrand();
      renderSettings(document.getElementById('main'));
      showToast('Logo updated ✓');
    };
    img.onerror = () => { showToast('Could not read that image'); };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
}

function clearLogo(){
  state.settings.logo = null;
  saveState();
  updateBrand();
  renderSettings(document.getElementById('main'));
  showToast('Logo cleared ✓');
}

/* ============ ICONS ============ */
function editIcon(){ return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>`; }
function trashIcon(){ return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14"/></svg>`; }
function receiptIcon(){ return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 2h9l3 3v17l-3-2-3 2-3-2-3 2V2Z"/><path d="M9 8h6M9 12h6"/></svg>`; }

/* ============ LANGUAGE & LOCALIZATION ============ */
let appLang = 'en';
const translations = {
  en: {
    dashboard: 'Dashboard',
    transactions: 'Transactions',
    accounts: 'Fund Accounts',
    donors: 'Donors',
    members: 'Members',
    receipts: 'Receipts',
    reports: 'Reports',
    settings: 'Temple Settings',
    donorList: 'Donors (General Contributors)',
    memberList: 'Members (Temple Maintainers)',
    printLanguage: 'Print Language:',
    printReceipt: 'Print Receipt',
    printReport: 'Print / Save as PDF',
    close: 'Close',
    selectLanguage: 'Select Language',
    dashboardSub: 'overview of the trust ledger',
    accountsSub: 'Separate ledgers e.g. General Fund, Annadhanam, Renovation, Hundi',
    newFundAccount: '+ New Fund Account',
    transactionsSub: 'Record donations (income) and temple expenses',
    addExpense: '+ Expense',
    addIncome: '+ Donation / Income',
    receiptsSub: 'Generate and print donation receipts for devotees',
    reportsSub: 'Summarised view of trust finances by category and period',
    settingsSub: 'Details shown on printed donation receipts',
    receiptDocTitle: 'Donation Receipt',
    receiptNoLabel: 'Receipt No.',
    dateLabel: 'Date',
    receivedFromLabel: 'Received From',
    purposeLabel: 'Purpose',
    fundAccountLabel: 'Fund Account',
    paymentModeLabel: 'Payment Mode',
    thankYouMsg: 'Thank you for your generous contribution to the Trust.',
    systemGenerated: 'This receipt is system-generated.',
    anonymousWalkin: 'Anonymous / Walk-in Devotee',
    regNoLabel: 'Reg. No:',
    phLabel: 'Ph:',
    reportDocTitle: 'Overall Financial Report',
    periodLabel: 'Period:',
    incomeLabel: 'Income',
    expensesLabel: 'Expenses',
    netLabel: 'Net',
    incomeByCategory: 'Income by Category',
    expensesByCategory: 'Expenses by Category',
    transactionDetail: 'Transaction Detail',
    noDataPeriod: 'No data in this period.',
    noTxnPeriod: 'No transactions in this period.',
    generatedOn: 'Generated on',
    systemGeneratedReport: 'system-generated report.',
    colDate: 'Date', colType: 'Type', colCategory: 'Category', colAccount: 'Account',
    colDonorPayee: 'Donor/Payee', colAmount: 'Amount',
    typeIncome: 'Income', typeExpense: 'Expense',
    allTime: 'All time', fromLabel: 'From', uptoLabel: 'Up to',
  },
  ta: {
    dashboard: 'முகப்பு',
    transactions: 'பரிவர்த்தனைகள்',
    accounts: 'நிதி கணக்குகள்',
    donors: 'நன்கொடையாளர்கள்',
    members: 'உறுப்பினர்கள்',
    receipts: 'ரசீதுகள்',
    reports: 'அறிக்கைகள்',
    settings: 'கோயில் அமைப்புகள்',
    donorList: 'நன்கொடையாளர்கள் (பொது பங்களிப்பு)',
    memberList: 'உறுப்பினர்கள் (கோயில் பராமரிப்பு)',
    printLanguage: 'அச்சு மொழி:',
    printReceipt: 'ரசீது அச்சு',
    printReport: 'அச்சு / PDF இல் சேமி',
    close: 'மூடு',
    selectLanguage: 'மொழி தேர்வு',
    dashboardSub: 'அறக்கட்டளை கணக்கு மேலோட்டம்',
    accountsSub: 'தனித்தனி கணக்குகள் — பொது நிதி, அன்னதானம், புதுப்பித்தல், உண்டியல்',
    newFundAccount: '+ புதிய நிதி கணக்கு',
    transactionsSub: 'நன்கொடைகள் (வரவு) மற்றும் கோயில் செலவுகளை பதிவு செய்யவும்',
    addExpense: '+ செலவு',
    addIncome: '+ நன்கொடை / வரவு',
    receiptsSub: 'பக்தர்களுக்கான நன்கொடை ரசீதுகளை உருவாக்கி அச்சிடவும்',
    reportsSub: 'வகை மற்றும் காலம் வாரியான அறக்கட்டளை நிதி சுருக்கம்',
    settingsSub: 'அச்சிடப்பட்ட நன்கொடை ரசீதுகளில் காட்டப்படும் விவரங்கள்',
    receiptDocTitle: 'நன்கொடை ரசீது',
    receiptNoLabel: 'ரசீது எண்.',
    dateLabel: 'தேதி',
    receivedFromLabel: 'பெறப்பட்டவர்',
    purposeLabel: 'நோக்கம்',
    fundAccountLabel: 'நிதி கணக்கு',
    paymentModeLabel: 'செலுத்தும் முறை',
    thankYouMsg: 'அறக்கட்டளைக்கு உங்கள் தாராள நன்கொடைக்கு நன்றி.',
    systemGenerated: 'இந்த ரசீது தானாக உருவாக்கப்பட்டது.',
    anonymousWalkin: 'அநாமதேய / நடைபயணி பக்தர்',
    regNoLabel: 'பதிவு எண்:',
    phLabel: 'தொலைபேசி:',
    reportDocTitle: 'முழுமையான நிதி அறிக்கை',
    periodLabel: 'காலம்:',
    incomeLabel: 'வரவு',
    expensesLabel: 'செலவுகள்',
    netLabel: 'நிகரம்',
    incomeByCategory: 'வகை வாரியான வரவு',
    expensesByCategory: 'வகை வாரியான செலவுகள்',
    transactionDetail: 'பரிவர்த்தனை விவரம்',
    noDataPeriod: 'இந்த காலத்தில் தரவு இல்லை.',
    noTxnPeriod: 'இந்த காலத்தில் பரிவர்த்தனைகள் இல்லை.',
    generatedOn: 'உருவாக்கப்பட்ட தேதி',
    systemGeneratedReport: 'தானாக உருவாக்கப்பட்ட அறிக்கை.',
    colDate: 'தேதி', colType: 'வகை', colCategory: 'வகைப்பாடு', colAccount: 'கணக்கு',
    colDonorPayee: 'நன்கொடையாளர்/பெறுநர்', colAmount: 'தொகை',
    typeIncome: 'வரவு', typeExpense: 'செலவு',
    allTime: 'எல்லா காலமும்', fromLabel: 'தொடக்கம்', uptoLabel: 'வரை',
  }
};

function t(key){ return translations[appLang]?.[key] || key; }
function tr(key, lang){ return translations[lang]?.[key] || translations.en[key] || key; }

function setAppLang(newLang){
  appLang = newLang;
  localStorage.setItem('appLang', newLang);
  render();
  showToast('Language changed ✓');
}

/* ============ INIT ============ */
(async function init(){
  appLang = localStorage.getItem('appLang') || 'en';
  await loadState();
  render();
})();
</script>
</body>
</html>