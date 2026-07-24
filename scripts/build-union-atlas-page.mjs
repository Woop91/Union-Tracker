import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const tokens = fs.readFileSync(path.join(root, 'design', 'solid-ground', 'tokens.css'), 'utf8');
const workerSource = fs.readFileSync(path.join(root, 'src', 'union-atlas-worker.js'), 'utf8');
const indexSource = fs.readFileSync(path.join(root, 'docs', 'union-atlas-index.js'), 'utf8');
const directoryPackage = fs.readFileSync(path.join(root, 'docs', 'union-atlas-data.json.gz'));
const detailsPackage = fs.readFileSync(path.join(root, 'docs', 'union-atlas-details.json.gz'));
const evidencePackage = fs.readFileSync(path.join(root, 'docs', 'union-atlas-evidence.json.gz'));
const outputPath = path.join(root, 'docs', 'union-atlas.html');
const standalonePath = path.join(root, 'docs', 'union-atlas-standalone.html');
const workerPath = path.join(root, 'docs', 'union-atlas-worker.js');

const page = String.raw`<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="theme-color" content="#0a2f26">
<link rel="icon" href="data:,">
<title>Common Ground — United States Union Atlas</title>
<style>
/* BEGIN SOLID GROUND TOKENS */
${tokens}
/* END SOLID GROUND TOKENS */
:root{
  --background:var(--sg-color-canvas);--foreground:var(--sg-color-ink);
  --card:var(--sg-color-card);--card-foreground:var(--sg-color-ink);
  --popover:var(--sg-color-card);--popover-foreground:var(--sg-color-ink);
  --primary:var(--sg-color-green);--primary-foreground:var(--sg-color-canvas);
  --secondary:var(--sg-color-card-2);--secondary-foreground:var(--sg-color-green);
  --muted:var(--sg-color-card-2);--muted-foreground:var(--sg-color-sec);
  --accent:var(--sg-color-gold);--accent-foreground:var(--sg-color-ink-on-gold);
  --destructive:var(--sg-color-danger);--destructive-foreground:var(--sg-color-card);
  --success:var(--sg-color-green-soft);--success-foreground:var(--sg-color-green);
  --warning:var(--sg-color-warn-soft);--warning-foreground:var(--sg-color-ink);
  --border:var(--sg-color-muted);--input:var(--sg-color-sec);--ring:var(--sg-color-green-text);
}
html[data-theme="dark"]{
  --background:var(--sg-color-forest);--foreground:var(--sg-color-canvas);
  --card:var(--sg-color-green);--card-foreground:var(--sg-color-canvas);
  --popover:var(--sg-color-pine);--popover-foreground:var(--sg-color-canvas);
  --primary:var(--sg-color-gold);--primary-foreground:var(--sg-color-ink-on-gold);
  --secondary:var(--sg-color-sidebar);--secondary-foreground:var(--sg-color-canvas);
  --muted:var(--sg-color-pine);--muted-foreground:var(--sg-color-side-text);
  --accent:var(--sg-color-gold);--accent-foreground:var(--sg-color-ink-on-gold);
  --destructive:var(--sg-color-danger-soft);--destructive-foreground:var(--sg-color-pine);
  --success:var(--sg-color-green-tint);--success-foreground:var(--sg-color-pine);
  --warning:var(--sg-color-gold-soft);--warning-foreground:var(--sg-color-ink-on-gold);
  --border:var(--sg-color-sage);--input:var(--sg-color-side-text);--ring:var(--sg-color-gold);
}
*{box-sizing:border-box}
html,body{margin:0;max-width:100%;overflow-x:hidden}
body{min-height:100vh;background:var(--background);color:var(--foreground);font-family:var(--sg-font-family-body);font-size:14px;line-height:1.5}
button,input,select{font:inherit;color:inherit;touch-action:manipulation}
button{cursor:pointer}
button:disabled{cursor:not-allowed;opacity:.62}
button:focus-visible,input:focus-visible,select:focus-visible,a:focus-visible,[tabindex]:focus-visible{outline:3px solid var(--ring);outline-offset:2px}
a{color:inherit;text-decoration-thickness:1px;text-underline-offset:3px}
.shell{max-width:1240px;margin:auto;padding:0 18px}
header{background:var(--sg-color-pine);color:var(--sg-color-canvas);border-bottom:3px solid var(--sg-color-gold);padding:max(16px,env(safe-area-inset-top)) 0 14px}
.headerrow{display:grid;grid-template-columns:1fr auto;gap:12px 20px;align-items:start}
.brand{font-family:var(--sg-font-family-display);font-size:26px;font-weight:900;letter-spacing:-.03em}
.brand b{color:var(--sg-color-gold)}
.kicker{font-family:var(--sg-font-family-mono);font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--sg-color-side-text);margin-top:3px}
.sourcebadge{grid-column:1/-1;display:inline-flex;width:max-content;max-width:100%;align-items:center;gap:7px;border:1px solid var(--sg-color-gold);border-radius:var(--sg-radius-pill);padding:6px 10px;font-size:11px;font-weight:700;color:var(--sg-color-gold)}
.sourcebadge i{width:8px;height:8px;border-radius:50%;background:var(--sg-color-success)}
.theme{min-height:44px;border:1px solid var(--sg-color-sage);border-radius:var(--sg-radius-pill);padding:8px 14px;background:transparent;color:var(--sg-color-canvas);font-weight:700}
.summary{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px;margin-top:14px}
.summary div{border-left:2px solid var(--sg-color-gold);padding-left:9px;min-width:0}
.summary strong{display:block;font-family:var(--sg-font-family-display);font-size:17px}
.summary span{display:block;font-size:11px;color:var(--sg-color-side-text);overflow-wrap:anywhere}
nav{display:flex;overflow-x:auto;overscroll-behavior-inline:contain;scroll-snap-type:x proximity;scrollbar-width:none;background:var(--card);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:10}
nav::-webkit-scrollbar{display:none}
nav .shell{display:flex;width:100%;padding:0}
nav button{flex:1;min-width:145px;min-height:50px;border:0;border-bottom:3px solid transparent;background:transparent;color:var(--muted-foreground);font-weight:800;scroll-snap-align:center}
nav button[aria-selected="true"]{color:var(--foreground);border-color:var(--accent);background:var(--secondary)}
main{padding:22px 0 44px}
.view{display:none}.view.on{display:block}
.hero{display:flex;justify-content:space-between;gap:20px;align-items:end;margin-bottom:16px}
.eyebrow{font-family:var(--sg-font-family-mono);font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted-foreground);font-weight:700}
h1,h2,h3{font-family:var(--sg-font-family-display);line-height:1.12;margin:0}
h1{font-size:clamp(25px,4vw,40px);letter-spacing:-.035em;margin-top:5px}
h2{font-size:21px}h3{font-size:16px}
.lede{max-width:760px;color:var(--muted-foreground);margin:8px 0 0}
.card{background:var(--card);color:var(--card-foreground);border:1px solid var(--border);border-radius:var(--sg-radius-card);box-shadow:var(--sg-shadow-lift);padding:18px;min-width:0}
.grid2{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(320px,.85fr);gap:16px}
.stack{display:grid;gap:14px}
.controls{display:grid;grid-template-columns:minmax(180px,1fr) 150px auto;gap:9px;margin-bottom:15px}
.filtergrid{display:grid;grid-template-columns:2fr repeat(4,1fr);gap:8px;margin-bottom:12px}
input,select{width:100%;min-height:44px;border:1px solid var(--input);border-radius:var(--sg-radius-sm);padding:9px 11px;background:var(--card);color:var(--card-foreground)}
::placeholder{color:var(--muted-foreground);opacity:1}
.btn{min-height:44px;border:1px solid var(--primary);border-radius:var(--sg-radius-sm);padding:8px 14px;background:var(--primary);color:var(--primary-foreground);font-weight:800}
.btn.secondary{background:var(--card);color:var(--card-foreground);border-color:var(--border)}
.btn.ghost{background:transparent;color:var(--foreground);border-color:var(--border)}
.notice{border-left:3px solid var(--accent);padding:10px 12px;background:var(--muted);border-radius:0 var(--sg-radius-sm) var(--sg-radius-sm) 0;color:var(--muted-foreground)}
.notice strong{color:var(--foreground)}
.notice.warning{background:var(--warning);color:var(--warning-foreground)}
.compass{aspect-ratio:1;max-height:580px;width:100%;display:block}
.ring{fill:none;stroke:var(--border);stroke-width:1.5}.axis{stroke:var(--border);stroke-width:1}
.axislabel{fill:var(--muted-foreground);font:700 14px var(--sg-font-family-mono)}
.distance{fill:var(--muted-foreground);font:11px var(--sg-font-family-mono)}
.point{fill:var(--accent);stroke:var(--accent-foreground);stroke-width:1.5}
.centerpoint{fill:var(--primary);stroke:var(--primary-foreground);stroke-width:2}
.resultlist{display:grid;gap:8px}
.result{border:1px solid var(--border);border-radius:var(--sg-radius-lg);padding:12px;background:var(--card)}
.resulthead{display:flex;justify-content:space-between;gap:12px;align-items:start}
.result h3{font-size:14px;overflow-wrap:anywhere}
.pill{display:inline-flex;align-items:center;white-space:nowrap;border-radius:var(--sg-radius-pill);padding:4px 8px;background:var(--secondary);color:var(--secondary-foreground);font:700 10px var(--sg-font-family-mono)}
.meta{display:flex;flex-wrap:wrap;gap:5px 10px;margin-top:7px;color:var(--muted-foreground);font-size:12px}
.tablewrap{overflow:auto;border:1px solid var(--border);border-radius:var(--sg-radius-lg);overscroll-behavior-inline:contain}
table{width:100%;border-collapse:collapse;min-width:800px;background:var(--card)}
th,td{text-align:left;padding:10px 11px;border-bottom:1px solid var(--border);vertical-align:top}
th{position:sticky;top:0;background:var(--secondary);color:var(--secondary-foreground);font:700 10px var(--sg-font-family-mono);letter-spacing:.08em;text-transform:uppercase}
tbody tr:hover{background:var(--muted)}tbody tr:last-child td{border-bottom:0}
td.num{text-align:right;font-family:var(--sg-font-family-mono)}
.tablefoot{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-top:10px;color:var(--muted-foreground);font-size:12px}
.pager{display:flex;gap:7px;align-items:center}.pager button{min-width:44px;min-height:44px}
.metricgrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
.metric{background:var(--card);border:1px solid var(--border);border-radius:var(--sg-radius-lg);padding:14px}
.metric strong{display:block;font-family:var(--sg-font-family-display);font-size:24px}.metric span{color:var(--muted-foreground);font-size:11px}
.barlist{display:grid;gap:10px;margin-top:12px}
.barrow{display:grid;grid-template-columns:110px 1fr 72px;gap:9px;align-items:center}
.track{height:12px;border-radius:var(--sg-radius-pill);background:var(--muted);overflow:hidden}.fill{height:100%;background:var(--accent);border-radius:inherit}
.source{display:grid;gap:10px;margin-top:12px}.source>div{padding:12px;border:1px solid var(--border);border-radius:var(--sg-radius-lg)}
.source code{display:block;margin-top:5px;color:var(--muted-foreground);overflow-wrap:anywhere;font-family:var(--sg-font-family-mono);font-size:11px}
.empty,.loading{padding:32px;text-align:center;color:var(--muted-foreground)}
.loading::before{content:"";display:block;width:32px;height:32px;margin:0 auto 12px;border:4px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .8s linear infinite}
.error{border-color:var(--destructive);color:var(--destructive)}
.detailbtn{margin-top:9px;min-height:44px}
.finding{margin:9px 0 0;padding-left:18px}.finding li{margin:6px 0}
.taglist{display:flex;flex-wrap:wrap;gap:6px;margin-top:9px}
dialog{width:min(720px,calc(100% - 24px));max-height:calc(100vh - 24px);overflow:auto;border:1px solid var(--border);border-radius:var(--sg-radius-card);padding:0;background:var(--popover);color:var(--popover-foreground);box-shadow:var(--sg-shadow-modal)}
dialog::backdrop{background:var(--sg-color-forest);opacity:.72}
.dialoghead{display:flex;align-items:start;justify-content:space-between;gap:12px;padding:18px;border-bottom:1px solid var(--border);position:sticky;top:0;background:var(--popover);z-index:1}
.dialogbody{padding:18px}.close{min-width:44px;min-height:44px}
.definition{display:grid;grid-template-columns:150px 1fr;gap:8px 12px;margin:14px 0}.definition dt{font-weight:800}.definition dd{margin:0;color:var(--muted-foreground);overflow-wrap:anywhere}
footer{border-top:1px solid var(--border);padding:18px 0 max(18px,env(safe-area-inset-bottom));color:var(--muted-foreground);font-size:11px}
@keyframes spin{to{transform:rotate(360deg)}}
@media(max-width:940px){
  .summary{grid-template-columns:repeat(3,1fr)}.grid2{grid-template-columns:1fr}
  .filtergrid{grid-template-columns:repeat(2,1fr)}.filtergrid input{grid-column:1/-1}
}
@media(max-width:768px){
  .shell{padding-left:14px;padding-right:14px}.summary{grid-template-columns:repeat(2,1fr)}
  .metricgrid{grid-template-columns:repeat(2,1fr)}.resultlist{max-height:none}
}
@media(max-width:520px){
  .shell{padding-left:12px;padding-right:12px}
  header{padding-top:max(12px,env(safe-area-inset-top));padding-bottom:11px}
  .headerrow{gap:10px}.brand{font-size:21px}.kicker{font-size:9px}.sourcebadge{font-size:10px}
  .summary{gap:6px;margin-top:10px}.summary div{padding-left:7px}.summary strong{font-size:15px}.summary span{font-size:10px}
  nav button{min-width:120px;min-height:48px;font-size:12px}
  main{padding-top:16px}.hero{display:block}.lede{font-size:13px}
  .card{padding:13px;border-radius:var(--sg-radius-lg)}
  .controls{grid-template-columns:1fr 112px}.controls .btn{grid-column:1/-1}
  .filtergrid{grid-template-columns:1fr 1fr}.filtergrid input{grid-column:1/-1}
  .metricgrid{grid-template-columns:1fr 1fr}.metric strong{font-size:20px}
  .barrow{grid-template-columns:80px 1fr 54px;font-size:11px}
  .resulthead{display:block}.resulthead .pill{margin-top:6px}
  .tablefoot{align-items:flex-start;flex-direction:column}.definition{grid-template-columns:1fr}.definition dd{margin-bottom:8px}
}
@media(max-width:360px){
  .summary,.metricgrid,.filtergrid{grid-template-columns:1fr}
  .filtergrid input{grid-column:auto}.controls{grid-template-columns:1fr}.controls .btn{grid-column:auto}
}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important;animation:none!important}}
</style>
<!-- ATLAS_INDEX_SCRIPT -->
</head>
<body>
<header><div class="shell">
  <div class="headerrow">
    <div><div class="brand">Common <b>Ground</b></div><div class="kicker">United States Union Atlas</div></div>
    <button class="theme" id="theme" type="button">Dark theme</button>
    <div class="sourcebadge"><i></i> Real OLMS, NLRB, BLS, research, and Census records</div>
  </div>
  <div class="summary" id="summary" aria-label="Dataset summary"></div>
</div></header>
<nav aria-label="Atlas views"><div class="shell" role="tablist">
  <button type="button" role="tab" data-view="near" aria-selected="true">Near you</button>
  <button type="button" role="tab" data-view="directory" aria-selected="false">Directory</button>
  <button type="button" role="tab" data-view="organizing" aria-selected="false">Organizing</button>
  <button type="button" role="tab" data-view="research" aria-selected="false">Research</button>
  <button type="button" role="tab" data-view="affiliations" aria-selected="false">Affiliations</button>
  <button type="button" role="tab" data-view="status" aria-selected="false">Data status</button>
</div></nav>
<main><div class="shell">
  <section class="view on" id="view-near"><div class="loading" role="status">Loading compressed union directory…</div></section>
  <section class="view" id="view-directory"></section>
  <section class="view" id="view-organizing"></section>
  <section class="view" id="view-research"></section>
  <section class="view" id="view-affiliations"></section>
  <section class="view" id="view-status"></section>
</div></main>
<footer><div class="shell" id="footer"></div></footer>
<dialog id="detail-dialog" aria-labelledby="detail-title"><div class="dialoghead"><h2 id="detail-title">Record details</h2><button class="btn secondary close" id="detail-close" type="button" aria-label="Close details">×</button></div><div class="dialogbody" id="detail-body"></div></dialog>
<!-- ATLAS_EMBEDDED_SCRIPT -->
<script>
(function(){
"use strict";
var INDEX=window.UNION_ATLAS_INDEX;
var EMBEDDED=window.UNION_ATLAS_EMBEDDED||null;
var state={view:"near",zip:"01103",radius:100,nearPage:0,directoryPage:0,query:"",stateFilter:"",affiliation:"",level:"",coverage:"",organizingPage:0,organizingQuery:"",organizingState:"",outcome:"",year:"",researchPage:0,researchQuery:"",themeFilter:""};
var pageSize=50,requestId=0,pending=new Map(),recordCache=new Map(),searchTimer=null,ready=false;
var whole=new Intl.NumberFormat("en-US"),compactNumber=new Intl.NumberFormat("en-US",{notation:"compact",maximumFractionDigits:1});
var esc=function(value){return String(value==null?"":value).replace(/[&<>"']/g,function(c){return({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"})[c]})};
var fmt=function(value){return Number.isFinite(value)?whole.format(value):"Not reported"};
var compact=function(value){return Number.isFinite(value)?compactNumber.format(value):"—"};
var pct=function(value,total){return total?Math.round(value/total*100):0};
var el=function(selector){return document.querySelector(selector)};
var all=function(selector){return Array.from(document.querySelectorAll(selector))};

if(!INDEX){fatal("Atlas index missing.");return}
var worker;
try{
  if(EMBEDDED){
    var blob=new Blob([EMBEDDED.worker],{type:"text/javascript"});
    worker=new Worker(URL.createObjectURL(blob));
  }else{
    worker=new Worker("union-atlas-worker.js");
  }
}catch(error){fatal(error.message);return}
worker.onmessage=function(event){
  var message=event.data||{},entry=pending.get(message.id);
  if(!entry)return;
  pending.delete(message.id);
  message.ok?entry.resolve(message.result):entry.reject(new Error(message.error));
};
worker.onerror=function(event){fatal(event.message||"Atlas worker failed.")};
function rpc(type,payload,transfers){
  return new Promise(function(resolve,reject){
    var id=++requestId;pending.set(id,{resolve:resolve,reject:reject});
    worker.postMessage({id:id,type:type,payload:payload||{}},transfers||[]);
  });
}
function base64Buffer(value){
  var binary=atob(value),bytes=new Uint8Array(binary.length);
  for(var i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
  return bytes.buffer;
}
function fatal(message){
  document.querySelector("main .shell").innerHTML='<div class="card error"><h1>Atlas could not load</h1><p>'+esc(message)+'</p><p>Use the standalone export for local file access, or serve the docs folder over HTTP.</p></div>';
}
function setTheme(theme){
  document.documentElement.dataset.theme=theme;
  el("#theme").textContent=theme==="dark"?"Light theme":"Dark theme";
  try{localStorage.setItem("union-atlas-theme",theme)}catch(_){}
}
var savedTheme=null;try{savedTheme=localStorage.getItem("union-atlas-theme")}catch(_){}
setTheme(savedTheme||"light");
el("#theme").addEventListener("click",function(){setTheme(document.documentElement.dataset.theme==="dark"?"light":"dark")});

function summary(){
  var q=INDEX.quality,d=INDEX.datasets;
  el("#summary").innerHTML=[
    [fmt(q.records),"union records"],
    [fmt(q.withAddressCoordinates),"address-matched HQs"],
    [fmt(d.organizing.records),"NLRB cases"],
    [fmt(d.research.articles),"research articles"],
    [fmt(d.research.statistics),"verified statistics"]
  ].map(function(x){return"<div><strong>"+x[0]+"</strong><span>"+x[1]+"</span></div>"}).join("");
  var src=INDEX.metadata.sources[0];
  el("#footer").innerHTML="Source commit <b>"+esc(src.commit.slice(0,12))+"</b> · "+esc(INDEX.metadata.privacy);
}
function loading(view,label){el("#view-"+view).innerHTML='<div class="loading" role="status">'+esc(label)+'</div>'}
function bindPager(prefix,result,render){
  var prev=el("#"+prefix+"-prev"),next=el("#"+prefix+"-next");
  if(prev)prev.addEventListener("click",function(){render(Math.max(0,result.page-1))});
  if(next)next.addEventListener("click",function(){render(Math.min(result.pages-1,result.page+1))});
}
function pager(prefix,result){
  return'<div class="tablefoot"><span>Showing '+fmt(result.rows.length)+' of '+fmt(result.total)+' records.</span><div class="pager"><button class="btn secondary" id="'+prefix+'-prev" type="button" '+(result.page===0?"disabled":"")+'>←</button><span class="pill">Page '+(result.page+1)+' / '+result.pages+'</span><button class="btn secondary" id="'+prefix+'-next" type="button" '+(result.page>=result.pages-1?"disabled":"")+'>→</button></div></div>';
}
function switchView(view){
  state.view=view;
  all(".view").forEach(function(node){node.classList.toggle("on",node.id==="view-"+view)});
  all("nav button").forEach(function(button){
    var active=button.dataset.view===view;button.setAttribute("aria-selected",String(active));
    if(active)button.scrollIntoView({block:"nearest",inline:"center"});
  });
  if(!ready)return;
  ({near:renderNear,directory:renderDirectory,organizing:renderOrganizing,research:renderResearch,affiliations:renderAffiliations,status:renderStatus})[view]();
}
all("nav button").forEach(function(button){button.addEventListener("click",function(){switchView(button.dataset.view)})});

function compassSvg(result){
  var size=560,c=280,max=220,radius=result.radius;
  var rings=[.25,.5,.75,1].map(function(scale){return'<circle class="ring" cx="'+c+'" cy="'+c+'" r="'+(max*scale)+'"/><text class="distance" x="'+(c+5)+'" y="'+(c-max*scale+14)+'">'+Math.round(radius*scale)+' mi</text>'}).join("");
  var points=result.compass.map(function(item,index){
    var angle=(item.bearing-90)*Math.PI/180,r=Math.min(max,item.distance/radius*max),x=c+Math.cos(angle)*r,y=c+Math.sin(angle)*r;
    return'<circle class="point" tabindex="0" role="img" aria-label="'+esc(item.name)+", "+item.distance.toFixed(1)+' miles" cx="'+x.toFixed(1)+'" cy="'+y.toFixed(1)+'" r="'+(index<8?7:5)+'"><title>'+esc(item.name)+" · "+item.distance.toFixed(1)+' mi</title></circle>';
  }).join("");
  return'<svg class="compass" viewBox="0 0 560 560" role="img" aria-labelledby="compass-title"><title id="compass-title">Nearest union headquarters around ZIP '+esc(result.zip)+'</title>'+rings+'<line class="axis" x1="280" y1="45" x2="280" y2="515"/><line class="axis" x1="45" y1="280" x2="515" y2="280"/><text class="axislabel" x="274" y="27">N</text><text class="axislabel" x="274" y="548">S</text><text class="axislabel" x="20" y="285">W</text><text class="axislabel" x="525" y="285">E</text><circle class="centerpoint" cx="280" cy="280" r="9"/>'+points+"</svg>";
}
function cacheRows(rows){rows.forEach(function(row){if(row.id)recordCache.set(row.id,row)})}
function unionCard(row){
  return'<article class="result"><div class="resulthead"><h3>'+esc(row.name)+'</h3>'+(Number.isFinite(row.distance)?'<span class="pill">'+row.distance.toFixed(1)+' mi · '+Math.round(row.bearing)+'°</span>':'')+'</div><div class="meta"><span>'+esc([row.city,row.state,row.zip].filter(Boolean).join(", "))+'</span><span>'+esc(row.affiliation||"Independent")+'</span><span>'+fmt(row.membership?.latest?.count)+' reported members</span><span>'+esc(row.geocodePrecision==="address"?"Address-matched":"ZIP-area point")+'</span></div><button class="btn secondary detailbtn" type="button" data-record="'+esc(row.id)+'">View source details</button></article>';
}
function bindDetails(){
  all("[data-record]").forEach(function(button){button.addEventListener("click",function(){showRecord(button.dataset.record)})});
}
async function showRecord(id){
  var row=recordCache.get(id);if(!row)return;
  el("#detail-title").textContent=row.name;
  el("#detail-body").innerHTML='<div class="loading" role="status">Loading source details…</div>';
  el("#detail-dialog").showModal();
  try{
    row=await rpc("record-detail",{id:id});recordCache.set(id,row);
  }catch(error){
    el("#detail-body").innerHTML='<div class="error">Could not load details: '+esc(error.message)+'</div>';return;
  }
  var latest=row.membership?.latest,finance=row.finances,c=row.publicRecordCoverage||{};
  el("#detail-title").textContent=row.name;
  el("#detail-body").innerHTML='<div class="taglist"><span class="pill">'+esc(row.affiliation||"Independent")+'</span><span class="pill">'+esc(row.level)+'</span><span class="pill">'+esc(row.geocodePrecision||"not geocoded")+'</span></div><dl class="definition"><dt>OLMS file</dt><dd>'+esc(row.id.replace("olms-",""))+'</dd><dt>Headquarters</dt><dd>'+esc([row.city,row.state,row.zip].filter(Boolean).join(", "))+'</dd><dt>Membership</dt><dd>'+fmt(latest?.count)+(latest?.year?" in "+latest.year:"")+' · '+fmt(row.membershipHistory?.length||0)+' annual records</dd><dt>Finances</dt><dd>'+ (finance?"Receipts "+compact(finance.totalReceipts)+" · disbursements "+compact(finance.totalDisbursements)+" · assets "+compact(finance.totalAssets)+" · "+esc(finance.form||"form not reported"):"Not reported")+'</dd><dt>Public record coverage</dt><dd>'+fmt(c.sources||0)+' cited sources · '+fmt(c.filedOfficers||0)+' filed officers · '+fmt(c.chapters||0)+' chapters · '+fmt(c.contracts||0)+' contracts</dd><dt>Updated</dt><dd>'+esc(row.updatedAt||"Not reported")+'</dd></dl>'+(row.background?'<h3>Background</h3><p>'+esc(row.background)+'</p>':'')+(row.website?'<p><a href="'+esc(row.website)+'" target="_blank" rel="noopener">Visit published website</a></p>':'<p class="notice">No published website in the source record.</p>');
}
el("#detail-close").addEventListener("click",function(){el("#detail-dialog").close()});
el("#detail-dialog").addEventListener("click",function(event){if(event.target===el("#detail-dialog"))el("#detail-dialog").close()});

async function renderNear(page){
  if(Number.isFinite(page))state.nearPage=page;
  loading("near","Calculating real headquarters distances…");
  try{
    var result=await rpc("near",{zip:state.zip,radius:state.radius,page:state.nearPage,pageSize:pageSize});
    if(result.error){
      el("#view-near").innerHTML='<div class="hero"><div><div class="eyebrow">Geographic discovery</div><h1>Nearest union headquarters</h1></div></div>'+nearControls()+'<div class="notice warning"><strong>ZIP not found.</strong> '+esc(result.error)+'</div>';
      bindNearControls();return;
    }
    cacheRows(result.rows);
    el("#view-near").innerHTML='<div class="hero"><div><div class="eyebrow">Geographic discovery</div><h1>Nearest union headquarters</h1><p class="lede">Address-range coordinates where Census matched the public filing address; ZIP-area representative points otherwise.</p></div></div>'+nearControls()+'<div class="grid2"><div class="card">'+compassSvg(result)+'<div class="notice"><strong>Compass:</strong> nearest '+Math.min(24,result.total)+' of '+fmt(result.total)+'. The paginated list contains every result.</div></div><div class="stack"><div class="notice"><strong>'+fmt(result.total)+' records</strong> within '+result.radius+' miles of ZIP '+esc(result.zip)+'.</div><div class="resultlist">'+(result.rows.length?result.rows.map(unionCard).join(""):'<div class="card empty">No headquarters in this radius.</div>')+'</div>'+pager("near",result)+'</div></div>';
    bindNearControls();bindPager("near",result,renderNear);bindDetails();
  }catch(error){fatal(error.message)}
}
function nearControls(){
  return'<div class="controls"><input id="zip" inputmode="numeric" maxlength="5" aria-label="Five digit ZIP code" value="'+esc(state.zip)+'" placeholder="ZIP code"><select id="radius" aria-label="Search radius">'+[25,50,100,250,500,1000].map(function(n){return'<option '+(n===state.radius?"selected":"")+' value="'+n+'">'+n+' miles</option>'}).join("")+'</select><button class="btn" id="locate" type="button">Plot records</button></div>';
}
function bindNearControls(){
  el("#locate").addEventListener("click",function(){state.zip=el("#zip").value.trim();state.radius=Number(el("#radius").value);state.nearPage=0;renderNear()});
  el("#zip").addEventListener("keydown",function(event){if(event.key==="Enter")el("#locate").click()});
}

async function renderDirectory(page){
  if(Number.isFinite(page))state.directoryPage=page;
  loading("directory","Searching canonical union records…");
  try{
    var result=await rpc("directory-search",{query:state.query,state:state.stateFilter,affiliation:state.affiliation,level:state.level,coverage:state.coverage,page:state.directoryPage,pageSize:pageSize});
    cacheRows(result.rows);
    var states=INDEX.states.map(function(row){return row.key}).sort(),affiliations=INDEX.affiliations.map(function(row){return row.key}).sort();
    el("#view-directory").innerHTML='<div class="hero"><div><div class="eyebrow">Canonical public records</div><h1>US union directory</h1><p class="lede">Full organization-level source coverage. Missing fields remain missing; private contacts and personal records stay excluded.</p></div></div><div class="card"><div class="filtergrid"><input id="query" type="search" value="'+esc(state.query)+'" placeholder="Name, local, sector, city, state, ZIP…"><select id="state-filter"><option value="">All states</option>'+options(states,state.stateFilter)+'</select><select id="aff-filter"><option value="">All affiliations</option>'+options(affiliations,state.affiliation)+'</select><select id="level-filter"><option value="">All levels</option>'+options(["national","intermediate","local","independent","unknown"],state.level)+'</select><select id="coverage-filter"><option value="">Any coverage</option>'+optionValues([["website","Has website"],["membership","Has membership"],["address","Address geocoded"]],state.coverage)+'</select></div><div class="tablewrap"><table><thead><tr><th>Union</th><th>Affiliation</th><th>HQ</th><th>Level</th><th>Members</th><th>Geography</th><th>Details</th></tr></thead><tbody>'+(result.rows.length?result.rows.map(function(row){return'<tr><td><strong>'+esc(row.name)+'</strong><br><span class="meta">'+esc(row.id.toUpperCase())+(row.localNumber?" · #"+esc(row.localNumber):"")+'</span></td><td>'+esc(row.affiliation||"Independent")+'</td><td>'+esc([row.city,row.state,row.zip].filter(Boolean).join(", "))+'</td><td>'+esc(row.level)+'</td><td class="num">'+fmt(row.membership?.latest?.count)+'</td><td>'+esc(row.geocodePrecision||"Unavailable")+'</td><td><button class="btn secondary" type="button" data-record="'+esc(row.id)+'">Open</button></td></tr>'}).join(""):'<tr><td colspan="7" class="empty">No records match.</td></tr>')+'</tbody></table></div>'+pager("directory",result)+'</div>';
    bindDirectoryFilters();bindPager("directory",result,renderDirectory);bindDetails();
  }catch(error){fatal(error.message)}
}
function options(values,current){return values.map(function(value){return'<option value="'+esc(value)+'" '+(value===current?"selected":"")+'>'+esc(value)+'</option>'}).join("")}
function optionValues(values,current){return values.map(function(pair){return'<option value="'+esc(pair[0])+'" '+(pair[0]===current?"selected":"")+'>'+esc(pair[1])+'</option>'}).join("")}
function bindDirectoryFilters(){
  el("#query").addEventListener("input",function(event){state.query=event.target.value;state.directoryPage=0;clearTimeout(searchTimer);searchTimer=setTimeout(function(){renderDirectory()},180)});
  [["#state-filter","stateFilter"],["#aff-filter","affiliation"],["#level-filter","level"],["#coverage-filter","coverage"]].forEach(function(pair){el(pair[0]).addEventListener("change",function(event){state[pair[1]]=event.target.value;state.directoryPage=0;renderDirectory()})});
}

async function renderOrganizing(page){
  if(Number.isFinite(page))state.organizingPage=page;
  loading("organizing","Loading NLRB organizing evidence…");
  try{
    var result=await rpc("organizing-search",{query:state.organizingQuery,state:state.organizingState,outcome:state.outcome,year:state.year,page:state.organizingPage,pageSize:pageSize});
    var years=[];for(var year=2026;year>=2016;year--)years.push(String(year));
    var states=INDEX.states.map(function(row){return row.key}).sort();
    el("#view-organizing").innerHTML='<div class="hero"><div><div class="eyebrow">Ten years of federal cases</div><h1>Organizing attempts</h1><p class="lede">Real NLRB representation cases. This does not observe drives that never filed, voluntary recognition, or workers outside NLRA jurisdiction.</p></div></div><div class="metricgrid"><div class="metric"><strong>'+fmt(result.summary.counts.representation_cases)+'</strong><span>representation cases</span></div><div class="metric"><strong>'+fmt(result.summary.counts.new_representation_attempts_rc)+'</strong><span>new RC attempts</span></div><div class="metric"><strong>'+fmt(result.summary.counts.rc_with_extracted_petitioner_organization)+'</strong><span>with petitioner organization</span></div><div class="metric"><strong>'+pct(result.summary.rc_decided_representation_outcomes.labor_organization_certified,result.summary.rc_decided_representation_outcomes.labor_organization_certified+result.summary.rc_decided_representation_outcomes.no_representative_certified)+'%</strong><span>certified among defined decided outcomes</span></div></div><div class="card" style="margin-top:16px"><div class="filtergrid"><input id="organizing-query" type="search" value="'+esc(state.organizingQuery)+'" placeholder="Employer, case, petitioner organization…"><select id="organizing-state"><option value="">All states</option>'+options(states,state.organizingState)+'</select><select id="outcome"><option value="">All outcomes</option>'+options(["pending","representative_certified","no_representative_certified","withdrawn_non_adjusted","withdrawn_adjusted","dismissed_non_adjusted","dismissed_adjusted"],state.outcome)+'</select><select id="year"><option value="">All years</option>'+options(years,state.year)+'</select></div><div class="resultlist">'+(result.rows.length?result.rows.map(function(row){return'<article class="result"><div class="resulthead"><h3>'+esc(row.employer||row.caseName)+'</h3><span class="pill">'+esc(row.outcome||row.status)+'</span></div><div class="meta"><span>'+esc(row.caseNumber)+'</span><span>'+esc([row.city,row.state].filter(Boolean).join(", "))+'</span><span>Filed '+esc(row.filedAt)+'</span><span>'+fmt(row.employeesOnPetition)+' employees on petition</span></div><p>'+esc(row.petitionerOrganizations.join("; ")||"Petitioner not extracted")+'</p><a href="'+esc(row.caseUrl)+'" target="_blank" rel="noopener">Open NLRB case</a></article>'}).join(""):'<div class="empty">No organizing cases match.</div>')+'</div>'+pager("organizing",result)+'</div><div class="notice warning" style="margin-top:14px"><strong>Interpretation boundary.</strong> '+esc(result.summary.rc_decided_representation_outcomes.warning)+'</div>';
    bindOrganizingFilters();bindPager("organizing",result,renderOrganizing);
  }catch(error){fatal(error.message)}
}
function bindOrganizingFilters(){
  el("#organizing-query").addEventListener("input",function(event){state.organizingQuery=event.target.value;state.organizingPage=0;clearTimeout(searchTimer);searchTimer=setTimeout(function(){renderOrganizing()},180)});
  [["#organizing-state","organizingState"],["#outcome","outcome"],["#year","year"]].forEach(function(pair){el(pair[0]).addEventListener("change",function(event){state[pair[1]]=event.target.value;state.organizingPage=0;renderOrganizing()})});
}

async function renderResearch(page){
  if(Number.isFinite(page))state.researchPage=page;
  loading("research","Loading verified research and national statistics…");
  try{
    var result=await rpc("research-search",{query:state.researchQuery,theme:state.themeFilter,page:state.researchPage,pageSize:20});
    var themes=[...new Set(result.rows.flatMap(function(row){return row.themes||[]}))].sort();
    var membership=result.statistics.filter(function(row){return row.topic==="union_membership"&&row.unit==="percent"&&row.label==="Union membership rate"}).sort(function(a,b){return String(a.period).localeCompare(String(b.period))});
    el("#view-research").innerHTML='<div class="hero"><div><div class="eyebrow">Evidence library</div><h1>Labor research</h1><p class="lede">Peer-reviewed studies, government publications, verified findings, limitations, and BLS national statistics from the pinned source.</p></div></div><div class="grid2"><div class="card"><h2>Union membership rate</h2><div class="barlist">'+membership.map(function(row){return'<div class="barrow"><strong>'+esc(row.period)+'</strong><div class="track"><div class="fill" style="width:'+Math.min(100,Number(row.value)*6)+'%"></div></div><span>'+Number(row.value).toFixed(1)+'%</span></div>'}).join("")+'</div></div><div class="card"><h2>Evidence coverage</h2><div class="metricgrid"><div class="metric"><strong>'+fmt(INDEX.datasets.research.articles)+'</strong><span>articles</span></div><div class="metric"><strong>'+fmt(INDEX.datasets.research.statistics)+'</strong><span>statistics</span></div></div><p class="notice">Each record retains source URLs, methods, evidence status, and limitations.</p></div></div><div class="card" style="margin-top:16px"><div class="filtergrid"><input id="research-query" type="search" value="'+esc(state.researchQuery)+'" placeholder="Title, author, finding, method…"><select id="theme-filter"><option value="">All themes</option>'+options(themes,state.themeFilter)+'</select></div><div class="resultlist">'+(result.rows.length?result.rows.map(function(row){return'<article class="result"><div class="resulthead"><h3>'+esc(row.title)+'</h3><span class="pill">'+esc(row.year||"Year not reported")+'</span></div><div class="meta"><span>'+esc((row.authors||[]).join(", "))+'</span><span>'+esc(row.journal||row.publication_type)+'</span><span>'+esc(row.evidence_status)+'</span></div><p>'+esc(row.method||"Method not reported")+'</p><h3>Key findings</h3><ul class="finding">'+(row.key_findings||[]).map(function(item){return"<li>"+esc(item)+"</li>"}).join("")+'</ul><h3>Limitations</h3><ul class="finding">'+(row.limitations||[]).map(function(item){return"<li>"+esc(item)+"</li>"}).join("")+'</ul><a href="'+esc(row.url)+'" target="_blank" rel="noopener">Open source record</a></article>'}).join(""):'<div class="empty">No research records match.</div>')+'</div>'+pager("research",result)+'</div>';
    el("#research-query").addEventListener("input",function(event){state.researchQuery=event.target.value;state.researchPage=0;clearTimeout(searchTimer);searchTimer=setTimeout(function(){renderResearch()},180)});
    el("#theme-filter").addEventListener("change",function(event){state.themeFilter=event.target.value;state.researchPage=0;renderResearch()});
    bindPager("research",result,renderResearch);
  }catch(error){fatal(error.message)}
}

function renderAffiliations(){
  var rows=INDEX.affiliations,max=rows[0]?.unions||1;
  el("#view-affiliations").innerHTML='<div class="hero"><div><div class="eyebrow">Filed hierarchy</div><h1>All affiliations</h1><p class="lede">Every filed affiliation code is shown. Membership totals can duplicate workers across hierarchy levels.</p></div></div><div class="grid2"><div class="card"><h2>Largest by filing count</h2><div class="barlist">'+rows.slice(0,25).map(function(row){return'<div class="barrow"><strong>'+esc(row.key)+'</strong><div class="track"><div class="fill" style="width:'+pct(row.unions,max)+'%"></div></div><span class="num">'+fmt(row.unions)+'</span></div>'}).join("")+'</div></div><div class="card"><h2>Complete source rollup</h2><div class="tablewrap"><table><thead><tr><th>Affiliation</th><th>Unions</th><th>Members</th><th>Address</th><th>ZIP fallback</th><th>Websites</th></tr></thead><tbody>'+rows.map(function(row){return'<tr><td><strong>'+esc(row.key)+'</strong></td><td class="num">'+fmt(row.unions)+'</td><td class="num">'+fmt(row.reportedMembers)+'</td><td class="num">'+fmt(row.addressGeocoded)+'</td><td class="num">'+fmt(row.zctaGeocoded)+'</td><td class="num">'+fmt(row.websites)+'</td></tr>'}).join("")+'</tbody></table></div></div></div>';
}
function renderStatus(){
  var q=INDEX.quality,src=INDEX.metadata.sources,p=INDEX.packages;
  el("#view-status").innerHTML='<div class="hero"><div><div class="eyebrow">Provenance and completeness</div><h1>Data status</h1><p class="lede">Exact denominators, source gaps, package size, dates, and limitations. No seeded replacements.</p></div></div><div class="metricgrid"><div class="metric"><strong>'+fmt(q.records)+'</strong><span>canonical unions</span></div><div class="metric"><strong>'+pct(q.withAddressCoordinates,q.records)+'%</strong><span>address matched</span></div><div class="metric"><strong>'+pct(q.withMembership,q.records)+'%</strong><span>membership reported</span></div><div class="metric"><strong>'+pct(q.withWebsites,q.records)+'%</strong><span>website reported</span></div></div><div class="grid2" style="margin-top:16px"><div class="card"><h2>Directory coverage</h2><div class="barlist">'+[["Address coordinates",q.withAddressCoordinates],["ZIP fallback",q.withZctaCoordinates],["Any coordinates",q.withCoordinates],["Membership",q.withMembership],["Website",q.withWebsites]].map(function(row){return'<div class="barrow"><strong>'+row[0]+'</strong><div class="track"><div class="fill" style="width:'+pct(row[1],q.records)+'%"></div></div><span>'+pct(row[1],q.records)+'%</span></div>'}).join("")+'</div><div class="notice warning" style="margin-top:14px"><strong>Source gaps, not synthetic fixes.</strong> '+fmt(q.withoutCoordinates)+' lack coordinates; '+fmt(q.records-q.withMembership)+' lack membership; '+fmt(q.records-q.withWebsites)+' lack websites.</div></div><div class="card"><h2>Runtime packages</h2><dl class="definition"><dt>Directory</dt><dd>'+compact(p.directory.bytes)+' compressed; loads first</dd><dt>Record details</dt><dd>'+compact(p.details.bytes)+' compressed; loads on first detail open</dd><dt>Evidence</dt><dd>'+compact(p.evidence.bytes)+' compressed; loads only when Organizing or Research opens</dd><dt>Search</dt><dd>Worker thread, 180 ms debounce, 50-row pages, no hidden result cap</dd><dt>Offline</dt><dd><a href="union-atlas-standalone.html">Open true single-file Atlas</a></dd></dl></div></div><div class="card" style="margin-top:16px"><h2>Traceable sources</h2><div class="source">'+src.map(function(item){return'<div><strong>'+esc(item.name)+'</strong><code>'+(item.repository?"Repository: "+esc(item.repository)+" · commit "+esc(item.commit):esc(item.url))+'</code><code>'+esc((item.paths||[]).join(", ")||item.purpose||"")+'</code></div>'}).join("")+'</div><div class="notice" style="margin-top:12px"><strong>Privacy boundary.</strong> '+esc(INDEX.metadata.privacy)+'</div></div>';
}

summary();
var initPromise;
if(EMBEDDED){
  var directoryBuffer=base64Buffer(EMBEDDED.directory),detailsBuffer=base64Buffer(EMBEDDED.details),evidenceBuffer=base64Buffer(EMBEDDED.evidence);
  initPromise=rpc("init-buffer",{packages:INDEX.packages,directory:directoryBuffer,details:detailsBuffer,evidence:evidenceBuffer},[directoryBuffer,detailsBuffer,evidenceBuffer]);
}else{
  initPromise=rpc("init-url",{packages:{directory:INDEX.packages.directory.path,details:INDEX.packages.details.path,evidence:INDEX.packages.evidence.path}});
}
initPromise.then(function(){ready=true;switchView(state.view)}).catch(function(error){fatal(error.message)});
})();
</script>
</body>
</html>`;

const indexTag = '<script src="union-atlas-index.js"></script>';
const normalPage = page
  .replace('<!-- ATLAS_INDEX_SCRIPT -->', indexTag)
  .replace('<!-- ATLAS_EMBEDDED_SCRIPT -->', '');
const standaloneIndex = `<script>${indexSource}</script>`;
const standaloneEmbedded = `<script>window.UNION_ATLAS_EMBEDDED=${JSON.stringify({
  worker: workerSource,
  directory: directoryPackage.toString('base64'),
  details: detailsPackage.toString('base64'),
  evidence: evidencePackage.toString('base64')
})};</script>`;
const standalonePage = page
  .replace('<!-- ATLAS_INDEX_SCRIPT -->', standaloneIndex)
  .replace('<!-- ATLAS_EMBEDDED_SCRIPT -->', standaloneEmbedded);

fs.writeFileSync(outputPath, normalPage, 'utf8');
fs.writeFileSync(standalonePath, standalonePage, 'utf8');
fs.writeFileSync(workerPath, workerSource, 'utf8');
console.log(`Built ${outputPath}`);
console.log(`Built ${standalonePath}`);
