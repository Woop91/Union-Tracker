import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const tokens = fs.readFileSync(path.join(root, 'design', 'solid-ground', 'tokens.css'), 'utf8');
const outputPath = path.join(root, 'docs', 'union-atlas.html');

const page = String.raw`<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<link rel="icon" href="data:,">
<title>Common Ground — Source-backed Union Atlas</title>
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
  --muted-foreground:var(--sg-color-sec);--muted-surface:var(--sg-color-card-2);
  --accent:var(--sg-color-gold);--accent-foreground:var(--sg-color-ink-on-gold);
  --destructive:var(--sg-color-danger);--destructive-foreground:var(--sg-color-card);
  --border:var(--sg-color-muted);--input:var(--sg-color-sec);--ring:var(--sg-color-green-text);
  --surface:var(--card);--surface-2:var(--secondary);--muted:var(--muted-foreground);
  --primary-ink:var(--primary-foreground);--accent-ink:var(--accent-foreground);
  --focus:var(--ring);--danger:var(--destructive);
}
html[data-theme="dark"]{
  --background:var(--sg-color-forest);--foreground:var(--sg-color-canvas);
  --card:var(--sg-color-green);--card-foreground:var(--sg-color-canvas);
  --popover:var(--sg-color-pine);--popover-foreground:var(--sg-color-canvas);
  --primary:var(--sg-color-gold);--primary-foreground:var(--sg-color-ink-on-gold);
  --secondary:var(--sg-color-sidebar);--secondary-foreground:var(--sg-color-canvas);
  --muted-foreground:var(--sg-color-side-text);--muted-surface:var(--sg-color-pine);
  --accent:var(--sg-color-gold);--accent-foreground:var(--sg-color-ink-on-gold);
  --destructive:var(--sg-color-danger-soft);--destructive-foreground:var(--sg-color-pine);
  --border:var(--sg-color-sage);--input:var(--sg-color-side-text);--ring:var(--sg-color-gold);
  --surface:var(--card);--surface-2:var(--secondary);--muted:var(--muted-foreground);
  --primary-ink:var(--primary-foreground);--accent-ink:var(--accent-foreground);
  --focus:var(--ring);--danger:var(--destructive);
}
*{box-sizing:border-box}
html,body{margin:0;max-width:100%;overflow-x:hidden}
body{min-height:100vh;background:var(--background);color:var(--foreground);font-family:var(--sg-font-family-body);font-size:14px;line-height:1.45}
button,input,select{font:inherit;color:inherit;touch-action:manipulation}
button{cursor:pointer}
button:focus-visible,input:focus-visible,select:focus-visible,a:focus-visible{outline:3px solid var(--focus);outline-offset:2px}
.shell{max-width:1180px;margin:auto;padding:0 18px}
header{background:var(--sg-color-pine);color:var(--sg-color-canvas);border-bottom:3px solid var(--sg-color-gold);padding:max(16px,env(safe-area-inset-top)) 0 14px}
.headerrow{display:flex;align-items:flex-start;justify-content:space-between;gap:20px}
.brand{font-family:var(--sg-font-family-display);font-size:24px;font-weight:900;letter-spacing:-.02em}
.brand b{color:var(--sg-color-gold)}
.kicker{font-family:var(--sg-font-family-mono);font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--sg-color-side-text);margin-top:3px}
.sourcebadge{margin-top:10px;display:inline-flex;align-items:center;gap:7px;border:1px solid var(--sg-color-gold);border-radius:var(--sg-radius-pill);padding:6px 10px;font-size:11px;font-weight:700;color:var(--sg-color-gold)}
.sourcebadge i{width:8px;height:8px;border-radius:50%;background:var(--sg-color-success)}
.theme{min-height:44px;border:1px solid var(--sg-color-sage);border-radius:var(--sg-radius-pill);padding:8px 14px;background:transparent;color:var(--sg-color-canvas);font-weight:700}
.summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-top:14px}
.summary div{border-left:2px solid var(--sg-color-gold);padding-left:9px}
.summary strong{display:block;font-family:var(--sg-font-family-display);font-size:17px}
.summary span{font-size:11px;color:var(--sg-color-side-text)}
nav{display:flex;overflow-x:auto;overscroll-behavior-inline:contain;scroll-snap-type:x proximity;scrollbar-width:none;background:var(--surface);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:10}
nav::-webkit-scrollbar{display:none}
nav .shell{display:flex;width:100%}
nav button{flex:1;min-width:170px;min-height:50px;border:0;border-bottom:3px solid transparent;background:transparent;color:var(--muted);font-weight:800;scroll-snap-align:center}
nav button[aria-selected="true"]{color:var(--foreground);border-color:var(--accent);background:var(--surface-2)}
main{padding:22px 0 44px}
.view{display:none}.view.on{display:block}
.hero{display:flex;justify-content:space-between;gap:20px;align-items:end;margin-bottom:16px}
.eyebrow{font-family:var(--sg-font-family-mono);font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);font-weight:700}
h1,h2,h3{font-family:var(--sg-font-family-display);line-height:1.12;margin:0}
h1{font-size:clamp(25px,4vw,38px);letter-spacing:-.03em;margin-top:5px}
h2{font-size:21px} h3{font-size:16px}
.lede{max-width:720px;color:var(--muted);margin:8px 0 0}
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--sg-radius-card);box-shadow:var(--sg-shadow-lift);padding:18px;min-width:0}
.grid2{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(300px,.85fr);gap:16px}
.stack{display:grid;gap:14px}
.controls{display:grid;grid-template-columns:1fr 150px auto;gap:9px;margin-bottom:15px}
input,select{width:100%;min-height:44px;border:1px solid var(--border);border-radius:var(--sg-radius-sm);padding:9px 11px;background:var(--surface);color:var(--foreground)}
.btn{min-height:44px;border:1px solid var(--primary);border-radius:var(--sg-radius-sm);padding:8px 14px;background:var(--primary);color:var(--primary-ink);font-weight:800}
.btn.secondary{background:var(--surface);color:var(--foreground);border-color:var(--border)}
.notice{border-left:3px solid var(--accent);padding:10px 12px;background:var(--surface-2);border-radius:0 var(--sg-radius-sm) var(--sg-radius-sm) 0;color:var(--muted)}
.notice strong{color:var(--foreground)}
.compass{aspect-ratio:1;max-height:560px;width:100%;display:block}
.ring{fill:none;stroke:var(--border);stroke-width:1.5}.axis{stroke:var(--border);stroke-width:1}
.axislabel{fill:var(--muted);font:700 14px var(--sg-font-family-mono)}
.distance{fill:var(--muted);font:11px var(--sg-font-family-mono)}
.point{fill:var(--accent);stroke:var(--accent-ink);stroke-width:1.5}
.centerpoint{fill:var(--primary);stroke:var(--primary-ink);stroke-width:2}
.resultlist{display:grid;gap:8px;max-height:560px;overflow:auto;padding-right:3px}
.result{border:1px solid var(--border);border-radius:var(--sg-radius-lg);padding:11px;background:var(--surface)}
.resulthead{display:flex;justify-content:space-between;gap:12px;align-items:start}
.result h3{font-size:14px;overflow-wrap:anywhere}.distancepill,.pill{white-space:nowrap;border-radius:var(--sg-radius-pill);padding:4px 8px;background:var(--surface-2);font:700 10px var(--sg-font-family-mono)}
.meta{display:flex;flex-wrap:wrap;gap:5px 10px;margin-top:7px;color:var(--muted);font-size:12px}
.directory-tools{display:grid;grid-template-columns:2fr repeat(3,1fr);gap:8px;margin-bottom:12px}
.tablewrap{overflow:auto;border:1px solid var(--border);border-radius:var(--sg-radius-lg)}
table{width:100%;border-collapse:collapse;min-width:760px;background:var(--surface)}
th,td{text-align:left;padding:10px 11px;border-bottom:1px solid var(--border);vertical-align:top}
th{position:sticky;top:0;background:var(--surface-2);font:700 10px var(--sg-font-family-mono);letter-spacing:.08em;text-transform:uppercase}
tbody tr:last-child td{border-bottom:0}
td.num{text-align:right;font-family:var(--sg-font-family-mono)}
.tablefoot{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-top:10px;color:var(--muted);font-size:12px}
.pager{display:flex;gap:7px}.pager button{min-width:44px;min-height:44px}
.metricgrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
.metric{background:var(--surface);border:1px solid var(--border);border-radius:var(--sg-radius-lg);padding:14px}
.metric strong{display:block;font-family:var(--sg-font-family-display);font-size:24px}.metric span{color:var(--muted);font-size:11px}
.barlist{display:grid;gap:10px;margin-top:12px}
.barrow{display:grid;grid-template-columns:90px 1fr 70px;gap:9px;align-items:center}
.track{height:12px;border-radius:var(--sg-radius-pill);background:var(--surface-2);overflow:hidden}.fill{height:100%;background:var(--accent);border-radius:inherit}
.source{display:grid;gap:10px;margin-top:12px}.source div{padding:12px;border:1px solid var(--border);border-radius:var(--sg-radius-lg)}
.source code{display:block;margin-top:5px;color:var(--muted);overflow-wrap:anywhere;font-family:var(--sg-font-family-mono);font-size:11px}
.empty{padding:24px;text-align:center;color:var(--muted)}
.error{border-color:var(--danger);color:var(--danger)}
footer{border-top:1px solid var(--border);padding:18px 0 max(18px,env(safe-area-inset-bottom));color:var(--muted);font-size:11px}
@media(max-width:820px){
  .shell{padding-left:14px;padding-right:14px}.summary{grid-template-columns:repeat(2,1fr)}
  .grid2{grid-template-columns:1fr}.resultlist{max-height:none}
  .metricgrid{grid-template-columns:repeat(2,1fr)}.directory-tools{grid-template-columns:repeat(2,1fr)}
  .directory-tools input{grid-column:1/-1}
}
@media(max-width:520px){
  .shell{padding-left:12px;padding-right:12px}
  header{padding-top:max(12px,env(safe-area-inset-top));padding-bottom:11px}
  .headerrow{display:grid;grid-template-columns:1fr auto;gap:10px}.brand{font-size:20px}.kicker{font-size:9px}
  .sourcebadge{grid-column:1/-1;margin-top:2px}.summary{gap:6px;margin-top:10px}
  .summary div{padding-left:7px}.summary strong{font-size:15px}.summary span{font-size:10px}
  nav button{min-width:130px;min-height:46px;font-size:12px}
  main{padding-top:16px}.hero{display:block}.lede{font-size:13px}
  .card{padding:13px;border-radius:var(--sg-radius-lg)}
  .controls{grid-template-columns:1fr 112px}.controls .btn{grid-column:1/-1}
  .directory-tools{grid-template-columns:1fr 1fr}.directory-tools select:last-child{grid-column:1/-1}
  .metricgrid{grid-template-columns:1fr 1fr}.metric strong{font-size:20px}
  .barrow{grid-template-columns:72px 1fr 54px;font-size:11px}
  .resulthead{display:block}.distancepill{display:inline-block;margin-top:5px}
  .tablefoot{align-items:flex-start;flex-direction:column}
}
@media(max-width:360px){
  .summary,.metricgrid,.directory-tools{grid-template-columns:1fr}
  .directory-tools input,.directory-tools select:last-child{grid-column:auto}
  .controls{grid-template-columns:1fr}.controls .btn{grid-column:auto}
}
@media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;transition:none!important}}
</style>
<script src="union-atlas-data.js"></script>
</head>
<body>
<header>
  <div class="shell">
    <div class="headerrow">
      <div><div class="brand">Common <b>Ground</b></div><div class="kicker">Source-backed United States Union Atlas</div></div>
      <button class="theme" id="theme" type="button">Dark theme</button>
      <div class="sourcebadge"><i></i> Public OLMS filing records + Census geography</div>
    </div>
    <div class="summary" id="summary" aria-label="Dataset summary"></div>
  </div>
</header>
<nav aria-label="Atlas views"><div class="shell">
  <button type="button" role="tab" data-view="near" aria-selected="true">Near you</button>
  <button type="button" role="tab" data-view="directory" aria-selected="false">Directory</button>
  <button type="button" role="tab" data-view="affiliations" aria-selected="false">Affiliations</button>
  <button type="button" role="tab" data-view="status" aria-selected="false">Data status</button>
</div></nav>
<main><div class="shell">
  <section class="view on" id="view-near"></section>
  <section class="view" id="view-directory"></section>
  <section class="view" id="view-affiliations"></section>
  <section class="view" id="view-status"></section>
</div></main>
<footer><div class="shell" id="footer"></div></footer>
<script>
(function(){
"use strict";
var DATA=window.UNION_ATLAS_DATA;
var state={view:"near",zip:"01103",radius:100,query:"",stateFilter:"",affiliation:"",level:"",page:0};
var pageSize=100;
var money=new Intl.NumberFormat("en-US",{notation:"compact",maximumFractionDigits:1});
var whole=new Intl.NumberFormat("en-US");
var esc=function(value){return String(value==null?"":value).replace(/[&<>"']/g,function(c){return({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"})[c]})};
var fmt=function(value){return Number.isFinite(value)?whole.format(value):"Not reported"};
var compact=function(value){return Number.isFinite(value)?money.format(value):"—"};
var pct=function(value,total){return total?Math.round(value/total*100):0};
var el=function(selector){return document.querySelector(selector)};
var all=function(selector){return Array.from(document.querySelectorAll(selector))};

function fail(message){
  document.querySelector("main .shell").innerHTML='<div class="card error"><h1>Atlas data could not load</h1><p>'+esc(message)+'</p></div>';
}
if(!DATA||!Array.isArray(DATA.directory)){fail("The source-backed data file is missing or invalid.");return}

function setTheme(theme){
  document.documentElement.dataset.theme=theme;
  el("#theme").textContent=theme==="dark"?"Light theme":"Dark theme";
  try{localStorage.setItem("union-atlas-theme",theme)}catch(_){}
}
var savedTheme=null;try{savedTheme=localStorage.getItem("union-atlas-theme")}catch(_){}
setTheme(savedTheme||"light");
el("#theme").addEventListener("click",function(){setTheme(document.documentElement.dataset.theme==="dark"?"light":"dark")});

function summary(){
  var q=DATA.quality;
  el("#summary").innerHTML=[
    [fmt(q.records),"OLMS union records"],
    [fmt(q.withZctaCoordinates),"Census-geocoded HQs"],
    [fmt(q.withMembership),"membership filings"],
    [fmt(q.withWebsites),"published websites"]
  ].map(function(x){return "<div><strong>"+x[0]+"</strong><span>"+x[1]+"</span></div>"}).join("");
  var src=DATA.metadata.sources[0];
  el("#footer").innerHTML="Directory grain: "+esc(DATA.metadata.recordGrain)+
    " · Source commit <b>"+esc(src.commit.slice(0,12))+"</b> · "+esc(DATA.metadata.privacy);
}

function switchView(view){
  state.view=view;
  all(".view").forEach(function(node){node.classList.toggle("on",node.id==="view-"+view)});
  all("nav button").forEach(function(button){
    var active=button.dataset.view===view;
    button.setAttribute("aria-selected",String(active));
    if(active)button.scrollIntoView({block:"nearest",inline:"center"});
  });
  ({near:renderNear,directory:renderDirectory,affiliations:renderAffiliations,status:renderStatus})[view]();
}
all("nav button").forEach(function(button){button.addEventListener("click",function(){switchView(button.dataset.view)})});

function radians(value){return value*Math.PI/180}
function distance(aLat,aLon,bLat,bLon){
  var earth=3958.7613,dLat=radians(bLat-aLat),dLon=radians(bLon-aLon);
  var a=Math.sin(dLat/2)**2+Math.cos(radians(aLat))*Math.cos(radians(bLat))*Math.sin(dLon/2)**2;
  return earth*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function bearing(aLat,aLon,bLat,bLon){
  var y=Math.sin(radians(bLon-aLon))*Math.cos(radians(bLat));
  var x=Math.cos(radians(aLat))*Math.sin(radians(bLat))-Math.sin(radians(aLat))*Math.cos(radians(bLat))*Math.cos(radians(bLon-aLon));
  return (Math.atan2(y,x)*180/Math.PI+360)%360;
}
function compassSvg(rows,radius){
  var size=560,c=280,max=220;
  var rings=[.25,.5,.75,1].map(function(scale){
    return '<circle class="ring" cx="'+c+'" cy="'+c+'" r="'+(max*scale)+'"/><text class="distance" x="'+(c+5)+'" y="'+(c-max*scale+14)+'">'+Math.round(radius*scale)+' mi</text>'
  }).join("");
  var points=rows.slice(0,24).map(function(item,index){
    var angle=radians(item.bearing-90),r=Math.min(max,item.distance/radius*max);
    var x=c+Math.cos(angle)*r,y=c+Math.sin(angle)*r;
    return '<circle class="point" tabindex="0" role="img" aria-label="'+esc(item.row.name)+", "+item.distance.toFixed(1)+' miles" cx="'+x.toFixed(1)+'" cy="'+y.toFixed(1)+'" r="'+(index<8?7:5)+'"><title>'+esc(item.row.name)+" · "+item.distance.toFixed(1)+' mi</title></circle>'
  }).join("");
  return '<svg class="compass" viewBox="0 0 560 560" role="img" aria-labelledby="compass-title">'+
    '<title id="compass-title">Nearest source-backed union headquarters by Census ZIP centroid</title>'+
    rings+'<line class="axis" x1="280" y1="45" x2="280" y2="515"/><line class="axis" x1="45" y1="280" x2="515" y2="280"/>'+
    '<text class="axislabel" x="274" y="27">N</text><text class="axislabel" x="274" y="548">S</text><text class="axislabel" x="20" y="285">W</text><text class="axislabel" x="525" y="285">E</text>'+
    '<circle class="centerpoint" cx="280" cy="280" r="9"/>'+points+"</svg>";
}
function unionCard(item){
  var row=item.row;
  return '<article class="result"><div class="resulthead"><h3>'+esc(row.name)+'</h3><span class="distancepill">'+item.distance.toFixed(1)+' mi · '+Math.round(item.bearing)+'°</span></div>'+
    '<div class="meta"><span>'+esc([row.city,row.state,row.zip].filter(Boolean).join(", "))+'</span>'+
    '<span>'+esc(row.affiliation||"Independent")+'</span><span>'+fmt(row.members)+' reported members'+(row.membershipYear?" ("+row.membershipYear+")":"")+'</span></div></article>';
}
function renderNear(){
  var root=el("#view-near");
  var anchor=DATA.zctas[state.zip];
  var nearby=[];
  if(anchor){
    nearby=DATA.directory.filter(function(row){return Number.isFinite(row.latitude)&&Number.isFinite(row.longitude)})
      .map(function(row){return{row:row,distance:distance(anchor[0],anchor[1],row.latitude,row.longitude),bearing:bearing(anchor[0],anchor[1],row.latitude,row.longitude)}})
      .filter(function(item){return item.distance<=state.radius})
      .sort(function(a,b){return a.distance-b.distance||a.row.name.localeCompare(b.row.name)});
  }
  root.innerHTML='<div class="hero"><div><div class="eyebrow">Geographic discovery</div><h1>Nearest union headquarters</h1>'+
    '<p class="lede">Distances and compass bearings use public OLMS headquarters ZIP codes joined to official Census ZCTA representative coordinates.</p></div></div>'+
    '<div class="controls"><input id="zip" inputmode="numeric" maxlength="5" aria-label="Five digit ZIP code" value="'+esc(state.zip)+'" placeholder="ZIP code">'+
    '<select id="radius" aria-label="Search radius">'+[25,50,100,250,500].map(function(n){return'<option '+(n===state.radius?"selected":"")+' value="'+n+'">'+n+' miles</option>'}).join("")+'</select>'+
    '<button class="btn" id="locate" type="button">Plot real records</button></div>'+
    (!anchor?'<div class="notice"><strong>ZIP not found.</strong> Enter a five-digit Census ZCTA.</div>':
    '<div class="grid2"><div class="card">'+compassSvg(nearby,state.radius)+'</div><div class="stack"><div class="notice"><strong>'+fmt(nearby.length)+' records</strong> within '+state.radius+' miles of ZIP '+esc(state.zip)+'. Repeated headquarters are retained because each OLMS filing is a distinct union record.</div>'+
    '<div class="resultlist">'+(nearby.length?nearby.slice(0,50).map(unionCard).join(""):'<div class="card empty">No geocoded union headquarters in this radius.</div>')+'</div></div></div>');
  el("#locate").addEventListener("click",function(){state.zip=el("#zip").value.trim();state.radius=Number(el("#radius").value);renderNear()});
  el("#zip").addEventListener("keydown",function(event){if(event.key==="Enter")el("#locate").click()});
}

function filteredDirectory(){
  var query=state.query.trim().toLowerCase();
  return DATA.directory.filter(function(row){
    var text=[row.name,row.affiliation,row.localNumber,row.city,row.state,row.zip].join(" ").toLowerCase();
    return(!query||text.includes(query))&&(!state.stateFilter||row.state===state.stateFilter)&&
      (!state.affiliation||row.affiliation===state.affiliation)&&(!state.level||row.level===state.level);
  });
}
function renderDirectory(){
  var rows=filteredDirectory(),pages=Math.max(1,Math.ceil(rows.length/pageSize));
  state.page=Math.min(state.page,pages-1);
  var slice=rows.slice(state.page*pageSize,(state.page+1)*pageSize);
  var states=DATA.states.map(function(row){return row.key}).sort();
  var affiliations=DATA.affiliations.map(function(row){return row.key}).sort();
  el("#view-directory").innerHTML='<div class="hero"><div><div class="eyebrow">Canonical records</div><h1>US union directory</h1><p class="lede">One row per canonical OLMS union record. No platform groups, campaigns, or member records.</p></div></div>'+
    '<div class="card"><div class="directory-tools"><input id="query" type="search" value="'+esc(state.query)+'" placeholder="Name, local number, city, state, ZIP…">'+
    '<select id="state-filter"><option value="">All states</option>'+states.map(function(x){return'<option '+(x===state.stateFilter?"selected":"")+'>'+esc(x)+'</option>'}).join("")+'</select>'+
    '<select id="aff-filter"><option value="">All affiliations</option>'+affiliations.map(function(x){return'<option '+(x===state.affiliation?"selected":"")+'>'+esc(x)+'</option>'}).join("")+'</select>'+
    '<select id="level-filter"><option value="">All levels</option>'+["national","intermediate","local","independent","unknown"].map(function(x){return'<option '+(x===state.level?"selected":"")+'>'+x+'</option>'}).join("")+'</select></div>'+
    '<div class="tablewrap"><table><thead><tr><th>Union</th><th>Affiliation</th><th>HQ</th><th>Level</th><th>Members</th><th>Filing</th></tr></thead><tbody>'+
    (slice.length?slice.map(function(row){return'<tr><td><strong>'+esc(row.name)+'</strong><br><span class="meta">OLMS '+row.olmsFileNumber+(row.localNumber?" · #"+esc(row.localNumber):"")+'</span></td>'+
      '<td>'+esc(row.affiliation||"Independent")+'</td><td>'+esc([row.city,row.state,row.zip].filter(Boolean).join(", "))+'</td><td>'+esc(row.level)+'</td>'+
      '<td class="num">'+fmt(row.members)+'</td><td>'+esc([row.filingForm,row.membershipYear].filter(Boolean).join(" · "))+'</td></tr>'}).join(""):'<tr><td colspan="6" class="empty">No source records match these filters.</td></tr>')+
    '</tbody></table></div><div class="tablefoot"><span>Showing '+fmt(slice.length)+' of '+fmt(rows.length)+' matching records.</span><div class="pager"><button class="btn secondary" id="prev" '+(state.page===0?"disabled":"")+'>←</button><span class="pill">Page '+(state.page+1)+' / '+pages+'</span><button class="btn secondary" id="next" '+(state.page>=pages-1?"disabled":"")+'>→</button></div></div></div>';
  el("#query").addEventListener("input",function(event){state.query=event.target.value;state.page=0;renderDirectory();var q=el("#query");q.focus();q.setSelectionRange(q.value.length,q.value.length)});
  el("#state-filter").addEventListener("change",function(event){state.stateFilter=event.target.value;state.page=0;renderDirectory()});
  el("#aff-filter").addEventListener("change",function(event){state.affiliation=event.target.value;state.page=0;renderDirectory()});
  el("#level-filter").addEventListener("change",function(event){state.level=event.target.value;state.page=0;renderDirectory()});
  el("#prev").addEventListener("click",function(){state.page=Math.max(0,state.page-1);renderDirectory()});
  el("#next").addEventListener("click",function(){state.page=Math.min(pages-1,state.page+1);renderDirectory()});
}

function renderAffiliations(){
  var rows=DATA.affiliations.slice(0,100),max=rows[0]?.unions||1;
  el("#view-affiliations").innerHTML='<div class="hero"><div><div class="eyebrow">Filed parent codes</div><h1>Affiliation coverage</h1><p class="lede">Counts are derived from each canonical record’s filed affiliation abbreviation. Reported membership is not de-duplicated across hierarchy levels.</p></div></div>'+
    '<div class="grid2"><div class="card"><h2>Largest affiliations by filing count</h2><div class="barlist">'+rows.slice(0,25).map(function(row){return'<div class="barrow"><strong>'+esc(row.key)+'</strong><div class="track"><div class="fill" style="width:'+pct(row.unions,max)+'%"></div></div><span class="num">'+fmt(row.unions)+'</span></div>'}).join("")+
    '</div></div><div class="card"><h2>Source rollup</h2><div class="tablewrap" style="margin-top:12px"><table><thead><tr><th>Affiliation</th><th>Unions</th><th>Reported members</th><th>Websites</th></tr></thead><tbody>'+
    rows.map(function(row){return'<tr><td><strong>'+esc(row.key)+'</strong></td><td class="num">'+fmt(row.unions)+'</td><td class="num">'+fmt(row.reportedMembers)+'</td><td class="num">'+fmt(row.websites)+'</td></tr>'}).join("")+
    '</tbody></table></div></div></div>';
}

function renderStatus(){
  var q=DATA.quality,src=DATA.metadata.sources;
  var statusRows=Object.entries(q.enrichmentStatuses);
  el("#view-status").innerHTML='<div class="hero"><div><div class="eyebrow">Provenance and completeness</div><h1>Data status</h1><p class="lede">This view reports actual row coverage. Missing values remain missing; the Atlas does not manufacture replacements.</p></div></div>'+
    '<div class="metricgrid"><div class="metric"><strong>'+fmt(q.records)+'</strong><span>canonical records</span></div><div class="metric"><strong>'+fmt(q.uniqueIds)+'</strong><span>unique stable IDs</span></div>'+
    '<div class="metric"><strong>'+pct(q.withMembership,q.records)+'%</strong><span>with membership filing</span></div><div class="metric"><strong>'+pct(q.withZctaCoordinates,q.records)+'%</strong><span>with Census ZIP coordinates</span></div></div>'+
    '<div class="grid2" style="margin-top:16px"><div class="card"><h2>Coverage checks</h2><div class="barlist">'+
    [["Membership",q.withMembership],["Website",q.withWebsites],["Coordinates",q.withZctaCoordinates]].map(function(row){return'<div class="barrow"><strong>'+row[0]+'</strong><div class="track"><div class="fill" style="width:'+pct(row[1],q.records)+'%"></div></div><span>'+pct(row[1],q.records)+'%</span></div>'}).join("")+
    '</div><h3 style="margin-top:18px">Enrichment state</h3><div class="meta">'+statusRows.map(function(row){return'<span class="pill">'+esc(row[0])+': '+fmt(row[1])+'</span>'}).join("")+'</div></div>'+
    '<div class="card"><h2>Traceable sources</h2><div class="source">'+src.map(function(item){return'<div><strong>'+esc(item.name)+'</strong><code>'+(item.repository?"Repository: "+esc(item.repository)+" · commit "+esc(item.commit):esc(item.url))+'</code><code>'+(item.path?"Path: "+esc(item.path):"")+'</code></div>'}).join("")+
    '</div><div class="notice" style="margin-top:12px"><strong>Privacy boundary.</strong> '+esc(DATA.metadata.privacy)+'</div></div></div>';
}

summary();switchView("near");
})();
</script>
</body>
</html>`;

fs.writeFileSync(outputPath, page, 'utf8');
console.log(`Built ${outputPath}`);
