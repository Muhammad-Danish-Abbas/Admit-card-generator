(function(){
'use strict';
var $=function(s){return document.querySelector(s)};
var LS={get:function(k,d){try{var v=localStorage.getItem(k);return v?JSON.parse(v):d}catch(e){return d}},set:function(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){}},del:function(k){try{localStorage.removeItem(k)}catch(e){}}};
var esc=function(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]})};

/* ---------- settings ---------- */
var DEFAULT_INS=[
 'Candidates are advised to carry their admit cards and stationery items in a transparent stationery case/pouch.',
 '**No candidate will be allowed** to enter the campus after **{time}**',
 'Logbooks, periodic tables, hardboard, formula sheets, and stencils are **NOT** allowed in the examination room.',
 'All types of smartwatches or any electronic gadgets are strictly **PROHIBITED** inside the examination room',
 'Refusal or failure to follow instructions, disruptive behavior, cheating, or **use of any unfair means** may lead to your **disqualification** from the examination session.'
].join('\n');
var DEFAULTS={school:'AGA KHAN HIGHER SECONDARY SCHOOL, KARACHI',exam:'MIDTERM EXAMINATION-2026',morning:'7:40 a.m',evening:'2:00 p.m',signTitle:'Vice Principal',sigMorning:'',sigEvening:'',instructions:DEFAULT_INS};
var settings=Object.assign({},DEFAULTS,LS.get('acg.settings',{}));
function saveSettings(){LS.set('acg.settings',settings)}

/* ---------- data ---------- */
var SAMPLE=[
 {id:'2724981',name:'ABDUL WASAY',father:'ASIF NIZAM',cls:'C1',grade:'11',shift:'Morning'},
 {id:'2725063',name:'AHMED SHAHZAIB',father:'SHAHZAIB',cls:'C1',grade:'11',shift:'Morning'},
 {id:'2723101',name:'HASSAN RAZA',father:'RAZA ALI',cls:'M1',grade:'11',shift:'Morning'},
 {id:'2724009',name:'AABAN PASHA SHAIKH',father:'SHAHID PASHA SHAIKH',cls:'EE6B',grade:'12',shift:'Evening'},
 {id:'2724010',name:'SARA KHAN',father:'IMRAN KHAN',cls:'EE6B',grade:'12',shift:'Evening'},
 {id:'2723544',name:'ZARA SHEIKH',father:'FAISAL SHEIKH',cls:'E2',grade:'11',shift:'Evening'}
];
var students=LS.get('acg.students',null), isSample=false;
if(!students||!students.length){students=SAMPLE.slice();isSample=true}
var photos=new Map();
var state={cls:'',q:'',sort:'name',shown:60,tab:'C',modalStudent:null};

/* ---------- photo storage (IndexedDB) ---------- */
function reqP(r){return new Promise(function(res,rej){r.onsuccess=function(){res(r.result)};r.onerror=function(){rej(r.error)}})}
function openDB(){return new Promise(function(res,rej){var r=indexedDB.open('acg-photos',1);r.onupgradeneeded=function(){r.result.createObjectStore('p')};r.onsuccess=function(){res(r.result)};r.onerror=function(){rej(r.error)}})}
async function dbSave(entries){try{var db=await openDB();var tx=db.transaction('p','readwrite');var st=tx.objectStore('p');entries.forEach(function(e){st.put(e[1],e[0])});await new Promise(function(res){tx.oncomplete=res;tx.onerror=res;tx.onabort=res})}catch(e){}}
async function dbLoad(){try{var db=await openDB();var st=db.transaction('p').objectStore('p');var keys=await reqP(st.getAllKeys());var vals=await reqP(st.getAll());keys.forEach(function(k,i){photos.set(String(k),URL.createObjectURL(vals[i]))})}catch(e){}}
async function dbClear(){try{var db=await openDB();var tx=db.transaction('p','readwrite');tx.objectStore('p').clear();await new Promise(function(res){tx.oncomplete=res;tx.onerror=res})}catch(e){}}

function photoFor(id){id=String(id||'').trim();return photos.get(id)||photos.get(id.replace(/\D/g,''))||''}

/* ---------- known classes (buttons always shown) ---------- */
var KNOWN_SIS=['11-C1-HSA','11-E3A-HSA','11-EE1-HSA','11-EE2-HSA','11-EM1-HSA','11-EM2-HSA','11-EM3-HSA','11-EM4-HSA','11-EM5-HSA','11-M1-HSA','11-M2-HSA','11-M3-HSA','11-M4-HSA','11-E3B-HSA','11-M5-HSA','11-E1A-HSA','11-E1B-HSA','11-E2A-HSA','11-E2B-HSA','11-EC1-CA-HSA','11-EC1-HB-HSA','11-PM-HSA','11-PE/SG-HSA','11-COM/HUM-HSA','11-EE3-HSA','12-C2-HSA','12-E4A-HSA','12-E6A-HSA','12-E6B-HSA','12-M6-HSA','12-M7-HSA','12-M8-HSA','12-EE4-HSA','12-EE5A-HSA','12-EE6A-HSA','12-EM6-HSA','12-EM7-HSA','12-EM8-HSA','12-EM9-HSA','12-EM10-HSA','12-M9-HSA','12-M10-HSA','12-E4B-HSA','12-E5B-HSA','12-E5A-HSA','12-EC2A-HSA','12-EC2B-HSA','12-EE5B-HSA','12-EE6B-HSA'];

/* ---------- Excel ---------- */
var COLS={id:['studentid','id','rollno','rollnumber'],name:['studentname','name'],father:['fathername','fathersname','father'],cls:['class','classname','classcode'],shift:['shift'],grade:['class1','grade','year'],sis:['sisfeeclass','sisclass','feeclass']};
function norm(h){return String(h).toLowerCase().replace(/[^a-z0-9]/g,'')}
function normShift(v,cls){var t=String(v||'').trim().toLowerCase();if(t.charAt(0)==='m')return 'Morning';if(t.charAt(0)==='e')return 'Evening';return /^E/.test(cls)?'Evening':'Morning'}
function cleanCls(v){return String(v||'').trim().toUpperCase().replace(/\s+/g,'').replace(/^\d{1,2}-/,'').replace(/-HSA$/,'')}
async function readExcel(files){
  var out=[],notes=[],used=0;
  for(var fi=0;fi<files.length;fi++){
    var f=files[fi];
    try{
      var wb=XLSX.read(await f.arrayBuffer(),{type:'array'});
      var got=false;
      wb.SheetNames.forEach(function(sn){
        var grid=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false});
        if(!grid.length)return;
        /* find the header row (it may not be the first row) */
        var hi=-1;
        for(var i=0;i<Math.min(grid.length,25);i++){
          var ns=grid[i].map(norm);
          if(ns.indexOf('studentid')>-1||(ns.indexOf('studentname')>-1)){hi=i;break}
        }
        if(hi<0){notes.push(f.name+' ('+sn+'): no "Student ID" / "Student Name" heading found. First row: '+grid[0].filter(String).slice(0,8).join(', '));return}
        var heads=grid[hi],map={};
        heads.forEach(function(h,ci){var n=norm(h);for(var k in COLS){if(!(k in map)&&COLS[k].indexOf(n)>-1)map[k]=ci}});
        if(map.id==null||map.name==null||(map.cls==null&&map.sis==null)){
          notes.push(f.name+' ('+sn+'): missing columns. Found: '+heads.filter(String).join(', '));return}
        for(var r=hi+1;r<grid.length;r++){
          var row=grid[r];
          var id=String(row[map.id]==null?'':row[map.id]).trim(),name=String(row[map.name]==null?'':row[map.name]).trim();
          if(!id||!name)continue;
          var cls=cleanCls(map.cls!=null?row[map.cls]:'')||cleanCls(map.sis!=null?row[map.sis]:'');
          if(!cls)continue;
          var gm=(map.grade!=null?String(row[map.grade]).match(/\d+/):null)||(map.sis!=null?String(row[map.sis]).match(/^\s*(\d{1,2})\s*-/):null);
          var grade=gm?(gm[1]||gm[0]):'';
          out.push({id:id,grade:grade,name:name.toUpperCase(),father:map.father!=null?String(row[map.father]).trim().toUpperCase():'',cls:cls,shift:normShift(map.shift!=null?row[map.shift]:'',cls)});
          got=true;
        }
      });
      if(got)used++;
    }catch(e){notes.push(f.name+': could not be read ('+(e&&e.message?e.message:e)+')')}
  }
  return {rows:out,notes:notes,used:used};
}
async function handleExcel(files){
  var st=$('#xlStatus');
  if(!files.length)return;
  if(!window.XLSX){st.className='status err';st.textContent='The Excel reader did not load. Check your internet and reload the page.';return}
  st.className='status';st.textContent='Reading files...';
  var r=await readExcel(files);
  if(!r.rows.length){st.className='status err';st.textContent=(r.notes.join(' | ')||'No students found.');return}
  var map=new Map();
  if(!isSample)students.forEach(function(s){map.set(s.id,s)});
  r.rows.forEach(function(s){map.set(s.id,s)});
  students=Array.from(map.values());isSample=false;
  LS.set('acg.students',students);
  state.cls='';
  st.className='status ok';
  st.textContent=r.rows.length+' students loaded from '+r.used+' file(s). Total: '+students.length+'.'+(r.notes.length?' Skipped: '+r.notes.join(' | '):'');
  render();
}

/* ---------- Photos ---------- */
async function handlePhotos(files){
  var st=$('#phStatus'),entries=[],skipped=0;
  files.forEach(function(f){
    if(!/^image\//.test(f.type)&&!/\.(jpe?g|png|webp|gif|bmp)$/i.test(f.name))return;
    var base=f.name.replace(/\.[^.]+$/,'').trim();
    var key=/^\d+$/.test(base)?base:((base.match(/\d{6,}/)||[])[0]);
    if(!key){skipped++;return}
    if(photos.has(key))URL.revokeObjectURL(photos.get(key));
    photos.set(key,URL.createObjectURL(f));
    entries.push([key,f]);
  });
  if(!entries.length){st.className='status err';st.textContent='No photo with a Student ID in its name was found.'+(skipped?' ('+skipped+' skipped)':'');return}
  st.className='status ok';st.textContent=entries.length+' photos added.'+(skipped?' '+skipped+' skipped because the file name has no Student ID.':'');
  await dbSave(entries);
  render();
}
function downloadMissing(){
  var miss=students.filter(function(s){return !photoFor(s.id)});
  var csv='Student ID,Student Name,Class,Shift\n'+miss.map(function(s){return [s.id,'"'+s.name.replace(/"/g,'""')+'"',s.cls,s.shift].join(',')}).join('\n');
  var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download='missing-photos.csv';document.body.appendChild(a);a.click();a.remove();
  $('#phStatus').className='status';$('#phStatus').textContent=miss.length+' students have no photo yet.';
}

/* ---------- Card ---------- */
var PLACEHOLDER='<svg viewBox="0 0 92 124" aria-hidden="true"><rect width="92" height="124" fill="#d5dde0"/><circle cx="46" cy="48" r="19" fill="#9fb0b6"/><path d="M8 124c2-30 18-44 38-44s36 14 38 44z" fill="#9fb0b6"/></svg>';
function fmtIns(line,time){return esc(line).replace(/\*\*(.+?)\*\*/g,'<b>$1</b>').replace(/\{time\}/g,esc(time))}
function cardHTML(s){
  var evening=s.shift==='Evening';
  var time=evening?settings.evening:settings.morning;
  var sig=evening?settings.sigEvening:settings.sigMorning;
  var ph=s.photo||photoFor(s.id);
  var lines=String(settings.instructions).split('\n').filter(function(l){return l.trim()});
  return '<div class="card">'+
   '<div class="c-head"><div>'+esc(settings.school)+'</div><div>'+esc(settings.exam)+'</div><div>ADMIT CARD</div></div>'+
   '<div class="c-photo">'+(ph?'<img src="'+esc(ph)+'" alt="">':PLACEHOLDER)+'</div>'+
   '<div class="c-grid">'+
     '<div><b>STUDENT NAME:</b><span>'+esc(s.name)+'</span></div>'+
     '<div><b>FATHER\'S NAME:</b><span>'+esc(s.father)+'</span></div>'+
     '<div><b>STUDENT ID:</b><span>'+esc(s.id)+'</span></div>'+
     '<div><b>CLASS:</b><span>'+esc(s.cls)+'</span></div>'+
   '</div>'+
   '<div class="c-ins"><span class="t">Instructions:</span><ol>'+lines.map(function(l){return '<li>'+fmtIns(l,time)+'</li>'}).join('')+'</ol></div>'+
   '<div class="c-sign">'+(sig?'<img src="'+esc(sig)+'" alt="">':'<div class="sp"></div>')+'<div>'+esc(settings.signTitle)+'</div><div>'+(evening?'Evening':'Morning')+' Shift</div></div>'+
  '</div>';
}

/* ---------- Print (browser): A5 landscape, 1 card per page ---------- */
function printCards(list){
  if(!list.length)return;
  $('#printArea').innerHTML=list.map(function(s){return '<div class="sheet"><div class="cw">'+cardHTML(s)+'</div></div>'}).join('');
  setTimeout(function(){window.print()},250);
}

/* ---------- PDF: A5 landscape, 1 card per page, drawn as vectors (small files, fast) ---------- */
var PW=210,PH=148,CWm=194,CHm=CWm*452/740,CXm=(PW-CWm)/2,CYm=(PH-CHm)/2,U=CWm/100; /* U = 1 percent of card width in mm */
function pt(c){return c*U*72/25.4}
var layoutCache={key:null,val:null};
function richWords(line,time){
  var segs=line.replace(/\{time\}/g,time).split('**'),words=[],space=false;
  segs.forEach(function(seg,i){
    seg.split(/(\s+)/).forEach(function(p){
      if(!p)return;
      if(/^\s+$/.test(p)){space=true;return}
      words.push({txt:p,b:i%2===1,sp:space&&words.length>0});space=false;
    });
  });
  return words;
}
function insLayout(doc,time){
  var key=settings.instructions+'|'+time;
  if(layoutCache.key===key)return layoutCache.val;
  var maxW=93.1*U,lh=2.77*U,out=[],cursor=26.6*U;
  doc.setFontSize(pt(2.05));
  String(settings.instructions).split('\n').filter(function(l){return l.trim()}).forEach(function(line,idx){
    doc.setFont('times','normal');var spw=doc.getTextWidth(' ');
    var rows=[],cur={w:0,runs:[]};
    richWords(line,time).forEach(function(w){
      doc.setFont('times',w.b?'bold':'normal');
      var ww=doc.getTextWidth(w.txt),gap=(cur.runs.length&&w.sp)?spw:0;
      if(cur.runs.length&&cur.w+gap+ww>maxW){rows.push(cur);cur={w:0,runs:[]};gap=0}
      var last=cur.runs[cur.runs.length-1];
      if(last&&last.b===w.b)last.txt+=(gap?' ':'')+w.txt;
      else cur.runs.push({x:cur.w+gap,txt:w.txt,b:w.b});
      cur.w+=gap+ww;
    });
    rows.push(cur);
    cursor+=.95*U;
    out.push({num:(idx+1)+'.',y:cursor+2.19*U,rows:rows,lh:lh});
    cursor+=rows.length*lh;
  });
  layoutCache={key:key,val:out};
  return out;
}
function drawCard(doc,s,ph,sig){
  var x0=CXm,y0=CYm,evening=s.shift==='Evening',time=evening?settings.evening:settings.morning;
  doc.setTextColor(0);doc.setDrawColor(0);doc.setLineWidth(.55*U);
  doc.rect(x0+.275*U,y0+.275*U,CWm-.55*U,CHm-.55*U,'S');
  /* heading */
  doc.setFont('times','bold');
  var hy=[4.68,7.98,11.28],maxHW=74.9*U;
  [settings.school,settings.exam,'ADMIT CARD'].forEach(function(t,i){
    var f=pt(2.75);doc.setFontSize(f);var w=doc.getTextWidth(t);
    if(w>maxHW){doc.setFontSize(f*maxHW/w)}
    doc.text(t,x0+48*U,y0+hy[i]*U,{align:'center'});
  });
  /* photo */
  var px=x0+85.85*U,py=y0+1.35*U,pw=9.2*U,phh=12.4*U;
  if(ph){doc.addImage(ph,'JPEG',px,py,pw,phh,undefined,'FAST')}
  else{
    doc.setFillColor(213,221,224);doc.rect(px,py,pw,phh,'F');
    doc.setFillColor(159,176,182);doc.circle(px+pw/2,py+phh*.39,pw*.2,'F');
    doc.lines([[pw*.16,-phh*.2],[pw*.5,0],[pw*.16,phh*.2]],px+pw*.09,py+phh,[1,1],'F',true);
  }
  /* name / father / id / class */
  function pair(label,val,lx,y,limit){
    doc.setFont('times','bold');doc.setFontSize(pt(2.1));doc.setTextColor(0);
    doc.text(label,lx,y);
    var vx=lx+doc.getTextWidth(label)+.9*U,f=pt(2.2);
    doc.setFont('times','normal');doc.setFontSize(f);doc.setTextColor(64);
    var w=doc.getTextWidth(val),avail=limit-vx;
    if(w>avail)doc.setFontSize(Math.max(f*.6,f*avail/w));
    doc.text(val,vx,y);
  }
  pair('STUDENT NAME:',s.name,x0+1.55*U,y0+17.4*U,x0+46.4*U);
  pair("FATHER'S NAME:",s.father,x0+49.05*U,y0+17.4*U,x0+97.6*U);
  pair('STUDENT ID:',s.id,x0+1.55*U,y0+21.4*U,x0+46.4*U);
  pair('CLASS:',s.cls,x0+49.05*U,y0+21.4*U,x0+97.6*U);
  /* instructions */
  doc.setTextColor(0);doc.setFont('times','bold');doc.setFontSize(pt(2.1));
  doc.text('Instructions:',x0+1.75*U,y0+25.9*U);
  var L=insLayout(doc,time);
  doc.setFontSize(pt(2.05));
  L.forEach(function(it){
    doc.setFont('times','normal');doc.text(it.num,x0+2.9*U,y0+it.y);
    it.rows.forEach(function(r,ri){
      r.runs.forEach(function(run){doc.setFont('times',run.b?'bold':'normal');doc.text(run.txt,x0+5.15*U+run.x,y0+it.y+ri*it.lh)});
    });
  });
  /* signature */
  if(sig){
    var hh=5.4*U,ww=hh*sig.ratio;if(ww>16*U){ww=16*U;hh=ww/sig.ratio}
    doc.addImage(sig.data,'PNG',x0+1.75*U,y0+54.3*U-hh,ww,hh,undefined,'FAST');
  }
  doc.setFont('times','bold');doc.setFontSize(pt(2.1));doc.setTextColor(0);
  doc.text(settings.signTitle,x0+1.75*U,y0+56.3*U);
  doc.text((evening?'Evening':'Morning')+' Shift',x0+1.75*U,y0+58.8*U);
}
function loadImg(url){return new Promise(function(res){var i=new Image();i.onload=function(){res(i)};i.onerror=function(){res(null)};i.src=url})}
async function photoJpeg(url){
  var im=await loadImg(url);if(!im)return null;
  var tw=240,th=323,c=document.createElement('canvas');c.width=tw;c.height=th;
  var g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,tw,th);
  var sr=im.width/im.height,tr=tw/th,sx=0,sy=0,sw=im.width,sh=im.height;
  if(sr>tr){sw=im.height*tr;sx=(im.width-sw)/2}else{sh=im.width/tr;sy=(im.height-sh)/2}
  g.drawImage(im,sx,sy,sw,sh,0,0,tw,th);
  return c.toDataURL('image/jpeg',.85);
}
async function sigInfo(url){if(!url)return null;var im=await loadImg(url);return im?{data:url,ratio:im.width/im.height}:null}

var job={cancel:false,done:0,total:0};
function tick(){return new Promise(function(r){setTimeout(r,0)})}
function showProg(title,total){
  job={cancel:false,done:0,total:total};
  $('#progTitle').textContent=title;$('#progBar').style.width='0';$('#progText').textContent='Starting...';
  $('#prog').classList.add('open');
}
function updProg(){
  var p=job.total?Math.round(job.done*100/job.total):0;
  $('#progBar').style.width=p+'%';$('#progText').textContent=job.done+' of '+job.total+' cards ('+p+'%)';
}
function hideProg(){$('#prog').classList.remove('open')}
$('#progCancel').addEventListener('click',function(){job.cancel=true;$('#progText').textContent='Cancelling...'});
function saveBlob(blob,name){
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();a.remove();
  setTimeout(function(){URL.revokeObjectURL(a.href)},60000);
}
function safeName(n){return String(n).replace(/[\\/:*?"<>|]+/g,'-')}
async function makePdf(list){
  var doc=new window.jspdf.jsPDF({orientation:'landscape',unit:'mm',format:'a5',compress:true});
  var sigs={Morning:await sigInfo(settings.sigMorning),Evening:await sigInfo(settings.sigEvening)};
  for(var i=0;i<list.length;i++){
    if(job.cancel)throw new Error('cancelled');
    var s=list[i];if(i>0)doc.addPage('a5','landscape');
    var url=s.photo||photoFor(s.id),ph=url?await photoJpeg(url):null;
    drawCard(doc,s,ph,sigs[s.shift==='Evening'?'Evening':'Morning']);
    job.done++;
    if(i%6===5){updProg();await tick()}
  }
  updProg();
  return doc.output('blob');
}
function exportMsg(t,cls){var p=$('#exportStatus');p.textContent=t;p.className=cls||''}
async function runExport(title,total,fn,okMsg){
  if(!window.jspdf||!window.JSZip){exportMsg('The PDF library did not load. Check your internet and reload the page.','err');return}
  exportMsg('');showProg(title,total);
  try{await fn();hideProg();exportMsg(okMsg,'ok')}
  catch(e){hideProg();exportMsg(e&&e.message==='cancelled'?'Cancelled.':'Could not make the PDF: '+(e&&e.message?e.message:e),'err')}
}
function sortedAll(){return students.slice().sort(function(a,b){return natural(a.grade||'',b.grade||'')||natural(a.cls,b.cls)||a.name.localeCompare(b.name)})}
function noPhotoNote(list){var n=list.filter(function(s){return !(s.photo||photoFor(s.id))}).length;return n?' ('+n+' cards have no photo.)':''}
function pdfFor(list,name){
  if(!list.length)return;
  runExport('Making '+name,list.length,async function(){saveBlob(await makePdf(list),name)},'Downloaded '+name+' with '+list.length+' cards.'+noPhotoNote(list));
}
function pdfAllClasses(){
  var all=sortedAll();if(!all.length)return;
  pdfFor(all,'All-Admit-Cards.pdf');
}
function zipAllClasses(){
  var all=sortedAll();if(!all.length)return;
  var groups={},order=[];
  all.forEach(function(s){var k=(s.grade||'')+'|'+s.cls;if(!groups[k]){groups[k]=[];order.push(k)}groups[k].push(s)});
  runExport('Making ZIP of class PDFs',all.length,async function(){
    var zip=new window.JSZip();
    for(var i=0;i<order.length;i++){
      var list=groups[order[i]],g=list[0].grade;
      zip.file((g?'Class '+g+'/':'')+safeName(list[0].cls)+'.pdf',await makePdf(list));
    }
    $('#progText').textContent='Packing ZIP...';await tick();
    saveBlob(await zip.generateAsync({type:'blob',compression:'STORE'}),'Admit-Cards-by-Class.zip');
  },'Downloaded Admit-Cards-by-Class.zip with '+order.length+' class PDFs ('+all.length+' cards).'+noPhotoNote(all));
}

/* ---------- Classes / navigation ---------- */
function natural(a,b){return a.localeCompare(b,undefined,{numeric:true})}
function classList(){
  var m={};
  students.forEach(function(s){var c=m[s.cls]||(m[s.cls]={cls:s.cls,grade:s.grade||'',count:0,m:0,e:0});c.count++;if(s.shift==='Evening')c.e++;else c.m++});
  KNOWN_SIS.forEach(function(x){
    var g=x.match(/^(\d{1,2})-/)[1],c=cleanCls(x);
    if(!m[c])m[c]={cls:c,grade:g,count:0,m:0,e:0,guessE:/^E/.test(c)};
  });
  return Object.keys(m).sort(natural).map(function(k){return m[k]});
}
function visible(){
  var q=state.q.trim().toLowerCase();
  var l=students.filter(function(s){
    if(q)return s.name.toLowerCase().indexOf(q)>-1||s.id.indexOf(q)>-1||s.father.toLowerCase().indexOf(q)>-1;
    return state.cls==='ALL'||s.cls===state.cls;
  });
  l.sort(function(a,b){
    if(state.sort==='id')return natural(a.id,b.id);
    if(state.sort==='class')return natural(a.cls,b.cls)||a.name.localeCompare(b.name);
    return a.name.localeCompare(b.name);
  });
  return l;
}
function renderStats(){
  var cl=classList(),withPh=students.filter(function(s){return photoFor(s.id)}).length;
  $('#stats').innerHTML=
   '<div class="stat s1"><b>'+students.length+'</b><span>Students</span></div>'+
   '<div class="stat s2"><b>'+cl.filter(function(c){return c.count>0}).length+'</b><span>Classes with students</span></div>'+
   '<div class="stat s3"><b>'+withPh+'</b><span>Photos matched</span></div>'+
   '<div class="stat s4"><b>'+(students.length-withPh)+'</b><span>Missing photos</span></div>';
}
function renderNav(){
  var cl=classList();
  if(!state.cls||(state.cls!=='ALL'&&!cl.some(function(c){return c.cls===state.cls}))){var first=cl.filter(function(c){return c.count>0})[0];state.cls=first?first.cls:'ALL'}
  var byGrade={},gOrder=[];
  cl.forEach(function(c){var g=c.grade||'';if(!byGrade[g]){byGrade[g]=[];gOrder.push(g)}byGrade[g].push(c)});
  gOrder.sort(natural);
  var h='<div class="group"><div class="glabel">ALL</div><div class="chips"><button type="button" class="chip" data-cls="ALL" aria-pressed="'+(state.cls==='ALL')+'">All classes <small>'+students.length+'</small></button></div></div>';
  gOrder.forEach(function(g){
    var list=byGrade[g],total=list.reduce(function(a,c){return a+c.count},0);
    h+='<div class="gtitle">'+(g?'Class '+esc(g):'Other')+' <small>'+list.length+' sections &middot; '+total+' students</small></div>';
    var groups={},order=[];
    list.forEach(function(c){var p=(c.cls.match(/^[A-Z]+/)||['Other'])[0];if(!groups[p]){groups[p]=[];order.push(p)}groups[p].push(c)});
    order.sort(natural).forEach(function(p){
      h+='<div class="group"><div class="glabel">'+esc(p)+'</div><div class="chips">'+groups[p].map(function(c){
        var dot=(c.count?c.e>c.m:c.guessE)?'e':'m';
        return '<button type="button" class="chip'+(c.count?'':' zero')+'" data-cls="'+esc(c.cls)+'" aria-pressed="'+(state.cls===c.cls)+'"><i class="'+dot+'"></i>'+esc(c.cls)+' <small>'+c.count+'</small></button>';
      }).join('')+'</div></div>';
    });
  });
  $('#nav').innerHTML=h;
}
function renderGrid(){
  var l=visible(),q=state.q.trim();
  var what=q?'results':(state.cls==='ALL'?'all':state.cls);
  $('#pdfClass').textContent='Download PDF: '+what+' ('+l.length+')';
  $('#printShown').textContent='Print '+what;
  $('#pdfClass').disabled=$('#printShown').disabled=!l.length;
  if(!l.length){$('#grid').innerHTML='<div class="empty">'+(q?'No students found.':'No students in this class yet. Upload its Excel file in Step 1.')+'</div>';$('#more').hidden=true;return}
  var part=l.slice(0,state.shown);
  $('#grid').innerHTML=part.map(function(s){
    return '<article class="item"><div class="cw" data-open="'+esc(s.id)+'" title="Click to enlarge">'+cardHTML(s)+'</div>'+
     '<div class="meta"><div><strong>'+esc(s.name)+'</strong><small>'+esc(s.id)+' &middot; '+esc(s.cls)+'<span class="tag '+(s.shift==='Evening'?'e':'m')+'">'+s.shift+'</span>'+(photoFor(s.id)?'':'<span class="tag nophoto">No photo</span>')+'</small></div>'+
     '<button class="btn small" type="button" data-print="'+esc(s.id)+'">Print</button></div></article>';
  }).join('');
  $('#more').hidden=l.length<=state.shown;
  $('#more').textContent='Show more ('+(l.length-state.shown)+' left)';
}
function render(){
  $('#banner').hidden=!isSample;
  $('#subtitle').textContent=settings.school;
  renderStats();renderNav();renderGrid();renderSingle();
}
function byId(id){return students.find(function(s){return s.id===id})}

/* ---------- events: lists ---------- */
$('#nav').addEventListener('click',function(e){var b=e.target.closest('[data-cls]');if(!b)return;state.cls=b.dataset.cls;state.q='';$('#q').value='';state.shown=60;renderNav();renderGrid()});
$('#grid').addEventListener('click',function(e){
  var p=e.target.closest('[data-print]');
  if(p){var s=byId(p.dataset.print);if(s)printCards([s]);return}
  var o=e.target.closest('[data-open]');
  if(o){var s2=byId(o.dataset.open);if(s2)openModal(s2)}
});
$('#q').addEventListener('input',function(e){state.q=e.target.value;state.shown=60;renderGrid()});
$('#sort').addEventListener('change',function(e){state.sort=e.target.value;state.shown=60;renderGrid()});
$('#more').addEventListener('click',function(){state.shown+=60;renderGrid()});
$('#printShown').addEventListener('click',function(){var l=visible();printCards(l)});
$('#pdfClass').addEventListener('click',function(){
  var l=visible(),q=state.q.trim();
  pdfFor(l,q?'Search-results.pdf':(state.cls==='ALL'?'All-Admit-Cards.pdf':safeName(state.cls)+'.pdf'));
});
$('#pdfAll').addEventListener('click',pdfAllClasses);
$('#zipAll').addEventListener('click',zipAllClasses);

/* ---------- modal ---------- */
function openModal(s){state.modalStudent=s;$('#mCard').innerHTML=cardHTML(s);$('#modal').classList.add('open')}
function closeModal(){$('#modal').classList.remove('open')}
$('#mClose').addEventListener('click',closeModal);
$('#mPrint').addEventListener('click',function(){if(state.modalStudent)printCards([state.modalStudent])});
$('#mPdf').addEventListener('click',function(){var s=state.modalStudent;if(s){closeModal();pdfFor([s],safeName(s.id)+'.pdf')}});
$('#modal').addEventListener('click',function(e){if(e.target.id==='modal')closeModal()});
document.addEventListener('keydown',function(e){if(e.key==='Escape')closeModal()});

/* ---------- tabs ---------- */
function setTab(t){
  state.tab=t;
  $('#viewC').hidden=t!=='C';$('#viewS').hidden=t!=='S';
  $('#tabC').setAttribute('aria-selected',t==='C');$('#tabS').setAttribute('aria-selected',t==='S');
  if(t==='S')renderSingle();
}
$('#tabC').addEventListener('click',function(){setTab('C')});
$('#tabS').addEventListener('click',function(){setTab('S')});

/* ---------- single card ---------- */
var manualPhoto='';
function singleData(){
  var id=$('#sc-id').value.trim();
  return {id:id||'0000000',name:($('#sc-name').value.trim()||'STUDENT NAME').toUpperCase(),father:($('#sc-father').value.trim()||"FATHER'S NAME").toUpperCase(),cls:($('#sc-class').value.trim()||'C1').toUpperCase().replace(/\s+/g,''),shift:$('#sc-shift').value,photo:manualPhoto};
}
function renderSingle(){$('#sc-prev').innerHTML=cardHTML(singleData())}
['sc-id','sc-name','sc-father','sc-class'].forEach(function(id){$('#'+id).addEventListener('input',renderSingle)});
$('#sc-shift').addEventListener('change',renderSingle);
$('#sc-id').addEventListener('change',function(){
  var s=byId($('#sc-id').value.trim());
  if(s){$('#sc-name').value=s.name;$('#sc-father').value=s.father;$('#sc-class').value=s.cls;$('#sc-shift').value=s.shift;manualPhoto=''}
  renderSingle();
});
$('#sc-photo').addEventListener('change',function(e){
  var f=e.target.files[0];if(!f)return;
  if(manualPhoto)URL.revokeObjectURL(manualPhoto);
  manualPhoto=URL.createObjectURL(f);renderSingle();
});
$('#sc-print').addEventListener('click',function(){printCards([singleData()])});
$('#sc-pdf').addEventListener('click',function(){var s=singleData();pdfFor([s],safeName(s.id)+'.pdf')});

/* ---------- settings UI ---------- */
var FIELDS={'set-school':'school','set-exam':'exam','set-morning':'morning','set-evening':'evening','set-title':'signTitle','set-ins':'instructions'};
function fillSettings(){
  Object.keys(FIELDS).forEach(function(id){$('#'+id).value=settings[FIELDS[id]]});
  [['sigMp','sigMorning'],['sigEp','sigEvening']].forEach(function(p){var im=$('#'+p[0]);if(settings[p[1]]){im.src=settings[p[1]];im.style.display='block'}else{im.removeAttribute('src');im.style.display='none'}});
}
Object.keys(FIELDS).forEach(function(id){
  $('#'+id).addEventListener('input',function(e){settings[FIELDS[id]]=e.target.value;saveSettings();$('#subtitle').textContent=settings.school;renderGrid();renderSingle()});
});
$('#resetSettings').addEventListener('click',function(){settings.instructions=DEFAULT_INS;settings.school=DEFAULTS.school;settings.exam=DEFAULTS.exam;settings.morning=DEFAULTS.morning;settings.evening=DEFAULTS.evening;settings.signTitle=DEFAULTS.signTitle;saveSettings();fillSettings();render()});
function readSig(file,key,imgId){
  if(!file)return;
  var rd=new FileReader();
  rd.onload=function(){
    var im=new Image();
    im.onload=function(){
      var w=Math.min(500,im.width),h=Math.round(im.height*w/im.width);
      var c=document.createElement('canvas');c.width=w;c.height=h;c.getContext('2d').drawImage(im,0,0,w,h);
      settings[key]=c.toDataURL('image/png');saveSettings();fillSettings();renderGrid();renderSingle();
    };
    im.src=rd.result;
  };
  rd.readAsDataURL(file);
}
$('#sigM').addEventListener('change',function(e){readSig(e.target.files[0],'sigMorning')});
$('#sigE').addEventListener('change',function(e){readSig(e.target.files[0],'sigEvening')});

/* ---------- inputs & drag/drop ---------- */
$('#xl').addEventListener('change',function(e){var f=Array.from(e.target.files);e.target.value='';handleExcel(f)});
$('#ph').addEventListener('change',function(e){var f=Array.from(e.target.files);e.target.value='';handlePhotos(f)});
$('#phDir').addEventListener('change',function(e){var f=Array.from(e.target.files);e.target.value='';handlePhotos(f)});
$('#missingBtn').addEventListener('click',downloadMissing);
function dnd(id,fn){
  var z=$(id);
  ['dragenter','dragover'].forEach(function(t){z.addEventListener(t,function(e){e.preventDefault();z.classList.add('over')})});
  ['dragleave','drop'].forEach(function(t){z.addEventListener(t,function(e){e.preventDefault();z.classList.remove('over')})});
  z.addEventListener('drop',function(e){fn(Array.from(e.dataTransfer.files))});
}
dnd('#dropXl',handleExcel);dnd('#dropPh',handlePhotos);
$('#clearAll').addEventListener('click',async function(){
  if(!window.__confirmClear){window.__confirmClear=true;this.textContent='Click again to confirm';var b=this;setTimeout(function(){window.__confirmClear=false;b.textContent='Clear all saved data'},4000);return}
  window.__confirmClear=false;this.textContent='Clear all saved data';
  LS.del('acg.students');LS.del('acg.settings');await dbClear();
  photos.forEach(function(u){URL.revokeObjectURL(u)});photos.clear();
  settings=Object.assign({},DEFAULTS);students=SAMPLE.slice();isSample=true;state.cls='';
  $('#xlStatus').textContent='';$('#phStatus').textContent='';
  fillSettings();render();
});

/* ---------- start ---------- */
$('#ver').textContent='v4 ready';
fillSettings();render();
dbLoad().then(function(){if(photos.size){$('#phStatus').className='status ok';$('#phStatus').textContent=photos.size+' saved photos loaded.';render()}});
window.__acg={readExcel:readExcel,cardHTML:cardHTML,normShift:normShift};
})();
