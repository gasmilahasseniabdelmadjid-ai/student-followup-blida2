import './style.css';
import * as XLSX from 'xlsx';
import {jsPDF} from 'jspdf';
import autoTable from 'jspdf-autotable';
import {Capacitor} from '@capacitor/core';
import {Filesystem,Directory} from '@capacitor/filesystem';
import {Share} from '@capacitor/share';
import {CapacitorNfc} from '@capgo/capacitor-nfc';
import {FilePicker} from '@capawesome/capacitor-file-picker';

const U='جامعة البليدة 2 لونيسي علي', KEY='student_followup_v3', VERSION='3.3.1';
const I=()=>({id:crypto.randomUUID(),course:'',level:'',department:'',specialty:'',group:'',teacher:'',semester:'',year:'2026/2027',session:0,sessions:Array.from({length:14},()=>''),students:[]});
let data=load(),lang=localStorage.getItem('sfc_lang')||'ar',nfcListener=null,nfcMode='attendance';
const T={ar:{title:'بطاقة متابعة الطلبة',new:'بطاقة جديدة',import:'استيراد Excel',pdf:'تصدير PDF',save:'حفظ',students:'الطلبة',add:'إضافة طالب',scan:'مسح NFC',course:'المقياس',level:'المستوى',dept:'القسم',spec:'التخصص',group:'الفوج',teacher:'الأستاذ',sem:'السداسي',year:'الموسم الجامعي',reg:'رقم التسجيل',name:'اللقب والاسم',abs:'الغيابات',conduct:'المواظبة /3',part:'المشاركة /3',work:'العمل الشخصي /4',exam:'الامتحان /10',final:'النهائية /20',session:'الحصة الحالية',link:'ربط البطاقة',choose:'اختر طالباً',present:'تم تسجيل الحضور',unknown:'البطاقة غير مرتبطة',unsupported:'NFC غير مدعوم',disabled:'NFC غير مفعّل'},fr:{title:'Fiche de suivi des étudiants',new:'Nouvelle fiche',importer:'Importer Excel',import:'Importer Excel',pdf:'Exporter PDF',save:'Enregistrer',students:'Étudiants',add:'Ajouter',scan:'Scanner NFC',course:'Module',level:'Niveau',dept:'Département',spec:'Spécialité',group:'Groupe',teacher:'Enseignant',sem:'Semestre',year:'Année universitaire',reg:'Matricule',name:'Nom et prénom',abs:'Absences',conduct:'Assiduité /3',part:'Participation /3',work:'Travail /4',exam:'Examen /10',final:'Finale /20',session:'Séance',link:'Associer',choose:'Choisir',present:'Présence enregistrée',unknown:'Carte non associée',unsupported:'NFC non pris en charge',disabled:'NFC désactivé'},en:{title:'Student Follow-up Card',new:'New card',import:'Import Excel',pdf:'Export PDF',save:'Save',students:'Students',add:'Add student',scan:'Scan NFC',course:'Course',level:'Level',dept:'Department',spec:'Specialty',group:'Group',teacher:'Teacher',sem:'Semester',year:'Academic year',reg:'Registration No.',name:'Name',abs:'Absences',conduct:'Conduct /3',part:'Participation /3',work:'Personal work /4',exam:'Exam /10',final:'Final /20',session:'Session',link:'Link card',choose:'Select student',present:'Attendance recorded',unknown:'Card not linked',unsupported:'NFC unsupported',disabled:'NFC disabled'}};
const tr=k=>T[lang][k]||T.en[k]||k;
function load(){try{return JSON.parse(localStorage.getItem(KEY))||{cards:[],active:null}}catch{return{cards:[],active:null}}}
function save(){localStorage.setItem(KEY,JSON.stringify(data))}
function card(){return data.cards.find(x=>x.id===data.active)}
function esc(x=''){return String(x).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function total(s){return Math.min(20,(+s.conduct||0)+(+s.part||0)+(+s.work||0)+(+s.exam||0))}
function newCard(){const c=I();data.cards.unshift(c);data.active=c.id;save();render()}
function render(){
 document.documentElement.lang=lang;document.documentElement.dir=lang==='ar'?'rtl':'ltr';
 const c=card();
 document.querySelector('#app').innerHTML=`<header><div><b>${U}</b><span>${tr('title')} · v${VERSION}</span></div><div class="actions"><button class="secondary" id="lang">${lang.toUpperCase()}</button><button id="new">${tr('new')}</button></div></header><main>${c?editor(c):welcome()}</main>`;
 document.querySelector('#new').onclick=newCard;
 document.querySelector('#lang').onclick=()=>{lang=lang==='ar'?'fr':lang==='fr'?'en':'ar';localStorage.setItem('sfc_lang',lang);render()};
 document.querySelector('#create')?.addEventListener('click',newCard);
 document.querySelectorAll('[data-id]').forEach(x=>x.onclick=()=>{data.active=x.dataset.id;save();render()});
 if(c) wire(c);
}
function showCards(){
 const m=document.createElement('div');m.className='modal';
 m.innerHTML=`<div class="modal-card cards-modal"><h2>📚 ${lang==='ar'?'بطاقاتي':lang==='fr'?'Mes fiches':'My cards'}</h2><div class="my-cards-list">${data.cards.map(x=>`<button class="my-card-item" data-card="${x.id}">${esc([x.course,x.specialty,x.group].filter(Boolean).join(' — ')||tr('new'))}</button>`).join('')}</div><button id="closeCards">إغلاق</button></div>`;
 document.body.appendChild(m);m.querySelector('#closeCards').onclick=()=>m.remove();
 m.querySelectorAll('[data-card]').forEach(b=>b.onclick=()=>{data.active=b.dataset.card;save();m.remove();render()});
}
function welcome(){return`<section class="welcome"><h1>${tr('title')}</h1><p>${U}</p><button class="big" id="create">＋ ${tr('new')}</button><div class="saved">${data.cards.map(c=>`<p><button data-id="${c.id}">${esc(c.course||tr('new'))} — ${esc(c.group)}</button></p>`).join('')}</div></section>`}
function editor(c){
 const present=c.students.filter(s=>s.attendance?.[c.session]==='P').length;
 return `<section class="stats"><div><b>${c.students.length}</b><span>${tr('students')}</span></div><div><b>${c.students.filter(s=>s.nfc).length}</b><span>NFC</span></div><div><b>${present}</b><span>${tr('present')}</span></div><div><b>${c.students.filter(s=>s.attendance?.[c.session]==='A').length}</b><span>${tr('abs')}</span></div><div><b>${c.students.filter(s=>s.attendance?.[c.session]==='E').length}</b><span>م</span></div></section>
 <section class="cardsbar"><button id="new2" class="card-action primary">＋ ${tr('new')}</button><button id="myCards" class="card-action secondary">📚 ${lang==='ar'?'بطاقاتي':lang==='fr'?'Mes fiches':'My cards'}</button><span class="active-card-name">${esc([c.course,c.specialty,c.group].filter(Boolean).join(' — ')||tr('new'))}</span></section><section class="toolbar"><button id="save">${tr('save')}</button><button id="pickExcel">📥 ${tr('import')}</button><input id="excel" type="file" accept=".xlsx,.xls,.csv,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style="display:none"><button id="pdf">${tr('pdf')}</button><button id="scan">📡 ${tr('scan')}</button><button id="link">${tr('link')}</button></section>
 <section class="cardform"><div class="fields">${field('department',tr('dept'),c.department)}${field('course',tr('course'),c.course)}${field('level',tr('level'),c.level)}${field('specialty',tr('spec'),c.specialty)}${field('group',tr('group'),c.group)}${field('teacher',tr('teacher'),c.teacher)}${field('semester',tr('sem'),c.semester)}${field('year',tr('year'),c.year)}</div></section>
 <section class="sessions"><b>${tr('session')}</b><select id="session" class="current-session">${c.sessions.map((d,i)=>`<option value="${i}" ${c.session===i?'selected':''}>ح${i+1} ${d||''}</option>`).join('')}</select><div class="sessiongrid">${c.sessions.map((d,i)=>`<label>ح${i+1}<input data-date="${i}" type="date" value="${d}"></label>`).join('')}</div></section>
 <section class="tablewrap"><button id="add">＋ ${tr('add')}</button><table><thead><tr><th>${tr('reg')}</th><th>${tr('name')}</th>${c.sessions.map((_,i)=>`<th>ح${i+1}</th>`).join('')}<th>${tr('abs')}</th><th>${tr('conduct')}</th><th>${tr('part')}</th><th>${tr('work')}</th><th>${tr('exam')}</th><th>${tr('final')}</th><th>NFC</th><th></th></tr></thead><tbody>${c.students.map((st,i)=>row(st,i,c)).join('')}</tbody></table></section>`;
}
function field(id,l,v){return`<label>${l}<input id="${id}" value="${esc(v)}"></label>`}
function row(s,i,c){return`<tr><td><input data-s="${i}" data-k="reg" value="${esc(s.reg)}"></td><td><input data-s="${i}" data-k="name" value="${esc(s.name)}"></td>${c.sessions.map((_,j)=>`<td><select data-s="${i}" data-a="${j}"><option value="">—</option><option ${s.attendance?.[j]==='P'?'selected':''} value="P">✓</option><option ${s.attendance?.[j]==='A'?'selected':''} value="A">غ</option><option ${s.attendance?.[j]==='E'?'selected':''} value="E">م</option></select></td>`).join('')}<td>${s.attendance?.filter(x=>x==='A').length||0}</td><td><input data-s="${i}" data-k="conduct" type="number" max="3" value="${s.conduct||''}"></td><td><input data-s="${i}" data-k="part" type="number" max="3" value="${s.part||''}"></td><td><input data-s="${i}" data-k="work" type="number" max="4" value="${s.work||''}"></td><td><input data-s="${i}" data-k="exam" type="number" max="10" value="${s.exam||''}"></td><td data-total>${total(s).toFixed(2)}</td><td>${s.nfc?'✓':'—'}</td><td><button class="danger" data-del="${i}">×</button></td></tr>`}
function wire(c){
 document.querySelector('#save').onclick=()=>{bind(c);toast(tr('save'))};
 const cardSelect=document.querySelector('#cardSelect'); if(cardSelect) cardSelect.onchange=e=>{bind(c);data.active=e.target.value;save();render()};
 document.querySelector('#new2').onclick=newCard; document.querySelector('#myCards').onclick=()=>showCards();
 document.querySelector('#add').onclick=()=>{c.students.push({reg:'',name:'',attendance:Array(14).fill(''),conduct:'',part:'',work:'',exam:'',nfc:''});save();render()};
 const picker=document.querySelector('#pickExcel'), input=document.querySelector('#excel');
 picker.onclick=async()=>{
  if(Capacitor.getPlatform()==='web'){input.click();return}
  try{
   const result=await FilePicker.pickFiles({types:['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/vnd.ms-excel','text/csv','application/octet-stream'],limit:1});
   const file=result.files?.[0]; if(!file)return;
   let blob=file.blob;
   if(!blob&&file.webPath){const response=await fetch(file.webPath);if(!response.ok)throw new Error('تعذر قراءة الملف المحدد');blob=await response.blob()}
   if(!blob)throw new Error('تعذر الوصول إلى محتوى ملف Excel');
   await importExcelBlob(blob,c,file.name||'Excel.xlsx');
  }catch(err){if(err?.message&&/cancel|dismiss|canceled|cancelled/i.test(err.message))return;alert('تعذر اختيار ملف Excel: '+(err?.message||err))}
 };
 input.onchange=e=>importExcel(e,c);
 document.querySelector('#pdf').onclick=()=>{bind(c);makePDF(c)};
 document.querySelector('#session').onchange=e=>{c.session=+e.target.value;save();render()};
 document.querySelectorAll('[data-date]').forEach(e=>e.onchange=()=>{c.sessions[+e.dataset.date]=e.value;save()});
 document.querySelectorAll('[data-s]').forEach(e=>e.oninput=()=>{const st=c.students[+e.dataset.s];if(e.dataset.a!==undefined)st.attendance[+e.dataset.a]=e.value;else st[e.dataset.k]=e.value;save();if(e.dataset.k) { const trEl=e.closest('tr'); if(trEl&&['conduct','part','work','exam'].includes(e.dataset.k)){const out=trEl.querySelector('[data-total]');if(out)out.textContent=total(st).toFixed(2)}}});
 document.querySelectorAll('[data-del]').forEach(e=>e.onclick=()=>{c.students.splice(+e.dataset.del,1);save();render()});
 document.querySelector('#scan').onclick=()=>nfc(c,'attendance');
 document.querySelector('#link').onclick=()=>nfc(c,'link');
}
function bind(c){['department','course','level','specialty','group','teacher','semester','year'].forEach(k=>{c[k]=document.querySelector('#'+k).value});save()}
async function importExcelBlob(blob,c,name='Excel.xlsx'){ const file=new File([blob],name,{type:blob.type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); await importExcel({target:{files:[file],value:name}},c); }
async function importExcel(e,c){
 const f=e.target.files?.[0]; if(!f)return;
 try{
  const wb=XLSX.read(await f.arrayBuffer(),{cellDates:false,raw:false});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
  if(!matrix.length)throw new Error('لم يتم العثور على بيانات في ملف Excel');

  const norm=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase()
    .replace(/[\s_.\/()\-]+/g,'').replace(/[إأآا]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه');

  // ملف الجامعة: السطر 1 عنوان الملف، السطر 2 رؤوس الأعمدة، ثم بيانات الطلبة.
  const headerIndex=matrix.findIndex((row,i)=>i<8 && row.some(v=>['matricule','رقمالتسجيل','nom','prenom','اللقب','الاسم','note'].includes(norm(v))));
  const h=headerIndex>=0?headerIndex:0;
  const rawHeader=(matrix[h]||[]).map(v=>String(v??'').trim());
  const isUniversityTemplate=rawHeader.length>=9 && norm(rawHeader[0])==='matricule' && norm(rawHeader[1])==='nom' && norm(rawHeader[2])==='prenom';
  const headers=isUniversityTemplate?['matricule','nom','prenom','note','absent','absencejustifiee','observation','section','groupe']:rawHeader.map((v,i)=>v||('COL'+i));
  const rows=matrix.slice(h+1).map(row=>{const o={};headers.forEach((k,i)=>o[k]=row[i]??'');return o}).filter(r=>Object.values(r).some(v=>String(v).trim()));

  const aliases={
   reg:['رقمالتسجيل','matricule','registrationno','registrationnumber','studentid','id','ninscription'],
   name:['اللقبوالاسم','nomprenom','nometprenom','name','fullname','studentname'],
   last:['اللقب','nom','lastname','surname'],
   first:['الاسم','prenom','firstname'],
   note:['note','mark','grade','العلامة','النقطة'],
   absent:['absent','absence','غياب','غائب'],
   justified:['absencejustifiee','absencedjustifiee','absjustifiee','absencejustified','غيابمبرر'],
   observation:['observation','remarque','ملاحظة'],
   section:['section','section'],
   group:['الفوج','groupe','group'],
   level:['المستوى','niveau','level'],
   spec:['التخصص','specialite','specialty'],
   nfc:['nfc','uid','nfcuid','rfid','rfiduid']
  };
  const find=(keys,want)=>keys.find(x=>aliases[want]?.includes(norm(x)))??null;
  const value=(r,want)=>{const k=find(Object.keys(r),want);return k==null?'':String(r[k]??'').trim()};

  // قراءة عنوان الملف في السطر الأول، مثل:
  // GASMI ... sociologie de loisir et du voyage/Semestre 2/Hôtels et restaurants/G2
  const title=String((matrix[0]||[]).find(v=>String(v).trim())||'').trim();
  if(title){
    const parts=title.split('/').map(x=>x.trim()).filter(Boolean);
    if(parts.length){
      const sem=parts.find(x=>/semestre|السداسي/i.test(x));
      const grp=parts.find(x=>/^G\\d+$/i.test(x)||/groupe|الفوج/i.test(x));
      const candidateCourse=parts.length>=4?parts[parts.length-3]:(parts.length>=2?parts[parts.length-2]:'');
      if(candidateCourse&&!c.course)c.course=candidateCourse;
      if(sem&&!c.semester)c.semester=sem;
      if(grp&&!c.group)c.group=grp.replace(/^groupe\\s*/i,'').trim();
    }
  }

  if(!rows.length)throw new Error('لم يتم العثور على بيانات الطلبة بعد صف العناوين');
  let added=0,updated=0,skipped=0;
  for(const r of rows){
   let reg=value(r,'reg');
   let name=value(r,'name');
   if(!name){const l=value(r,'last'),f2=value(r,'first');name=[l,f2].filter(Boolean).join(' / ')}
   const uid=value(r,'nfc').toUpperCase().replace(/[^0-9A-F]/g,'');
   if(isUniversityTemplate){
     const cells=Object.values(r).map(v=>String(v??'').trim());
     const positionalReg=cells[0]||'';
     const positionalName=[cells[1]||'',cells[2]||''].filter(Boolean).join(' / ');
     if(!reg) reg=positionalReg;
     if(!name) name=positionalName;
   }
   const group=value(r,'group'),level=value(r,'level'),spec=value(r,'spec');
   const note=value(r,'note'),absent=value(r,'absent'),justified=value(r,'justified'),observation=value(r,'observation'),section=value(r,'section');
   if(!reg&&!name){skipped++;continue}

   let st=c.students.find(x=>reg&&String(x.reg).trim()===reg)||c.students.find(x=>name&&String(x.name).trim()===name);
   if(!st){
     st={reg,name,group,level,specialty:spec,section,note,absent,justified,observation,
       attendance:Array(14).fill(''),conduct:'',part:'',work:'',exam:'',nfc:uid};
     c.students.push(st);added++;
   }else{
     st.name=name||st.name; if(reg)st.reg=reg; if(uid)st.nfc=uid;
     if(group)st.group=group;if(level)st.level=level;if(spec)st.specialty=spec;
     if(section)st.section=section;if(note)st.note=note;if(absent)st.absent=absent;
     if(justified)st.justified=justified;if(observation)st.observation=observation;
     updated++;
   }
   if(group&&!c.group)c.group=group;if(level&&!c.level)c.level=level;if(spec&&!c.specialty)c.specialty=spec;
   if(section&&!c.department)c.department=section;
  }
  save();render();
  toast(`${tr('import')}: +${added} / ${updated}${skipped?' — '+skipped+' متجاهل':''}`);
 }catch(err){alert('فشل استيراد Excel: '+(err?.message||err))}
 finally{e.target.value=''}
}
async function nfc(c,mode){
 const modal=document.createElement('div');modal.className='modal';
 modal.innerHTML=`<div class="modal-card"><h2>📡 NFC</h2><p>${mode==='link'?tr('link'):tr('scan')}</p>${mode==='link'?'<select id="nfcStudent">'+c.students.map((st,i)=>`<option value="${i}">${esc(st.reg)} — ${esc(st.name)}</option>`).join('')+'</select>':''}<div class="status" id="nfcStatus">جاري فحص دعم NFC…</div><button id="settings" style="display:none">⚙️ إعدادات NFC</button> <button id="close">إغلاق</button></div>`;
 document.body.appendChild(modal);
 modal.querySelector('#close').onclick=()=>stopNfc(modal);
 modal.querySelector('#settings').onclick=()=>CapacitorNfc.showSettings().catch(()=>{});
 try{
  if(Capacitor.getPlatform()==='web'){modal.querySelector('#nfcStatus').textContent='افتح نسخة Android المدمجة لاستخدام NFC';return}
  const info=await CapacitorNfc.isSupported();
  if(!info.supported){modal.querySelector('#nfcStatus').textContent=tr('unsupported');return}
  const st=await CapacitorNfc.getStatus();
  if(st.status==='NFC_DISABLED'){modal.querySelector('#nfcStatus').textContent=tr('disabled');modal.querySelector('#settings').style.display='inline-block';return}
  if(nfcListener)await nfcListener.remove();
  nfcListener=await CapacitorNfc.addListener('nfcEvent',ev=>{
   const raw=ev?.tag?.id;
   const uid=Array.isArray(raw)?raw.map(x=>{const n=typeof x==='string'?parseInt(x,16):Number(x);return Number.isFinite(n)?n.toString(16).padStart(2,'0'):''}).join('').toUpperCase():String(raw||'').replace(/[^0-9A-F]/gi,'').toUpperCase();
   if(!uid)return;
   const st2=mode==='link'?c.students[+(modal.querySelector('#nfcStudent')?.value||0)]:c.students.find(x=>String(x.nfc||'').replace(/[^0-9A-F]/gi,'').toUpperCase()===uid);
   if(mode==='link'&&st2){st2.nfc=uid;modal.querySelector('#nfcStatus').textContent='✓ '+tr('link')+': '+st2.name}
   else if(st2){st2.attendance[c.session]='P';modal.querySelector('#nfcStatus').textContent='✓ '+st2.name+' — '+tr('present')}
   else modal.querySelector('#nfcStatus').textContent=tr('unknown')+' ('+uid+')';
   save();render();
  });
  modal.querySelector('#nfcStatus').textContent='قرّب بطاقة NFC من ظهر الهاتف…';
  await CapacitorNfc.startScanning({invalidateAfterFirstRead:false,iosSessionType:'tag',alertMessage:'قرّب بطاقة الطالب من الهاتف'});
 }catch(err){modal.querySelector('#nfcStatus').textContent='خطأ NFC: '+(err?.message||String(err))}
}
async function stopNfc(m){try{await CapacitorNfc.stopScanning();if(nfcListener)await nfcListener.remove()}catch{}m.remove()}
async function makePDF(c){
 const d=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
 const add=async(n,f,st)=>{const r=await fetch('/fonts/'+f);if(!r.ok)throw new Error('font');const b=await r.arrayBuffer();let x='';const a=new Uint8Array(b);for(let i=0;i<a.length;i+=8192)x+=String.fromCharCode(...a.subarray(i,i+8192));d.addFileToVFS(f,x);d.addFont(f,n,st)};
 try{await add('NotoNaskh','NotoNaskhArabic-Regular.ttf','normal');await add('NotoNaskh','NotoNaskhArabic-Bold.ttf','bold')}catch{}
 const rtl=s=>{const v=String(s??'');return typeof d.processArabic==='function'?d.processArabic(v):v};
 const cellText=v=>rtl(v);
 d.setFont('NotoNaskh','normal');d.setFontSize(12);
 try{const lr=await fetch('/logo.png');if(lr.ok){const lb=await lr.blob();const reader=new FileReader();const logo=await new Promise(res=>{reader.onload=()=>res(reader.result);reader.readAsDataURL(lb)});d.addImage(logo,'PNG',12,6,24,24)}}catch{}
 d.text(cellText(U),148,10,{align:'center'});d.setFontSize(10);d.text(cellText('كلية العلوم الإنسانية والاجتماعية'),148,16,{align:'center'});d.text(cellText('قسم علم الاجتماع وعلم السكان'),148,22,{align:'center'});d.setFontSize(11);d.text(cellText('بطاقة متابعة الطلبة'),148,29,{align:'center'});

 const meta=[
  [cellText('المقياس'),cellText(c.course||'—'),cellText('المستوى'),cellText(c.level||'—'),cellText('القسم'),cellText(c.department||'—')],
  [cellText('التخصص'),cellText(c.specialty||'—'),cellText('الفوج'),cellText(c.group||'—'),cellText('الأستاذ'),cellText(c.teacher||'—')],
  [cellText('السداسي'),cellText(c.semester||'—'),cellText('الموسم الجامعي'),cellText(c.year||'—'),cellText('عدد الطلبة'),String(c.students.length)]
 ];
 autoTable(d,{startY:34,body:meta,theme:'grid',styles:{font:'NotoNaskh',fontStyle:'normal',fontSize:7,halign:'right',cellPadding:1.2},columnStyles:{0:{cellWidth:24,fontStyle:'bold'},1:{cellWidth:63},2:{cellWidth:24,fontStyle:'bold'},3:{cellWidth:63},4:{cellWidth:24,fontStyle:'bold'},5:{cellWidth:63}},didParseCell:data=>{if(typeof data.cell.text==='string')data.cell.text=[cellText(data.cell.text)];else data.cell.text=data.cell.text.map(x=>cellText(x))}});

 const head=[[cellText('رقم التسجيل'),cellText('اللقب والاسم'),cellText('القسم'),cellText('الفوج'),...c.sessions.map((_,i)=>cellText('ح'+(i+1))),cellText('الغيابات'),cellText('المواظبة /3'),cellText('المشاركة /3'),cellText('العمل /4'),cellText('الامتحان /10'),cellText('النهائية /20')]];
 const body=c.students.map(st=>[
  String(st.reg||''),
  cellText(st.name||[st.last,st.first].filter(Boolean).join(' / ')||'—'),
  cellText(st.section||c.department||'—'),
  cellText(st.group||c.group||'—'),
  ...c.sessions.map((_,i)=>st.attendance?.[i]==='P'?'✓':st.attendance?.[i]==='A'?cellText('غ'):st.attendance?.[i]==='E'?cellText('م'):''),
  String(st.attendance?.filter(x=>x==='A').length||0),
  String(st.conduct||''),String(st.part||''),String(st.work||''),String(st.exam||''),total(st).toFixed(2)
 ]);

 const tableY=(d.lastAutoTable?.finalY||34)+5;
 autoTable(d,{startY:tableY,head,body,theme:'grid',styles:{font:'NotoNaskh',fontStyle:'normal',fontSize:5.2,cellPadding:.65,halign:'center',overflow:'linebreak'},headStyles:{font:'NotoNaskh',fontStyle:'bold',fontSize:5.2},columnStyles:{0:{cellWidth:24},1:{cellWidth:34,halign:'right'},2:{cellWidth:18},3:{cellWidth:14}},didParseCell:data=>{if(Array.isArray(data.cell.text))data.cell.text=data.cell.text.map(x=>cellText(x));else data.cell.text=[cellText(data.cell.text)]}});

 const y=(d.lastAutoTable?.finalY||195)+5;
 d.setFontSize(7);d.text(cellText('ملاحظات: العلامة النهائية = المواظبة + المشاركة + العمل الشخصي + الامتحان.'),12,Math.min(y,202),{align:'left'});
 const name='بطاقة-متابعة-'+(c.course||'طلبة')+'.pdf';
 if(Capacitor.getPlatform()==='web'){d.save(name);return}
 const uri=d.output('datauristring').split(',')[1];const saved=await Filesystem.writeFile({path:name,data:uri,directory:Directory.Cache});await Share.share({title:'بطاقة متابعة الطلبة',url:saved.uri});
}
function toast(x){const t=document.createElement('div');t.className='toast';t.textContent=x;document.body.appendChild(t);setTimeout(()=>t.remove(),1800)}
render();