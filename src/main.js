import './style.css';
import * as XLSX from 'xlsx';
import {Capacitor} from '@capacitor/core';
import {Filesystem,Directory} from '@capacitor/filesystem';
import {Share} from '@capacitor/share';
import {CapacitorNfc} from '@capgo/capacitor-nfc';
import {FilePicker} from '@capawesome/capacitor-file-picker';
import {Document,Packer,Paragraph,Table,TableRow,TableCell,TextRun,WidthType,AlignmentType,ImageRun} from 'docx';
import JsBarcode from 'jsbarcode';

const U='جامعة البليدة 2 لونيسي علي', UNI={ar:'جامعة البليدة 2 لونيسي علي',fr:'Université de Blida 2 Lounici Ali',en:'University of Blida 2 Lounici Ali'}, KEY='student_followup_v3', VERSION='3.6.0';
const I=()=>({id:crypto.randomUUID(),course:'',level:'',department:'',specialty:'',group:'',teacher:'',semester:'',year:'2026/2027',session:0,sessions:Array.from({length:14},()=>''),students:[]});
let data=load(),lang=localStorage.getItem('sfc_lang')||'ar',nfcListener=null,nfcMode='attendance';
const T={ar:{title:'بطاقة متابعة الطلبة',new:'بطاقة جديدة',import:'استيراد Excel',word:'تصدير Word',archive:'تصدير أرشيف',deleteAll:'حذف جميع البيانات',save:'حفظ',students:'الطلبة',add:'إضافة طالب',scan:'مسح NFC',course:'المقياس',level:'المستوى',dept:'القسم',spec:'التخصص',group:'الفوج',teacher:'الأستاذ',sem:'السداسي',year:'الموسم الجامعي',reg:'رقم التسجيل',name:'اللقب والاسم',abs:'الغيابات',conduct:'المواظبة /3',part:'المشاركة /3',work:'العمل الشخصي /4',exam:'الامتحان /10',final:'النهائية /20',session:'الحصة الحالية',link:'ربط البطاقة',choose:'اختر طالباً',present:'تم تسجيل الحضور',unknown:'البطاقة غير مرتبطة',unsupported:'NFC غير مدعوم',disabled:'NFC غير مفعّل'},fr:{title:'Fiche de suivi des étudiants',new:'Nouvelle fiche',importer:'Importer Excel',import:'Importer Excel',word:'Exporter Word',archive:'Exporter archive',deleteAll:'Supprimer toutes les données',save:'Enregistrer',students:'Étudiants',add:'Ajouter',scan:'Scanner NFC',course:'Module',level:'Niveau',dept:'Département',spec:'Spécialité',group:'Groupe',teacher:'Enseignant',sem:'Semestre',year:'Année universitaire',reg:'Matricule',name:'Nom et prénom',abs:'Absences',conduct:'Assiduité /3',part:'Participation /3',work:'Travail /4',exam:'Examen /10',final:'Finale /20',session:'Séance',link:'Associer',choose:'Choisir',present:'Présence enregistrée',unknown:'Carte non associée',unsupported:'NFC non pris en charge',disabled:'NFC désactivé'},en:{title:'Student Follow-up Card',new:'New card',import:'Import Excel',word:'Export Word',archive:'Export archive',deleteAll:'Delete all data',save:'Save',students:'Students',add:'Add student',scan:'Scan NFC',course:'Course',level:'Level',dept:'Department',spec:'Specialty',group:'Group',teacher:'Teacher',sem:'Semester',year:'Academic year',reg:'Registration No.',name:'Name',abs:'Absences',conduct:'Conduct /3',part:'Participation /3',work:'Personal work /4',exam:'Exam /10',final:'Final /20',session:'Session',link:'Link card',choose:'Select student',present:'Attendance recorded',unknown:'Card not linked',unsupported:'NFC unsupported',disabled:'NFC disabled'}};
const tr=k=>T[lang][k]||T.en[k]||k;
function load(){try{return JSON.parse(localStorage.getItem(KEY))||{cards:[],active:null}}catch{return{cards:[],active:null}}}
function save(){localStorage.setItem(KEY,JSON.stringify(data))}
function card(){return data.cards.find(x=>x.id===data.active)}
function esc(x=''){return String(x).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function total(s){return Math.min(20,(+s.conduct||0)+(+s.part||0)+(+s.work||0)+(+s.exam||0))}
function hasArabic(x=''){return /[\u0600-\u06FF]/.test(String(x))}
function displayName(st){
 const ar=String(st.nameAr||[st.lastAr,st.firstAr].filter(Boolean).join(' / ')||'').trim();
 const lat=String(st.nameLat||[st.lastLat,st.firstLat].filter(Boolean).join(' / ')||'').trim();
 return lang==='ar' ? ar : lat;
}
function cardBarcodeValue(c){return 'B2-'+String(c.id||'').replace(/-/g,'').slice(0,18).toUpperCase()}
async function blobBytes(blob){return new Uint8Array(await blob.arrayBuffer())}
async function fetchAssetBytes(path){const r=await fetch(path);if(!r.ok)throw new Error('asset');return new Uint8Array(await r.arrayBuffer())}
async function barcodePngBytes(value){
 const canvas=document.createElement('canvas');
 JsBarcode(canvas,value,{format:'CODE128',displayValue:true,font:'Arial',fontSize:12,height:46,width:2,margin:6,textMargin:3});
 const blob=await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(new Error('barcode')), 'image/png'));
 return blobBytes(blob);
}
function newCard(){const c=I();data.cards.unshift(c);data.active=c.id;save();render()}
function render(){
 document.documentElement.lang=lang;document.documentElement.dir=lang==='ar'?'rtl':'ltr';
 const c=card();
 document.querySelector('#app').innerHTML=`<header><div><b>${UNI[lang]}</b><span>${tr('title')} · v${VERSION}</span></div><div class="actions"><button class="secondary" id="lang">${lang.toUpperCase()}</button><button id="new">${tr('new')}</button></div></header><main>${c?editor(c):welcome()}</main>`;
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
 <section class="cardsbar"><button id="new2" class="card-action primary">＋ ${tr('new')}</button><button id="myCards" class="card-action secondary">📚 ${lang==='ar'?'بطاقاتي':lang==='fr'?'Mes fiches':'My cards'}</button><span class="active-card-name">${esc([c.course,c.specialty,c.group].filter(Boolean).join(' — ')||tr('new'))}</span></section><section class="toolbar"><button id="save">${tr('save')}</button><button id="pickExcel">📥 ${tr('import')}</button><input id="excel" type="file" accept=".xlsx,.xls,.csv,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style="display:none"><button id="word">📝 ${tr('word')}</button><button id="archive">🗃️ ${tr('archive')}</button><button id="deleteAll" class="danger">🗑️ ${tr('deleteAll')}</button><button id="scan">📡 ${tr('scan')}</button><button id="link">${tr('link')}</button></section>
 <section class="cardform"><div class="fields">${field('department',tr('dept'),c.department)}${field('course',tr('course'),c.course)}${field('level',tr('level'),c.level)}${field('specialty',tr('spec'),c.specialty)}${field('group',tr('group'),c.group)}${field('teacher',tr('teacher'),c.teacher)}${field('semester',tr('sem'),c.semester)}${field('year',tr('year'),c.year)}</div></section>
 <section class="sessions"><b>${tr('session')}</b><select id="session" class="current-session">${c.sessions.map((d,i)=>`<option value="${i}" ${c.session===i?'selected':''}>ح${i+1} ${d||''}</option>`).join('')}</select><div class="sessiongrid">${c.sessions.map((d,i)=>`<label>ح${i+1}<input data-date="${i}" type="date" value="${d}"></label>`).join('')}</div></section>
 <section class="tablewrap"><button id="add">＋ ${tr('add')}</button><table><thead><tr><th>${tr('reg')}</th><th>${tr('name')}</th>${c.sessions.map((_,i)=>`<th>ح${i+1}</th>`).join('')}<th>${tr('abs')}</th><th>${tr('conduct')}</th><th>${tr('part')}</th><th>${tr('work')}</th><th>${tr('exam')}</th><th>${tr('final')}</th><th>NFC</th><th></th></tr></thead><tbody>${c.students.map((st,i)=>row(st,i,c)).join('')}</tbody></table></section>`;
}
function field(id,l,v){return`<label>${l}<input id="${id}" value="${esc(v)}"></label>`}
function row(s,i,c){return`<tr><td><input data-s="${i}" data-k="reg" value="${esc(s.reg)}"></td><td><input data-s="${i}" data-k="name" value="${esc(displayName(s))}"></td>${c.sessions.map((_,j)=>`<td><select data-s="${i}" data-a="${j}"><option value="">—</option><option ${s.attendance?.[j]==='P'?'selected':''} value="P">✓</option><option ${s.attendance?.[j]==='A'?'selected':''} value="A">غ</option><option ${s.attendance?.[j]==='E'?'selected':''} value="E">م</option></select></td>`).join('')}<td>${s.attendance?.filter(x=>x==='A').length||0}</td><td><input data-s="${i}" data-k="conduct" type="number" max="3" value="${s.conduct||''}"></td><td><input data-s="${i}" data-k="part" type="number" max="3" value="${s.part||''}"></td><td><input data-s="${i}" data-k="work" type="number" max="4" value="${s.work||''}"></td><td><input data-s="${i}" data-k="exam" type="number" max="10" value="${s.exam||''}"></td><td data-total>${total(s).toFixed(2)}</td><td>${s.nfc?'✓':'—'}</td><td><button class="danger" data-del="${i}">×</button></td></tr>`}
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
 document.querySelector('#word').onclick=async()=>{bind(c);try{await makeWord(c)}catch(err){console.error('Word export failed',err);alert((lang==='ar'?'تعذر تصدير بطاقة Word: ':lang==='fr'?'Échec de l’export Word : ':'Word export failed: ')+(err?.message||String(err)))}};
 document.querySelector('#archive').onclick=()=>exportArchive();
 document.querySelector('#deleteAll').onclick=()=>deleteAllData();
 document.querySelector('#session').onchange=e=>{c.session=+e.target.value;save();render()};
 document.querySelectorAll('[data-date]').forEach(e=>e.onchange=()=>{c.sessions[+e.dataset.date]=e.value;save()});
 document.querySelectorAll('[data-s]').forEach(e=>e.oninput=()=>{const st=c.students[+e.dataset.s];if(e.dataset.a!==undefined)st.attendance[+e.dataset.a]=e.value;else if(e.dataset.k==='name'){st.name=e.value;if(lang==='ar')st.nameAr=e.value;else st.nameLat=e.value}else st[e.dataset.k]=e.value;save();if(e.dataset.k) { const trEl=e.closest('tr'); if(trEl&&['conduct','part','work','exam'].includes(e.dataset.k)){const out=trEl.querySelector('[data-total]');if(out)out.textContent=total(st).toFixed(2)}}});
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
   nameAr:['الاسمبالعربية','الاسمبالعربي','nomprenomarabe','nometprenomarabe','arabicname','arabicfullname'],
   lastAr:['اللقببالعربية','اللقببالعربي','nomarabe','arabiclastname'],
   firstAr:['الاسمالعربيه','الاسمبالعربي','prenomarabe','arabicfirstname'],
   nameLat:['الاسمواللقبباللاتينية','الاسمواللقبباللاتينيه','nomprenomlatin','nometprenomlatin','latinname','latinfullname','nameen','namefr'],
   lastLat:['اللقبباللاتينية','اللقبباللاتينيه','nomlatin','latinlastname'],
   firstLat:['الاسماللاتيني','الاسمباللاتينية','prenomlatin','latinfirstname'],
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
   let last=value(r,'last');
   let first=value(r,'first');
   let name=value(r,'name');
   let nameAr=value(r,'nameAr')||[value(r,'lastAr'),value(r,'firstAr')].filter(Boolean).join(' / ');
   let nameLat=value(r,'nameLat')||[value(r,'lastLat'),value(r,'firstLat')].filter(Boolean).join(' / ');
   if(!name)name=[last,first].filter(Boolean).join(' / ');
   if(!nameAr&&hasArabic(name))nameAr=name;
   if(!nameLat&&name&&!hasArabic(name))nameLat=name;
   const uid=value(r,'nfc').toUpperCase().replace(/[^0-9A-F]/g,'');
   if(isUniversityTemplate){
     const cells=Object.values(r).map(v=>String(v??'').trim());
     const positionalReg=cells[0]||'';
     const positionalLast=cells[1]||'';
     const positionalFirst=cells[2]||'';
     const positionalName=[positionalLast,positionalFirst].filter(Boolean).join(' / ');
     if(!reg) reg=positionalReg;
     if(!last) last=positionalLast;
     if(!first) first=positionalFirst;
     if(!name) name=positionalName;
   }
   const group=value(r,'group'),level=value(r,'level'),spec=value(r,'spec');
   const note=value(r,'note'),absent=value(r,'absent'),justified=value(r,'justified'),observation=value(r,'observation'),section=value(r,'section');
   if(!reg&&!name){skipped++;continue}

   let st=c.students.find(x=>reg&&String(x.reg).trim()===reg)||c.students.find(x=>name&&String(x.name).trim()===name);
   if(!st){
     st={reg,name,nameAr,nameLat,last,first,lastAr:value(r,'lastAr'),firstAr:value(r,'firstAr'),lastLat:value(r,'lastLat')||last,firstLat:value(r,'firstLat')||first,group,level,specialty:spec,section,note,absent,justified,observation,
       attendance:Array(14).fill(''),conduct:'',part:'',work:'',exam:'',nfc:uid};
     c.students.push(st);added++;
   }else{
     st.name=name||st.name; if(nameAr)st.nameAr=nameAr; if(nameLat)st.nameLat=nameLat; if(last)st.last=last; if(first)st.first=first; if(nameAr)st.lastAr=value(r,'lastAr')||st.lastAr; if(nameAr)st.firstAr=value(r,'firstAr')||st.firstAr; if(nameLat)st.lastLat=value(r,'lastLat')||st.lastLat; if(nameLat)st.firstLat=value(r,'firstLat')||st.firstLat; if(reg)st.reg=reg; if(uid)st.nfc=uid;
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
 modal.innerHTML=`<div class="modal-card"><h2>📡 NFC</h2><p>${mode==='link'?tr('link'):tr('scan')}</p>${mode==='link'?'<select id="nfcStudent">'+c.students.map((st,i)=>`<option value="${i}">${esc(st.reg)} — ${esc(displayName(st))}</option>`).join('')+'</select>':''}<div class="status" id="nfcStatus">جاري فحص دعم NFC…</div><button id="settings" style="display:none">⚙️ إعدادات NFC</button> <button id="close">إغلاق</button></div>`;
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

function blobToBase64(blob){return new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(String(r.result).split(',')[1]);r.onerror=reject;r.readAsDataURL(blob)})}
async function deliverFile(blob,filename,mime){
 if(!blob||blob.size===0)throw new Error('تم إنشاء ملف Word فارغ');
 if(Capacitor.getPlatform()==='web'){
   const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename;document.body.appendChild(a);a.click();a.remove();
   setTimeout(()=>URL.revokeObjectURL(a.href),1500);return;
 }
 const base64=await blobToBase64(blob);
 const safeName=(filename||'Student-Follow-Up.docx').replace(/[\\/:*?"<>|]/g,'-');
 const saved=await Filesystem.writeFile({path:safeName,data:base64,directory:Directory.Cache,recursive:true});
 if(!saved?.uri)throw new Error('تعذر حفظ ملف Word على الهاتف');
 const canShare=await Share.canShare().catch(()=>({value:true}));
 if(canShare?.value===false)throw new Error('المشاركة غير متاحة على هذا الهاتف');
 await Share.share({title:safeName,url:saved.uri,dialogTitle:lang==='ar'?'مشاركة بطاقة Word':lang==='fr'?'Partager la fiche Word':'Share Word card'});
}
async function makeWord(c){
 const L={ar:{title:'بطاقة متابعة الطلبة',faculty:'كلية العلوم الإنسانية والاجتماعية',dept:'قسم علم الاجتماع وعلم السكان',reg:'رقم التسجيل',name:'اللقب والاسم',abs:'عدد الغيابات',conduct:'علامة المواظبة والسلوك /3',part:'علامة المشاركة /3',work:'علامة العمل الشخصي /4',exam:'التقييم /10',finale:'العلامة النهائية /20',course:'المقياس',level:'المستوى',department:'القسم',specialty:'التخصص',group:'الفوج',teacher:'الأستاذ',semester:'السداسي',year:'الموسم الجامعي',sessions:'تاريخ الحصص',note:'ملاحظة'},fr:{title:'Fiche de suivi des étudiants',faculty:'Faculté des sciences humaines et sociales',dept:'Département de sociologie et de démographie',reg:'Matricule',name:'Nom et prénom',abs:'Nombre d’absences',conduct:'Assiduité /3',part:'Participation /3',work:'Travail personnel /4',exam:'Évaluation /10',finale:'Note finale /20',course:'Module',level:'Niveau',department:'Département',specialty:'Spécialité',group:'Groupe',teacher:'Enseignant',semester:'Semestre',year:'Année universitaire',sessions:'Dates des séances',note:'Note'},en:{title:'Student Follow-up Card',faculty:'Faculty of Humanities and Social Sciences',dept:'Department of Sociology and Demography',reg:'Registration No.',name:'Student name',abs:'Absences',conduct:'Conduct /3',part:'Participation /3',work:'Personal work /4',exam:'Assessment /10',finale:'Final grade /20',course:'Course',level:'Level',department:'Department',specialty:'Specialty',group:'Group',teacher:'Teacher',semester:'Semester',year:'Academic year',sessions:'Session dates',note:'Note'}}[lang];
 const rtl=lang==='ar';
 let logo=null,barcode=null;
 try{logo=await fetchAssetBytes('/university-logo.jpg')}catch(err){console.warn('University logo unavailable',err)}
 try{barcode=await barcodePngBytes(cardBarcodeValue(c))}catch(err){console.warn('Barcode unavailable',err)}

 // A4 landscape printable area: 16838 - 600 = 16238 DXA.
 // The previous export used >22,000 DXA, which caused Word to clip/overflow columns.
 const W=[1050,2450,...Array(14).fill(520),700,700,700,700,700,700];
 const totalW=W.reduce((x,y)=>x+y,0); // 14980 DXA
 const run=(v,o={})=>new TextRun({text:String(v??'—'),font:'Arial',size:o.size||16,bold:!!o.bold,rtl});
 const para=(v,o={})=>new Paragraph({children:[run(v,o)],alignment:o.align||(rtl?AlignmentType.RIGHT:AlignmentType.LEFT),bidirectional:rtl,spacing:{after:0,before:0}});
 const imgPara=(im,a=AlignmentType.CENTER)=>new Paragraph({children:[im],alignment:a,spacing:{after:0,before:0}});
 const cell=(v,w,i,{bold=false,size=9,center=false}={})=>new TableCell({
   width:{size:w,type:WidthType.DXA},
   margins:{top:45,bottom:45,left:35,right:35},
   children:[para(v,{bold,size,align:center?AlignmentType.CENTER:(rtl?AlignmentType.RIGHT:AlignmentType.LEFT)})]
 });
 const meta=[[L.course,c.course,L.level,c.level],[L.department,c.department,L.specialty,c.specialty],[L.group,c.group,L.teacher,c.teacher],[L.semester,c.semester,L.year,c.year]];
 const metaWidths=[1450,2600,1450,2600];
 const metaTable=new Table({
   width:{size:8100,type:WidthType.DXA},
   alignment:AlignmentType.CENTER,
   rows:meta.map(a=>new TableRow({children:(rtl?[a[2],a[3],a[0],a[1]]:a).map((v,i)=>cell(v,metaWidths[i],i,{bold:i%2===0,size:10}) )}))
 });
 const headers=[L.reg,L.name,...c.sessions.map((_,i)=>rtl?'ح'+(i+1):'S'+(i+1)),L.abs,L.conduct,L.part,L.work,L.exam,L.finale];
 const dates=['','',...c.sessions.map(d=>d||''),'','','','','',''];
 const ordered=a=>rtl?[...a].reverse():a;
 const makeHeaderRow=a=>new TableRow({children:ordered(a).map((v,i)=>{
   const logicalIndex=rtl?a.length-1-i:i;
   return cell(v,W[logicalIndex],logicalIndex,{bold:true,size:8,center:true});
 })});
 const makeDataRow=a=>new TableRow({children:ordered(a).map((v,i)=>{
   const logicalIndex=rtl?a.length-1-i:i;
   return cell(v,W[logicalIndex],logicalIndex,{size:8,center:true});
 })});
 const rows=[makeHeaderRow(headers),makeDataRow(dates)];
 for(const st of c.students){
   const vals=[st.reg,displayName(st),
     ...c.sessions.map((_,i)=>st.attendance?.[i]==='P'?'✓':st.attendance?.[i]==='A'?(rtl?'غ':'A'):st.attendance?.[i]==='E'?(rtl?'م':'E'):''),
     st.attendance?.filter(x=>x==='A').length||0,st.conduct||'',st.part||'',st.work||'',st.exam||'',total(st).toFixed(2)];
   rows.push(makeDataRow(vals));
 }
 const studentTable=new Table({
   width:{size:totalW,type:WidthType.DXA},
   alignment:AlignmentType.CENTER,
   rows,
   columnWidths:W
 });

 const logoRun=logo?new ImageRun({type:'jpg',data:logo,transformation:{width:62,height:62}}):null;
 const barcodeRun=barcode?new ImageRun({type:'png',data:barcode,transformation:{width:150,height:49}}):null;
 const headerLeft=new TableCell({width:{size:1450,type:WidthType.DXA},children:logoRun?[imgPara(logoRun)]:[para('')]});
 const headerText=new TableCell({width:{size:10000,type:WidthType.DXA},children:[
   para(UNI[lang],{bold:true,size:19,align:AlignmentType.CENTER}),
   para(L.faculty,{size:11,align:AlignmentType.CENTER}),
   para(L.dept,{size:11,align:AlignmentType.CENTER}),
   para(L.title,{bold:true,size:15,align:AlignmentType.CENTER})
 ]});
 const headerCells=rtl?[headerText,headerLeft]:[headerLeft,headerText];
 const titleTable=new Table({width:{size:11450,type:WidthType.DXA},alignment:AlignmentType.CENTER,rows:[new TableRow({children:headerCells})]});
 const barcodeTable=new Table({
   width:{size:totalW,type:WidthType.DXA},alignment:AlignmentType.CENTER,
   rows:[new TableRow({children:[
     new TableCell({width:{size:totalW-1900,type:WidthType.DXA},children:[para('')]}),
     new TableCell({width:{size:1900,type:WidthType.DXA},children:barcode?[imgPara(barcodeRun,rtl?AlignmentType.LEFT:AlignmentType.RIGHT)]:[para('')]})
   ]})]
 });
 const notes=rtl
   ? 'ملاحظة: علامة الامتحان لا تزيد عن /10، والعلامة النهائية هي حصيلة مجموع العلامات. تُحتسب الغيابات وفق النظام المعتمد.'
   : lang==='fr'
   ? 'Note : l’évaluation est sur 10 et la note finale correspond à la somme des notes renseignées. Les absences sont comptabilisées selon le règlement en vigueur.'
   : 'Note: assessment is out of 10 and the final grade is the sum of the entered grades. Absences are counted according to the applicable regulations.';

 const doc=new Document({
   styles:{default:{document:{run:{font:'Arial',size:16}}}},
   sections:[{
     properties:{page:{size:{width:16838,height:11906,orientation:'landscape'},margin:{top:250,right:300,bottom:250,left:300}}},
     children:[
       titleTable,
       new Paragraph({children:[],spacing:{after:25,before:0}}),
       barcodeTable,
       new Paragraph({children:[],spacing:{after:20,before:0}}),
       metaTable,
       para(L.sessions,{bold:true,size:10,align:rtl?AlignmentType.RIGHT:AlignmentType.LEFT}),
       para(c.sessions.map((d,i)=>'ح'+(i+1)+': '+(d||'—')).join('   |   '),{size:9,align:rtl?AlignmentType.RIGHT:AlignmentType.LEFT}),
       new Paragraph({children:[],spacing:{after:20,before:0}}),
       studentTable,
       new Paragraph({children:[],spacing:{after:10,before:0}}),
       para(notes,{size:8,align:rtl?AlignmentType.RIGHT:AlignmentType.LEFT}),
       para((rtl?'رمز البطاقة: ':lang==='fr'?'Code de la fiche : ':'Card code: ')+cardBarcodeValue(c),{size:7,align:rtl?AlignmentType.RIGHT:AlignmentType.LEFT})
     ]
   }]
 });
 const blob=await Packer.toBlob(doc);
 if(!blob||blob.size<1000)throw new Error('تعذر إنشاء مستند Word صالح');
 const safe=(c.course||'student-follow-up').replace(/[\\/:*?"<>|]/g,'-').slice(0,80);
 const filename=(lang==='ar'?'بطاقة-متابعة-':lang==='fr'?'Fiche-suivi-':'Student-Follow-Up-')+safe+'.docx';
 await deliverFile(blob,filename,'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
 toast(tr('word'));
}
async function exportArchive(){const archive={format:'StudentFollowUpArchive',version:VERSION,university:UNI,language:lang,exportedAt:new Date().toISOString(),cards:data.cards};const blob=new Blob([JSON.stringify(archive,null,2)],{type:'application/json;charset=utf-8'});const filename='StudentFollowUp-Blida2-Archive-'+new Date().toISOString().slice(0,10)+'.json';await deliverFile(blob,filename,'application/json');toast(tr('archive'))}
function deleteAllData(){const msg=lang==='ar'?'سيتم حذف جميع البطاقات والطلبة والربط ببطاقات NFC من هذا الهاتف. لا يمكن التراجع عن العملية. هل تريد المتابعة؟':lang==='fr'?'Toutes les fiches, étudiants et associations NFC seront supprimés de cet appareil. Cette action est irréversible. Continuer ?':'All cards, students and NFC associations will be deleted from this device. This cannot be undone. Continue?';if(!confirm(msg))return;if(!confirm(lang==='ar'?'تأكيد نهائي: حذف كل البيانات؟':lang==='fr'?'Confirmation finale : supprimer toutes les données ?':'Final confirmation: delete all data?'))return;data={cards:[],active:null};save();render();toast(lang==='ar'?'تم حذف جميع البيانات':lang==='fr'?'Toutes les données ont été supprimées':'All data deleted')}

function toast(x){const t=document.createElement('div');t.className='toast';t.textContent=x;document.body.appendChild(t);setTimeout(()=>t.remove(),1800)}
render();