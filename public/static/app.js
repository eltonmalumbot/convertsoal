(()=>{
'use strict';
const $=id=>document.getElementById(id);
const escapeHTML=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const strip=s=>String(s||'').replace(/<[^>]*>/g,' ').replace(/\s+/g,' ').trim();
const state={questions:[],selected:0,source:'',warnings:[]};
const types={single:'Pilihan tunggal',multiple:'Pilihan ganda kompleks',essay:'Esai',matrix:'Tabel Benar/Salah',cloze:'Cloze (XML)',truefalse:'Benar/Salah'};
const toast=(message,error=false)=>{const x=$('toast');x.textContent=message;x.className='toast'+(error?' error':'');x.hidden=false;setTimeout(()=>x.hidden=true,6500)};
function issues(q){
 const x=[];
 if(!strip(q.stem))x.push('Teks soal kosong');
 if(['single','multiple','truefalse'].includes(q.type)){
  if((q.options||[]).length<2)x.push('Minimal 2 opsi');
  const correct=(q.options||[]).filter(o=>o.correct).length;
  if(!correct)x.push('Kunci belum ditentukan');
  if(q.type==='single'&&correct>1)x.push('Pilih hanya 1 jawaban');
 }
 if(q.type==='matrix'&&(!(q.rows||[]).length||(q.rows||[]).some(r=>!r.answer)))x.push('Kunci tabel belum lengkap');
 return x;
}
function render(){
 const q=state.questions;
 $('workspace').hidden=!q.length;
 $('upload-panel').hidden=!!q.length;
 if(!q.length)return;
 $('source-line').textContent=state.source;
 $('count-total').textContent=q.length;
 $('count-ready').textContent=q.filter(x=>!issues(x).length).length;
 $('count-review').textContent=q.filter(x=>issues(x).length).length;
 $('count-types').textContent=new Set(q.map(x=>x.type)).size;
 $('list-counter').textContent=q.length+' soal';
 const warnings=state.warnings||[];
 $('warnings').hidden=!warnings.length;
 $('warnings').textContent=warnings.join(' • ');
 renderList();renderEditor();
}
function renderList(){
 const search=$('search').value.toLowerCase(),status=$('status-filter').value;
 const el=$('question-list');
 const matches=state.questions.map((q,i)=>({q,i})).filter(({q})=>{
  const review=issues(q).length>0;
  return (status==='all'||(status==='review'?review:!review)) &&
    (q.name+' '+strip(q.stem)).toLowerCase().includes(search);
 });
 el.innerHTML=matches.length?'':'<div class="empty-list">Tidak ada soal yang cocok.</div>';
 for(const {q,i} of matches){
  const button=document.createElement('button');button.className='question-item'+(i===state.selected?' active':'');
  button.innerHTML='<span class="question-number">'+(i+1)+'</span><span class="question-body"><b>'+escapeHTML(q.name||'Soal '+(i+1))+'</b><small>'+escapeHTML(strip(q.stem).slice(0,105))+'</small></span><span class="status-dot '+(issues(q).length?'review':'ready')+'" title="'+(issues(q).length?'Perlu review':'Siap')+'"></span>';
  button.addEventListener('click',()=>{state.selected=i;renderList();renderEditor()});
  el.append(button);
 }
}
function renderEditor(){
 const el=$('editor');const q=state.questions[state.selected];
 $('editor-empty').hidden=!!q;el.hidden=!q;if(!q)return;
 const errs=issues(q);
 el.innerHTML='<div class="editor-head"><div><span class="section-label">SOAL '+(state.selected+1)+' / '+state.questions.length+'</span><h3>'+escapeHTML(q.name)+'</h3></div><span class="review-chip '+(errs.length?'bad':'good')+'">'+(errs.length?'Perlu diperiksa':'Siap ekspor')+'</span></div>'+
 '<div class="form-grid"><label>Nama soal<input id="edit-name" maxlength="250" value="'+escapeHTML(q.name)+'"/></label><label>Jenis soal<select id="edit-type">'+Object.keys(types).map(k=>'<option value="'+k+'" '+(q.type===k?'selected':'')+'>'+types[k]+'</option>').join('')+'</select></label></div>'+
 '<label class="form-label">Teks pertanyaan (HTML dari Word/XML diperbolehkan)<textarea id="edit-stem" rows="8"></textarea></label>'+
 '<div class="question-preview"><b>Pratinjau materi</b><div id="stem-preview"></div></div>'+
 '<div id="answer-editor"></div>'+
 '<div class="form-grid"><label>Nilai default<input type="number" min="0.01" step="0.25" id="edit-grade" value="'+escapeHTML(q.grade||1)+'"/></label></div>'+
 (errs.length?'<div class="edit-warning">⚠ '+escapeHTML(errs.join(' · '))+'</div>':'')+
 (q.notes?.length?'<div class="muted-note">'+escapeHTML(q.notes.join(' • '))+'</div>':'');
 $('edit-stem').value=q.stem||'';
 $('stem-preview').textContent=strip(q.stem);
 $('edit-name').addEventListener('input',e=>{q.name=e.target.value;renderList()});
 $('edit-type').addEventListener('change',e=>{q.type=e.target.value;render()});
 $('edit-stem').addEventListener('input',e=>{q.stem=e.target.value;$('stem-preview').textContent=strip(q.stem);renderList()});
 $('edit-grade').addEventListener('change',e=>{q.grade=Number(e.target.value)||1});
 renderAnswers(q);
}
function refreshEditor(){render();} 
function renderAnswers(q){
 const box=$('answer-editor');box.innerHTML='';
 if(['single','multiple','truefalse'].includes(q.type)){
  box.innerHTML='<div class="section-title"><b>Pilihan dan kunci jawaban</b><small>Pilih '+(q.type==='multiple'?'semua jawaban benar':'satu jawaban benar')+'</small></div>';
  const opts=document.createElement('div');opts.className='options';
  (q.options||[]).forEach((op,i)=>{
   const row=document.createElement('div');row.className='option-row';
   const checked=op.correct?'checked':'';
   row.innerHTML='<label class="key-check"><input type="'+(q.type==='multiple'?'checkbox':'radio')+'" name="answer-key" '+checked+' aria-label="Jawaban benar '+(i+1)+'"/><span>'+String.fromCharCode(65+i)+'</span></label><textarea rows="2" aria-label="Opsi '+(i+1)+'"></textarea><button class="remove-option" title="Hapus opsi">×</button>';
   row.querySelector('textarea').value=op.text||'';
   row.querySelector('textarea').addEventListener('change',e=>{op.text=e.target.value;renderList()});
   row.querySelector('input').addEventListener('change',e=>{
    if(q.type!=='multiple')q.options.forEach(a=>a.correct=false);
    op.correct=e.target.checked;refreshEditor();
   });
   row.querySelector('button').addEventListener('click',()=>{q.options.splice(i,1);refreshEditor()});
   opts.append(row);
  });
  box.append(opts);
  const add=document.createElement('button');add.className='secondary small';add.textContent='+ Tambah pilihan';add.disabled=q.options.length>=8;
  add.onclick=()=>{q.options.push({text:'',correct:false});refreshEditor()};box.append(add);
 }else if(q.type==='matrix'){
  box.innerHTML='<div class="section-title"><b>Pernyataan tabel</b><small>Pilih jawaban per baris</small></div>';
  (q.rows||[]).forEach((r,i)=>{
   const el=document.createElement('div');el.className='matrix-row';
   el.innerHTML='<label>Pernyataan '+(i+1)+'<textarea rows="2"></textarea></label><label>Jawaban<select><option value="">— Belum dipilih —</option>'+r.choices.map(c=>'<option value="'+escapeHTML(c)+'">'+escapeHTML(c)+'</option>').join('')+'</select></label>';
   el.querySelector('textarea').value=r.text;
   el.querySelector('textarea').addEventListener('change',e=>r.text=e.target.value);
   el.querySelector('select').value=r.answer||'';
   el.querySelector('select').onchange=e=>{r.answer=e.target.value;renderList()};box.append(el);
  });
 }else if(q.type==='cloze'){
  box.innerHTML='<p class="muted-note">Sintaks Cloze disimpan di teks soal. Pertahankan kode {1:MULTICHOICE:...} saat menyunting.</p>';
 }else{
  box.innerHTML='<p class="muted-note">Soal esai diperiksa secara manual oleh guru. Tidak memerlukan kunci pilihan.</p>';
 }
}
async function parseFile(file){
 if(!file)return;
 if(file.size>3800000)return toast('Ukuran file maksimal 3,8 MB pada Vercel.',true);
 $('dropzone').classList.add('loading');
 try{
  const data=new FormData();data.append('file',file);
  const response=await fetch('/api/parse',{method:'POST',body:data});
  const result=await response.json();
  if(!response.ok)throw Error(result.error||'Gagal membaca dokumen');
  state.questions=result.questions||[];state.source=result.source||file.name;
  state.warnings=result.warnings||[];state.selected=0;$('category').value=result.category||'Bank Soal';
  render();toast(state.questions.length+' soal berhasil dibaca. Periksa kunci sebelum ekspor.');
 }catch(err){toast(err.message||'Gagal mengunggah berkas',true)}
 finally{$('dropzone').classList.remove('loading');$('file-input').value=''}
}
async function exportFile(){
 const format=$('format').value;
 if(!state.questions.length)return toast('Belum ada soal.',true);
 const btn=$('export-btn');btn.disabled=true;btn.textContent='Menyiapkan file…';
 try{
  const res=await fetch('/api/export',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({questions:state.questions,format,category:$('category').value||'Bank Soal'})});
  if(!res.ok){let err=await res.json();throw Error(err.error||'Ekspor gagal')}
  const blob=await res.blob();
  const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;
  a.download='moodle-question-studio.'+(format==='xml'?'xml':format==='csv'?'csv':'txt');
  document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  toast('File berhasil dibuat. Silakan periksa sebelum impor Moodle.');
 }catch(e){toast(e.message,true)}finally{btn.disabled=false;btn.textContent='↓ Unduh hasil konversi'}
}
$('dropzone').onclick=()=>$('file-input').click();
$('dropzone').onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('file-input').click()}};
$('file-input').onchange=e=>parseFile(e.target.files?.[0]);
$('dropzone').ondragover=e=>{e.preventDefault();$('dropzone').classList.add('drag-over')};
$('dropzone').ondragleave=()=>$('dropzone').classList.remove('drag-over');
$('dropzone').ondrop=e=>{e.preventDefault();$('dropzone').classList.remove('drag-over');parseFile(e.dataTransfer.files?.[0])};
$('new-file-btn').onclick=()=>{state.questions=[];state.selected=0;render();$('file-input').click()};
$('search').oninput=renderList;$('status-filter').onchange=renderList;
$('export-btn').onclick=exportFile;
$('demo-btn').onclick=()=>{
 state.questions=[{name:'Contoh soal 01',type:'single',stem:'<p>Format apa yang paling lengkap untuk impor bank soal Moodle?</p>',options:[{text:'Moodle XML',correct:true},{text:'Aiken',correct:false},{text:'CSV biasa',correct:false},{text:'TXT polos',correct:false}],rows:[],notes:[],media:[],grade:1,feedback:''}];
 state.source='Contoh interaktif — bukan file unggahan';state.warnings=['Data contoh untuk mencoba editor.'];state.selected=0;render();
};
})();