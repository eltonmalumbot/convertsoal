import base64
import csv
import html
import io
import re
import zipfile
from collections import Counter
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
from lxml import etree

class ConversionError(Exception):
    pass

Q_MARKER = re.compile(r'^\s*(\d{1,3})[.)]\s+(.*)$', re.S)
ANSWER_LINE = re.compile(r'^\s*(?:jawab(?:an|annya)?|jawban|kunci\s*jawaban)\s*[:=]?\s*(.*)$', re.I | re.S)
LETTER_OPT = re.compile(r'^\s*([A-H])[.)]\s+(.*)$', re.I | re.S)
PAREN_OPT = re.compile(r'^\s*\[\s*[…\.\s]*\]\s*', re.S)
NOTE_WORDS = re.compile(r'^\s*(?:bacalah|cermati(?:lah)?|perhatikan|kutipan teks|teks \d+|sumber\s*:|rabu,|soal to\.)\b', re.I)

def escape(s):
    return html.escape(str(s or ''), quote=True)

def para_html(text):
    return '<p>' + escape(text).replace('\n', '<br/>') + '</p>'

def media_html(media):
    return ''.join('<p><img src="@@PLUGINFILE@@/' + escape(re.sub(r'[^a-zA-Z0-9._-]', '_', f['name'])[:120]) + '" alt="Gambar soal" /></p>' for f in media)

def strip_markup(s):
    return re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]*>', ' ', str(s or '')))).strip()

def new_question(name, stem='', typ='single'):
    return {'name': str(name), 'type': typ, 'stem': stem, 'options': [], 'rows': [],
            'notes': [], 'media': [], 'grade': 1, 'feedback': '', 'needsReview': False}

def blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)

def num_id(par):
    pr = par._p.pPr
    n = pr.numPr if pr is not None else None
    return int(n.numId.val) if n is not None and n.numId is not None else None

def media_for_paragraph(par, idx):
    result = []
    for r in par._p.xpath('.//a:blip'):
        rid = r.get(qn('r:embed'))
        part = par.part.related_parts.get(rid) if rid else None
        if not part or not getattr(part, 'content_type', '').startswith('image/'):
            continue
        ext = part.content_type.split('/')[-1].replace('jpeg', 'jpg')
        result.append({'name':f'word-image-{idx}-{len(result)+1}.{ext}',
                       'mime':part.content_type, 'data':base64.b64encode(part.blob).decode('ascii')})
    return result

def parse_key_text(s):
    s = (s or '').strip().upper()
    letters = re.findall(r'(?<![A-Z])[A-H](?![A-Z])', s)
    if not letters:
        letters = [chr(64+int(n)) for n in re.findall(r'(?<!\d)\d{1,2}(?!\d)', s) if 1 <= int(n) <= 8]
    return sorted(set(letters))

def apply_answer(q, key):
    if not q: return
    raw = (key or '').upper().strip()
    if q['rows']:
        abbr = re.findall(r'(?<![A-Z])(TS|S|B|T|F)(?![A-Z])', raw)
        if len(abbr) == len(q['rows']):
            for r,a in zip(q['rows'], abbr):
                opts=r['choices']
                if a in ('B','T') or (a=='S' and str(opts[0]).lower().startswith('setuju')):
                    r['answer']=opts[0]
                elif a in ('TS','F') or (a=='S' and str(opts[-1]).lower().startswith('salah')):
                    r['answer']=opts[-1]
            return
        good=parse_key_text(raw)
        if good:
            for i,r in enumerate(q['rows']):
                r['answer']=r['choices'][0] if chr(65+i) in good else r['choices'][-1]
            return
    letters=parse_key_text(raw)
    if not letters: return
    count=0
    for i,op in enumerate(q['options']):
        op['correct']=chr(65+i) in letters
        count+=int(op['correct'])
    if count>1: q['type']='multiple'

def validate(q):
    q['needsReview']=False
    q['notes']=list(dict.fromkeys(q.get('notes', [])))
    if q['type'] in ('single','multiple','truefalse'):
        if len(q['options'])<2:
            q['notes'].append('Pilihan jawaban belum lengkap.')
            q['needsReview']=True
        n=sum(bool(x.get('correct')) for x in q['options'])
        if not n:
            q['notes'].append('Kunci jawaban tidak ditemukan. Pilih jawaban benar sebelum ekspor.')
            q['needsReview']=True
        elif q['type']=='single' and n>1: q['type']='multiple'
    elif q['type']=='matrix':
        if not q['rows'] or any(not r.get('answer') for r in q['rows']):
            q['notes'].append('Beberapa jawaban tabel belum ditetapkan.')
            q['needsReview']=True
    return q

def parse_docx(data, filename='questions.docx'):
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise ConversionError('DOCX bukan ZIP Office yang valid.')
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if len(z.infolist())>5000 or sum(x.file_size for x in z.infolist())>65*1024*1024:
            raise ConversionError('DOCX terlalu besar setelah diekstrak.')
    d=Document(io.BytesIO(data))
    items=list(blocks(d))
    nums=Counter(num_id(b) for b in items if isinstance(b,Paragraph) and num_id(b))
    question_ids={n for n,count in nums.items() if count>=10}
    if not question_ids: question_ids={n for n,count in nums.items() if count>=4 and n==1}
    questions=[]
    current=None
    before=[]
    before_media=[]
    shared_html=''
    shared_media=[]
    share_remaining=0
    option_id=None
    question_counter=0

    def start(text,media):
        nonlocal current,before,before_media,shared_html,shared_media,share_remaining,option_id,question_counter
        if current: questions.append(validate(current))
        question_counter+=1
        if before:
            shared_html=''.join(before[-12:])
            shared_media=before_media[:]
            joint=strip_markup(shared_html)
            match=re.search(r'soal\s*(?:nomor|no|n0)?\s*(\d+)\s*(?:dan|[-–])\s*(\d+)',joint,re.I)
            share_remaining=max(0,min(5,int(match.group(2))-int(match.group(1)))) if match else 0
            before,before_media=[],[]
        elif share_remaining>0: share_remaining-=1
        else: shared_html,shared_media='',[]
        current=new_question(f'Soal {question_counter:02}',shared_html+para_html(text)+media_html(media))
        current['media']=list(shared_media)+list(media)
        option_id=None
        if re.search(r'(?:jawaban\s+benar\s+lebih\s+dari\s+satu|jawaban\s+dapat\s+lebih\s+dari\s+satu)',text,re.I):
            current['type']='multiple'
        elif re.search(r'\b(?:tuliskan|jelaskan|uraikan|mengapa|setujukah|setujukan)\b',text,re.I) and not re.search(r'(?:pilih|centang)',text,re.I):
            current['type']='essay'

    for block_index,b in enumerate(items):
        if isinstance(b,Table):
            rows=[[c.text.strip() for c in r.cells] for r in b.rows]
            media=[]
            for row in b.rows:
                for cell in row.cells:
                    for par in cell.paragraphs:
                        media.extend(media_for_paragraph(par,block_index))
            if len(rows)>=3 and len(rows[0])>=3 and current:
                label_offset=2 if len(rows[0])>=4 else 1
                labels=[re.sub(r'\s+',' ',x).strip() for x in rows[0][label_offset:] if x.strip()]
                if len(labels)>=2:
                    current['type']='matrix'
                    current['rows']=[]
                    for row in rows[1:]:
                        statement=' '.join(x for x in row[:label_offset] if x and not re.fullmatch('[A-H][.]?',x))
                        if not statement.strip(): continue
                        answer=''
                        for j,mark in enumerate(row[label_offset:label_offset+len(labels)]):
                            if '√' in mark or '✓' in mark or mark.strip().upper() in ('V','X'):
                                answer=labels[j]
                        current['rows'].append({'text':statement,'choices':labels[:2],'answer':answer})
                    current['media'].extend(media)
                    continue
            if len(rows)>=3 and len(rows[0])==2 and current and ('PILIHAN' in rows[0][1].upper() or 'PERNYATAAN' in rows[0][0].upper()):
                current['type']='multiple'
                current['options']=[{'text':row[0],'correct':False} for row in rows[1:] if row[0].strip()]
                current['media'].extend(media)
                continue
            if rows and (len(rows)==1 or len(rows[0])==1):
                content='\n'.join(' '.join(x for x in row if x) for row in rows).strip()
                if content:
                    before.append(para_html(content)+media_html(media))
                    before_media.extend(media)
            elif current:
                current['notes'].append('Tabel kompleks terdeteksi; periksa sebelum ekspor.')
                current['media'].extend(media)
            continue
        txt=b.text.strip()
        if not current and re.match(r'^(?:SOAL\s+TO\.|(?:Senin|Selasa|Rabu|Kamis|Jumat|Sabtu|Minggu),\s*\d{1,2}\s+\w+\s+\d{4})',txt,re.I):
            continue
        media=media_for_paragraph(b,block_index)
        n=num_id(b)
        if not txt and not media: continue
        am=ANSWER_LINE.match(txt)
        if am and current:
            apply_answer(current,am.group(1))
            leftover=re.split(r'\n\s*(?=Perhatikan|Bacalah|Cermati)',txt,maxsplit=1,flags=re.I)
            if len(leftover)>1: before.append(para_html(leftover[1]))
            continue
        manual=Q_MARKER.match(txt)
        question_by_num=n in question_ids
        if question_by_num and (NOTE_WORDS.match(txt) or txt.lower().startswith('sumber:')):
            before.append(para_html(txt)+media_html(media))
            before_media.extend(media)
            continue
        if (not question_by_num and current and (current['options'] or current['rows']) and
            re.search(r'^(?:Berdasarkan penggalan tersebut, pilihlah pernyataan|Pernyataan berikut yang sesuai dengan isi paragraf)',txt,re.I)):
            start(txt,media); continue
        if question_by_num or (manual and len(manual.group(2).strip())>8):
            start(manual.group(2) if manual else txt,media); continue
        if current and n is not None and n not in question_ids and current['type']!='essay':
            if option_id is None or option_id==n:
                option_id=n
                current['options'].append({'text':PAREN_OPT.sub('',txt),'correct':False})
                continue
        mm=LETTER_OPT.match(txt)
        if current and mm and current['type']!='essay':
            current['options'].append({'text':mm.group(2).strip(),'correct':False}); continue
        if current and txt.startswith('[….]') and current['options']:
            current['options'].append({'text':PAREN_OPT.sub('',txt),'correct':False}); continue
        if before or NOTE_WORDS.match(txt):
            before.append(para_html(txt)+media_html(media))
            before_media.extend(media)
        elif current and not current['options'] and not current['rows']:
            current['stem']+=para_html(txt)+media_html(media)
            current['media'].extend(media)
        else:
            before.append(para_html(txt)+media_html(media))
            before_media.extend(media)
    if current: questions.append(validate(current))
    return {'source':filename,'questions':questions,'category':re.sub(r'\.[^.]+$','',filename),
            'warnings':['Word: deteksi soal berdasarkan penomoran dan struktur dokumen; tinjau hasil sebelum impor Moodle.',
                        'Soal bertabel/gambar dan kunci yang tidak tertulis jelas membutuhkan pemeriksaan manual.']}

def xtext(elem,path,default=''):
    x=elem.find(path)
    return x.text if x is not None and x.text is not None else default

def parse_moodle_xml(data,filename='bank.xml'):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise ConversionError('XML dengan DTD/ENTITY tidak diperbolehkan.')
    try:
        root=etree.fromstring(data,parser=etree.XMLParser(resolve_entities=False,no_network=True,huge_tree=False))
    except etree.XMLSyntaxError:
        raise ConversionError('XML tidak valid atau bukan dokumen Moodle.')
    if root.tag!='quiz': raise ConversionError('Akar XML harus <quiz> (format Moodle XML).')
    questions=[]
    category='Bank Soal'
    for node in root.findall('question'):
        typ=node.get('type','description')
        if typ=='category':
            category=xtext(node,'category/text',category).split('/')[-1]; continue
        if typ not in ('multichoice','essay','truefalse','cloze','description'): continue
        name=xtext(node,'name/text',f'Soal {len(questions)+1:02}')
        stem=xtext(node,'questiontext/text')
        qtype='essay' if typ=='essay' else 'cloze' if typ in ('cloze','description') else 'single'
        if typ=='multichoice' and xtext(node,'single','true')=='false': qtype='multiple'
        q=new_question(name,stem,qtype)
        q['grade']=float(xtext(node,'defaultgrade','1') or 1)
        q['feedback']=xtext(node,'generalfeedback/text')
        for ans in node.findall('answer'):
            val=float(ans.get('fraction','0'))
            q['options'].append({'text':strip_markup(xtext(ans,'text')),'correct':val>0})
        for f in node.findall('questiontext/file'):
            if f.get('encoding')=='base64':
                q['media'].append({'name':f.get('name','image.png'),'mime':'','data':(f.text or '').strip()})
        if qtype=='cloze': q['notes'].append('Soal Cloze dipertahankan dari XML; periksa jika menyunting sintaks {…}.')
        questions.append(validate(q))
    if not questions: raise ConversionError('Tidak ditemukan soal yang didukung di dalam XML.')
    return {'source':filename,'category':category,'questions':questions,
            'warnings':['Moodle XML terbaca. Tipe plugin di luar multichoice/essay/truefalse/cloze/description belum didukung.']}

def parse_csv(data,filename='questions.csv'):
    try: txt=data.decode('utf-8-sig')
    except UnicodeDecodeError: raise ConversionError('CSV harus ber-encoding UTF-8.')
    rdr=csv.DictReader(io.StringIO(txt))
    if not {'question','type','correct'}.issubset(set(rdr.fieldnames or [])):
        raise ConversionError('CSV memerlukan kolom question,type,correct dan option_a dst.')
    qs=[]
    for i,r in enumerate(rdr):
        typ=(r.get('type') or 'single').lower()
        if typ not in ('single','multiple','essay'):
            raise ConversionError(f'Tipe CSV baris {i+2} tidak dikenal: {typ}')
        q=new_question(r.get('name') or f'Soal {i+1:02}',para_html(r.get('question','')),typ)
        key=parse_key_text(r.get('correct',''))
        for j in range(8):
            val=(r.get('option_'+chr(97+j)) or '').strip()
            if val: q['options'].append({'text':val,'correct':chr(65+j) in key})
        qs.append(validate(q))
    return {'source':filename,'category':'Import CSV','questions':qs,
            'warnings':['Format CSV khusus Moodle Question Studio, bukan CSV bawaan Moodle.']}

def add_text(parent,tag,text,html_content=False,fmt=None):
    el=etree.SubElement(parent,tag)
    if fmt: el.set('format',fmt)
    t=etree.SubElement(el,'text')
    if html_content: t.text=etree.CDATA(str(text or '').replace(']]>',']]&gt;'))
    else: t.text=str(text or '')
    return el

def question_errors(q):
    typ=q.get('type')
    errors=[]
    if not strip_markup(q.get('stem')): errors.append('teks pertanyaan kosong')
    if typ in ('single','multiple','truefalse'):
        opts=q.get('options',[])
        n=sum(bool(x.get('correct')) for x in opts)
        if len(opts)<2: errors.append('minimal dua opsi')
        if n==0: errors.append('belum ada kunci jawaban')
        if typ=='single' and n!=1: errors.append('pilihan tunggal harus satu jawaban benar')
    if typ=='matrix' and (not q.get('rows') or any(not r.get('answer') for r in q['rows'])):
        errors.append('kunci tabel belum lengkap')
    if typ not in ('single','multiple','truefalse','essay','matrix','cloze'):
        errors.append('tipe tidak didukung')
    return errors

def moodle_xml(questions,category):
    quiz=etree.Element('quiz')
    cat=etree.SubElement(quiz,'question',type='category')
    add_text(cat,'category','$course$/'+category.strip().replace('/','-')[:100])
    for q in questions:
        typ=q.get('type','single')
        moodle_type={'single':'multichoice','multiple':'multichoice','truefalse':'multichoice',
                     'matrix':'cloze','essay':'essay','cloze':'cloze'}[typ]
        n=etree.SubElement(quiz,'question',type=moodle_type)
        add_text(n,'name',q.get('name','Soal'))
        stem=q.get('stem','')
        if typ=='matrix':
            parts=['<table border="1"><tbody><tr><th>Pernyataan</th><th>Jawaban</th></tr>']
            for row in q['rows']:
                choices=row.get('choices',['Benar','Salah'])
                def cl(x): return re.sub(r'([~=#:{}\\])',r'\\\1',x)
                code='{1:MULTICHOICE:'+'~'.join(('=' if x==row['answer'] else '')+cl(x) for x in choices)+'}'
                parts.append('<tr><td>'+escape(row.get('text',''))+'</td><td>'+code+'</td></tr>')
            stem+=''.join(parts)+'</tbody></table>'
        qt=add_text(n,'questiontext',stem,html_content=True,fmt='html')
        for img in q.get('media',[]):
            if not isinstance(img,dict) or not img.get('name') or not img.get('data'): continue
            name=re.sub(r'[^a-zA-Z0-9._-]','_',img['name'])[:120]
            etree.SubElement(qt,'file',name=name,path='/',encoding='base64').text=img['data']
        add_text(n,'generalfeedback',q.get('feedback',''),fmt='html',html_content=True)
        if typ!='cloze': etree.SubElement(n,'defaultgrade').text=str(float(q.get('grade',1)))
        etree.SubElement(n,'penalty').text='0.0000000'
        etree.SubElement(n,'hidden').text='0'
        if typ in ('single','multiple','truefalse'):
            etree.SubElement(n,'single').text='true' if typ!='multiple' else 'false'
            etree.SubElement(n,'shuffleanswers').text='false'
            etree.SubElement(n,'answernumbering').text='ABCD'
            good=sum(bool(x.get('correct')) for x in q['options'])
            for a in q['options']:
                frac=100/good if a.get('correct') else (-100 if typ=='multiple' else 0)
                ans=etree.SubElement(n,'answer',fraction=f'{frac:.5f}',format='html')
                etree.SubElement(ans,'text').text=etree.CDATA('<p>'+escape(a.get('text',''))+'</p>')
                add_text(ans,'feedback','',fmt='html',html_content=True)
        elif typ=='essay':
            etree.SubElement(n,'responseformat').text='editor'
            etree.SubElement(n,'responserequired').text='1'
            etree.SubElement(n,'responsefieldlines').text='10'
            etree.SubElement(n,'attachments').text='0'
    return etree.tostring(quiz,xml_declaration=True,encoding='UTF-8',pretty_print=True)

def gift_escape(text):
    return re.sub(r'([~=#{}:\\])',r'\\\1',strip_markup(text))

def gift(questions):
    lines=[]
    for q in questions:
        typ=q.get('type')
        if typ in ('matrix','cloze'):
            raise ConversionError('GIFT tidak mendukung tabel/Cloze pada versi ini. Gunakan Moodle XML.')
        name=gift_escape(q.get('name','Soal'))
        stem=gift_escape(q.get('stem',''))
        if typ=='essay': lines.append(f'::{name}::{stem} {{}}')
        elif typ in ('single','truefalse'):
            opts=[('=' if a.get('correct') else '~')+gift_escape(a.get('text','')) for a in q['options']]
            lines.append(f'::{name}::{stem} {{{" ".join(opts)}}}')
        elif typ=='multiple':
            good=sum(bool(a.get('correct')) for a in q['options'])
            opts=[('~%'+str(round(100/good,4))+'%' if a.get('correct') else '~%-100%')+gift_escape(a.get('text','')) for a in q['options']]
            lines.append(f'::{name}::{stem} {{{" ".join(opts)}}}')
    return ('\n\n'.join(lines)+'\n').encode('utf-8')

def aiken(questions):
    lines=[]
    for q in questions:
        if q.get('type') not in ('single','truefalse'):
            raise ConversionError('Aiken hanya untuk pilihan tunggal. Gunakan Moodle XML untuk tipe campuran.')
        lines.append(strip_markup(q['stem']))
        for i,a in enumerate(q['options']):
            lines.append(f'{chr(65+i)}. {strip_markup(a.get("text",""))}')
        correct=next(i for i,a in enumerate(q['options']) if a.get('correct'))
        lines.append('ANSWER: '+chr(65+correct)+'\n')
    return ('\n'.join(lines)+'\n').encode('utf-8')

def export_csv(questions):
    out=io.StringIO()
    fields=['name','type','question','correct']+[f'option_{chr(97+i)}' for i in range(8)]
    writer=csv.DictWriter(out,fieldnames=fields)
    writer.writeheader()
    for q in questions:
        row={'name':q.get('name',''),'type':q.get('type',''),'question':strip_markup(q.get('stem','')),
             'correct':','.join(chr(65+i) for i,a in enumerate(q.get('options',[])) if a.get('correct'))}
        for i,a in enumerate(q.get('options',[])[:8]): row[f'option_{chr(97+i)}']=strip_markup(a.get('text',''))
        writer.writerow(row)
    return ('\ufeff'+out.getvalue()).encode('utf-8')

def export_data(questions,fmt,category):
    if not questions: raise ConversionError('Belum ada soal untuk diekspor.')
    if fmt not in ('xml','gift','aiken','csv'): raise ConversionError('Format ekspor tidak didukung.')
    if fmt!='csv':
        for i,q in enumerate(questions):
            errs=question_errors(q)
            if errs: raise ConversionError(f'Soal #{i+1} ({q.get("name","Soal")}): '+', '.join(errs)+'. Lengkapi di editor.')
    if fmt=='xml': return moodle_xml(questions,category),'application/xml','xml'
    if fmt=='gift': return gift(questions),'text/plain; charset=utf-8','txt'
    if fmt=='aiken': return aiken(questions),'text/plain; charset=utf-8','txt'
    return export_csv(questions),'text/csv; charset=utf-8','csv'
