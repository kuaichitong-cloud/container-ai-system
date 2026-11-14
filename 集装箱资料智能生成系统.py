# -*- coding: utf-8 -*-
"""集装箱资料智能生成系统 - 完整版（Windows）
ZIP批量导入、OCR箱号/字段提取、照片自动分类、人工复核、套用DOC/DOCX模板生成逐箱DOCX。
业务数值只从可识别文字中提取；不猜测缺失的件数/重量等数据。
"""
import os, re, sys, zipfile, shutil, tempfile, subprocess, threading, json, csv
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageOps
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

APP_TITLE='集装箱资料智能生成系统'
IMG_EXT={'.jpg','.jpeg','.png','.bmp','.tif','.tiff','.webp'}
CATS=['箱号/箱门','空箱','1/3装箱','2/3装箱','满箱','铅封','货物实物','其他']
FIELDS=[('container','集装箱箱号'),('from_station','发站'),('to_station','到站'),('overseas','境外到站'),('mixed','是否混装'),('shipper','托运人'),('sender','发货人'),('goods','货物名称'),('quantity','货物件数'),('weight','货物总重量'),('max_weight','最大件重'),('max_size','最大件重尺寸'),('box_type','自备箱型/类')]
BOX_RE=re.compile(r'(?<![A-Z0-9])([A-Z]{4}\s?\d{7})(?!\d)',re.I)

def clean_box(s):
    m=BOX_RE.search((s or '').upper().replace(' ','')); return m.group(1).upper() if m else ''

OCR_LAST_ERROR = ''


def resolve_tesseract():
    """Find bundled or commonly installed Tesseract and its language data."""
    roots = []
    app_root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    roots.append(app_root / 'tesseract')
    for env_name in ('ProgramFiles', 'ProgramFiles(x86)'):
        base = os.environ.get(env_name)
        if base:
            roots.append(Path(base) / 'Tesseract-OCR')
    found = shutil.which('tesseract')
    if found:
        roots.append(Path(found).resolve().parent)

    checked = []
    for folder in roots:
        exe = folder / 'tesseract.exe'
        data = folder / 'tessdata'
        checked.append(str(folder))
        if exe.is_file() and (data / 'chi_sim.traineddata').is_file() and (data / 'eng.traineddata').is_file():
            return exe, data
    raise RuntimeError(
        '找不到可用的 Tesseract OCR 或语言包。已检查：' + '；'.join(checked) +
        '。请确认 C:\\Program Files\\Tesseract-OCR\\tesseract.exe、'
        'tessdata\\chi_sim.traineddata 和 tessdata\\eng.traineddata 均存在。'
    )


def ocr_image(path):
    global OCR_LAST_ERROR
    try:
        import pytesseract
        exe, tessdata = resolve_tesseract()
        pytesseract.pytesseract.tesseract_cmd = str(exe)
        im = Image.open(path).convert('RGB')
        if max(im.size) < 1800:
            im = im.resize((im.width * 2, im.height * 2))
        config = '--psm 11 --tessdata-dir "' + str(tessdata) + '"'
        return pytesseract.image_to_string(im, lang='chi_sim+eng', config=config)
    except Exception as e:
        if not OCR_LAST_ERROR:
            OCR_LAST_ERROR = str(e)
        return ''

def classify(text, name):
    s=(text+' '+name).lower()
    rules=[(('seal','铅封','封条'),'铅封'),(('empty','空箱','tare'),'空箱'),(('1/3','1／3','1-3','三分之一'),'1/3装箱'),(('2/3','2／3','2-3','三分之二'),'2/3装箱'),(('满箱','full load','fully loaded'),'满箱'),(('箱号','集装箱号码','container no','container number','foru'),'箱号/箱门')]
    for keys,cat in rules:
        if any(k in s for k in keys): return cat
    # 图片内容通常无字，给出候选而不假装已准确识别
    return '其他'

def extract_fields(text):
    out={k:'' for k,_ in FIELDS}
    box=clean_box(text)
    if box: out['container']=box
    patterns={
      'from_station':r'发站\s*[:：]?\s*([^\s,，]+)', 'to_station':r'到站\s*[:：]?\s*([^\s,，]+)',
      'overseas':r'境外到站\s*[:：]?\s*([^\s,，]+)', 'mixed':r'是否混装\s*[:：]?\s*(是|否)',
      'shipper':r'托运人\s*[:：]?\s*(.{3,35}?)(?=发货人|货物名称|\n|$)',
      'sender':r'发货人\s*[:：]?\s*(.{3,35}?)(?=货物名称|货物件数|\n|$)',
      'goods':r'货物名称\s*[:：]?\s*(.{2,40}?)(?=货物件数|货物总重量|\n|$)',
      'quantity':r'货物件数\s*[:：]?\s*([\d,，]+\s*(?:件|箱|台|套|托)?)',
      'weight':r'货物总重量\s*[:：]?\s*([\d,.，]+\s*(?:KG|kg|吨|T)?)',
      'max_weight':r'最大件重\s*[:：]?\s*([\d,.，]+\s*(?:KG|kg)?)',
      'max_size':r'最大件重尺寸\s*[:：]?\s*([\d*×xX.\s]+\s*(?:mm|MM|cm|CM)?)',
      'box_type':r'(?:自备箱型/类|箱型)\s*[:：]?\s*([A-Z0-9-]+)'}
    for k,p in patterns.items():
        m=re.search(p,text,re.I)
        if m: out[k]=re.sub(r'\s+',' ',m.group(1)).strip(' :：')
    return out

def convert_template(path, work):
    p=Path(path)
    if p.suffix.lower()=='.docx': return str(p)
    # LibreOffice first, Microsoft Word COM on Windows fallback
    out=Path(work)/ (p.stem+'.docx')
    exe=shutil.which('soffice') or shutil.which('libreoffice')
    if exe:
        subprocess.run([exe,'--headless','--convert-to','docx','--outdir',str(work),str(p)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
        if out.exists(): return str(out)
    if os.name=='nt':
        try:
            import win32com.client
            word=win32com.client.Dispatch('Word.Application'); word.Visible=False
            doc=word.Documents.Open(str(p.resolve())); doc.SaveAs2(str(out.resolve()),FileFormat=16); doc.Close(); word.Quit()
            return str(out)
        except Exception as e: raise RuntimeError('旧版DOC转DOCX失败。请用Word打开模板并另存为DOCX。\n'+str(e))
    raise RuntimeError('无法转换旧版DOC模板。请先用Word另存为DOCX，或安装LibreOffice。')

def replace_paragraph_text(paragraph, replacements):
    # 保留段落格式和首个run样式
    old=paragraph.text
    new=old
    for a,b in replacements.items(): new=new.replace(a,b)
    if new!=old:
        if paragraph.runs:
            paragraph.runs[0].text=new
            for r in paragraph.runs[1:]: r.text=''
        else: paragraph.add_run(new)

def add_heading_compat(doc, text, level=1):
    p=doc.add_paragraph(); r=p.add_run(text); r.bold=True; r.font.size=__import__('docx').shared.Pt(14 if level==1 else 12); return p

def make_docx(template, data, photos, output, tempdir):
    doc=Document(template)
    # Original sample-template values are replaced in-place; blank recognized values clear sample data.
    repl={'TEFU4028297':data.get('container',''), '全自动滚筒洗衣机':data.get('goods',''),
          '铜山':data.get('from_station',''), '二连':data.get('to_station',''), '俄罗斯':data.get('overseas',''),
          '江苏省国际货运班列有限公司':data.get('shipper',''), '251':data.get('quantity',''),
          '发货人原模板占位':data.get('sender',''),
          '13774.88 KG':data.get('weight',''), '76KG':data.get('max_weight',''),
          '705*675*885mm':data.get('max_size',''), '40HC':data.get('box_type','')}
    # Replace text while retaining layout, then replace known template photo blobs in document order
    for p in doc.paragraphs: replace_paragraph_text(p,repl)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: replace_paragraph_text(p,repl)
    # Update photo placeholder media, preserving exact inline placement/size in template
    media_dir=Path(tempdir)/'media'; media_dir.mkdir(exist_ok=True)
    ordered=[]
    for cat in CATS:
        ordered.extend([(p,c) for p,c in photos if c==cat])
    try:
        rels=[]
        for rel in doc.part.rels.values():
            if 'image' in rel.reltype: rels.append(rel)
        # Replace embedded pictures by the matching ordered photos; image count beyond placeholders appended below.
        for i,rel in enumerate(rels):
            if i>=len(ordered): break
            blob=Path(ordered[i][0]).read_bytes()
            part=rel.target_part
            part._blob=blob
    except Exception: pass
    # Update visible text placeholders / container number wherever split across runs
    for p in doc.paragraphs:
        if '集装箱号码' in p.text or '集装箱箱号' in p.text:
            replace_paragraph_text(p,{'TEFU4028297':data.get('container','')})
    # Add photos beyond the template's embedded photo slots, grouped with clear labels
    existing_count=0
    try: existing_count=sum(1 for r in doc.part.rels.values() if 'image' in r.reltype)
    except: pass
    if len(ordered)>existing_count:
        doc.add_page_break(); add_heading_compat(doc,'补充装箱照片',1)
        for path,cat in ordered[existing_count:]:
            add_heading_compat(doc,cat,2)
            try:
                im=Image.open(path); im.thumbnail((1600,1600)); tmp=Path(tempdir)/(Path(path).stem+'.jpg'); im.convert('RGB').save(tmp,quality=90)
                doc.add_picture(str(tmp),width=Inches(5.8))
                doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
            except Exception: continue
    # Append machine-read data audit sheet only for missing/ambiguous fields, keep source template untouched otherwise
    missing=[label for key,label in FIELDS if not data.get(key,'').strip()]
    if missing:
        doc.add_page_break(); add_heading_compat(doc,'自动识别结果核对（待补充字段）',1)
        doc.add_paragraph('以下字段未能从图片中可靠识别，请依据业务单据核实后补充：'+'、'.join(missing))
    doc.save(output)

class App:
 def __init__(self,root):
    self.root=root; root.title(APP_TITLE+' - 完整版'); root.geometry('1250x820'); root.minsize(1050,700)
    self.work=Path(tempfile.mkdtemp(prefix='container_ai_')); self.items=[]; self.groups={}; self.template=''; self.preview=None; self.vars={}; self.build()
 def build(self):
    top=ttk.Frame(self.root,padding=8); top.pack(fill='x')
    ttk.Label(top,text=APP_TITLE+'（完整版）',font=('Microsoft YaHei',18,'bold')).pack(side='left')
    ttk.Button(top,text='导入ZIP压缩包（支持批量）',command=self.import_zip).pack(side='right',padx=4)
    ttk.Button(top,text='选择Word模板（DOC/DOCX）',command=self.pick_template).pack(side='right',padx=4)
    body=ttk.Panedwindow(self.root,orient='horizontal'); body.pack(fill='both',expand=True,padx=8,pady=6)
    left=ttk.Frame(body,padding=6); right=ttk.Frame(body,padding=8); body.add(left,weight=2); body.add(right,weight=3)
    ttk.Label(left,text='识别照片清单（可逐张复核）',font=('Microsoft YaHei',11,'bold')).pack(anchor='w')
    self.list=tk.Listbox(left,width=48,height=30); self.list.pack(fill='both',expand=True); self.list.bind('<<ListboxSelect>>',self.select)
    nav=ttk.Frame(left); nav.pack(fill='x',pady=5)
    ttk.Button(nav,text='重新识别全部',command=self.reocr).pack(side='left'); ttk.Button(nav,text='导出识别报告',command=self.report).pack(side='right')
    self.pic=ttk.Label(right,text='先导入ZIP压缩包',anchor='center'); self.pic.pack(fill='both',expand=True)
    bar=ttk.Frame(right); bar.pack(fill='x',pady=4)
    ttk.Label(bar,text='照片类别').pack(side='left'); self.cat=tk.StringVar(value='其他'); ttk.Combobox(bar,textvariable=self.cat,values=CATS,state='readonly',width=16).pack(side='left',padx=5)
    ttk.Button(bar,text='保存当前分类',command=self.save_cat).pack(side='left')
    ttk.Label(right,text='集装箱资料（自动提取后请核对；不确定的数值不会自动编造）',font=('Microsoft YaHei',11,'bold')).pack(anchor='w',pady=(8,2))
    frm=ttk.Frame(right); frm.pack(fill='x')
    for i,(k,label) in enumerate(FIELDS):
      r=i//2;c=(i%2)*2; ttk.Label(frm,text=label).grid(row=r,column=c,sticky='w',padx=3,pady=3)
      v=tk.StringVar(); self.vars[k]=v; ttk.Entry(frm,textvariable=v,width=30).grid(row=r,column=c+1,sticky='ew',padx=3,pady=3)
    frm.columnconfigure(1,weight=1);frm.columnconfigure(3,weight=1)
    bottom=ttk.Frame(right);bottom.pack(fill='x',pady=10)
    ttk.Button(bottom,text='按箱号批量生成Word',command=self.generate_all).pack(side='right')
    self.status=tk.StringVar(value='等待导入ZIP文件'); ttk.Label(bottom,textvariable=self.status).pack(side='left')
 def pick_template(self):
    p=filedialog.askopenfilename(filetypes=[('Word模板','*.doc *.docx')]);
    if p:self.template=p; self.status.set('模板：'+Path(p).name)
 def import_zip(self):
    paths=filedialog.askopenfilenames(filetypes=[('ZIP压缩包','*.zip')]);
    if not paths:return
    self.items=[]; self.groups={}; self.list.delete(0,'end')
    for zi,zp in enumerate(paths):
      dest=self.work/f'zip_{zi}';dest.mkdir(exist_ok=True)
      try:
       with zipfile.ZipFile(zp) as z:
        for info in z.infolist():
         if info.is_dir():continue
         name=Path(info.filename)
         if name.suffix.lower() not in IMG_EXT:continue
         # defend against zip-slip
         target=(dest/f'{len(self.items):05d}_{name.name}'); target.parent.mkdir(parents=True,exist_ok=True)
         with z.open(info) as src,open(target,'wb') as out: shutil.copyfileobj(src,out)
         self.items.append({'path':str(target),'name':info.filename,'text':'','cat':'其他','box':''})
      except Exception as e: messagebox.showerror('压缩包读取失败',f'{zp}\n{e}')
    self.status.set(f'已导入 {len(self.items)} 张图片，正在OCR识别…')
    threading.Thread(target=self.scan,daemon=True).start()
 def scan(self):
    global OCR_LAST_ERROR
    OCR_LAST_ERROR = ''
    for i,item in enumerate(self.items):
      item['text']=ocr_image(item['path']); item['box']=clean_box(item['text'])
      item['cat']=classify(item['text'],item['name'])
      # container folder or archive parent as secondary clue
      if not item['box']:
       b=clean_box(item['name']); item['box']=b
      if not item['box']:
       b=clean_box(Path(item['name']).parent.name); item['box']=b
      if not item['box']: item['box']='未识别箱号'
      self.root.after(0,self.refresh_row,i)
    self.root.after(0,self.finish_scan)
 def refresh_row(self,i):
    if i<len(self.items):
      it=self.items[i]; self.list.insert('end',f"[{it['box']}] [{it['cat']}] {Path(it['name']).name}")
 def finish_scan(self):
    # Create groups; prefill current form from OCR of the first group
    self.groups={}
    for it in self.items:self.groups.setdefault(it['box'],[]).append(it)
    if self.items:
      merged='\n'.join(x['text'] for x in self.items)
      data=extract_fields(merged)
      boxes=[b for b in self.groups if b!='未识别箱号']
      if len(boxes)==1:data['container']=boxes[0]
      for k,v in data.items():
       if v:self.vars[k].set(v)
      self.list.selection_set(0);self.select()
    if OCR_LAST_ERROR:
      self.status.set(f'OCR异常：{len(self.items)}张图片未能正常识别。{OCR_LAST_ERROR[:180]}')
    else:
      self.status.set(f'OCR完成：{len(self.items)}张图片，识别到 {len(self.groups)} 个箱号分组（含未识别组）')
 def select(self,event=None):
    sel=self.list.curselection()
    if not sel:return
    it=self.items[sel[0]];self.cat.set(it['cat'])
    try:
      im=Image.open(it['path']); im.thumbnail((820,390)); self.preview=ImageTk.PhotoImage(im); self.pic.configure(image=self.preview,text='')
    except Exception:self.pic.configure(image='',text='无法预览图片')
 def save_cat(self):
    sel=self.list.curselection()
    if not sel:return
    it=self.items[sel[0]];it['cat']=self.cat.get();self.list.delete(sel[0]);self.list.insert(sel[0],f"[{it['box']}] [{it['cat']}] {Path(it['name']).name}");self.list.selection_set(sel[0])
 def reocr(self):
    if not self.items:return
    self.status.set('重新识别中…');threading.Thread(target=self.scan,daemon=True).start()
 def report(self):
    if not self.items:return
    p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')],initialfile='集装箱识别报告.csv')
    if not p:return
    with open(p,'w',newline='',encoding='utf-8-sig') as f:
      w=csv.writer(f);w.writerow(['箱号分组','图片文件','自动分类','OCR识别文字'])
      for x in self.items:w.writerow([x['box'],x['name'],x['cat'],x['text']])
    messagebox.showinfo('完成','识别报告已导出')
 def generate_all(self):
    if not self.items:messagebox.showwarning('缺少资料','请先导入ZIP');return
    if not self.template:messagebox.showwarning('缺少模板','请先选择原始Word模板');return
    out= filedialog.askdirectory(title='选择Word输出文件夹')
    if not out:return
    try: template=convert_template(self.template,self.work)
    except Exception as e:messagebox.showerror('模板处理失败',str(e));return
    globaldata={k:v.get().strip() for k,v in self.vars.items()}
    made=[]
    for box,items in self.groups.items():
      # Multi-container batches must never inherit another container's cargo fields.
      data=dict(globaldata) if len(self.groups)==1 else {k:'' for k,_ in FIELDS}
      data['container']=box if box!='未识别箱号' else (data.get('container') or globaldata.get('container',''))
      # OCR each group independently; per-container text takes precedence.
      gd=extract_fields('\n'.join(x['text'] for x in items))
      for k,v in gd.items():
       if v: data[k]=v
      if not data.get('container'):continue
      photos=[(x['path'],x['cat']) for x in items]
      dest=Path(out)/(data['container']+'.docx')
      try:make_docx(template,data,photos,str(dest),self.work);made.append(str(dest))
      except Exception as e:messagebox.showerror('生成失败',f'{box}\n{e}')
    if made:messagebox.showinfo('生成完成',f'已生成 {len(made)} 份Word文件：\n'+'\n'.join(made[:12]))
    else:messagebox.showwarning('未生成','没有有效箱号。请在资料栏补充箱号或核对OCR结果。')

if __name__=='__main__':
 root=tk.Tk();App(root);root.mainloop()
