import streamlit as st, zipfile, io, re, csv
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from collections import defaultdict

st.set_page_config(page_title='COE Dummy Matcher', page_icon='🎓', layout='centered')
st.markdown('''<style>
.block-container{padding:1rem .7rem 3rem;max-width:1100px} .stButton>button,.stDownloadButton>button{width:100%;min-height:3rem;white-space:normal}
[data-testid="stFileUploader"]{border:1px solid #9aa4b2;border-radius:12px;padding:10px}
@media(max-width:600px){h1{font-size:1.65rem!important} h2{font-size:1.3rem!important}}
</style>''',unsafe_allow_html=True)
st.title('🎓 COE Dummy Matcher')
st.caption('Mobile-first • Recursive ZIP scanning • Multi-source reconciliation • Auditable downloads')
st.warning('Use institution-approved hosting for confidential student records. Test with copies before official release.')
CODE=re.compile(r'(?<![A-Z0-9])\d{3}[A-Z]{2,4}\d{2,3}(?![A-Z0-9])',re.I)
def code_from_name(name):
 m=CODE.search(name.upper()); return m.group(0).upper() if m else ''
def norm(v):
 if v is None:return ''
 s=str(v).strip()
 if s.endswith('.0') and s[:-2].isdigit():s=s[:-2]
 return re.sub(r'\s+','',s).upper()
def read_csv_bytes(b):
 for enc in ['utf-8-sig','utf-8','cp1252']:
  try:return pd.read_csv(io.BytesIO(b),dtype=str,encoding=enc).fillna('')
  except Exception:pass
 return pd.DataFrame()
def find_col(cols, terms):
 for c in cols:
  n=re.sub(r'[^a-z]','',str(c).lower())
  if any(t in n for t in terms):return c
 return None
def scan_zip(upload, is_source):
 z=zipfile.ZipFile(upload); records=[]; errors=[]; files=0
 for info in z.infolist():
  if info.is_dir():continue
  ext=info.filename.lower().rsplit('.',1)[-1] if '.' in info.filename else ''
  if ext not in (['csv','xlsx','xlsm'] if is_source else ['xlsx','xlsm']):continue
  files+=1; code=code_from_name(info.filename)
  try:
   raw=z.read(info)
   if ext=='csv':
    df=read_csv_bytes(raw)
    if df.empty:continue
    dcol=find_col(df.columns,['dummy','dummyid','dummycode']); rcol=find_col(df.columns,['registrationnumber','regno','registrationno'])
    ccol=find_col(df.columns,['subjectcode'])
    if ccol:
     for c in df[ccol].dropna().unique():
      if CODE.fullmatch(str(c).strip()): code=str(c).strip().upper(); break
    if dcol and rcol and code:
     for i,row in df.iterrows():
      d,r=norm(row[dcol]),norm(row[rcol])
      if d and r: records.append((code,d,r,info.filename,int(i)+2))
   else:
    wb=load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
    for ws in wb.worksheets:
     rows=ws.iter_rows(values_only=True); head=None; hrow=0
     for ri,row in enumerate(rows,1):
      vals=list(row); normheads=[re.sub(r'[^a-z]','',str(x or '').lower()) for x in vals]
      di=next((j for j,x in enumerate(normheads) if 'dummy' in x),None)
      rg=next((j for j,x in enumerate(normheads) if 'registrationnumber' in x or x in ('regno','registrationno')),None)
      if di is not None and rg is not None: head=(di,rg); hrow=ri; break
     if not head:continue
     di,rg=head
     for ri,row in enumerate(rows,hrow+1):
      vals=list(row); d=norm(vals[di] if di<len(vals) else ''); r=norm(vals[rg] if rg<len(vals) else '')
      if d and r and code:records.append((code,d,r,info.filename+'!'+ws.title,ri))
    wb.close()
  except Exception as e:errors.append((info.filename,str(e)))
 return records,files,errors

def target_code(name):return code_from_name(name)
st.header('1. Upload your archives')
source=st.file_uploader('Daywise source ZIP (nested folders supported)',type=['zip'],key='src')
target=st.file_uploader('Valuation workbooks ZIP',type=['zip'],key='tgt')
if source and target and st.button('🔎 Scan and reconcile',type='primary'):
 with st.spinner('Scanning every nested source file and reconciling mappings…'):
  rec,nfiles,errs=scan_zip(source,True)
  mapping=defaultdict(set); evidence=defaultdict(list)
  for code,d,r,path,row in rec:
   mapping[(code,d)].add(r); evidence[(code,d)].append(f'{path} row {row}')
  zin=zipfile.ZipFile(target); outputs=[]; audit=[]; status=[]; total=matched=conflict=unmatched=0; seen=set()
  for info in zin.infolist():
   if info.is_dir() or not info.filename.lower().endswith(('.xlsx','.xlsm')):continue
   code=target_code(info.filename)
   if not code:status.append({'Workbook':info.filename,'Subject code':'','Status':'Needs subject code review'});continue
   try:
    data=zin.read(info); wb=load_workbook(io.BytesIO(data),keep_vba=info.filename.lower().endswith('.xlsm'))
    sheet_hit=0; sheet_mapped=0
    for ws in wb.worksheets:
     if ws.title.strip().lower()=='completed':continue
     header_row=None; dcol=None
     for rr in range(1,min(ws.max_row,40)+1):
      for cc in range(1,ws.max_column+1):
       v=str(ws.cell(rr,cc).value or '').lower()
       if 'dummy' in v:header_row,dcol=rr,cc;break
      if dcol:break
     if not dcol:continue
     sheet_hit+=1; headers={str(ws.cell(header_row,c).value or '').strip().lower():c for c in range(1,ws.max_column+1)}
     regcol=next((c for k,c in headers.items() if k in ('reg.no','reg no','registration no','registration number')),None)
     editcol=next((c for k,c in headers.items() if k in ('reg edit','reg.no edited','reg no edited')),None)
     if not regcol:regcol=ws.max_column+1;ws.cell(header_row,regcol,'Reg.No')
     if not editcol:editcol=max(ws.max_column,regcol)+1;ws.cell(header_row,editcol,'Reg Edit')
     for rr in range(header_row+1,ws.max_row+1):
      d=norm(ws.cell(rr,dcol).value)
      if not d:continue
      total+=1; key=(code,d); regs=mapping.get(key,set()); state=''
      if len(regs)==1:
       r=next(iter(regs)); state='Matched'; matched+=1;sheet_mapped+=1
      elif len(regs)>1:r='';state='Conflict';conflict+=1
      else:r='';state='Unmatched';unmatched+=1
      ws.cell(rr,regcol,r);ws.cell(rr,editcol,r.split('-',1)[-1] if '-' in r else r)
      audit.append({'Workbook':info.filename,'Subject':code,'Sheet':ws.title,'Excel row':rr,'Dummy':d,'Reg.No':r,'Status':state,'Evidence':' | '.join(evidence.get(key,[])[:5])})
    if sheet_hit:
     b=io.BytesIO();wb.save(b); wb2=load_workbook(io.BytesIO(b.getvalue()),read_only=True,data_only=False); assert wb2.sheetnames; wb2.close()
     name=info.filename.rsplit('.',1)[0]+'_processed.'+info.filename.rsplit('.',1)[-1]
     if name in seen:name=f'{len(outputs)+1}_'+name
     seen.add(name);outputs.append((name,b.getvalue()))
     status.append({'Workbook':info.filename,'Subject code':code,'Status':'Processed' if sheet_mapped else 'Processed — no matches','Rows matched':sheet_mapped})
    else:status.append({'Workbook':info.filename,'Subject code':code,'Status':'No dummy column found'})
   except Exception as e:status.append({'Workbook':info.filename,'Subject code':code,'Status':'ERROR: '+str(e)})
  st.session_state['result']=(outputs,pd.DataFrame(audit),pd.DataFrame(status),nfiles,errs,(total,matched,conflict,unmatched),len(mapping))
if 'result' in st.session_state:
 outs,audit,stat,nfiles,errs,counts,mapcount=st.session_state['result'];total,matched,conflict,unmatched=counts
 st.header('2. Reconciliation results')
 a,b,c,d=st.columns(2);a.metric('Source files scanned',nfiles);b.metric('Mappings found',mapcount);c.metric('Rows matched',matched);d.metric('Conflicts',conflict)
 st.dataframe(stat,use_container_width=True,hide_index=True)
 st.dataframe(audit.head(500),use_container_width=True,hide_index=True)
 st.download_button('⬇️ Download row audit CSV',audit.to_csv(index=False).encode('utf-8-sig'),'COE_row_audit.csv','text/csv')
 st.download_button('⬇️ Download subject status CSV',stat.to_csv(index=False).encode('utf-8-sig'),'COE_subject_status.csv','text/csv')
 if outs:
  zbuf=io.BytesIO()
  with zipfile.ZipFile(zbuf,'w',zipfile.ZIP_DEFLATED) as z:
   for name,data in outs:z.writestr(name,data)
  zbuf.seek(0)
  st.success(f'{len(outs)} workbook(s) saved and reopened successfully. {matched}/{total} rows matched; {unmatched} unmatched; {conflict} conflicts.')
  st.download_button('⬇️ Download processed workbooks ZIP',zbuf.getvalue(),'COE_processed_workbooks.zip','application/zip',type='primary')
 else:st.error('No valid workbook was generated. No empty ZIP download is offered.')
 if errs:
  with st.expander(f'{len(errs)} source file read errors'):st.dataframe(pd.DataFrame(errs,columns=['File','Error']))
