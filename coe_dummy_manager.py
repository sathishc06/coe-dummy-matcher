import io, re, zipfile, copy
from pathlib import PurePosixPath
from collections import defaultdict
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

st.set_page_config(page_title='COE Subject-wise Dummy Matcher', page_icon='🎓', layout='wide')
st.title('🎓 COE Subject-wise Dummy Number Matcher')
st.caption('Upload the original Valuation May 2026 ZIP and Daywise ZIP. The app matches Dummy No. → complete Reg.No. using subject context and records exact source evidence.')

EXCEL_EXT={'.xlsx','.xlsm'}
DATA_EXT=EXCEL_EXT|{'.csv'}
DUMMY_ALIASES=['Dummy No','Dummy Number','Dummy','Dummy Reg No','Dummy ID']
REG_ALIASES=['Registration_Number','Registration Number','Registration No','Reg.No','Reg No','Register Number','Register No']
SUBJECT_ALIASES=['Subject_Code','Subject Code','Sub Code','Subject']


def norm(v):
    if v is None: return ''
    s=str(v).strip()
    if s.lower() in {'nan','none'}: return ''
    if s.endswith('.0') and s[:-2].isdigit(): s=s[:-2]
    return re.sub(r'[^A-Z0-9]+','',s.upper())


def path_excluded(path):
    p=path.replace('\\','/').lower()
    parts=[x for x in p.split('/') if x]
    return any(x.strip().lower() in {'completed','marks to department','marks to department folder'} for x in parts)


def zip_files(upload):
    out=[]
    with zipfile.ZipFile(upload) as z:
        for info in z.infolist():
            if info.is_dir() or info.filename.startswith('__MACOSX/') or '/.' in info.filename: continue
            if PurePosixPath(info.filename).suffix.lower() in DATA_EXT:
                out.append((info.filename,z.read(info.filename)))
    return out


def find_header(ws, aliases, scan=40):
    aliases=[norm(x) for x in aliases]
    for r in range(1,min(ws.max_row,scan)+1):
        vals={norm(ws.cell(r,c).value):c for c in range(1,ws.max_column+1) if ws.cell(r,c).value is not None}
        for a in aliases:
            if a in vals: return r,vals[a]
        for k,c in vals.items():
            if k and any(len(a)>=5 and a in k for a in aliases): return r,c
    return None,None


def find_pair(ws, scan=40):
    for r in range(1,min(ws.max_row,scan)+1):
        vals={norm(ws.cell(r,c).value):c for c in range(1,ws.max_column+1) if ws.cell(r,c).value is not None}
        def pick(aliases):
            aa=[norm(x) for x in aliases]
            for a in aa:
                if a in vals: return vals[a]
            for k,c in vals.items():
                if k and any(len(a)>=5 and a in k for a in aa): return c
            return None
        d,rcol=pick(DUMMY_ALIASES),pick(REG_ALIASES)
        if d and rcol: return r,d,rcol
    return None,None,None


def extract_code_from_text(text):
    # Codes used by the institution are typically 6-8 alphanumeric chars such as 225PPI05.
    vals=re.findall(r'(?<![A-Z0-9])\d{3}[A-Z]{2,4}\d{2,3}(?![A-Z0-9])',str(text).upper())
    return [norm(x) for x in vals]


def valuation_subject_code(ws, path):
    # Prefer the actual "Sub. Code & Sub. Name" cell in the workbook.
    for r in range(1,min(ws.max_row,15)+1):
        row=[ws.cell(r,c).value for c in range(1,min(ws.max_column,8)+1)]
        if any('SUB. CODE' in str(v).upper() for v in row if v is not None):
            for v in row:
                codes=extract_code_from_text(v)
                if codes: return codes[0]
    codes=extract_code_from_text(path)
    return codes[0] if codes else ''


def source_codes_from_row(row, path):
    vals=[]
    for col in ['Subject_Code','Subject Code','Sub Code']:
        if col in row.index:
            vals.extend(extract_code_from_text(row[col]))
    # Filename can contain a primary code and an additional integrated code.
    vals.extend(extract_code_from_text(path))
    return sorted(set(vals))


def read_daywise_sources(files):
    # index[(subject_code,dummy)] -> list of {reg,path,row,generated}
    idx=defaultdict(list); scan=[]; source_count=0
    for path,data in files:
        if path_excluded(path):
            scan.append({'Source':path,'Status':'Skipped: Completed/Marks to Department'}); continue
        ext=PurePosixPath(path).suffix.lower()
        try:
            if ext=='.csv':
                df=pd.read_csv(io.BytesIO(data),dtype=str,keep_default_na=False)
                cols={norm(c):c for c in df.columns}
                dc=next((cols[norm(x)] for x in ['Dummy_ID','Dummy No','Dummy Number'] if norm(x) in cols),None)
                rc=next((cols[norm(x)] for x in ['Registration_Number','Registration Number','Registration No','Reg.No'] if norm(x) in cols),None)
                sc=next((cols[norm(x)] for x in ['Subject_Code','Subject Code','Sub Code'] if norm(x) in cols),None)
                if not (dc and rc and sc):
                    scan.append({'Source':path,'Status':'Skipped: required Dummy/Reg.No/Subject_Code columns not found'}); continue
                count=0
                for i,row in df.iterrows():
                    d=norm(row[dc]); reg=str(row[rc]).strip(); codes=extract_code_from_text(row[sc])
                    if not d or not reg or not codes: continue
                    for code in codes:
                        idx[(code,d)].append({'reg':reg,'source':path,'location':f'{path} | CSV row {i+2}','generated':str(row.get('Generated_Date',''))})
                    count+=1
                source_count+=1; scan.append({'Source':path,'Status':'Indexed','Rows':count,'Type':'CSV'})
            elif ext in EXCEL_EXT:
                wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True)
                total=0
                for ws in wb.worksheets:
                    hr,dc,rc=find_pair(ws)
                    if not hr: continue
                    # Subject code may be in a Subject_Code column or workbook cells/path.
                    sh_code=''
                    for r in range(1,min(ws.max_row,20)+1):
                        for c in range(1,min(ws.max_column,20)+1):
                            sh_code=(extract_code_from_text(ws.cell(r,c).value) or [sh_code])[0]
                            if sh_code: break
                        if sh_code: break
                    sh_code=sh_code or (extract_code_from_text(path) or [''])[0]
                    for r in range(hr+1,ws.max_row+1):
                        d=norm(ws.cell(r,dc).value); reg=ws.cell(r,rc).value
                        if not d or reg is None or not str(reg).strip() or not sh_code: continue
                        idx[(sh_code,d)].append({'reg':str(reg).strip(),'source':path,'location':f'{path} | {ws.title} | row {r}','generated':''}); total+=1
                wb.close(); source_count+=1; scan.append({'Source':path,'Status':'Indexed','Rows':total,'Type':'Excel'})
        except Exception as e:
            scan.append({'Source':path,'Status':f'ERROR: {e}'})
    return idx,pd.DataFrame(scan),source_count


def reg_edit(reg):
    return reg.split('-',1)[1] if '-' in reg else ''


def copy_row(ws, src, dst, maxc):
    for c in range(1,maxc+1):
        s=ws.cell(src,c); d=ws.cell(dst,c); val=s.value
        if isinstance(val,str) and val.startswith('=') and src!=dst:
            try: val=Translator(val,origin=s.coordinate).translate_formula(d.coordinate)
            except Exception: pass
        d.value=val
        if s.has_style: d._style=s._style
        d.number_format=s.number_format
        d.alignment=s.alignment
        d.protection=s.protection
        if s.hyperlink: d._hyperlink=s.hyperlink
        if s.comment: d.comment=s.comment


def process_valuation(files, idx):
    outputs=[]; audit=[]; subject_log=[]; skipped=0
    for path,data in files:
        if path_excluded(path):
            skipped+=1; continue
        if PurePosixPath(path).suffix.lower() not in EXCEL_EXT: continue
        try:
            wb=load_workbook(io.BytesIO(data),keep_vba=PurePosixPath(path).suffix.lower()=='.xlsm')
            workbook_has_subject=False
            for ws in wb.worksheets:
                if ws.title.strip().lower()=='completed':
                    audit.append({'Valuation workbook':path,'Sheet':ws.title,'Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':'Skipped: Completed sheet','Source evidence':''}); continue
                # Explicitly skip sheets/folders intended for departmental distribution.
                if 'MARKS TO DEPARTMENT' in ws.title.upper():
                    audit.append({'Valuation workbook':path,'Sheet':ws.title,'Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':'Skipped: Marks to Department sheet','Source evidence':''}); continue
                hr,dc=find_header(ws,DUMMY_ALIASES)
                if not hr: continue
                code=valuation_subject_code(ws,path)
                if not code:
                    audit.append({'Valuation workbook':path,'Sheet':ws.title,'Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':'BLOCKED: subject code not detected','Source evidence':''}); continue
                workbook_has_subject=True
                # Add/replace Reg.No and Reg Edit immediately after the dummy column.
                headers={norm(ws.cell(hr,c).value):c for c in range(1,ws.max_column+1) if ws.cell(hr,c).value is not None}
                regcol=headers.get(norm('Reg.No'))
                if not regcol: regcol=dc+1; ws.insert_cols(regcol,1)
                ws.cell(hr,regcol).value='Reg.No'
                editcol=headers.get(norm('Reg Edit'))
                if editcol and editcol==regcol: editcol=regcol+1
                if not editcol: editcol=regcol+1; ws.insert_cols(editcol,1)
                ws.cell(hr,editcol).value='Reg Edit'
                for c in (regcol,editcol):
                    src=ws.cell(hr,dc); dst=ws.cell(hr,c)
                    if src.has_style: dst._style=src._style
                rows=[]
                for r in range(hr+1,ws.max_row+1):
                    d=norm(ws.cell(r,dc).value)
                    if not d: continue
                    hits=idx.get((code,d),[])
                    regs=sorted(set(h['reg'] for h in hits))
                    if len(regs)==1:
                        full=regs[0]; edit=reg_edit(full); status='Matched'
                    elif len(regs)>1:
                        full=edit=''; status='Conflict: same subject+dummy maps to multiple Reg.No values'
                    else:
                        full=edit=''; status='Unmatched: no verified subject+dummy source record'
                    ws.cell(r,regcol).value=full; ws.cell(r,editcol).value=edit
                    audit.append({'Valuation workbook':path,'Sheet':ws.title,'Excel row':r,'Dummy No.':d,'Reg.No':full,'Reg Edit':edit,'Status':status,'Source evidence':'; '.join(sorted(set(h['location'] for h in hits)))})
                    rows.append(r)
                # Sort all populated data rows by Reg.No, blanks last; formulas are translated.
                if rows:
                    start=hr+1; end=ws.max_row; maxc=ws.max_column
                    payload=[]
                    for r in range(start,end+1):
                        key=ws.cell(r,regcol).value or ''
                        payload.append((key,r))
                    payload.sort(key=lambda x:(x[0]=='',str(x[0]).upper(),x[1]))
                    # Snapshot first, then write in sorted order.
                    snapshots={r:[(ws.cell(r,c).value,copy.copy(ws.cell(r,c)._style),ws.cell(r,c).number_format,copy.copy(ws.cell(r,c).alignment),copy.copy(ws.cell(r,c).protection),ws.cell(r,c).hyperlink,copy.copy(ws.cell(r,c).comment)) for c in range(1,maxc+1)] for _,r in payload}
                    for newr,(_,oldr) in enumerate(payload,start):
                        for c,(val,style,nfmt,align,prot,hyper,comment) in enumerate(snapshots[oldr],1):
                            dest=ws.cell(newr,c)
                            if isinstance(val,str) and val.startswith('=') and oldr!=newr:
                                try: val=Translator(val,origin=ws.cell(oldr,c).coordinate).translate_formula(dest.coordinate)
                                except Exception: pass
                            dest.value=val
                            if style: dest._style=style
                            dest.number_format=nfmt; dest.alignment=align; dest.protection=prot
                            if hyper: dest._hyperlink=hyper
                            if comment: dest.comment=comment
            if workbook_has_subject:
                out=io.BytesIO(); wb.save(out)
                name=PurePosixPath(path).name
                outputs.append((path,name.rsplit('.',1)[0]+'_processed.'+name.rsplit('.',1)[-1],out.getvalue()))
                subject_log.append({'Valuation workbook':path,'Subject code':code,'Status':'Processed'})
            else:
                skipped+=1
            wb.close()
        except Exception as e:
            audit.append({'Valuation workbook':path,'Sheet':'','Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':f'ERROR: {e}','Source evidence':''})
    return outputs,pd.DataFrame(audit),pd.DataFrame(subject_log),skipped


def run(val_files,day_files):
    idx,scan,source_count=read_daywise_sources(day_files)
    outputs,audit,subjects,skipped=process_valuation(val_files,idx)
    return outputs,audit,scan,subjects,source_count,skipped

with st.sidebar:
    st.header('1. Upload original archives')
    val_zip=st.file_uploader('Valuation MAY 2026.zip',type=['zip'])
    day_zip=st.file_uploader('Daywise.zip',type=['zip'])
    st.markdown('**Rules:** Completed and Marks to Department are excluded. No filename-only guessing. A match requires subject code + dummy number in Daywise data.')

if not val_zip or not day_zip:
    st.info('Upload both ZIP files. Then press **Start matching**.')
    st.stop()

if st.button('🔎 Start matching',type='primary',use_container_width=True):
    try:
        vf=zip_files(val_zip); df=zip_files(day_zip)
        st.session_state['run']=run(vf,df)
    except Exception as e:
        st.error(f'Processing failed: {e}')

if 'run' in st.session_state:
    outputs,audit,scan,subjects,source_count,skipped=st.session_state['run']
    status=audit.get('Status',pd.Series(dtype=str)).astype(str)
    matched=int((status=='Matched').sum()); conflicts=int(status.str.startswith('Conflict').sum()); unmatched=int(status.str.startswith('Unmatched').sum()); errors=int(status.str.startswith('ERROR').sum()); blocked=int(status.str.startswith('BLOCKED').sum())
    a,b,c,d,e=st.columns(5); a.metric('Matched',matched); b.metric('Unmatched',unmatched); c.metric('Conflicts',conflicts); d.metric('Errors/Blocked',errors+blocked); e.metric('Output workbooks',len(outputs))
    st.info(f'Daywise source files indexed: {source_count}. Valuation files/sheets excluded by Completed/Marks-to-Department rules: {skipped}.')
    with st.expander('Subject processing log'): st.dataframe(subjects,use_container_width=True)
    with st.expander('Daywise source scan'): st.dataframe(scan,use_container_width=True,height=300)
    st.subheader('Detailed audit')
    st.dataframe(audit,use_container_width=True,height=450)
    st.download_button('⬇️ Download audit CSV',audit.to_csv(index=False).encode('utf-8-sig'),'COE_dummy_matching_audit.csv','text/csv',use_container_width=True)
    zbuf=io.BytesIO()
    with zipfile.ZipFile(zbuf,'w',zipfile.ZIP_DEFLATED) as z:
        for original,name,data in outputs: z.writestr(str(PurePosixPath(original).parent/name),data)
    if conflicts or unmatched or errors or blocked:
        st.warning('Release remains BLOCKED because unresolved/conflicting/error rows exist. Review the audit before using any output officially.')
    else:
        st.success('All audited dummy rows matched uniquely.')
        st.download_button('⬇️ Download processed valuation ZIP',zbuf.getvalue(),'COE_processed_valuation.zip','application/zip',use_container_width=True)
