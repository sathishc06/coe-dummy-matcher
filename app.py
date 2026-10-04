"""COE Subject-wise Dummy Number Matcher. Run: streamlit run app.py"""
import io, re, zipfile, os, csv
from pathlib import PurePosixPath
from collections import defaultdict
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

st.set_page_config(page_title='COE Dummy Matcher', page_icon='🎓', layout='centered', initial_sidebar_state='collapsed')
st.markdown('''<style>
.block-container{max-width:1100px;padding:1rem 1rem 3rem!important}
.stButton>button,.stDownloadButton>button{min-height:3rem;width:100%;white-space:normal;font-weight:600}
[data-testid=stFileUploader]{border:1px solid #d5dbe5;border-radius:12px;padding:10px}
[data-testid=stMetric]{background:#f4f7fb;border-radius:12px;padding:12px}
@media(max-width:640px){.block-container{padding:.75rem .65rem 2rem!important}h1{font-size:1.55rem!important}h2{font-size:1.25rem!important}h3{font-size:1.1rem!important}[data-testid=stDataFrame]{font-size:12px}}
</style>''',unsafe_allow_html=True)
st.title('🎓 COE Dummy Matcher')
st.write('Upload valuation and Daywise ZIPs. Workbooks with a uniquely identified source are processed; subjects without source data are excluded from the output ZIP and listed in the status report.')
st.warning('For confidential examination data, use only an institution-approved, access-controlled hosting environment. Uploaded files are processed on the server running this app.')

EXCEL_EXT={'.xlsx','.xlsm'}
DATA_EXT=EXCEL_EXT|{'.csv'}
def norm(v):
    if v is None: return ''
    s=str(v).strip()
    if s.lower()=='nan': return ''
    if s.endswith('.0') and s[:-2].isdigit(): s=s[:-2]
    return re.sub(r'[^A-Z0-9]+','',s.upper())
def clean_name(s): return re.sub(r'[^A-Z0-9]+',' ',str(s).upper()).strip()
def unzip_upload(upload):
    files=[]
    try:
        with zipfile.ZipFile(upload) as z:
            for info in z.infolist():
                if info.is_dir() or info.filename.startswith('__MACOSX/') or '/.' in info.filename: continue
                if PurePosixPath(info.filename).suffix.lower() in DATA_EXT:
                    files.append((info.filename,z.read(info)))
    except zipfile.BadZipFile: st.error(f'{upload.name} is not a valid ZIP file.')
    return files

def detect_header(ws, aliases, scan=40):
    aliases=[norm(x) for x in aliases]
    for r in range(1,min(ws.max_row,scan)+1):
        vals={norm(ws.cell(r,c).value):c for c in range(1,ws.max_column+1) if ws.cell(r,c).value is not None}
        for a in aliases:
            if a in vals: return r, vals[a]
        for k,c in vals.items():
            if k and any(a in k for a in aliases if len(a)>=5): return r,c
    return None,None

def detect_pair(ws, scan=40):
    """Find dummy and registration headers on the same row only."""
    da=[norm(x) for x in DUMMY_ALIASES]; ra=[norm(x) for x in REG_ALIASES]
    for r in range(1,min(ws.max_row,scan)+1):
        vals={norm(ws.cell(r,c).value):c for c in range(1,ws.max_column+1) if ws.cell(r,c).value is not None}
        def find(aliases):
            for a in aliases:
                if a in vals: return vals[a]
            for k,c in vals.items():
                if k and any(a in k for a in aliases if len(a)>=5): return c
            return None
        dc,rc=find(da),find(ra)
        if dc and rc: return r,dc,rc
    return None,None,None
DUMMY_ALIASES=['Dummy No','Dummy Number','Dummy','Dummy Reg No','Dummy ID']
REG_ALIASES=['Registration_Number','Registration Number','Registration No','Reg.No','Reg No','Register Number','Register No']

def subject_tokens(path):
    # Use path and filename terms; omit generic folder/file words and date fragments.
    text=clean_name(PurePosixPath(path).stem)
    words=[w for w in text.split() if w not in {'DAYWISE','VALUATION','MAY','2026','MARKS','MARK','EXAM','EXAMINATION','COMPLETED','FINAL','DUMMY'} and not re.fullmatch(r'\d{1,2}',w)]
    return words

def subject_match(target_path, source_path):
    tw=subject_tokens(target_path); sw=subject_tokens(source_path)
    if not tw or not sw: return 0
    # exact distinctive code match takes priority
    tset=set(tw); sset=set(sw)
    codes_t=[x for x in tw if re.search(r'[A-Z]',x) and re.search(r'\d',x)]
    codes_s=[x for x in sw if re.search(r'[A-Z]',x) and re.search(r'\d',x)]
    common_codes=set(codes_t)&set(codes_s)
    if common_codes: return 100+len(common_codes)
    common=tset&sset
    if not common: return 0
    return 10*len(common)/max(1,len(tset))

def load_source_records(source_files):
    # index: normalized subject identifier -> dummy -> list(reg, provenance)
    records=defaultdict(lambda:defaultdict(list)); scan=[]; src_entries=[]
    for path,data in source_files:
        ext=PurePosixPath(path).suffix.lower()
        if ext=='.csv':
            try:
                import pandas as pd
                df=pd.read_csv(io.BytesIO(data),dtype=str,keep_default_na=False)
                cols={norm(c):c for c in df.columns}
                dc=next((cols[norm(a)] for a in ['Dummy_ID','Dummy No','Dummy Number'] if norm(a) in cols),None)
                rc=next((cols[norm(a)] for a in ['Registration_Number','Registration Number','Reg.No'] if norm(a) in cols),None)
                sc=next((cols[norm(a)] for a in ['Subject_Code','Subject Code'] if norm(a) in cols),None)
                if dc and rc and sc:
                    for i,row in df.iterrows():
                        code=norm(row[sc]); d=norm(row[dc]); reg=str(row[rc]).strip()
                        if code and d and reg: records[path+'|'+code][d].append((reg,f'{path} | row {i+2}'))
                    scan.append({'Source workbook':path,'Sheet':'CSV','Mapping rows':len(df)})
                else: scan.append({'Source workbook':path,'Sheet':'CSV','Mapping rows':0,'Error':'Required Subject_Code, Dummy_ID, Registration_Number columns missing'})
            except Exception as e: scan.append({'Source workbook':path,'Sheet':'CSV','Mapping rows':0,'Error':str(e)})
            continue
        if ext not in EXCEL_EXT: continue
        try:
            wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True)
            for ws in wb.worksheets:
                hr,dc,rc=detect_pair(ws)
                if not hr or not dc or not rc:
                    continue
                count=0
                for r in range(hr+1,ws.max_row+1):
                    d=norm(ws.cell(r,dc).value); raw=ws.cell(r,rc).value
                    reg=str(raw).strip() if raw is not None else ''
                    if d and reg:
                        records[path][d].append((reg,f'{path} | {ws.title} | row {r}')); count+=1
                scan.append({'Source workbook':path,'Sheet':ws.title,'Mapping rows':count})
            wb.close()
        except Exception as e: scan.append({'Source workbook':path,'Sheet':'','Mapping rows':0,'Error':str(e)})
    return records,scan

with st.sidebar:
    st.header('Upload archives')
    val_zip=st.file_uploader('1. Valuation May 2026 ZIP',type=['zip'],key='valzip')
    day_zip=st.file_uploader('2. Daywise ZIP (date folders + DAYWISE MAY 2026.xlsx)',type=['zip'],key='dayzip')
    st.caption('The app scans ZIP contents directly. You do not need to upload individual subject files.')

if not val_zip or not day_zip:
    st.info('Upload both ZIP files to begin. No files are processed until you press Start processing.')
    st.stop()
val_files=unzip_upload(val_zip); day_files=unzip_upload(day_zip)
st.metric('Valuation Excel files found',len(val_files)); st.metric('Daywise Excel files found',len(day_files))
if not val_files or not day_files: st.error('One or both ZIP archives contain no supported .xlsx/.xlsm files.'); st.stop()

st.subheader('Processing rules')
st.markdown('- Match **subject first**, then dummy number within that subject.\n- `Reg.No` retains the complete registration value exactly as found in Daywise.\n- Sort complete student rows by full `Reg.No` ascending.\n- `Reg Edit` contains only the text after the first hyphen.\n- Conflicts and unmatched records remain blank and are reported; the app never guesses.')
if st.button('🔎 Start processing valuation and Daywise ZIPs',type='primary',use_container_width=True):
    with st.spinner('Scanning Daywise files and processing subject-wise valuation workbooks…'):
        source_records,scanlog=load_source_records(day_files)
        outputs=[]; audit=[]; subject_log=[]; seen_names=set()
        for tpath,tdata in val_files:
            if PurePosixPath(tpath).suffix.lower() not in EXCEL_EXT: continue
            # Conservative selection: only a single uniquely best source workbook is eligible.
            target_codes=[x for x in subject_tokens(tpath) if re.search(r'[A-Z]',x) and re.search(r'\d',x)]
            exact=[sp for sp in source_records if any(sp.endswith('|'+code) for code in target_codes)]
            ranked=sorted(((subject_match(tpath,sp.split('|')[0]),sp) for sp in source_records if '|' not in sp),key=lambda x:(-x[0],x[1]))
            best=ranked[0][0] if ranked else 0
            candidates=exact if exact else [sp for score,sp in ranked if score==best and score>0]
            chosen=candidates if len(candidates)==1 else []
            if not chosen:
                subject_log.append({'Valuation workbook':tpath,'Matched Daywise workbook(s)':'','Status':'NO DAYWISE DATA: excluded from processed output; review subject/date/session'})
                audit.append({'Valuation workbook':tpath,'Sheet':'','Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':'NO DAYWISE DATA: workbook not processed'})
                continue
            else:
                subject_log.append({'Valuation workbook':tpath,'Matched Daywise workbook(s)':'; '.join(chosen),'Status':'Candidate source selected; index/date verification required'})
            mapping=defaultdict(list)
            for sp in chosen:
                for d,vals in source_records[sp].items(): mapping[d].extend(vals)
            try:
                keep_vba=PurePosixPath(tpath).suffix.lower()=='.xlsm'
                wb=load_workbook(io.BytesIO(tdata),keep_vba=keep_vba)
                for ws in wb.worksheets:
                    if ws.title.strip().lower()=='completed':
                        audit.append({'Valuation workbook':tpath,'Sheet':ws.title,'Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':'Skipped: Completed sheet'}); continue
                    hr,dc=detect_header(ws,DUMMY_ALIASES)
                    if not hr or not dc:
                        audit.append({'Valuation workbook':tpath,'Sheet':ws.title,'Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':'Skipped: dummy header not detected in first 40 rows'}); continue
                    headers={norm(ws.cell(hr,c).value):c for c in range(1,ws.max_column+1) if ws.cell(hr,c).value is not None}
                    regcol=headers.get(norm('Reg.No'))
                    if not regcol:
                        regcol=dc+1
                        # insert adjacent, shifting existing columns right
                        ws.insert_cols(regcol,1)
                    ws.cell(hr,regcol).value='Reg.No'
                    editcol=regcol+1
                    if norm(ws.cell(hr,editcol).value)!=norm('Reg Edit'):
                        ws.insert_cols(editcol,1)
                    ws.cell(hr,editcol).value='Reg Edit'
                    # Keep header style from dummy column
                    for c in (regcol,editcol):
                        src=ws.cell(hr,dc); dst=ws.cell(hr,c)
                        if src.has_style: dst._style=src._style
                    row_records=[]
                    for r in range(hr+1,ws.max_row+1):
                        d=norm(ws.cell(r,dc).value)
                        if not d: continue
                        vals=mapping.get(d,[])
                        regs=sorted(set(x[0] for x in vals))
                        if len(regs)==1:
                            full=regs[0]; edited=full.split('-',1)[1] if '-' in full else ''
                            status='Matched'
                        elif len(regs)>1:
                            full=''; edited=''; status='Conflict: same subject dummy maps to multiple registrations'
                        else:
                            full=''; edited=''; status='Unmatched: no subject-specific source record'
                        ws.cell(r,regcol).value=full; ws.cell(r,editcol).value=edited
                        audit.append({'Valuation workbook':tpath,'Sheet':ws.title,'Excel row':r,'Dummy No.':d,'Reg.No':full,'Reg Edit':edited,'Status':status,'Source evidence':'; '.join(x[1] for x in vals)})
                        row_records.append(r)
                    # Sort full rows by full Reg.No; formulas translated to their new row.
                    if row_records:
                        start=hr+1; end=ws.max_row; maxc=ws.max_column
                        rows=[]
                        for r in range(start,end+1):
                            cells=[]
                            for c in range(1,maxc+1):
                                cell=ws.cell(r,c); val=cell.value
                                cells.append((val,cell._style,cell.number_format,cell.hyperlink,cell.comment))
                            key=ws.cell(r,regcol).value or '\uffff'
                            rows.append((key,r,cells))
                        rows.sort(key=lambda x:(x[0]=='\uffff',str(x[0]).upper()))
                        for newr,(_,oldr,cells) in enumerate(rows,start):
                            for c,(val,style,nfmt,hyper,comment) in enumerate(cells,1):
                                dest=ws.cell(newr,c)
                                if isinstance(val,str) and val.startswith('=') and oldr!=newr:
                                    try: val=Translator(val,origin=ws.cell(oldr,c).coordinate).translate_formula(ws.cell(newr,c).coordinate)
                                    except Exception: pass
                                dest.value=val
                                if style: dest._style=style
                                dest.number_format=nfmt
                                if hyper: dest._hyperlink=hyper
                                if comment: dest.comment=comment
                out=io.BytesIO(); wb.save(out)
                base=PurePosixPath(tpath).name
                outname=base.rsplit('.',1)[0]+'_processed.'+base.rsplit('.',1)[-1]
                if outname in seen_names: outname=f'{len(outputs)+1}_{outname}'
                seen_names.add(outname); outputs.append((tpath,outname,out.getvalue()))
            except Exception as e:
                audit.append({'Valuation workbook':tpath,'Sheet':'','Excel row':'','Dummy No.':'','Reg.No':'','Reg Edit':'','Status':f'ERROR: {e}'})
    st.session_state['coe_result']={'outputs':outputs,'audit':pd.DataFrame(audit),'scan':pd.DataFrame(scanlog),'subjects':pd.DataFrame(subject_log)}

if 'coe_result' in st.session_state:
    res=st.session_state['coe_result']; aud=res['audit']; outs=res['outputs']
    matched=int((aud.get('Status',pd.Series(dtype=str))=='Matched').sum())
    conflicts=int(aud.get('Status',pd.Series(dtype=str)).astype(str).str.startswith('Conflict').sum())
    unmatched=int(aud.get('Status',pd.Series(dtype=str)).astype(str).str.startswith('Unmatched').sum())
    c1,c2,c3=st.columns(3); c1.metric('Matched rows',matched); c2.metric('Conflicts',conflicts); c3.metric('Unmatched',unmatched)
    with st.expander('Source scan details'): st.dataframe(res['scan'],use_container_width=True)
    with st.expander('Subject-to-source selection'): st.dataframe(res['subjects'],use_container_width=True)
    st.subheader('Detailed audit'); st.dataframe(aud,use_container_width=True,height=400)
    st.download_button('⬇️ Download audit CSV',aud.to_csv(index=False).encode('utf-8-sig'),'COE_subject_mapping_audit.csv','text/csv',use_container_width=True)
    st.download_button('⬇️ Download subject status CSV',res['subjects'].to_csv(index=False).encode('utf-8-sig'),'COE_subject_status.csv','text/csv',use_container_width=True)
    no_data = res['subjects'].get('Status',pd.Series(dtype=str)).astype(str).str.startswith('NO DAYWISE DATA').sum()
    st.info(f'{no_data} valuation workbook(s) have no uniquely identified Daywise source. They are excluded from the processed ZIP and listed in the subject status report.')
    zbuf=io.BytesIO()
    with zipfile.ZipFile(zbuf,'w',zipfile.ZIP_DEFLATED) as z:
        for original,outname,data in outs: z.writestr(str(PurePosixPath(original).parent/outname),data)
    
    blocked = conflicts > 0 or unmatched > 0 or aud.get('Status',pd.Series(dtype=str)).astype(str).str.startswith(('ERROR','BLOCKED','No subject')).any() or any('BLOCKED' in str(x) for x in res['subjects'].get('Status',[]))
    if blocked:
        st.error('RELEASE BLOCKED: unresolved matches, conflicts, or processing errors remain. Download is disabled to prevent accidental official use.')
    else:
        st.warning('Filename-only source selection is not an official date-index verification. Confirm against the Daywise date index before official use.')
        st.download_button('⬇️ Download all processed workbooks (ZIP)',zbuf.getvalue(),'COE_processed_valuation.zip','application/zip',use_container_width=True)
