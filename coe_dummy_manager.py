"""COE Dummy Number Manager — multi-source, multi-sheet processing and audit."""
import io, re, zipfile, hashlib
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
from copy import copy

st.set_page_config(page_title="COE Dummy Number Manager", page_icon="🎓", layout="wide")
st.title("🎓 COE Dummy Number Manager")
st.caption("Combine source workbooks, process target sheets, and audit every registration match.")
st.warning("For confidential examination records, use only an institution-approved, access-controlled deployment. Uploaded files are processed by the server hosting this app.")

ALIASES_DUMMY = ("dummy", "dummy no", "dummy number", "dummy reg no", "dummy register no")
ALIASES_REG = ("reg no", "reg.no", "registration no", "registration number", "register no", "register number", "university no", "roll no")

def norm(v):
    if v is None or (isinstance(v, float) and pd.isna(v)): return ""
    s=str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit(): s=s[:-2]
    return re.sub(r"\s+", "", s).upper()

def canon(v): return re.sub(r"[^a-z0-9]", "", str(v or "").lower())

def find_header(ws, aliases=ALIASES_DUMMY, scan_rows=40):
    keys=[canon(x) for x in aliases]
    for r in range(1, min(ws.max_row, scan_rows)+1):
        vals=[canon(ws.cell(r,c).value) for c in range(1, ws.max_column+1)]
        if any(k and (v==k or k in v) for v in vals for k in keys): return r
    return None

def find_column(ws, header_row, aliases):
    vals=[(c,canon(ws.cell(header_row,c).value)) for c in range(1,ws.max_column+1)]
    keys=[canon(x) for x in aliases]
    for c,v in vals:
        if v in keys: return c
    for c,v in vals:
        if any(k in v for k in keys if k): return c
    return None

def source_records(uploaded_files):
    records=[]; issues=[]
    for fi,f in enumerate(uploaded_files,1):
        try:
            wb=load_workbook(f, read_only=True, data_only=True)
            for ws in wb.worksheets:
                hr=find_header(ws)
                if not hr:
                    issues.append(f"{f.name} / {ws.title}: dummy header not found in first 40 rows")
                    continue
                dc=find_column(ws,hr,ALIASES_DUMMY); rc=find_column(ws,hr,ALIASES_REG)
                if not dc or not rc or dc==rc:
                    issues.append(f"{f.name} / {ws.title}: could not identify distinct dummy and registration columns")
                    continue
                for r in range(hr+1,ws.max_row+1):
                    d=norm(ws.cell(r,dc).value); reg=norm(ws.cell(r,rc).value)
                    if d and reg: records.append((d,reg,f.name,ws.title,r))
            wb.close()
        except Exception as e: issues.append(f"{f.name}: {e}")
    return records,issues

def build_mapping(records):
    evidence={}
    for d,r,fn,sn,row in records: evidence.setdefault(d,[]).append((r,f"{fn} | {sn} | row {row}"))
    mapping={}; conflicts={}
    for d,items in evidence.items():
        regs=sorted(set(r for r,_ in items))
        if len(regs)==1: mapping[d]=regs[0]
        else: conflicts[d]=regs
    return mapping,conflicts,evidence

def unique_name(name, used):
    if name not in used: used.add(name); return name
    stem=Path(name).stem; ext=Path(name).suffix; i=2
    while f"{stem}_{i}{ext}" in used: i+=1
    out=f"{stem}_{i}{ext}"; used.add(out); return out

def copy_cell(src,dst,formula_origin=None, formula_dest=None):
    val=src.value
    if isinstance(val,str) and val.startswith("=") and formula_origin and formula_dest:
        try: val=Translator(val, origin=formula_origin).translate_formula(formula_dest)
        except Exception: pass
    dst.value=val
    if src.has_style: dst._style=copy(src._style)
    if src.number_format: dst.number_format=src.number_format
    if src.hyperlink: dst._hyperlink=copy(src.hyperlink)
    if src.comment: dst.comment=copy(src.comment)

def sort_sheet_rows(ws, header_row, reg_col):
    """Sort data rows while moving whole rows and translating relative formulas."""
    start=header_row+1
    if ws.max_row < start+1: return
    max_col=ws.max_column
    row_data=[]
    for r in range(start,ws.max_row+1):
        key=norm(ws.cell(r,reg_col).value)
        row_data.append((r,key,[copy(ws.cell(r,c)) for c in range(1,max_col+1)]))
    row_data.sort(key=lambda x:(not bool(x[1]),x[1]))
    # Snapshot row dimensions before writing.
    dims={r:copy(ws.row_dimensions[r]) for r in range(start,ws.max_row+1)}
    for dest_r,(src_r,_,cells) in enumerate(row_data,start):
        for c,src in enumerate(cells,1):
            dst=ws.cell(dest_r,c)
            copy_cell(src,dst,src.coordinate,dst.coordinate)
        if src_r in dims:
            d=ws.row_dimensions[dest_r]; old=dims[src_r]
            d.height=old.height; d.hidden=old.hidden; d.outlineLevel=old.outlineLevel
    # Repositioning merged cells is ambiguous; fail explicitly rather than corrupting them.
    # Existing merges spanning data rows are not expected in mark sheets.

def process_targets(target_files,mapping,conflicts,reg_header,edited_header,sort_rows,skip_completed):
    outputs=[]; audit=[]; errors=[]; used=set(); totals={"rows":0,"matched":0,"unmatched":0,"conflicts":0,"skipped":0}
    for f in target_files:
        try:
            f.seek(0)
            keep_vba=f.name.lower().endswith(".xlsm")
            wb=load_workbook(f,keep_vba=keep_vba)
            for ws in wb.worksheets:
                if skip_completed and ws.title.strip().lower()=="completed":
                    totals["skipped"]+=1; audit.append({"Workbook":f.name,"Sheet":ws.title,"Excel Row":"","Dummy No.":"","Reg.No":"","Status":"Skipped: Completed sheet"}); continue
                hr=find_header(ws)
                if not hr:
                    audit.append({"Workbook":f.name,"Sheet":ws.title,"Excel Row":"","Dummy No.":"","Reg.No":"","Status":"Skipped: dummy header not found"}); continue
                dc=find_column(ws,hr,ALIASES_DUMMY)
                if not dc:
                    audit.append({"Workbook":f.name,"Sheet":ws.title,"Excel Row":"","Dummy No.":"","Reg.No":"","Status":"Skipped: dummy column not found"}); continue
                existing={canon(ws.cell(hr,c).value):c for c in range(1,ws.max_column+1) if ws.cell(hr,c).value is not None}
                oc=existing.get(canon(reg_header)) or ws.max_column+1
                if canon(reg_header) not in existing: ws.cell(hr,oc,reg_header)
                ec=existing.get(canon(edited_header)) or max(ws.max_column,oc)+1
                if canon(edited_header) not in existing: ws.cell(hr,ec,edited_header)
                # Record mappings before sorting so audit row numbers refer to original file locations.
                for r in range(hr+1,ws.max_row+1):
                    d=norm(ws.cell(r,dc).value)
                    if not d: continue
                    totals["rows"]+=1
                    if d in conflicts: reg=""; status="Conflict: dummy maps to multiple registrations"; totals["conflicts"]+=1
                    elif d in mapping: reg=mapping[d]; status="Matched"; totals["matched"]+=1
                    else: reg=""; status="Unmatched: no source mapping"; totals["unmatched"]+=1
                    ws.cell(r,oc,reg); ws.cell(r,ec,reg)
                    audit.append({"Workbook":f.name,"Sheet":ws.title,"Excel Row":r,"Dummy No.":d,"Reg.No":reg,"Status":status})
                if sort_rows: sort_sheet_rows(ws,hr,oc)
            buf=io.BytesIO(); wb.save(buf)
            outname=unique_name(Path(f.name).stem+"_processed"+Path(f.name).suffix,used)
            outputs.append((outname,buf.getvalue()))
        except Exception as e: errors.append(f"{f.name}: {e}")
    return outputs,pd.DataFrame(audit),totals,errors

with st.sidebar:
    st.header("1. Source files")
    source_files=st.file_uploader("Upload all Daywise source workbooks",type=["xlsx","xlsm"],accept_multiple_files=True,key="sources")
    st.header("2. Target files")
    target_files=st.file_uploader("Upload mark-sheet workbooks",type=["xlsx","xlsm"],accept_multiple_files=True,key="targets")

st.markdown("#### Processing options")
a,b=st.columns(2)
with a: sort_rows=st.checkbox("Sort each sheet by Reg.No (blank last)",value=False)
with b: skip_completed=st.checkbox("Skip sheets named Completed",value=True)
reg_header=st.text_input("Registration output header","Reg.No")
edited_header=st.text_input("Edited registration header","Reg.No Edited")

if source_files:
    records,source_issues=source_records(source_files)
    mapping,conflicts,evidence=build_mapping(records)
    st.metric("Source mapping entries",len(records))
    st.metric("Unique dummy numbers",len(evidence))
    if source_issues:
        with st.expander(f"Source warnings ({len(source_issues)})"):
            for issue in source_issues: st.warning(issue)
    if conflicts:
        st.error(f"{len(conflicts)} conflicting dummy numbers found. Conflicted numbers will be left blank.")
        st.dataframe(pd.DataFrame([{"Dummy No.":d,"Conflicting Reg.No":", ".join(rs),"Evidence": " ; ".join(x[1] for x in evidence[d])} for d,rs in conflicts.items()]),use_container_width=True)
    if target_files and st.button("⚙️ Process, audit and prepare downloads",type="primary",use_container_width=True):
        outputs,audit,totals,errors=process_targets(target_files,mapping,conflicts,reg_header,edited_header,sort_rows,skip_completed)
        st.session_state["coe_outputs"]=outputs; st.session_state["coe_audit"]=audit; st.session_state["coe_totals"]=totals; st.session_state["coe_errors"]=errors

if "coe_audit" in st.session_state:
    totals=st.session_state["coe_totals"]
    st.subheader("Processing summary")
    st.write(f"Rows: {totals['rows']} | Matched: {totals['matched']} | Unmatched: {totals['unmatched']} | Conflicts: {totals['conflicts']} | Skipped sheets: {totals['skipped']}")
    for e in st.session_state["coe_errors"]: st.error(e)
    audit=st.session_state["coe_audit"]
    st.dataframe(audit,use_container_width=True,height=400)
    st.download_button("⬇️ Download audit CSV",audit.to_csv(index=False).encode("utf-8-sig"),"COE_mapping_audit.csv","text/csv",use_container_width=True)
    outs=st.session_state["coe_outputs"]
    if len(outs)==1:
        n,data=outs[0]; st.download_button(f"⬇️ Download {n}",data,n,use_container_width=True)
    elif outs:
        zbuf=io.BytesIO()
        with zipfile.ZipFile(zbuf,"w",zipfile.ZIP_DEFLATED) as z:
            for n,data in outs: z.writestr(n,data)
        st.download_button("⬇️ Download all processed workbooks (ZIP)",zbuf.getvalue(),"COE_processed_workbooks.zip","application/zip",use_container_width=True)
