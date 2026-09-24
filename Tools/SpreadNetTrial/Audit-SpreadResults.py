"""Read-only comparison of isolated engine value probes; never edits workbooks."""
import argparse
import json
from pathlib import Path

ROOT = (Path(__file__).resolve().parents[2] / "obj/AsposeTrial").resolve()

def private(value):
    p=Path(value).resolve()
    if ROOT not in p.parents:
        raise ValueError("Private trial paths only")
    return p

def compare(reference, candidate):
    rows=[]
    for sheet,a in reference.items():
        b=candidate[sheet]
        numeric=types=text=errors=compared=0
        samples=[]
        for r in range(max(a["rows"],b["rows"])):
            for c in range(max(a["columns"],b["columns"])):
                x=a["values"][r][c] if r<a["rows"] and c<a["columns"] else None
                y=b["values"][r][c] if r<b["rows"] and c<b["columns"] else None
                x=None if x=="" else x
                y=None if y=="" else y
                if x is None and y is None: continue
                compared+=1
                if isinstance(y,str) and y.startswith("#"): errors+=1
                nx=isinstance(x,(int,float)) and not isinstance(x,bool)
                ny=isinstance(y,(int,float)) and not isinstance(y,bool)
                different=False
                if nx and ny:
                    if abs(x-y)>max(1e-6,abs(x)*1e-10):numeric+=1;different=True
                elif isinstance(x,str) and isinstance(y,str):
                    if x!=y:text+=1;different=True
                elif x!=y or type(x)!=type(y):types+=1;different=True
                if different and len(samples)<5:samples.append(dict(row=r+1,column=c+1,reference=x,candidate=y))
        rows.append(dict(sheet=sheet,compared=compared,numericDifferences=numeric,typeDifferences=types,
                         textDifferences=text,candidateErrors=errors,samples=samples))
    return rows

if __name__=="__main__":
    p=argparse.ArgumentParser()
    p.add_argument("reference");p.add_argument("candidate");p.add_argument("report")
    p.add_argument("--edits",action="store_true")
    a=p.parse_args()
    ref=json.loads(private(a.reference).read_text(encoding="utf-8-sig"))
    cand=json.loads(private(a.candidate).read_text(encoding="utf-8-sig"))
    if ref["inputHash"]!=cand["inputHash"]:raise ValueError("Different input bytes")
    stages=["baseline","editedIncremental","editedFull","restoredIncremental"] if a.edits else ["calculatedProbes"]
    result=dict(reference=a.reference,candidate=a.candidate,inputHash=ref["inputHash"],
                referenceCompleted=ref.get("success"),candidateCompleted=cand.get("success"),
                stages={s:compare(ref[s],cand[s]) for s in stages},
                limits="Compared measured probes only; a failed later mutation still invalidates that benchmark run.")
    with private(a.report).open("x",encoding="utf-8") as f:json.dump(result,f,indent=2)
    print(json.dumps(result,indent=2))
