import sys, openpyxl, collections, statistics as st
def load(path):
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    rows=list(wb['Respuestas'].iter_rows(values_only=True))[1:]
    P=collections.defaultdict(lambda: {'ans':{},'meta':None})
    for r in rows:
        doc=str(r[1]); p=P[doc]
        p['meta']=dict(nombre=r[0],empresa=r[2],sede=r[3],area=r[4],perfil=r[5],estado=r[7])
        p['ans'][r[12]]=(r[16],r[19],r[20])
    return P
def scores(p):
    a=p['ans']
    def pct(codes):
        v=[1 if a[c][1]=='Sí' else 0 for c in codes if c in a and a[c][1] in ('Sí','No')]
        return 100*sum(v)/len(v) if len(v)==len(codes) else None
    S=pct(['S_1','S_2','S_3','S_4','S_5']); D=pct(['D_1','D_2','D_3','D_4','D_5'])
    us=[a[c][2] for c in ('U1','U4','U5') if c in a and a[c][2] is not None]
    U=sum(us)/3 if len(us)==3 else None
    ms=[]
    for c in ['M1','M2','M3','M4','M5','M6']:
        if c in a and a[c][0]: ms.append(int(str(a[c][0])[0]))
    M=None
    if len(ms)==6: M=ms
    dom=round(.4*S+.35*D+.25*U,1) if None not in (S,D,U) else None
    return S,D,U,dom,ms
