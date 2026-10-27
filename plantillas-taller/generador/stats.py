import sys, json, collections, statistics as st
sys.path.insert(0,'/tmp/claude-0/-home-user-Super-Humanos/e34f4cca-9372-52e7-aff6-e3600275411a/scratchpad')
import calc
P=calc.load(sys.argv[1])
out={}
people=[]
for doc,p in P.items():
    m=p['meta']
    if m['estado']!='Terminado': continue
    S,D,U,dom,ms=calc.scores(p)
    if dom is None or len(ms)!=6: continue
    people.append(dict(doc=doc,**m,S=S,D=D,U=U,dom=dom,ment=round((sum(ms)/6-1)/4*100,1),a=p['ans']))
print('completos',len(people), collections.Counter(x['perfil'] for x in people), collections.Counter(x['area'] for x in people if x['area'] in ('Dirección del programa','—',None,'')))
def q(vals,p):
    v=sorted(vals); k=(len(v)-1)*p; f=int(k); c=min(f+1,len(v)-1); return v[f]+(v[c]-v[f])*(k-f)
for perf in ['comercial','administrativo']:
    G=[x for x in people if x['perfil']==perf]; n=len(G); o={'n':n}
    for k in ['S','D','U','dom','ment']: o[k]=round(st.mean(x[k] for x in G),1)
    items={}
    for c in ['S_1','S_2','S_3','S_4','S_5','D_1','D_2','D_3','D_4','D_5']:
        v=[x['a'][c][1]=='Sí' for x in G if c in x['a']]; items[c]=round(100*sum(v)/len(v),1)
    o['items']=items
    def dist(c): return collections.Counter(x['a'][c][0] for x in G if c in x['a']).most_common()
    for c in ['U1','U4','U5','U6','U7','U8','U9','U10','U11','H1','H2','H3','H4','H5','D7']: o[c]=dist(c)
    mo=collections.Counter()
    for x in G:
        if 'MO1' in x['a'] and x['a']['MO1'][0]:
            for opt in str(x['a']['MO1'][0]).split(' | ') if ' | ' in str(x['a']['MO1'][0]) else str(x['a']['MO1'][0]).split(', '): mo[opt.strip()]+=1
    o['MO1']=mo.most_common(); o['MO1_raw']=collections.Counter(str(x['a']['MO1'][0]) for x in G if 'MO1' in x['a']).most_common(3)
    ms={c:round(st.mean(int(str(x['a'][c][0])[0]) for x in G),2) for c in ['M1','M2','M3','M4','M5','M6']}
    low={c:sum(1 for x in G if int(str(x['a'][c][0])[0])<=2) for c in ms}
    o['M']=ms; o['M_low']=low
    doms=[x['dom'] for x in G]; cuts=[round(q(doms,p),1) for p in (.25,.5,.75)]; o['cuts']=cuts
    gr=collections.Counter()
    for x in G:
        g='G1' if x['dom']<cuts[0] else 'G2' if x['dom']<cuts[1] else 'G3' if x['dom']<cuts[2] else 'G4'; x['g']=g; gr[g]+=1
    o['grupos']=sorted(gr.items())
    tr=collections.Counter('Chispa' if x['dom']<40 else 'Impulso' if x['dom']<60 else 'Fuerza' if x['dom']<80 else 'Grandeza' for x in G); o['tramos']=tr
    weekly={'Una vez por semana','Varias veces por semana','Todos los días'}
    cand=[x for x in G if x['dom']>=cuts[2] and x['ment']>=80 and x['a']['U1'][0] in weekly]; o['candidatos']=len(cand)
    o['empresas']=collections.Counter(x['empresa'] for x in G).most_common()
    out[perf]=o
allp=people
out['global']=dict(n=len(allp),dom=round(st.mean(x['dom'] for x in allp),1),U=round(st.mean(x['U'] for x in allp),1),S=round(st.mean(x['S'] for x in allp),1),D=round(st.mean(x['D'] for x in allp),1),ment=round(st.mean(x['ment'] for x in allp),1))
json.dump(out,open(sys.argv[2],'w'),ensure_ascii=False,indent=1,default=str)
import pickle; pickle.dump([{k:v for k,v in x.items() if k!='a'}|{'U1':x['a']['U1'][0]} for x in people],open(sys.argv[2]+'.pkl','wb'))
print(json.dumps(out,ensure_ascii=False,default=str)[:6000])
