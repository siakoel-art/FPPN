# -*- coding: utf-8 -*-
import json
import pandas as pd
"""
FPPN - Sinistralité lue dans les deux sens (v15, 11/11/2026).
Étape 1 (lecture) : écritures des comptes de prestations (601…, 60800007, 6021/6024) des grands livres 2021 au 31/08/2026.
Étape 2 (classement) : chaque écriture maladie (6011, 6012, 6013, 6018, 6080) est classée par son libellé.
Étape 3 (indicateurs) : met à jour assets/sinistralite_v15.json (catégories, soins seuls, taux) et écrit
assets/prestations_hors_soins_v15.csv (une ligne par écriture étrangère aux soins), lus par v12/etape9.py et classeur_v14_fppn.py.
Utilisation : python3 sinistralite_gl_fppn.py <dossier des grands livres> <dossier de travail>
Les cotisations, les frais affectés et le ratio réglementaire (sin_num, sin_den) de sinistralite_v15.json viennent des balances.
"""
# ---- étape 1 : lecture des grands livres
import sys, re, collections, pickle
sys.path.insert(0,'/home/user/FPPN/scripts')
import extraction_fppn as X
G = sys.argv[1].rstrip("/") + "/"
F={2022:G+"2022/Nouveau/GRAND LIVRE DES COMPTES 2022.xlsx",2023:G+"2023/GRAND LIVRE DES COMPTES FPPN 2023.xlsx",2024:G+"Nouveaux fichiers/GRAND LIVRE FPPN 2024.xlsx",
   2025:G+"2025/GRAND LIVRE DES COMPTES FPPN 2025.xlsx",2026:G+"2026/GRAND LIVRE DES COMPTES FPPN DU 01-01-2026 AU 31-08-2026.xlsx",2021:G+"Nouveaux fichiers/GRAND LIVRE 2021.xlsx"}
COMPTES=re.compile(r'^(601\d{5}|60800007|602[14]\d{4})$')
out=[]
for an,f in F.items():
    d=open(f,'rb').read()
    try: feuilles=list(X.lire_calamine(d))
    except Exception as e:
        try: feuilles=list(X.lire_openpyxl(d))
        except Exception as e2: print(an,'ERR',e,e2); continue
    cur=None
    for nom,lignes in feuilles:
        for l in lignes:
            c0=str(l[0]).split('.')[0] if l and l[0] is not None else ''
            if re.fullmatch(r'\d{8}',c0) and (len(l)<2 or not l[1]): cur=c0; continue
            if cur and COMPTES.match(cur) and l and l[0] is not None and re.match(r'\d{4}-\d\d-\d\d',str(l[0])):
                num=lambda v: float(v) if isinstance(v,(int,float)) else 0.0
                lib=str(l[5] or '') if len(l)>5 else ''
                out.append(dict(an=an,compte=cur,date=str(l[0])[:10],journal=l[1],piece=l[2],lib=lib,debit=num(l[19]) if len(l)>19 else 0.0,credit=num(l[21]) if len(l)>21 else 0.0))
    print(an,len([o for o in out if o['an']==an]))
pickle.dump(out, open(sys.argv[2] + "/gl_prest.pkl", "wb"))

# ---- étapes 2 et 3 : classement et indicateurs
S = sys.argv[2]
d=pd.DataFrame(pickle.load(open(S+'/gl_prest.pkl','rb'))); d['m']=d.debit-d.credit
d=d[(d.journal!='RAN')&d.compte.str[:4].isin(['6011','6012','6013','6018','6080'])].copy()
L=d.lib.str.upper()
def cat(l):
    if re.search(r"\bDON\b", l) and re.search(r"SOIN|DIALYSE|DON MED|ANTI-AN|KIT D.ACCOUCH|DON DE SANG|COUVERTURE FRAIS|PHCIE", l): return "Autres prestataires de soins"
    if re.search(r"\bDON\b|K'DO|CADEAU|PANIER", l): return "Dons (autorités, élus, délégués)"
    if re.search(r"FUTURE VISI|GADGET|SAC KABA|IMPRIMERIE|IMPRIME|NEW PRINT|CARTE", l): return "Gadgets, imprimés et cartes"
    if re.search(r"00444\d{3}", l): return "Honoraires d'un cabinet (factures trimestrielles)"
    if re.search(r"RECLASSEMENT FRAIS DE RECOUVREMENT|PROV CPTE", l): return "Régularisations comptables"
    if re.search(r"CMP|CLINIQUE MODERNE|CLI.?IQUE MODER", l): return "Clinique du FPPN (CMP)"
    return "Autres prestataires de soins"
d['cat'] = d.lib.str.upper().map(cat)
x=d[~d.cat.isin(['Autres prestataires de soins','Clinique du FPPN (CMP)'])][['an','date','journal','piece','compte','lib','m','cat']].sort_values(['an','cat','date'])
x['date']=x.date.astype(str).str[:10]
x.to_csv('assets/prestations_hors_soins_v15.csv',sep=';',index=False,encoding='utf-8-sig')
t=d.pivot_table(index='cat',columns='an',values='m',aggfunc='sum').fillna(0)
J=json.load(open('assets/sinistralite_v15.json'))
for a,v in J.items():
    a_=int(a); v['cats']={c:float(t.loc[c,a_]) for c in t.index}
    assert abs(sum(v['cats'].values())-v['prest'])<1
    v['cmp']=v['cats']['Clinique du FPPN (CMP)']
    v['non_soins']=sum(val for c,val in v['cats'].items() if c not in ('Autres prestataires de soins','Clinique du FPPN (CMP)'))
    v['soins']=v['prest']-v['non_soins']; v['sc_soins']=v['soins']/v['cot']; v['part_cmp']=v['cmp']/v['soins'] if v['soins'] else 0
    v['comb_soins']=(v['soins']+v['frais'])/v['cot']; v['sin_soins']=(v['sin_num']-v['non_soins'])/v['sin_den']
    print(a, round(v['non_soins']), 'part',round(100*v['non_soins']/v['prest'],2),'scs', round(100*v['sc_soins'],2),'cmp', round(100*v['part_cmp'],2),'cbs', round(100*v['comb_soins'],2),'sins', round(100*v['sin_soins'],2), {k[:8]:round(w) for k,w in v['cats'].items()})
json.dump(J,open('assets/sinistralite_v15.json','w'),ensure_ascii=False,indent=1)
print(len(x))
