# Builds the two chart-ready CSVs for CC3 from the ONS regional GDP workbooks.
# Chart 1 (nations_gdp.csv): real GDP per head, 1998 = 100, for Scotland, Wales, Northern Ireland and England excl. London.
# Chart 2 (mayoral_growth.csv): average annual real GDP per head growth, 1998-2016 vs 2016-2023.
# Real GDP per head = real GDP (Table 10, 2023 prices) / population (Table 6), summed over the areas in each group.
# ONS notes that real (chained volume) figures are not exactly additive, so grouped areas are approximations.
# Run it in the folder that holds the two .xlsx files:  python3 cc3_build_data.py

import pandas as pd, numpy as np
ITL='regionalgrossdomesticproductgdpbyallinternationalterritoriallevelitlregions.xlsx'
LA ='regionalgrossdomesticproductgdplocalauthorities.xlsx'
def tab(path, sheet):
    d = pd.read_excel(path, sheet_name=sheet, header=1); d.columns=[str(c).strip() for c in d.columns]
    yrs=[c for c in d.columns if c.isdigit()]
    for c in yrs: d[c]=pd.to_numeric(d[c], errors='coerce')
    return d, yrs
# ---------- chart 1: nations, real GDP (2023 prices) per head ----------
r10,yrs = tab(ITL,'Table 10'); pop,_ = tab(ITL,'Table 6')
r10=r10.set_index('ITL code'); pop=pop.set_index('ITL code')
groups = {'Scotland':['TLM'],'Wales':['TLL'],'Northern Ireland':['TLN'],
          'England excl. London':['TLC','TLD','TLE','TLF','TLG','TLH','TLJ','TLK'],'UK':['UKX']}
rows=[]
for k,c in groups.items():
    ph = r10.loc[c,yrs].sum()/pop.loc[c,yrs].sum()
    for y in yrs: rows.append(dict(year=int(y), area=k, gdp_per_head_gbp=round(ph[y]), index_1998=round(ph[y]/ph['1998']*100,1)))
n=pd.DataFrame(rows); n.to_csv('nations_gdp_with_uk.csv',index=False)
n[n.area!='UK'].to_csv('nations_gdp.csv',index=False)   # chart 1 file (four nations)
print(n.pivot(index='year',columns='area',values='index_1998').iloc[[0,1,10,18,25,26]])
# ---------- chart 2: mayoral vs no-mayor, pre/post growth ----------
r10l,_=tab(LA,'Table 10'); popl,_=tab(LA,'Table 6')
r10l['LA name']=r10l['LA name'].str.strip(); popl['LA name']=popl['LA name'].str.strip()
r10l=r10l.set_index('LA code'); popl=popl.set_index('LA code')
nm = r10l['LA name']
def codes(names):
    out=[]; 
    for n_ in names:
        m=nm[nm==n_].index.tolist()
        if not m: print('NOT FOUND',n_)
        out+=m
    return out
CA={'Greater Manchester':['Bolton','Bury','Manchester','Oldham','Rochdale','Salford','Stockport','Tameside','Trafford','Wigan'],
 'West Midlands':['Birmingham','Coventry','Dudley','Sandwell','Solihull','Walsall','Wolverhampton'],
 'Liverpool City Region':['Halton','Knowsley','Liverpool','St. Helens','Sefton','Wirral'],
 'West Yorkshire':['Bradford','Calderdale','Kirklees','Leeds','Wakefield'],
 'South Yorkshire':['Barnsley','Doncaster','Rotherham','Sheffield'],
 'Tees Valley':['Hartlepool','Middlesbrough','Redcar and Cleveland','Stockton-on-Tees','Darlington']}
other_mayor=['Cambridge','East Cambridgeshire','Fenland','Huntingdonshire','South Cambridgeshire','Peterborough',
             'Bath and North East Somerset','Bristol, City of','South Gloucestershire','Newcastle upon Tyne','North Tyneside','Northumberland']
ca_codes={k:codes(v) for k,v in CA.items()}
six=sum(ca_codes.values(),[]); oth=codes(other_mayor)
reg = r10l['ITL1 Region'].str.strip()
eng = reg[~reg.isin(['Scotland','Wales','Northern Ireland','London'])].index.tolist()
complete = [c for c in eng if r10l.loc[c,yrs].notna().all()]
nomayor=[c for c in complete if c not in six and c not in oth]
print('English non-London LAs',len(eng),'complete',len(complete),'no-mayor pool',len(nomayor))
def ph(c): return r10l.loc[c,yrs].sum()/popl.loc[c,yrs].sum()
def growth(c):
    s=ph(c); pre=((s['2016']/s['1998'])**(1/18)-1)*100; post=((s['2023']/s['2016'])**(1/7)-1)*100
    return pre,post
rows=[]
def add(label,c,group,pooled):
    pre,post=growth(c); rows.append(dict(area=label,group=group,pooled=pooled,growth_pre=round(pre,2),growth_post=round(post,2),change=round(post-pre,2)))
for k,c in ca_codes.items(): add(k,c,'Mayoral combined authority',0)
add('Six mayoral areas (pooled)',six,'Mayoral combined authority',1)
top=popl.loc[nomayor,'2016'].sort_values(ascending=False).head(6).index.tolist()
for c in top: add(nm[c],[c],'No mayor',0)
add('Rest of England, no mayor (pooled)',nomayor,'No mayor',1)
d=pd.DataFrame(rows); d.to_csv('mayoral_growth.csv',index=False)
print(d.to_string(index=False))
print('pop of no-mayor pool 2016 (m):', round(popl.loc[nomayor,'2016'].sum()/1e6,1), ' six:', round(popl.loc[six,'2016'].sum()/1e6,1))