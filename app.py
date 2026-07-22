import streamlit as st
import pandas as pd
import numpy as np
import re, io, base64, warnings, os
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import FancyBboxPatch
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')

# ═══════════════════════════════════════════════
# MARIANO EVENTS — DESIGN SYSTEM
# ═══════════════════════════════════════════════
C = dict(
    black  = '#0C0C0C',
    dark   = '#141414',
    card   = '#1A1A1A',
    border = '#2C2C2C',
    beige  = '#F5EDD8',
    beige2 = '#EDE0C4',
    beige3 = '#E2D0A8',
    cream  = '#FAF6EE',
    gold   = '#F0B43C',
    gold2  = '#D9980E',
    teal   = '#288C8C',
    teal2  = '#1A6060',
    purple = '#8C1478',
    purple2= '#6A0E5A',
    white  = '#FFFFFF',
    gray   = '#888888',
    lgray  = '#555555',
    green  = '#27AE60',
    red    = '#E74C3C',
    amber  = '#F39C12',
)

# ═══════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════
st.set_page_config(
    page_title="Mariano Events — Campaign Analyzer",
    page_icon="📊", layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════
# LOGO
# ═══════════════════════════════════════════════
def get_b64(path):
    if os.path.exists(path):
        with open(path,'rb') as f: return base64.b64encode(f.read()).decode()
    return None

logo_path = os.path.join(os.path.dirname(__file__), 'me_logo.png')
LOGO = get_b64(logo_path)

# ═══════════════════════════════════════════════
# CSS — Black + Beige + Brand
# ═══════════════════════════════════════════════
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

*, html, body, [class*="css"] {{
  font-family: 'Inter', sans-serif !important;
  background-color: {C['black']};
  color: {C['white']};
  box-sizing: border-box;
}}
.main {{ background: {C['black']}; }}
.block-container {{ padding: 0 1.6rem 2rem 1.6rem; max-width: 100%; }}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
  background: {C['dark']} !important;
  border-right: 1px solid {C['border']};
  padding-top: 0;
}}
section[data-testid="stSidebar"] > div {{ padding-top: 0; }}
section[data-testid="stSidebar"] * {{ color: {C['white']} !important; }}
section[data-testid="stSidebar"] .stSelectbox > div > div,
section[data-testid="stSidebar"] .stMultiSelect > div > div {{
  background: {C['card']} !important;
  border: 1px solid {C['border']} !important;
  border-radius: 8px !important;
  color: {C['white']} !important;
}}
section[data-testid="stSidebar"] input {{
  background: {C['card']} !important;
}}

/* ── Header ── */
.me-header {{
  background: {C['dark']};
  border-bottom: 1px solid {C['border']};
  padding: 16px 24px;
  display: flex; align-items: center; justify-content: space-between;
  margin: 0 -1.6rem 1.4rem -1.6rem;
  position: sticky; top: 0; z-index: 100;
}}
.me-logo-area {{ display: flex; align-items: center; gap: 16px; }}
.me-wordmark {{
  font-size: 1.1rem; font-weight: 800; letter-spacing: 3px;
  color: {C['white']}; text-transform: uppercase;
}}
.me-tagline {{ font-size: 0.6rem; color: {C['gray']}; letter-spacing: 2px; margin-top: 2px; }}
.me-right {{ text-align: right; }}
.me-right-label {{ font-size: 0.58rem; color: {C['lgray']}; letter-spacing: 2px; }}
.me-right-val {{ font-size: 0.75rem; font-weight: 700; color: {C['gold']}; letter-spacing: 1px; }}

/* ── KPI Cards ── */
.kpi-outer {{ display: grid; grid-template-columns: repeat(8,1fr); gap: 8px; margin-bottom: 8px; }}
.kpi-outer-4 {{ display: grid; grid-template-columns: repeat(4,1fr); gap: 8px; margin-bottom: 14px; }}
.kpi {{
  background: {C['card']};
  border: 1px solid {C['border']};
  border-radius: 10px; padding: 14px 12px;
  position: relative; overflow: hidden;
}}
.kpi::after {{
  content: ''; position: absolute;
  bottom: 0; left: 0; right: 0; height: 2px;
  background: {C['gold']};
}}
.kpi.teal::after {{ background: {C['teal']}; }}
.kpi.purple::after {{ background: {C['purple']}; }}
.kpi.green::after {{ background: {C['green']}; }}
.kpi.red::after {{ background: {C['red']}; }}
.kpi.amber::after {{ background: {C['amber']}; }}
.kpi-v {{
  font-size: 1.6rem; font-weight: 900; line-height: 1;
  color: {C['gold']}; letter-spacing: -0.5px;
}}
.kpi-v.teal {{ color: {C['teal']}; }}
.kpi-v.purple {{ color: #B040A0; }}
.kpi-v.green {{ color: {C['green']}; }}
.kpi-v.red {{ color: {C['red']}; }}
.kpi-v.amber {{ color: {C['amber']}; }}
.kpi-v.white {{ color: {C['white']}; }}
.kpi-l {{ font-size: 0.58rem; color: {C['gray']}; letter-spacing: 1.5px; text-transform: uppercase; margin-top: 5px; font-weight: 600; }}
.kpi-s {{ font-size: 0.65rem; color: {C['lgray']}; margin-top: 2px; }}

/* ── Section headers ── */
.sh {{
  display: flex; align-items: center; gap: 10px;
  padding: 9px 14px; margin: 18px 0 10px 0;
  background: {C['card']};
  border-left: 3px solid {C['gold']};
  border-radius: 0 8px 8px 0;
}}
.sh.teal {{ border-left-color: {C['teal']}; }}
.sh.purple {{ border-left-color: {C['purple']}; }}
.sh.green {{ border-left-color: {C['green']}; }}
.sh.red {{ border-left-color: {C['red']}; }}
.sh-t {{ font-size: 0.82rem; font-weight: 700; color: {C['gold']}; letter-spacing: 1px; text-transform: uppercase; }}
.sh.teal .sh-t {{ color: {C['teal']}; }}
.sh.purple .sh-t {{ color: #B040A0; }}
.sh.green .sh-t {{ color: {C['green']}; }}
.sh.red .sh-t {{ color: {C['red']}; }}

/* ── Insight box ── */
.ibox {{
  background: {C['card']}; border: 1px solid {C['border']};
  border-radius: 10px; padding: 14px 16px; margin: 10px 0;
  font-size: 0.8rem; color: {C['gray']}; line-height: 1.7;
}}
.ibox b {{ color: {C['gold']}; }}
.ibox.green {{ border-left: 3px solid {C['green']}; }}
.ibox.red {{ border-left: 3px solid {C['red']}; }}
.ibox.teal {{ border-left: 3px solid {C['teal']}; }}

/* ── Stock cards ── */
.scard {{
  background: #1E1212; border: 1px solid {C['red']}33;
  border-left: 3px solid {C['red']}; border-radius: 8px;
  padding: 10px 14px; margin-bottom: 8px;
}}
.scard.warn {{ background: #1E1A10; border-left-color: {C['amber']}; border-color: {C['amber']}33; }}
.scard.ok {{ background: #101E12; border-left-color: {C['green']}; border-color: {C['green']}33; }}
.scard-t {{ font-size: 0.85rem; font-weight: 700; color: {C['white']}; }}
.scard-d {{ font-size: 0.75rem; color: {C['gray']}; margin-top: 3px; line-height: 1.5; }}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {{
  background: {C['card']}; border-radius: 10px;
  border: 1px solid {C['border']}; padding: 3px; gap: 2px;
}}
.stTabs [data-baseweb="tab"] {{
  color: {C['gray']}; border-radius: 7px;
  font-size: 0.75rem; font-weight: 600; letter-spacing: 0.3px;
  padding: 6px 10px;
}}
.stTabs [aria-selected="true"] {{
  background: {C['dark']} !important;
  color: {C['gold']} !important;
  border-bottom: 2px solid {C['gold']} !important;
}}

/* ── Buttons ── */
.stButton > button {{
  background: {C['gold']}; color: {C['black']};
  border: none; border-radius: 8px; font-weight: 700;
}}
.stDownloadButton > button {{
  background: {C['teal']}; color: {C['white']};
  border: none; border-radius: 8px; font-weight: 700; width: 100%;
}}

/* ── Sidebar logo area ── */
.sb-logo {{
  background: {C['black']};
  padding: 20px 16px 14px 16px;
  border-bottom: 1px solid {C['border']};
  text-align: center; margin-bottom: 16px;
}}
.sb-section {{ font-size: 0.6rem; color: {C['gold']}; letter-spacing: 2px; font-weight: 700; text-transform: uppercase; display: block; margin-bottom: 4px; }}

/* ── Dataframe ── */
.stDataFrame {{ border-radius: 10px; overflow: hidden; }}
iframe {{ border-radius: 10px; }}

/* ── Filter pills ── */
.pill {{
  display: inline-block; background: {C['gold']}18;
  border: 1px solid {C['gold']}44; color: {C['gold']};
  border-radius: 20px; padding: 2px 10px;
  font-size: 0.68rem; font-weight: 600; margin: 2px;
}}

hr {{ border-color: {C['border']}; margin: 14px 0; }}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# CHART HELPERS  — beige panels, dark bg
# ═══════════════════════════════════════════════
BEIGE_BG   = '#F5EDD8'
BEIGE_GRID = '#E0D0B0'
BEIGE_TEXT = '#3A2E1E'
BEIGE_SUB  = '#7A6A50'

def fig_setup(nrows=1, ncols=1, w=14, h=4):
    fig, axes = plt.subplots(nrows, ncols, figsize=(w, h))
    fig.patch.set_facecolor('#111111')
    axl = [axes] if (nrows==1 and ncols==1) else \
          list(axes) if not hasattr(axes,'flat') else list(axes.flat)
    for ax in axl:
        ax.set_facecolor(BEIGE_BG)
        ax.tick_params(colors=BEIGE_TEXT, labelsize=8.5)
        ax.xaxis.label.set_color(BEIGE_SUB)
        ax.yaxis.label.set_color(BEIGE_SUB)
        ax.title.set_color(BEIGE_TEXT)
        ax.title.set_fontsize(10.5)
        ax.title.set_fontweight('bold')
        for spine in ax.spines.values():
            spine.set_edgecolor(BEIGE_GRID)
        ax.grid(axis='y', color=BEIGE_GRID, linewidth=0.6, linestyle='--')
        ax.grid(axis='x', visible=False)
        ax.set_axisbelow(True)
    plt.tight_layout()
    return fig, axl

def hax(ax):
    ax.grid(axis='x', color=BEIGE_GRID, linewidth=0.6, linestyle='--')
    ax.grid(axis='y', visible=False)

def lbl(ax, bars, fmt='{:.1f}', horiz=False, color=BEIGE_TEXT, fs=8.5):
    for bar in bars:
        v = bar.get_width() if horiz else bar.get_height()
        if v > 0:
            if horiz:
                ax.text(v + v*0.012 + 0.05, bar.get_y()+bar.get_height()/2,
                        fmt.format(v), va='center', color=color, fontsize=fs, fontweight='bold')
            else:
                ax.text(bar.get_x()+bar.get_width()/2, v + v*0.015 + 0.05,
                        fmt.format(v), ha='center', va='bottom', color=color, fontsize=fs, fontweight='bold')

def cbuf(fig):
    b = io.BytesIO()
    fig.savefig(b, format='png', dpi=130, bbox_inches='tight', facecolor='#111111')
    b.seek(0); plt.close(fig); return b

# ═══════════════════════════════════════════════
# DATA HELPERS
# ═══════════════════════════════════════════════
def parse_coupons(x):
    if pd.isna(x): return 0
    n = re.findall(r'\d+', str(x)); return int(n[0]) if n else 0

def bmatch(cols, hints):
    cl = [c.lower().strip() for c in cols]
    for h in hints:
        for i,c in enumerate(cl):
            if h.lower() in c: return cols[i]
    return None

def guess_map(cols):
    return {
        'date':      bmatch(cols,['date','event date']),
        'campaign':  bmatch(cols,['campaign name','campaign','program']),
        'ba_name':   bmatch(cols,['ba name','ambassador','rep name','staff name']),
        'venue':     bmatch(cols,['venue name','store name','location name']),
        'city':      bmatch(cols,['city']),
        'state':     bmatch(cols,['state']),
        'units':     bmatch(cols,['# total sales','total sales','units sold','# of units']),
        'revenue':   bmatch(cols,['$ total sales','revenue','dollar']),
        'sampled':   bmatch(cols,['total sampled','sampled','samples given']),
        'hours':     bmatch(cols,['ba hours','hours worked','hours']),
        'coupons':   bmatch(cols,['coupon']),
        'sku1':      bmatch(cols,['pineapple','pnapple','guava']),
        'sku2':      bmatch(cols,['lemon lime','lemon']),
        'sku3':      bmatch(cols,['orange mango','mango']),
        'shelf':     bmatch(cols,['enough product','shelf','stock support']),
        'inventory': bmatch(cols,['inventory concern','inventory']),
        'impact':    bmatch(cols,['impacted','impact','affect']),
        'comments':  bmatch(cols,['customer comments','comments','feedback']),
        'recommend': bmatch(cols,['recommend','likelihood']),
        'taste':     bmatch(cols,['taste','flavor rating']),
    }

def prep(df, m):
    r = df.copy()
    r['_u']  = pd.to_numeric(r[m['units']],  errors='coerce').fillna(0)
    r['_s']  = pd.to_numeric(r[m['sampled']],errors='coerce').fillna(0)
    r['_rev']= pd.to_numeric(r[m['revenue']],errors='coerce').fillna(0) if m.get('revenue') else 0
    r['_h']  = pd.to_numeric(r[m['hours']],  errors='coerce').fillna(0) if m.get('hours') else 0
    r['_c']  = r[m['coupons']].apply(parse_coupons) if m.get('coupons') else 0
    r['_dt'] = pd.to_datetime(r[m['date']], errors='coerce')
    r['_dow']= r['_dt'].dt.day_name()
    r['_cv'] = (r['_u'] / r['_s'].replace(0,np.nan) * 100).round(1)
    for k in ['sku1','sku2','sku3']:
        r[f'_{k}'] = pd.to_numeric(r[m[k]],errors='coerce').fillna(0) if m.get(k) else 0
    r['_camp']  = r[m['campaign']].fillna('Unknown').astype(str) if m.get('campaign') else 'Campaign'
    r['_state'] = r[m['state']].fillna('Unknown').astype(str)   if m.get('state')    else 'Unknown'
    r['_city']  = r[m['city']].fillna('Unknown').astype(str)    if m.get('city')     else 'Unknown'
    r['_ba']    = r[m['ba_name']].fillna('Unknown').astype(str) if m.get('ba_name')  else 'Unknown'
    r['_venue'] = r[m['venue']].fillna('Unknown').astype(str)   if m.get('venue')    else 'Unknown'
    def blob(row):
        return ' '.join(str(row.get(m[k],'')) for k in ['shelf','inventory','impact','comments'] if m.get(k) and pd.notna(row.get(m.get(k,''),''))).lower()
    r['_blob']  = r.apply(blob, axis=1)
    r['_sold']  = r['_blob'].str.contains('sold out|ran out|out of stock|selling out|sold through', na=False)
    r['_low']   = r['_blob'].str.contains('low stock|low inventory|restock|not enough|limited stock|running low', na=False)
    r['_noshelf']= (r[m['shelf']] == 'No') if m.get('shelf') else False
    r['_stock'] = r['_sold'] | r['_low'] | r['_noshelf']
    return r

def kpis(r):
    n=len(r); u=r['_u'].sum(); s=r['_s'].sum()
    rev=r['_rev'].sum(); h=r['_h'].sum()
    conv=round(u/s*100,2) if s>0 else 0
    au=round(u/n,2) if n>0 else 0; as_=round(s/n,1) if n>0 else 0
    uph=round(u/h,2) if h>0 else 0; ar=round(rev/n,2) if n>0 else 0
    cp=r['_c'].sum(); cpe=(r['_c']>0).sum()
    wc=r[r['_c']>0]['_u'].mean() if (r['_c']>0).any() else 0
    nc=r[r['_c']==0]['_u'].mean() if (r['_c']==0).any() else 0
    s1=r['_sku1'].sum(); s2=r['_sku2'].sum(); s3=r['_sku3'].sum()
    st=r['_stock'].sum(); so=r['_sold'].sum()
    return dict(n=int(n),u=float(u),s=float(s),rev=float(rev),h=float(h),
                conv=conv,au=au,as_=as_,uph=uph,ar=ar,
                cp=float(cp),cpe=int(cpe),wc=round(wc,2),nc=round(nc,2),
                s1=float(s1),s2=float(s2),s3=float(s3),st=int(st),so=int(so))

def kcard(col, val, label, color='', sub=''):
    sub_h = f'<div class="kpi-s">{sub}</div>' if sub else ''
    col.markdown(f'<div class="kpi {color}"><div class="kpi-v {color}">{val}</div><div class="kpi-l">{label}</div>{sub_h}</div>',
                 unsafe_allow_html=True)

def section(title, color='', icon=''):
    cls = f'sh {color}'
    t_cls = f'sh-t'
    st.markdown(f'<div class="{cls}"><span class="{t_cls}">{icon}&nbsp; {title}</span></div>',
                unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════
with st.sidebar:
    logo_html = f'<img src="data:image/png;base64,{LOGO}" style="width:120px;margin:0 auto;display:block;">' if LOGO else ''
    st.markdown(f'<div class="sb-logo">{logo_html}<div style="font-size:0.58rem;color:{C["gray"]};letter-spacing:2px;text-align:center;margin-top:10px;">CAMPAIGN ANALYZER</div></div>',
                unsafe_allow_html=True)

    st.markdown(f'<span class="sb-section">📁 Upload Data</span>', unsafe_allow_html=True)
    uploaded = st.file_uploader("", type=['xlsx','xls','csv'],
                                 accept_multiple_files=True,
                                 help="Upload one or more Promomash Excel files")
    if not uploaded:
        st.markdown(f'<div style="background:{C["card"]};border:1px dashed {C["border"]};border-radius:10px;padding:24px;text-align:center;margin-top:8px;"><div style="font-size:2rem;">📂</div><div style="color:{C["gray"]};font-size:0.75rem;margin-top:8px;">Drop your Promomash file here</div></div>',
                    unsafe_allow_html=True)
        st.stop()

# ── Load files ─────────────────────────────────
@st.cache_data
def load_file(name, data):
    try:
        xl = pd.ExcelFile(io.BytesIO(data))
        df = xl.parse('Raw') if 'Raw' in xl.sheet_names else xl.parse(xl.sheet_names[0])
    except: df = pd.read_csv(io.BytesIO(data))
    return df[df.iloc[:,0].notna()].copy()

all_dfs = {}
for uf in uploaded:
    df = load_file(uf.name, uf.read())
    all_dfs[uf.name.rsplit('.',1)[0]] = df

combined_raw = pd.concat(all_dfs.values(), ignore_index=True)
cols = combined_raw.columns.tolist()

# ── Column mapping ─────────────────────────────
if 'mapping' not in st.session_state:
    st.session_state.mapping = guess_map(cols)
m = st.session_state.mapping

with st.sidebar:
    st.markdown(f'<hr style="border-color:{C["border"]}">', unsafe_allow_html=True)
    if st.checkbox("⚙️ Edit Column Mapping",
                    value=any(not m.get(k) for k in ['date','ba_name','units','sampled'])):
        none_o = '— Not in file —'; opts = [none_o]+cols
        new_m = {}
        FIELDS = [('date','Date ✱',True),('ba_name','BA Name ✱',True),
                  ('units','Units Sold ✱',True),('sampled','Total Sampled ✱',True),
                  ('revenue','Revenue ($)',False),('hours','BA Hours',False),
                  ('coupons','Coupons',False),('campaign','Campaign Name',False),
                  ('venue','Store / Venue',False),('city','City',False),('state','State',False),
                  ('sku1','SKU — Pineapple Guava',False),('sku2','SKU — Lemon Lime',False),
                  ('sku3','SKU — Orange Mango',False),('shelf','Enough Product on Shelf?',False),
                  ('inventory','Inventory Concerns',False),('impact','What Impacted Sales?',False),
                  ('comments','Customer Comments',False)]
        for key,label,req in FIELDS:
            g = m.get(key); idx = opts.index(g) if g and g in opts else 0
            sel = st.selectbox(label, opts, index=idx, key=f'c_{key}')
            new_m[key] = None if sel==none_o else sel
        if st.button("✅ Save"):
            st.session_state.mapping = new_m; m=new_m; st.rerun()

    st.markdown(f'<hr style="border-color:{C["border"]}">', unsafe_allow_html=True)
    st.markdown(f'<span class="sb-section">🎯 Filters</span>', unsafe_allow_html=True)

# ── Prepare data ───────────────────────────────
r_full = combined_raw[pd.to_numeric(combined_raw[m['units']],errors='coerce').notna()].copy()
r_full = prep(r_full, m)

# ── Sidebar filters ─────────────────────────────
with st.sidebar:
    camps  = sorted(r_full['_camp'].unique())
    sel_c  = st.multiselect("Campaign", camps, default=camps)
    mn,mx  = r_full['_dt'].min().date(), r_full['_dt'].max().date()
    dr     = st.date_input("Date Range", value=(mn,mx), min_value=mn, max_value=mx)
    states = sorted(r_full['_state'].unique())
    sel_s  = st.multiselect("State", states, default=states)
    bas    = sorted(r_full['_ba'].unique())
    sel_b  = st.multiselect("BA / Ambassador", bas, default=bas)

    st.markdown(f'<hr style="border-color:{C["border"]}">', unsafe_allow_html=True)
    st.markdown(f'<span class="sb-section">⚙️ Display</span>', unsafe_allow_html=True)
    top_n  = st.slider("Top / Bottom N", 5, 20, 10)
    min_ev = st.slider("Min events for BA rank", 1, 5, 2)

# ── Apply filters ──────────────────────────────
r = r_full[r_full['_camp'].isin(sel_c)].copy()
if len(dr)==2:
    r = r[(r['_dt'].dt.date>=dr[0]) & (r['_dt'].dt.date<=dr[1])]
r = r[r['_state'].isin(sel_s) & r['_ba'].isin(sel_b)]
if len(r)==0:
    st.warning("No data matches current filters."); st.stop()

k = kpis(r)

# ═══════════════════════════════════════════════
# HEADER BAR
# ═══════════════════════════════════════════════
logo_hdr = f'<img src="data:image/png;base64,{LOGO}" style="height:44px;">' if LOGO else '📊'
st.markdown(f"""
<div class="me-header">
  <div class="me-logo-area">
    {logo_hdr}
    <div>
      <div class="me-wordmark">Mariano Events</div>
      <div class="me-tagline">CAMPAIGN INTELLIGENCE PLATFORM</div>
    </div>
  </div>
  <div class="me-right">
    <div class="me-right-label">POWERED BY</div>
    <div class="me-right-val">PROMOMASH ANALYTICS</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# KPI DASHBOARD
# ═══════════════════════════════════════════════
st.markdown(f'<div style="color:{C["gray"]};font-size:0.75rem;margin-bottom:10px;">Showing <b style="color:{C["gold"]}">{k["n"]:,} events</b> · {dr[0] if len(dr)>0 else "—"} to {dr[1] if len(dr)>1 else "—"} · {len(sel_c)} campaign(s)</div>',
            unsafe_allow_html=True)

c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
kcard(c1, f"{k['n']:,}",      "TOTAL EVENTS",      "")
kcard(c2, f"{k['s']:,.0f}",   "SAMPLES GIVEN",     "teal")
kcard(c3, f"{k['u']:,.0f}",   "UNITS SOLD (BA)",   "purple")
kcard(c4, f"${k['rev']:,.0f}","REVENUE",            "")
kcard(c5, f"{k['conv']}%",    "CONVERSION RATE",   "teal")
kcard(c6, f"{k['au']}",       "AVG UNITS/DEMO",    "")
kcard(c7, f"{k['as_']}",      "AVG SAMPLED/DEMO",  "teal")
kcard(c8, f"{k['uph']}",      "UNITS/BA HOUR",     "purple")

st.markdown("<div style='margin-top:8px;'></div>", unsafe_allow_html=True)
lift = round((k['wc']/k['nc']-1)*100,1) if k['nc']>0 else 0
spct = round(k['st']/k['n']*100,1) if k['n']>0 else 0

c1,c2,c3,c4 = st.columns(4)
kcard(c1, f"{k['cpe']}", "EVENTS WITH COUPONS","amber", f"of {k['n']} total")
kcard(c2, f"+{lift}%",   "COUPON LIFT",        "green", f"{k['wc']:.1f} vs {k['nc']:.1f} avg units")
kcard(c3, f"{k['st']}",  "STOCK ISSUE EVENTS", "red",   f"{spct}% of events")
kcard(c4, f"{k['so']}",  "SELL-OUTS",          "red",   "demand signal")

st.markdown(f'<hr style="border-color:{C["border"]};margin:16px 0;">', unsafe_allow_html=True)

# ═══════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════
tabs = st.tabs(["📅 Trends","🏆 BA Performance","🏪 Stores",
                "🗺️ Geography","🎯 Sampling & Coupons",
                "📦 Stock & Inventory","🥤 SKUs",
                "📊 Campaign Compare","📥 Download"])
CE = {}  # charts for excel

# ─── TAB 1: TRENDS ────────────────────────────
with tabs[0]:
    daily = r.groupby('_dt').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
    daily['avg'] = (daily.u/daily.n).round(1)

    section("Daily Performance Trend","","📅")
    fig, axs = fig_setup(1,2,14,4)
    xi = range(len(daily))
    axs[0].fill_between(xi, daily['u'], alpha=0.2, color=C['gold'])
    axs[0].plot(xi, daily['u'], color=C['gold'], lw=2.2, marker='o', ms=3)
    axs[0].set_title('Total Units Sold per Day')
    axs[0].set_xticks(list(xi)[::max(1,len(xi)//8)])
    axs[0].set_xticklabels([str(d)[:10] for d in daily['_dt'].iloc[::max(1,len(xi)//8)]],
                            rotation=35, ha='right', fontsize=7.5, color=BEIGE_TEXT)

    axs[1].fill_between(xi, daily['avg'], alpha=0.2, color=C['teal'])
    axs[1].plot(xi, daily['avg'], color=C['teal'], lw=2.2, marker='o', ms=3)
    axs[1].set_title('Avg Units per Demo per Day')
    axs[1].set_xticks(list(xi)[::max(1,len(xi)//8)])
    axs[1].set_xticklabels([str(d)[:10] for d in daily['_dt'].iloc[::max(1,len(xi)//8)]],
                            rotation=35, ha='right', fontsize=7.5, color=BEIGE_TEXT)
    st.pyplot(fig); CE['By Day'] = cbuf(fig)

    section("Day of Week Performance","teal","📆")
    order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    dow = r.groupby('_dow').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
    dow['avg']=(dow.u/dow.n).round(2); dow['conv']=(dow.u/dow.s*100).round(2)
    dow['sort']=dow['_dow'].map({d:i for i,d in enumerate(order)}); dow=dow.sort_values('sort')

    fig, axs = fig_setup(1,2,14,4)
    wknd_c = [C['gold'] if d in ['Friday','Saturday','Sunday'] else C['teal'] for d in dow['_dow']]
    bars=axs[0].bar(dow['_dow'],dow['avg'],color=wknd_c,edgecolor='none',width=0.65)
    axs[0].set_title('Avg Units/Demo by Day of Week')
    lbl(axs[0],bars,'{:.1f}')
    plt.setp(axs[0].xaxis.get_majorticklabels(),rotation=25,ha='right',color=BEIGE_TEXT)

    bars=axs[1].bar(dow['_dow'],dow['conv'],color=wknd_c,edgecolor='none',width=0.65,alpha=0.85)
    axs[1].set_title('Conversion % by Day of Week')
    lbl(axs[1],bars,'{:.1f}%')
    plt.setp(axs[1].xaxis.get_majorticklabels(),rotation=25,ha='right',color=BEIGE_TEXT)
    st.pyplot(fig)

    st.dataframe(dow[['_dow','n','u','avg','conv']].rename(columns={
        '_dow':'Day','n':'Events','u':'Units','avg':'Avg/Demo','conv':'Conv %'}),
        use_container_width=True, hide_index=True)

# ─── TAB 2: BA ────────────────────────────────
with tabs[1]:
    ba = r.groupby('_ba').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum'),rev=('_rev','sum')).reset_index()
    ba['avg']=(ba.u/ba.n).round(2); ba['conv']=(ba.u/ba.s*100).round(2)
    baq = ba[ba.n>=min_ev]
    top_ba = baq.sort_values('avg',ascending=False).head(top_n)
    low_ba = baq.sort_values('avg').head(top_n)

    section(f"Top {top_n} Brand Ambassadors","","🥇")
    fig, axs = fig_setup(1,2,14,max(4,top_n*0.48))
    hax(axs[0]); hax(axs[1])
    tc=[C['gold'] if i==0 else C['teal'] if i==1 else C['beige3'] for i in range(len(top_ba))]
    bars=axs[0].barh(top_ba['_ba'],top_ba['avg'],color=tc,edgecolor='none')
    axs[0].set_title(f'Top {top_n} BAs — Avg Units/Demo'); axs[0].invert_yaxis()
    lbl(axs[0],bars,'{:.2f}',horiz=True)

    bars=axs[1].barh(low_ba['_ba'],low_ba['avg'],color=C['red'],alpha=0.75,edgecolor='none')
    axs[1].set_title(f'Bottom {top_n} BAs — Avg Units/Demo'); axs[1].invert_yaxis()
    lbl(axs[1],bars,'{:.2f}',horiz=True)
    st.pyplot(fig); CE['By BA'] = cbuf(fig)

    co1,co2=st.columns(2)
    with co1:
        st.markdown(f'<div style="color:{C["gold"]};font-size:0.75rem;font-weight:700;margin-bottom:6px;">TOP {top_n} BAs</div>',unsafe_allow_html=True)
        st.dataframe(top_ba[['_ba','n','u','avg','conv']].rename(columns={'_ba':'BA','n':'Events','u':'Units','avg':'Avg/Demo','conv':'Conv %'}),use_container_width=True,hide_index=True)
    with co2:
        st.markdown(f'<div style="color:{C["red"]};font-size:0.75rem;font-weight:700;margin-bottom:6px;">BOTTOM {top_n} BAs</div>',unsafe_allow_html=True)
        st.dataframe(low_ba[['_ba','n','u','avg','conv']].rename(columns={'_ba':'BA','n':'Events','u':'Units','avg':'Avg/Demo','conv':'Conv %'}),use_container_width=True,hide_index=True)

# ─── TAB 3: STORES ────────────────────────────
with tabs[2]:
    sdf=r.groupby(['_venue','_city','_state']).agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum'),rev=('_rev','sum')).reset_index()
    sdf['avg']=(sdf.u/sdf.n).round(2); sdf['conv']=(sdf.u/sdf.s*100).round(2)
    top_st=sdf.sort_values('u',ascending=False).head(top_n).copy()
    low_st=sdf.sort_values('u').head(top_n).copy()
    top_st['lbl']=top_st['_city']+', '+top_st['_state']
    low_st['lbl']=low_st['_city']+', '+low_st['_state']

    section(f"Top & Bottom {top_n} Stores","","🏪")
    fig, axs = fig_setup(1,2,14,max(4,top_n*0.48))
    hax(axs[0]); hax(axs[1])
    tc2=[C['gold'] if i==0 else C['teal'] if i==1 else C['beige3'] for i in range(len(top_st))]
    bars=axs[0].barh(top_st['lbl'],top_st['u'],color=tc2,edgecolor='none')
    axs[0].set_title(f'Top {top_n} Stores — Total Units'); axs[0].invert_yaxis()
    lbl(axs[0],bars,'{:.0f}',horiz=True)

    bars=axs[1].barh(low_st['lbl'],low_st['u'],color=C['red'],alpha=0.75,edgecolor='none')
    axs[1].set_title(f'Bottom {top_n} Stores — Total Units'); axs[1].invert_yaxis()
    lbl(axs[1],bars,'{:.0f}',horiz=True)
    st.pyplot(fig); CE['Top 20 Stores'] = cbuf(fig)

    co1,co2=st.columns(2)
    with co1:
        st.dataframe(top_st[['lbl','_venue','n','u','avg','conv']].rename(columns={'lbl':'Location','_venue':'Store','n':'Events','u':'Units','avg':'Avg/Demo','conv':'Conv %'}),use_container_width=True,hide_index=True)
    with co2:
        st.dataframe(low_st[['lbl','_venue','n','u','avg','conv']].rename(columns={'lbl':'Location','_venue':'Store','n':'Events','u':'Units','avg':'Avg/Demo','conv':'Conv %'}),use_container_width=True,hide_index=True)

# ─── TAB 4: GEOGRAPHY ─────────────────────────
with tabs[3]:
    sg=r.groupby('_state').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum'),rev=('_rev','sum')).reset_index()
    sg['avg']=(sg.u/sg.n).round(2); sg['conv']=(sg.u/sg.s*100).round(2)
    sg=sg.sort_values('u',ascending=False); avg_l=sg['avg'].mean()

    section("State Performance","teal","🗺️")
    fig, axs = fig_setup(1,2,14,4)
    sc=[C['green'] if v>=avg_l else C['red'] for v in sg['avg']]
    bars=axs[0].bar(sg['_state'],sg['avg'],color=sc,edgecolor='none',width=0.7)
    axs[0].axhline(avg_l,color=C['gold'],lw=1.5,linestyle='--',label=f'Avg {avg_l:.1f}')
    axs[0].legend(facecolor=BEIGE_BG,labelcolor=BEIGE_TEXT,fontsize=8)
    axs[0].set_title('Avg Units/Demo by State')
    plt.setp(axs[0].xaxis.get_majorticklabels(),rotation=45,ha='right',fontsize=8,color=BEIGE_TEXT)

    bars=axs[1].bar(sg['_state'],sg['n'],color=C['teal'],edgecolor='none',width=0.7,alpha=0.85)
    axs[1].set_title('Events by State')
    lbl(axs[1],bars,'{:.0f}',fs=8)
    plt.setp(axs[1].xaxis.get_majorticklabels(),rotation=45,ha='right',fontsize=8,color=BEIGE_TEXT)
    st.pyplot(fig); CE['By State'] = cbuf(fig)
    st.dataframe(sg.rename(columns={'_state':'State','n':'Events','u':'Units','s':'Sampled','rev':'Revenue ($)','avg':'Avg/Demo','conv':'Conv %'}),use_container_width=True,hide_index=True)

# ─── TAB 5: SAMPLING & COUPONS ────────────────
with tabs[4]:
    r2=r.copy()
    r2['tier']=pd.cut(r2['_s'],bins=[0,50,65,80,300],labels=['≤50','51-65','66-80','80+'])
    sg2=r2.groupby('tier',observed=True).agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
    sg2['avg']=(sg2.u/sg2.n).round(2)

    section("Sampling Volume Impact","teal","🎯")
    fig, axs = fig_setup(1,2,12,4)
    tc3=[C['red'],C['amber'],C['teal'],C['green']]
    bars=axs[0].bar(sg2['tier'].astype(str),sg2['avg'],color=tc3,edgecolor='none',width=0.65)
    axs[0].set_title('Avg Units/Demo by Sampling Volume')
    lbl(axs[0],bars,'{:.2f}',fs=10)
    axs[1].set_facecolor(BEIGE_BG)
    wedges,texts,autos=axs[1].pie(sg2['n'],labels=sg2['tier'].astype(str),colors=tc3,
        autopct='%1.0f%%',startangle=90,textprops={'color':BEIGE_TEXT,'fontsize':10})
    axs[1].set_title('Events by Sampling Tier')
    st.pyplot(fig); CE['Sampling'] = cbuf(fig)

    section("Coupon Impact","","🎟️")
    wc_v=r[r['_c']>0]['_u'].mean() if (r['_c']>0).any() else 0
    nc_v=r[r['_c']==0]['_u'].mean() if (r['_c']==0).any() else 0
    lift2=round((wc_v/nc_v-1)*100,1) if nc_v>0 else 0

    fig, axs = fig_setup(1,2,12,4)
    bars=axs[0].bar(['With Coupons','Without Coupons'],[wc_v,nc_v],
                    color=[C['green'],C['red']],edgecolor='none',width=0.5)
    axs[0].set_title('Coupon Impact on Avg Units/Demo')
    lbl(axs[0],bars,'{:.2f}',fs=11)
    axs[0].text(0.5,0.93,f'Lift: +{lift2}%',transform=axs[0].transAxes,
                ha='center',color=C['gold'],fontsize=11,fontweight='bold')

    cd=r.groupby('_dt').apply(lambda g: pd.Series({
        'wc': g[g['_c']>0]['_u'].mean() if (g['_c']>0).any() else np.nan,
        'nc': g[g['_c']==0]['_u'].mean() if (g['_c']==0).any() else np.nan})).reset_index()
    xi2=range(len(cd))
    if cd['wc'].notna().any():
        axs[1].fill_between(xi2,cd['wc'].fillna(0),alpha=0.2,color=C['green'])
        axs[1].plot(xi2,cd['wc'],color=C['green'],lw=2,label='With coupons')
    if cd['nc'].notna().any():
        axs[1].fill_between(xi2,cd['nc'].fillna(0),alpha=0.2,color=C['red'])
        axs[1].plot(xi2,cd['nc'],color=C['red'],lw=2,label='No coupons')
    axs[1].legend(facecolor=BEIGE_BG,labelcolor=BEIGE_TEXT,fontsize=8)
    axs[1].set_title('Daily Units: Coupon vs No Coupon')
    st.pyplot(fig); CE['Coupons'] = cbuf(fig)

# ─── TAB 6: STOCK ─────────────────────────────
with tabs[5]:
    section("Stock & Inventory Analysis","red","📦")
    stock_ev=r[r['_stock']]; no_st=r[~r['_stock']]
    aws=stock_ev['_u'].mean() if len(stock_ev)>0 else 0
    ans=no_st['_u'].mean() if len(no_st)>0 else 0
    diff=aws-ans; pos=diff>0

    c1,c2,c3,c4=st.columns(4)
    kcard(c1,f"{k['st']}","STOCK ISSUE EVENTS","red",f"{spct}% of all events")
    kcard(c2,f"{k['so']}","CONFIRMED SELL-OUTS","red","ran out mid-demo")
    kcard(c3,f"{aws:.1f}","AVG UNITS — STOCK ISSUE","amber","see comparison below")
    kcard(c4,f"{ans:.1f}","AVG UNITS — FULL SHELF","teal","baseline")

    cls_ = 'green' if pos else 'red'
    msg = (f"Events with a stock issue sold <b>{abs(diff):.2f} units MORE</b> than full shelf events. These stores are selling out because the product moves fast — a demand signal, not a problem. Prioritise larger opening orders at these locations."
           if pos else
           f"Events with a stock issue sold <b>{abs(diff):.2f} units FEWER</b> than full shelf events. Low inventory is hurting performance. Pre-event inventory checks are recommended.")
    st.markdown(f'<div class="ibox {cls_}">{msg}</div>', unsafe_allow_html=True)

    fig, axs = fig_setup(1,2,12,4)
    bars=axs[0].bar(['Stock Issue','Full Shelf'],[aws,ans],
                    color=[C['amber'],C['teal']],edgecolor='none',width=0.5)
    axs[0].set_title('Avg Units: Stock Issue vs Full Shelf')
    lbl(axs[0],bars,'{:.2f}',fs=12)

    if m.get('state'):
        sg3=r.groupby('_state').agg(tot=('_stock','count'),iss=('_stock','sum')).reset_index()
        sg3['rate']=(sg3.iss/sg3.tot*100).round(1)
        sg3=sg3[sg3.tot>=2].sort_values('rate',ascending=False).head(12)
        sc3=[C['red'] if v>30 else C['amber'] if v>15 else C['teal'] for v in sg3['rate']]
        bars=axs[1].bar(sg3['_state'],sg3['rate'],color=sc3,edgecolor='none',width=0.7)
        axs[1].set_title('Stock Issue Rate by State (%)')
        lbl(axs[1],bars,'{:.0f}%',fs=8)
        plt.setp(axs[1].xaxis.get_majorticklabels(),rotation=45,ha='right',fontsize=8,color=BEIGE_TEXT)
    st.pyplot(fig); CE['Stock'] = cbuf(fig)

    co1,co2=st.columns(2)
    with co1:
        section("Sell-Out Events","red","🔴")
        so_df=r[r['_sold']][['_ba','_city','_state','_u','_s']].copy()
        if len(so_df)>0:
            so_df.columns=['BA','City','State','Units','Sampled']
            st.dataframe(so_df.sort_values('Units',ascending=False),use_container_width=True,hide_index=True)
            st.markdown(f'<div class="scard ok"><div class="scard-t">✅ Sell-outs sold {(aws/ans-1)*100:.0f}% more than full shelf events</div><div class="scard-d">These stores need larger opening orders, not fewer events.</div></div>',unsafe_allow_html=True)
        else:
            st.info("No sell-out events in current filter.")

    with co2:
        section("Low Stock Events","amber","🟡")
        ls_df=r[r['_low']][['_ba','_city','_state','_u']].copy()
        if len(ls_df)>0:
            ls_df.columns=['BA','City','State','Units']
            st.dataframe(ls_df.sort_values('Units',ascending=False),use_container_width=True,hide_index=True)
            st.markdown(f'<div class="scard warn"><div class="scard-t">⚠️ {len(ls_df)} events flagged low stock mid-demo</div><div class="scard-d">BA confirmed stock issues during the event. Pre-event morning calls can prevent this.</div></div>',unsafe_allow_html=True)
        else:
            st.info("No low stock events in current filter.")

    # Field quotes
    def get_quotes(r):
        rows=[]
        for col in ['impact','inventory','comments']:
            if m.get(col):
                sub=r[r['_stock']][[m[col],'_city','_state']].dropna(subset=[m[col]])
                sub=sub[sub[m[col]].astype(str).str.lower().str.contains('stock|sold out|inventory|shelf|ran out|low|restock',na=False)]
                for _,row in sub.iterrows():
                    txt=str(row[m[col]]).strip()
                    if 10<len(txt)<300 and txt.lower() not in ['none','n/a','na','no']:
                        rows.append({'Location':f"{row['_city']}, {row['_state']}","Quote":txt})
        return pd.DataFrame(rows).drop_duplicates('Quote').head(8)
    q=get_quotes(r)
    if len(q)>0:
        section("Field Reports — Stock Comments","gold","💬")
        for _,row in q.iterrows():
            st.markdown(f'<div class="scard"><div class="scard-t">{row["Location"]}</div><div class="scard-d">"{row["Quote"]}"</div></div>',unsafe_allow_html=True)

# ─── TAB 7: SKUs ──────────────────────────────
with tabs[6]:
    s1=k['s1']; s2=k['s2']; s3=k['s3']; tot=s1+s2+s3
    if tot>0:
        section("SKU / Flavor Performance","purple","🥤")
        c1,c2,c3=st.columns(3)
        kcard(c1,f"{s1:.0f}","PINEAPPLE GUAVA","gold",f"{s1/tot*100:.1f}% share")
        kcard(c2,f"{s2:.0f}","LEMON LIME","teal",f"{s2/tot*100:.1f}% share")
        kcard(c3,f"{s3:.0f}","ORANGE MANGO","purple",f"{s3/tot*100:.1f}% share")

        fig, axs = fig_setup(1,2,12,4.5)
        skc=['#A8C4E0','#1F2260','#8B7362']
        axs[0].set_facecolor(BEIGE_BG)
        axs[0].pie([s1,s2,s3],labels=['Pineapple Guava','Lemon Lime','Orange Mango'],
               colors=skc,autopct='%1.1f%%',startangle=90,
               textprops={'color':BEIGE_TEXT,'fontsize':10})
        axs[0].set_title('Unit Share by Flavor')

        bars=axs[1].bar(['Pineapple Guava','Lemon Lime','Orange Mango'],[s1,s2,s3],
                        color=skc,edgecolor='none',width=0.6)
        axs[1].set_title('Total Units by SKU')
        lbl(axs[1],bars,'{:.0f}',fs=11)
        st.pyplot(fig); CE['SKU'] = cbuf(fig)
    else:
        st.info("Map SKU columns to see this analysis.")

# ─── TAB 8: CAMPAIGN COMPARE ──────────────────
with tabs[7]:
    if r['_camp'].nunique()>1:
        section("Campaign vs Campaign","","📊")
        rows=[]
        for camp in r['_camp'].unique():
            rc=r[r['_camp']==camp]; kc=kpis(rc)
            rows.append({'Campaign':camp,'Events':kc['n'],
                'Units':f"{kc['u']:.0f}",'Samples':f"{kc['s']:.0f}",
                'Revenue':f"${kc['rev']:,.0f}",'Conv %':f"{kc['conv']}%",
                'Avg Units/Demo':kc['au'],'Units/Hour':kc['uph']})
        comp=pd.DataFrame(rows)
        st.dataframe(comp,use_container_width=True,hide_index=True)

        fig, axs = fig_setup(1,3,14,4)
        camp_colors=[C['gold'],C['teal'],C['purple'],C['amber'],C['green']]
        for ax,(col,title) in zip(axs,[('Avg Units/Demo','Avg Units/Demo'),
                                        ('Conv %','Conversion %'),('Units/Hour','Units/Hour')]):
            vals=[float(str(v).replace('%','')) for v in comp[col]]
            clrs=[camp_colors[i%len(camp_colors)] for i in range(len(comp))]
            bars=ax.bar(comp['Campaign'],vals,color=clrs,edgecolor='none',width=0.6)
            ax.set_title(title)
            lbl(ax,bars,'{:.2f}',fs=9)
            plt.setp(ax.xaxis.get_majorticklabels(),rotation=20,ha='right',fontsize=8,color=BEIGE_TEXT)
        st.pyplot(fig)
    else:
        st.info("Upload or select more than one campaign to see comparison.")

# ─── TAB 9: DOWNLOAD ──────────────────────────
with tabs[8]:
    section("Download Full Report","teal","📥")
    st.markdown(f'<div style="color:{C["gray"]};font-size:0.8rem;margin-bottom:16px;">All tables + charts embedded in the Excel file.</div>',unsafe_allow_html=True)

    def build_xl(r, m, charts):
        buf=io.BytesIO()
        k2=kpis(r)
        sheets={}
        sheets['KPI Summary']=pd.DataFrame([
            ['Total Events',k2['n'],'COUNT of rows'],
            ['Total Samples',f"{k2['s']:,.0f}",'SUM Sampled'],
            ['Total Units (BA)',f"{k2['u']:,.0f}",'SUM Units'],
            ['Total Revenue',f"${k2['rev']:,.2f}",'SUM Revenue'],
            ['Conversion Rate',f"{k2['conv']}%",'Units/Sampled x100'],
            ['Avg Units/Demo',k2['au'],'Units/Events'],
            ['Avg Sampled/Demo',k2['as_'],'Sampled/Events'],
            ['Units/BA Hour',k2['uph'],'Units/Hours'],
            ['Coupon Lift',f"+{round((k2['wc']/k2['nc']-1)*100,1)}%" if k2['nc']>0 else 'N/A','(With/Without-1)x100'],
            ['Stock Issue Events',f"{k2['st']} ({spct}%)","Events with inventory flag"],
            ['Sell-Outs',k2['so'],'Events confirmed sold out'],
        ],columns=['Metric','Value','How Calculated'])

        rc=[m[v] for v in ['date','ba_name','venue','city','state','units','revenue','sampled','hours','coupons'] if m.get(v)]
        cl=r[rc].copy(); cl['Conv %']=r['_cv'].values; sheets['All Events']=cl

        sg_=r.groupby('_state').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum'),rev=('_rev','sum')).reset_index()
        sg_['avg']=(sg_.u/sg_.n).round(2); sg_['conv']=(sg_.u/sg_.s*100).round(2)
        sg_.columns=['State','Events','Units','Sampled','Revenue','Avg/Demo','Conv %']
        sheets['By State']=sg_.sort_values('Units',ascending=False)

        ba_=r.groupby('_ba').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
        ba_['avg']=(ba_.u/ba_.n).round(2); ba_['conv']=(ba_.u/ba_.s*100).round(2)
        ba_.columns=['BA','Events','Units','Sampled','Avg/Demo','Conv %']
        sheets['By BA']=ba_.sort_values('Avg/Demo',ascending=False)

        dow_=r.groupby('_dow').agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
        dow_['avg']=(dow_.u/dow_.n).round(2); dow_['conv']=(dow_.u/dow_.s*100).round(2)
        dow_['so']=dow_['_dow'].map({d:i for i,d in enumerate(['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'])})
        dow_=dow_.sort_values('so').drop('so',axis=1)
        dow_.columns=['Day','Events','Units','Sampled','Avg/Demo','Conv %']
        sheets['By Day']=dow_

        st_=r.groupby(['_venue','_city','_state']).agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
        st_['avg']=(st_.u/st_.n).round(2); st_['conv']=(st_.u/st_.s*100).round(2)
        st_.columns=['Store','City','State','Events','Units','Sampled','Avg/Demo','Conv %']
        sheets['Top 20 Stores']=st_.sort_values('Units',ascending=False).head(20)
        sheets['Low 20 Stores']=st_.sort_values('Units').head(20)

        sk_=r[r['_stock']][['_ba','_city','_state','_u','_s','_sold','_low']].copy()
        sk_.columns=['BA','City','State','Units','Sampled','Sold Out','Low Stock']
        sheets['Stock Issues']=sk_.sort_values('Units',ascending=False)

        r3=r.copy(); r3['tier']=pd.cut(r3['_s'],bins=[0,50,65,80,300],labels=['≤50','51-65','66-80','80+'])
        sg4=r3.groupby('tier',observed=True).agg(n=('_u','size'),u=('_u','sum'),s=('_s','sum')).reset_index()
        sg4['avg']=(sg4.u/sg4.n).round(2); sg4.columns=['Tier','Events','Units','Sampled','Avg/Demo']
        sheets['Sampling Tiers']=sg4

        s1_=k2['s1']; s2_=k2['s2']; s3_=k2['s3']; tot_=s1_+s2_+s3_
        if tot_>0:
            sheets['SKU']=pd.DataFrame({'SKU':['Pineapple Guava','Lemon Lime','Orange Mango'],
                'Units':[s1_,s2_,s3_],'Share %':[round(s1_/tot_*100,1),round(s2_/tot_*100,1),round(s3_/tot_*100,1)]})

        with pd.ExcelWriter(buf,engine='openpyxl') as wr:
            for sn,df in sheets.items(): df.to_excel(wr,sheet_name=sn,index=False)

        buf.seek(0); wb=load_workbook(buf)
        from openpyxl.drawing.image import Image as XLI
        for sn,ib in charts.items():
            ws=wb[sn] if sn in wb.sheetnames else wb.create_sheet(sn)
            ib.seek(0); img=XLI(ib); img.width=900; img.height=350
            ws.add_image(img,f'A{ws.max_row+3}')

        gf=PatternFill('solid',fgColor='F0B43C')
        a1=PatternFill('solid',fgColor='1A1A1A'); a2=PatternFill('solid',fgColor='222222')
        th=Border(left=Side(style='thin',color='2C2C2C'),right=Side(style='thin',color='2C2C2C'),
                  top=Side(style='thin',color='2C2C2C'),bottom=Side(style='thin',color='2C2C2C'))
        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.fill=gf; cell.font=Font(bold=True,color='000000',size=10)
                cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=th
            ws.row_dimensions[1].height=26
            for i,row in enumerate(ws.iter_rows(min_row=2),1):
                bg=a1 if i%2 else a2
                for cell in row:
                    cell.fill=bg; cell.border=th
                    cell.alignment=Alignment(vertical='center',wrap_text=True)
                    cell.font=Font(color='EEEEEE',size=9)
            for j in range(1,ws.max_column+1):
                ws.column_dimensions[get_column_letter(j)].width=20
            ws.freeze_panes='A2'

        fin=io.BytesIO(); wb.save(fin); fin.seek(0); return fin

    with st.spinner("Building report with charts..."):
        xl=build_xl(r,m,CE)

    st.download_button("⬇️  Download Excel Report", data=xl,
        file_name="Mariano_Events_Campaign_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    items=["✅ KPI Summary — all metrics with formulas",
           "✅ All Events — clean raw data",
           "✅ By State + embedded chart",
           "✅ By BA Ambassador + embedded chart",
           "✅ By Day of Week + chart",
           "✅ Top 20 & Low 20 Stores + chart",
           "✅ Stock Issues list",
           "✅ Sampling Tiers + chart",
           "✅ SKU Breakdown + chart",
           "✅ Coupon Analysis + chart"]
    for it in items:
        st.markdown(f'<div style="color:{C["gray"]};font-size:0.78rem;padding:2px 0;">{it}</div>',unsafe_allow_html=True)
