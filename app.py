import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import warnings
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings('ignore')

# ── Page config ────────────────────────────────────────────────
st.set_page_config(
    page_title="Campaign Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Brand colors ───────────────────────────────────────────────
ORG  = '#F0B43C'
GLD  = '#D9980E'
GRN  = '#3FA34D'
RED  = '#C0392B'
TEAL = '#288C8C'
BLK  = '#0D0D0D'
DARK = '#1A1A1A'
PUR  = '#8C1478'
PALETTE = [ORG, TEAL, PUR, GLD, '#50B0B0', '#B050A0', GRN, RED,
           '#F0D080','#E74C3C','#1ABC9C','#3498DB','#F39C12','#2ECC71']

# ── CSS ───────────────────────────────────────────────────────
st.markdown(f"""
<style>
    html, body, [class*="css"] {{ background-color: {BLK}; color: #EEEEEE; }}
    .main {{ background-color: {BLK}; }}
    .block-container {{ padding: 1.5rem 2rem; }}
    section[data-testid="stSidebar"] {{ background-color: #111111; border-right: 1px solid #333; }}
    section[data-testid="stSidebar"] * {{ color: #EEEEEE !important; }}
    h1 {{ color: {ORG}; font-family: Arial Black; letter-spacing: 2px; font-size: 2rem; }}
    h2 {{ color: {GLD}; font-family: Arial; font-size: 1.3rem; }}
    h3 {{ color: #EEEEEE; font-family: Arial; font-size: 1.1rem; }}
    .kpi-card {{
        background: linear-gradient(135deg, {ORG}, {GLD});
        border-radius: 12px; padding: 18px 12px; text-align: center;
        color: white; margin-bottom: 8px;
    }}
    .kpi-val {{ font-size: 1.8rem; font-weight: 900; font-family: Arial Black; }}
    .kpi-lbl {{ font-size: 0.72rem; font-weight: bold; margin-top: 4px; opacity: 0.92; letter-spacing: 1px; }}
    .kpi-card-teal {{
        background: linear-gradient(135deg, {TEAL}, #1A5A55);
        border-radius: 12px; padding: 18px 12px; text-align: center;
        color: white; margin-bottom: 8px;
    }}
    .stDataFrame {{ border-radius: 8px; }}
    .stButton > button {{
        background: linear-gradient(135deg, {ORG}, {GLD});
        color: white; border: none; border-radius: 8px;
        font-weight: bold; padding: 0.5rem 1.5rem;
    }}
    .stDownloadButton > button {{
        background: linear-gradient(135deg, {GRN}, #1A6B2A);
        color: white; border: none; border-radius: 8px;
        font-weight: bold; padding: 0.5rem 1.5rem; width: 100%;
    }}
    .section-header {{
        background: linear-gradient(90deg, {ORG}22, transparent);
        border-left: 4px solid {ORG}; padding: 8px 14px;
        border-radius: 0 8px 8px 0; margin: 18px 0 10px 0;
        color: {GLD}; font-weight: bold; font-size: 1.05rem;
    }}
    div[data-testid="metric-container"] {{ background: {DARK}; border-radius: 8px; padding: 8px; }}
    .stSelectbox > div {{ background: {DARK}; }}
    .stMultiSelect > div {{ background: {DARK}; }}
    hr {{ border-color: #333; }}
    .stTabs [data-baseweb="tab-list"] {{ background-color: {DARK}; border-radius: 8px; }}
    .stTabs [data-baseweb="tab"] {{ color: #AAAAAA; }}
    .stTabs [aria-selected="true"] {{ color: {ORG} !important; border-bottom: 2px solid {ORG}; }}
</style>
""", unsafe_allow_html=True)

# ── Chart styling ──────────────────────────────────────────────
def setup_ax(ax, horizontal=False):
    ax.set_facecolor('#1A1A1A')
    ax.figure.patch.set_facecolor('#0D0D0D')
    ax.tick_params(colors='#CCCCCC', labelsize=9)
    ax.xaxis.label.set_color('#CCCCCC')
    ax.yaxis.label.set_color('#CCCCCC')
    ax.title.set_color(GLD)
    for spine in ax.spines.values():
        spine.set_edgecolor('#333333')
    if horizontal:
        ax.grid(axis='x', color='#2A2A2A', linewidth=0.6)
        ax.grid(axis='y', visible=False)
    else:
        ax.grid(axis='y', color='#2A2A2A', linewidth=0.6)
        ax.grid(axis='x', visible=False)
    ax.set_axisbelow(True)

def bar_labels(ax, bars, fmt='{:.0f}', horizontal=False, color='white', fontsize=9):
    for bar in bars:
        val = bar.get_width() if horizontal else bar.get_height()
        if val > 0:
            if horizontal:
                ax.text(val + val*0.01 + 0.05, bar.get_y() + bar.get_height()/2,
                        fmt.format(val), va='center', color=color, fontsize=fontsize, fontweight='bold')
            else:
                ax.text(bar.get_x() + bar.get_width()/2, val + val*0.01 + 0.05,
                        fmt.format(val), ha='center', va='bottom', color=color, fontsize=fontsize, fontweight='bold')

# ── Helper: parse coupons ──────────────────────────────────────
def parse_coupons(x):
    if pd.isna(x): return 0
    nums = re.findall(r'\d+', str(x))
    return int(nums[0]) if nums else 0

# ── Helper: auto-detect columns ───────────────────────────────
def best_match(cols, hints):
    cl = [c.lower().strip() for c in cols]
    for h in hints:
        for i, c in enumerate(cl):
            if h.lower() in c:
                return cols[i]
    return None

def guess_mapping(cols):
    return {
        'date':        best_match(cols, ['date','event date']),
        'campaign':    best_match(cols, ['campaign name','campaign','program','event name']),
        'ba_name':     best_match(cols, ['ba name','ambassador','rep name','staff name']),
        'venue':       best_match(cols, ['venue name','store name','location name','venue']),
        'city':        best_match(cols, ['city']),
        'state':       best_match(cols, ['state']),
        'region':      best_match(cols, ['region','territory','area','zone']),
        'units':       best_match(cols, ['# total sales','total sales','units sold','# of units']),
        'revenue':     best_match(cols, ['$ total sales','revenue','dollar','$ sales']),
        'sampled':     best_match(cols, ['total sampled','sampled','samples given','samples handed']),
        'hours':       best_match(cols, ['ba hours','hours worked','hours']),
        'coupons':     best_match(cols, ['coupon']),
        'sku1':        best_match(cols, ['pineapple','pnapple','guava']),
        'sku2':        best_match(cols, ['lemon lime','lemon']),
        'sku3':        best_match(cols, ['orange mango','mango']),
        'familiarity': best_match(cols, ['familiar','brand awareness','heard of']),
        'recommend':   best_match(cols, ['recommend','likelihood']),
        'taste':       best_match(cols, ['taste','flavor rating']),
        'shelf':       best_match(cols, ['enough product','shelf']),
        'impact':      best_match(cols, ['impacted','impact','affect']),
        'comments':    best_match(cols, ['customer comments','comments','feedback']),
    }

# ── Helper: prepare dataframe ──────────────────────────────────
def prepare(df, m):
    r = df.copy()
    r['_units']    = pd.to_numeric(r[m['units']], errors='coerce').fillna(0)
    r['_sampled']  = pd.to_numeric(r[m['sampled']], errors='coerce').fillna(0)
    r['_revenue']  = pd.to_numeric(r[m['revenue']], errors='coerce').fillna(0) if m.get('revenue') else 0
    r['_hours']    = pd.to_numeric(r[m['hours']], errors='coerce').fillna(0) if m.get('hours') else 0
    r['_coupons']  = r[m['coupons']].apply(parse_coupons) if m.get('coupons') else 0
    r['_date']     = pd.to_datetime(r[m['date']], errors='coerce')
    r['_dow']      = r['_date'].dt.day_name()
    r['_month']    = r['_date'].dt.strftime('%b %Y')
    r['_conv']     = (r['_units'] / r['_sampled'].replace(0, np.nan) * 100).round(1)
    for k in ['sku1','sku2','sku3']:
        r[f'_{k}'] = pd.to_numeric(r[m[k]], errors='coerce').fillna(0) if m.get(k) else 0
    r['_campaign'] = r[m['campaign']].fillna('Unknown').astype(str) if m.get('campaign') else 'Campaign'
    r['_state']    = r[m['state']].fillna('Unknown').astype(str) if m.get('state') else 'Unknown'
    r['_city']     = r[m['city']].fillna('Unknown').astype(str) if m.get('city') else 'Unknown'
    r['_ba']       = r[m['ba_name']].fillna('Unknown').astype(str) if m.get('ba_name') else 'Unknown'
    r['_venue']    = r[m['venue']].fillna('Unknown').astype(str) if m.get('venue') else 'Unknown'
    return r

# ── Helper: compute KPIs ───────────────────────────────────────
def compute_kpis(r):
    n     = len(r)
    units = r['_units'].sum()
    samp  = r['_sampled'].sum()
    rev   = r['_revenue'].sum()
    hrs   = r['_hours'].sum()
    conv  = round(units/samp*100, 2) if samp > 0 else 0
    avg_u = round(units/n, 2)  if n > 0 else 0
    avg_s = round(samp/n, 1)   if n > 0 else 0
    uph   = round(units/hrs, 2) if hrs > 0 else 0
    avg_r = round(rev/n, 2)     if n > 0 else 0
    coup  = r['_coupons'].sum()
    coup_e= (r['_coupons']>0).sum()
    wc    = r[r['_coupons']>0]['_units'].mean() if (r['_coupons']>0).any() else 0
    nc    = r[r['_coupons']==0]['_units'].mean() if (r['_coupons']==0).any() else 0
    s1=r['_sku1'].sum(); s2=r['_sku2'].sum(); s3=r['_sku3'].sum()
    return dict(n=int(n), units=float(units), samp=float(samp), rev=float(rev),
                hrs=float(hrs), conv=conv, avg_u=avg_u, avg_s=avg_s, uph=uph,
                avg_r=avg_r, coup=float(coup), coup_e=int(coup_e),
                wc=round(wc,2), nc=round(nc,2), s1=float(s1), s2=float(s2), s3=float(s3))

def kpi_card(col, val, label, teal=False):
    cls = 'kpi-card-teal' if teal else 'kpi-card'
    col.markdown(f'<div class="{cls}"><div class="kpi-val">{val}</div><div class="kpi-lbl">{label}</div></div>', unsafe_allow_html=True)

def section(title):
    st.markdown(f'<div class="section-header">📊 {title}</div>', unsafe_allow_html=True)

# ── Excel export ───────────────────────────────────────────────
def build_excel(r, m, charts_dict):
    buf = io.BytesIO()
    ORG_F = PatternFill('solid', fgColor='E8641C')
    ALT1  = PatternFill('solid', fgColor='FFFDF6')
    ALT2  = PatternFill('solid', fgColor='FEF3E4')
    GRN_F = PatternFill('solid', fgColor='D5F5E3')
    RED_F = PatternFill('solid', fgColor='FADBD8')
    thin  = Border(left=Side(style='thin',color='DDDDDD'), right=Side(style='thin',color='DDDDDD'),
                   top=Side(style='thin',color='DDDDDD'),  bottom=Side(style='thin',color='DDDDDD'))

    def style_ws(ws):
        for cell in ws[1]:
            cell.fill = ORG_F; cell.font = Font(bold=True, color='FFFFFF', size=10)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin
        ws.row_dimensions[1].height = 26
        for i, row in enumerate(ws.iter_rows(min_row=2), 1):
            bg = ALT1 if i%2 else ALT2
            for cell in row:
                cell.fill = bg; cell.border = thin
                cell.alignment = Alignment(vertical='center', wrap_text=True)
        for j in range(1, ws.max_column+1):
            ws.column_dimensions[get_column_letter(j)].width = 20
        ws.freeze_panes = 'A2'

    k = compute_kpis(r)
    sheets = {}

    # KPIs
    sheets['KPI Summary'] = pd.DataFrame([
        ['Total Events', k['n'], 'COUNT of all rows'],
        ['Total Samples', f"{k['samp']:,.0f}", 'SUM Sampled column'],
        ['Total Units Sold', f"{k['units']:,.0f}", 'SUM Units column'],
        ['Total Revenue', f"${k['rev']:,.2f}", 'SUM Revenue column'],
        ['Conversion Rate', f"{k['conv']}%", 'Units / Sampled x 100'],
        ['Avg Units / Demo', k['avg_u'], 'Units / Events'],
        ['Avg Sampled / Demo', k['avg_s'], 'Sampled / Events'],
        ['Avg Revenue / Demo', f"${k['avg_r']:.2f}", 'Revenue / Events'],
        ['Units / BA Hour', k['uph'], 'Units / Hours'],
        ['Total Coupons Given', f"{k['coup']:,.0f}", 'SUM Coupons'],
        ['Events with Coupons', f"{k['coup_e']} of {k['n']}", 'Count Coupons > 0'],
        ['Avg Units WITH Coupons', k['wc'], 'AVERAGE where Coupons > 0'],
        ['Avg Units WITHOUT Coupons', k['nc'], 'AVERAGE where Coupons = 0'],
        ['Coupon Lift', f"+{round((k['wc']/k['nc']-1)*100,1)}%" if k['nc']>0 else 'N/A', '(With/Without-1)x100'],
    ], columns=['Metric', 'Value', 'How Calculated'])

    # All events
    raw_cols = [m[v] for v in ['date','ba_name','venue','city','state','units','revenue','sampled','hours','coupons'] if m.get(v)]
    clean = r[raw_cols].copy()
    clean['Conversion %'] = r['_conv'].values
    sheets['All Events'] = clean

    # By State
    g = r.groupby('_state').agg(Events=('_units','size'), Units=('_units','sum'),
        Revenue=('_revenue','sum'), Sampled=('_sampled','sum')).reset_index()
    g['Avg Units/Demo'] = (g.Units/g.Events).round(2)
    g['Conv %'] = (g.Units/g.Sampled*100).round(2)
    g.columns = ['State','Events','Units','Revenue ($)','Sampled','Avg Units/Demo','Conv %']
    sheets['By State'] = g.sort_values('Units', ascending=False)

    # By BA
    g = r.groupby('_ba').agg(Events=('_units','size'), Units=('_units','sum'),
        Revenue=('_revenue','sum'), Sampled=('_sampled','sum')).reset_index()
    g['Avg Units/Demo'] = (g.Units/g.Events).round(2)
    g['Conv %'] = (g.Units/g.Sampled*100).round(2)
    g.columns = ['BA Name','Events','Units','Revenue ($)','Sampled','Avg Units/Demo','Conv %']
    sheets['By BA'] = g.sort_values('Avg Units/Demo', ascending=False)

    # By Day
    order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    g = r.groupby('_dow').agg(Events=('_units','size'), Units=('_units','sum'),
        Sampled=('_sampled','sum')).reset_index()
    g['Avg Units/Demo'] = (g.Units/g.Events).round(2)
    g['Conv %'] = (g.Units/g.Sampled*100).round(2)
    g['sort'] = g['_dow'].map({d:i for i,d in enumerate(order)})
    g = g.sort_values('sort').drop('sort',axis=1)
    g.columns = ['Day','Events','Units','Sampled','Avg Units/Demo','Conv %']
    sheets['By Day'] = g

    # Top/Low Stores
    g = r.groupby(['_venue','_city','_state']).agg(Events=('_units','size'),
        Units=('_units','sum'), Sampled=('_sampled','sum'), Revenue=('_revenue','sum')).reset_index()
    g['Avg Units/Demo'] = (g.Units/g.Events).round(2)
    g['Conv %'] = (g.Units/g.Sampled*100).round(2)
    g.columns = ['Store','City','State','Events','Units','Sampled','Revenue ($)','Avg Units/Demo','Conv %']
    sheets['Top 20 Stores'] = g.sort_values('Units', ascending=False).head(20)
    sheets['Low 20 Stores'] = g.sort_values('Units').head(20)

    # Sampling tiers
    r2 = r.copy()
    r2['tier'] = pd.cut(r2['_sampled'], bins=[0,50,65,80,300], labels=['50 or fewer','51-65','66-80','80+'])
    g = r2.groupby('tier', observed=True).agg(Events=('_units','size'),
        Units=('_units','sum'), Sampled=('_sampled','sum')).reset_index()
    g['Avg Units/Demo'] = (g.Units/g.Events).round(2)
    g.columns = ['Sampling Tier','Events','Units','Sampled','Avg Units/Demo']
    sheets['Sampling Tiers'] = g

    # SKU
    s1=k['s1']; s2=k['s2']; s3=k['s3']; tot=s1+s2+s3
    if tot > 0:
        sheets['SKU Breakdown'] = pd.DataFrame({
            'SKU': ['Pineapple Guava','Lemon Lime','Orange Mango'],
            'Units': [s1,s2,s3],
            'Share %': [round(s1/tot*100,1), round(s2/tot*100,1), round(s3/tot*100,1)]
        })

    # Coupons
    sheets['Coupon Analysis'] = pd.DataFrame({
        'Group': ['With Coupons','Without Coupons'],
        'Events': [(r['_coupons']>0).sum(), (r['_coupons']==0).sum()],
        'Avg Units/Demo': [k['wc'], k['nc']],
        'Lift': [f"+{round((k['wc']/k['nc']-1)*100,1)}%" if k['nc']>0 else 'N/A','baseline']
    })

    # Write all sheets
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for sname, df in sheets.items():
            df.to_excel(writer, sheet_name=sname, index=False)

    buf.seek(0)
    wb = load_workbook(buf)

    # Add charts to Excel
    from openpyxl.drawing.image import Image as XLImage
    for sname, img_buf in charts_dict.items():
        if sname in wb.sheetnames:
            ws = wb[sname]
        else:
            ws = wb.create_sheet(sname)
        img_buf.seek(0)
        img = XLImage(img_buf)
        img.width  = 900
        img.height = 350
        # Place chart below the data
        last_row = ws.max_row + 3
        ws.add_image(img, f'A{last_row}')

    for ws in wb.worksheets:
        style_ws(ws)

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final

# ── Generate chart and return as image buffer ──────────────────
def make_chart_buf(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=120, bbox_inches='tight',
                facecolor=BLK, edgecolor='none')
    buf.seek(0)
    plt.close(fig)
    return buf

# ═══════════════════════════════════════════════════════════════
# MAIN APP
# ═══════════════════════════════════════════════════════════════

import base64, os
_lp = os.path.join(os.path.dirname(__file__), 'me_logo.png')
_logo64 = base64.b64encode(open(_lp,'rb').read()).decode() if os.path.exists(_lp) else None
if _logo64:
    st.markdown(f"""<div style="display:flex;align-items:center;gap:18px;margin-bottom:6px;">
      <img src="data:image/png;base64,{_logo64}" style="height:64px;">
      <div>
        <div style="font-size:1.7rem;font-weight:900;letter-spacing:2px;color:{ORG};font-family:Arial Black;">MARIANO EVENTS</div>
        <div style="font-size:0.72rem;color:#888;letter-spacing:2px;">CAMPAIGN INTELLIGENCE PLATFORM</div>
      </div>
    </div>""", unsafe_allow_html=True)
else:
    st.markdown("# 📊 CAMPAIGN ANALYZER")
st.markdown("Upload your Promomash data, filter by campaign, date or state, and get instant insights.")
st.markdown("---")

# ── SIDEBAR ────────────────────────────────────────────────────
with st.sidebar:
    if _logo64:
        st.markdown(f'<div style="text-align:center;padding:8px 0 14px 0;"><img src="data:image/png;base64,{_logo64}" style="width:110px;"></div>', unsafe_allow_html=True)
    st.markdown(f"## ⚙️ Controls")
    st.markdown("---")

    uploaded_files = st.file_uploader(
        "Upload Promomash file(s)",
        type=['xlsx','xls','csv'],
        accept_multiple_files=True,
        help="Upload one file per campaign, or one file with all campaigns"
    )

    if not uploaded_files:
        st.info("Upload a file above to get started.")
        st.stop()

# ── Load all uploaded files ────────────────────────────────────
@st.cache_data
def load_file(fname, data):
    try:
        xl = pd.ExcelFile(io.BytesIO(data))
        df = xl.parse('Raw') if 'Raw' in xl.sheet_names else xl.parse(xl.sheet_names[0])
    except:
        df = pd.read_csv(io.BytesIO(data))
    return df[df.iloc[:,0].notna()].copy()

all_dfs = {}
for uf in uploaded_files:
    df = load_file(uf.name, uf.read())
    cname = uf.name.replace('.xlsx','').replace('.xls','').replace('.csv','')
    all_dfs[cname] = df

combined_raw = pd.concat(all_dfs.values(), ignore_index=True)
cols = combined_raw.columns.tolist()

# ── Column mapping ─────────────────────────────────────────────
if 'mapping' not in st.session_state:
    st.session_state.mapping = guess_mapping(cols)
m = st.session_state.mapping

# Check required fields and show mapping UI if needed
required = ['date','ba_name','units','sampled']
missing_required = [k for k in required if not m.get(k)]

if missing_required or st.sidebar.checkbox("⚙️ Edit column mapping"):
    st.markdown("## ⚙️ Column Mapping")
    st.markdown("Map your file columns to the analysis fields. Required fields are marked ✱")
    none_opt = '— Not in file —'
    all_opts = [none_opt] + cols
    left, right = st.columns(2)
    new_m = {}
    fields = list({
        'date':('Date ✱',True), 'ba_name':('BA Name ✱',True),
        'units':('Units Sold ✱',True), 'sampled':('Total Sampled ✱',True),
        'revenue':('Revenue ($)',False), 'hours':('BA Hours',False),
        'coupons':('Coupons',False), 'campaign':('Campaign Name',False),
        'venue':('Store / Venue',False), 'city':('City',False),
        'state':('State',False), 'region':('Region',False),
        'sku1':('SKU 1 - Pineapple Guava',False),
        'sku2':('SKU 2 - Lemon Lime',False),
        'sku3':('SKU 3 - Orange Mango',False),
        'familiarity':('Brand Familiarity',False),
        'recommend':('Would Recommend',False),
        'taste':('Taste Rating',False),
        'impact':('What Impacted Sales',False),
        'comments':('Customer Comments',False),
    }.items())
    for i,(key,(label,req)) in enumerate(fields):
        container = left if i < len(fields)//2 else right
        guess = m.get(key)
        idx = all_opts.index(guess) if guess and guess in all_opts else 0
        sel = container.selectbox(label, all_opts, index=idx, key=f'col_{key}')
        new_m[key] = None if sel == none_opt else sel
    if st.button("✅ Save Mapping"):
        st.session_state.mapping = new_m
        m = new_m
        st.rerun()
    st.stop()

# ── Prepare data ───────────────────────────────────────────────
r_full = combined_raw[pd.to_numeric(combined_raw[m['units']], errors='coerce').notna()].copy()
r_full = prepare(r_full, m)

# ── SIDEBAR FILTERS ────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🎯 Filters")

    # Campaign filter
    campaigns = sorted(r_full['_campaign'].unique().tolist())
    if len(campaigns) > 1:
        sel_campaigns = st.multiselect("Campaign", campaigns, default=campaigns,
                                        help="Select one or more campaigns")
    else:
        sel_campaigns = campaigns

    # Date filter
    min_d = r_full['_date'].min().date()
    max_d = r_full['_date'].max().date()
    date_range = st.date_input("Date Range", value=(min_d, max_d),
                                min_value=min_d, max_value=max_d)

    # State filter
    states = sorted(r_full['_state'].unique().tolist())
    sel_states = st.multiselect("State", states, default=states)

    # BA filter
    bas = sorted(r_full['_ba'].unique().tolist())
    sel_bas = st.multiselect("BA / Ambassador", bas, default=bas)

    st.markdown("---")
    st.markdown("### 📊 Chart Settings")
    top_n = st.slider("Top / Bottom N stores and BAs", 5, 20, 10)
    min_events = st.slider("Min events for BA ranking", 1, 5, 2,
                            help="BAs with fewer than this many events are excluded from rankings")

# ── Apply filters ──────────────────────────────────────────────
r = r_full.copy()
r = r[r['_campaign'].isin(sel_campaigns)]
if len(date_range) == 2:
    r = r[(r['_date'].dt.date >= date_range[0]) & (r['_date'].dt.date <= date_range[1])]
r = r[r['_state'].isin(sel_states)]
r = r[r['_ba'].isin(sel_bas)]

if len(r) == 0:
    st.warning("No data matches the current filters. Please adjust the filters in the sidebar.")
    st.stop()

# ── KPI CARDS ──────────────────────────────────────────────────
k = compute_kpis(r)
st.markdown(f"### Showing **{k['n']:,} events** from **{date_range[0]}** to **{date_range[1] if len(date_range)==2 else '...'}**")

c1,c2,c3,c4,c5,c6,c7,c8 = st.columns(8)
kpi_card(c1, f"{k['n']:,}",      "EVENTS")
kpi_card(c2, f"{k['samp']:,.0f}","SAMPLES")
kpi_card(c3, f"{k['units']:,.0f}","UNITS SOLD")
kpi_card(c4, f"${k['rev']:,.0f}","REVENUE")
kpi_card(c5, f"{k['conv']}%",    "CONVERSION", teal=True)
kpi_card(c6, f"{k['avg_u']}",    "AVG UNITS/DEMO", teal=True)
kpi_card(c7, f"{k['avg_s']}",    "AVG SAMPLED", teal=True)
kpi_card(c8, f"{k['uph']}",      "UNITS/HOUR", teal=True)

# ── TABS ───────────────────────────────────────────────────────
st.markdown("---")
tabs = st.tabs(["📅 Trends", "🏆 Top & Low BAs", "🏪 Stores", "🗺️ Geography",
                "🎯 Sampling & Coupons", "🥤 SKUs", "📋 Campaign Compare", "📥 Download"])

charts_for_excel = {}

# ── TAB 1: TRENDS ─────────────────────────────────────────────
with tabs[0]:
    section("Daily Performance Trend")
    daily = r.groupby('_date').agg(events=('_units','size'), units=('_units','sum'),
                                    sampled=('_sampled','sum')).reset_index()
    daily['avg'] = (daily.units/daily.events).round(1)

    fig, axes = plt.subplots(1,2,figsize=(14,4))
    fig.patch.set_facecolor(BLK)
    ax = axes[0]; setup_ax(ax)
    ax.bar(daily['_date'], daily['units'], color=ORG, alpha=0.85, width=0.7)
    ax.set_title('Total Units Sold per Day', fontsize=11, pad=8)
    ax.set_ylabel('Units Sold')
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=8)

    ax = axes[1]; setup_ax(ax)
    ax.bar(daily['_date'], daily['avg'], color=GLD, alpha=0.85, width=0.7)
    ax.set_title('Avg Units per Demo per Day', fontsize=11, pad=8)
    ax.set_ylabel('Avg Units / Demo')
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)
    charts_for_excel['By State'] = make_chart_buf(fig)

    section("Day of Week Performance")
    order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    dow = r.groupby('_dow').agg(events=('_units','size'), units=('_units','sum'),
                                 sampled=('_sampled','sum')).reset_index()
    dow['avg']  = (dow.units/dow.events).round(2)
    dow['conv'] = (dow.units/dow.sampled*100).round(2)
    dow['sort'] = dow['_dow'].map({d:i for i,d in enumerate(order)})
    dow = dow.sort_values('sort')

    fig, axes = plt.subplots(1,2,figsize=(14,4))
    fig.patch.set_facecolor(BLK)
    ax = axes[0]; setup_ax(ax)
    colors = [GRN if d in ['Friday','Saturday','Sunday'] else ORG for d in dow['_dow']]
    bars = ax.bar(dow['_dow'], dow['avg'], color=colors, edgecolor='none')
    ax.set_title('Avg Units/Demo by Day of Week', fontsize=11, pad=8)
    ax.set_ylabel('Avg Units / Demo')
    bar_labels(ax, bars, '{:.1f}')
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha='right')

    ax = axes[1]; setup_ax(ax)
    bars = ax.bar(dow['_dow'], dow['conv'], color=TEAL, edgecolor='none')
    ax.set_title('Conversion % by Day of Week', fontsize=11, pad=8)
    ax.set_ylabel('Conversion %')
    bar_labels(ax, bars, '{:.1f}%')
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=30, ha='right')
    plt.tight_layout()
    st.pyplot(fig)
    charts_for_excel['By Day'] = make_chart_buf(fig)

    st.dataframe(dow[['_dow','events','units','avg','conv']].rename(columns={
        '_dow':'Day','events':'Events','units':'Units','avg':'Avg Units/Demo','conv':'Conv %'
    }), use_container_width=True, hide_index=True)

# ── TAB 2: TOP & LOW BAs ──────────────────────────────────────
with tabs[1]:
    ba = r.groupby('_ba').agg(events=('_units','size'), units=('_units','sum'),
        revenue=('_revenue','sum'), sampled=('_sampled','sum')).reset_index()
    ba['avg']  = (ba.units/ba.events).round(2)
    ba['conv'] = (ba.units/ba.sampled*100).round(2)
    ba['rev_per'] = (ba.revenue/ba.events).round(2)
    ba_q = ba[ba.events >= min_events]
    top_ba  = ba_q.sort_values('avg', ascending=False).head(top_n)
    low_ba  = ba_q.sort_values('avg').head(top_n)

    section(f"Top {top_n} Brand Ambassadors")
    fig, axes = plt.subplots(1,2,figsize=(14,max(4, top_n*0.45)))
    fig.patch.set_facecolor(BLK)
    ax = axes[0]; setup_ax(ax, horizontal=True)
    colors_top = [ORG if i==0 else GLD if i==1 else '#C8A87A' for i in range(len(top_ba))]
    bars = ax.barh(top_ba['_ba'], top_ba['avg'], color=colors_top, edgecolor='none')
    ax.set_title(f'Top {top_n} BAs — Avg Units/Demo', fontsize=11, pad=8)
    ax.set_xlabel('Avg Units / Demo')
    bar_labels(ax, bars, '{:.2f}', horizontal=True)
    ax.invert_yaxis()

    ax = axes[1]; setup_ax(ax, horizontal=True)
    bars = ax.barh(low_ba['_ba'], low_ba['avg'], color=RED, alpha=0.8, edgecolor='none')
    ax.set_title(f'Bottom {top_n} BAs — Avg Units/Demo', fontsize=11, pad=8)
    ax.set_xlabel('Avg Units / Demo')
    bar_labels(ax, bars, '{:.2f}', horizontal=True)
    ax.invert_yaxis()
    plt.tight_layout()
    st.pyplot(fig)
    charts_for_excel['By BA'] = make_chart_buf(fig)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Top {top_n} BAs**")
        st.dataframe(top_ba[['_ba','events','units','avg','conv']].rename(columns={
            '_ba':'BA Name','events':'Events','units':'Units','avg':'Avg Units/Demo','conv':'Conv %'
        }), use_container_width=True, hide_index=True)
    with col2:
        st.markdown(f"**Bottom {top_n} BAs**")
        st.dataframe(low_ba[['_ba','events','units','avg','conv']].rename(columns={
            '_ba':'BA Name','events':'Events','units':'Units','avg':'Avg Units/Demo','conv':'Conv %'
        }), use_container_width=True, hide_index=True)

# ── TAB 3: STORES ─────────────────────────────────────────────
with tabs[2]:
    st_df = r.groupby(['_venue','_city','_state']).agg(
        events=('_units','size'), units=('_units','sum'),
        revenue=('_revenue','sum'), sampled=('_sampled','sum')).reset_index()
    st_df['avg']  = (st_df.units/st_df.events).round(2)
    st_df['conv'] = (st_df.units/st_df.sampled*100).round(2)
    top_st = st_df.sort_values('units', ascending=False).head(top_n)
    low_st = st_df.sort_values('units').head(top_n)
    top_st['label'] = top_st['_city'] + ', ' + top_st['_state']
    low_st['label'] = low_st['_city'] + ', ' + low_st['_state']

    section(f"Top & Bottom {top_n} Stores")
    fig, axes = plt.subplots(1,2,figsize=(14,max(4,top_n*0.45)))
    fig.patch.set_facecolor(BLK)
    ax = axes[0]; setup_ax(ax, horizontal=True)
    c_top = [ORG if i==0 else GLD if i==1 else '#C8A87A' for i in range(len(top_st))]
    bars = ax.barh(top_st['label'], top_st['units'], color=c_top, edgecolor='none')
    ax.set_title(f'Top {top_n} Stores — Total Units', fontsize=11, pad=8)
    ax.set_xlabel('Total Units Sold')
    bar_labels(ax, bars, '{:.0f}', horizontal=True)
    ax.invert_yaxis()

    ax = axes[1]; setup_ax(ax, horizontal=True)
    bars = ax.barh(low_st['label'], low_st['units'], color=RED, alpha=0.8, edgecolor='none')
    ax.set_title(f'Bottom {top_n} Stores — Total Units', fontsize=11, pad=8)
    ax.set_xlabel('Total Units Sold')
    bar_labels(ax, bars, '{:.0f}', horizontal=True)
    ax.invert_yaxis()
    plt.tight_layout()
    st.pyplot(fig)
    charts_for_excel['Top 20 Stores'] = make_chart_buf(fig)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Top {top_n} Stores**")
        st.dataframe(top_st[['label','_venue','events','units','avg','conv']].rename(columns={
            'label':'Location','_venue':'Store','events':'Events','units':'Units',
            'avg':'Avg Units/Demo','conv':'Conv %'
        }), use_container_width=True, hide_index=True)
    with col2:
        st.markdown(f"**Bottom {top_n} Stores**")
        st.dataframe(low_st[['label','_venue','events','units','avg','conv']].rename(columns={
            'label':'Location','_venue':'Store','events':'Events','units':'Units',
            'avg':'Avg Units/Demo','conv':'Conv %'
        }), use_container_width=True, hide_index=True)

# ── TAB 4: GEOGRAPHY ──────────────────────────────────────────
with tabs[3]:
    sg = r.groupby('_state').agg(events=('_units','size'), units=('_units','sum'),
        sampled=('_sampled','sum'), revenue=('_revenue','sum')).reset_index()
    sg['avg']  = (sg.units/sg.events).round(2)
    sg['conv'] = (sg.units/sg.sampled*100).round(2)
    sg = sg.sort_values('units', ascending=False)
    avg_line = sg['avg'].mean()

    section("State Performance")
    fig, axes = plt.subplots(1,2,figsize=(14,4))
    fig.patch.set_facecolor(BLK)
    ax = axes[0]; setup_ax(ax)
    colors = [GRN if v >= avg_line else RED for v in sg['avg']]
    bars = ax.bar(sg['_state'], sg['avg'], color=colors, edgecolor='none')
    ax.axhline(avg_line, color=GLD, linewidth=1.5, linestyle='--',
               label=f'Avg {avg_line:.1f}')
    ax.legend(facecolor=DARK, labelcolor='white', fontsize=9)
    ax.set_title('Avg Units/Demo by State', fontsize=11, pad=8)
    ax.set_ylabel('Avg Units / Demo')
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=8)

    ax = axes[1]; setup_ax(ax)
    bars = ax.bar(sg['_state'], sg['events'], color=TEAL, edgecolor='none')
    ax.set_title('Events by State', fontsize=11, pad=8)
    ax.set_ylabel('Number of Events')
    bar_labels(ax, bars, '{:.0f}', fontsize=8)
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right', fontsize=8)
    plt.tight_layout()
    st.pyplot(fig)

    st.dataframe(sg.rename(columns={'_state':'State','events':'Events','units':'Units',
        'sampled':'Sampled','revenue':'Revenue ($)','avg':'Avg Units/Demo','conv':'Conv %'}),
        use_container_width=True, hide_index=True)

# ── TAB 5: SAMPLING & COUPONS ─────────────────────────────────
with tabs[4]:
    section("Sampling Volume Impact")
    r2 = r.copy()
    r2['tier'] = pd.cut(r2['_sampled'], bins=[0,50,65,80,300],
                         labels=['50 or fewer','51-65','66-80','80+'])
    sg2 = r2.groupby('tier', observed=True).agg(events=('_units','size'),
        units=('_units','sum'), sampled=('_sampled','sum')).reset_index()
    sg2['avg']  = (sg2.units/sg2.events).round(2)
    sg2['conv'] = (sg2.units/sg2.sampled*100).round(2)

    fig, axes = plt.subplots(1,2,figsize=(14,4))
    fig.patch.set_facecolor(BLK)
    ax = axes[0]; setup_ax(ax)
    tier_colors = [RED, GLD, ORG, GRN]
    bars = ax.bar(sg2['tier'].astype(str), sg2['avg'], color=tier_colors, edgecolor='none')
    ax.set_title('Avg Units/Demo by Sampling Volume', fontsize=11, pad=8)
    ax.set_xlabel('Customers Sampled per Event')
    ax.set_ylabel('Avg Units / Demo')
    bar_labels(ax, bars, '{:.2f}', fontsize=10)

    ax = axes[1]; setup_ax(ax)
    ax.pie(sg2['events'], labels=sg2['tier'].astype(str), colors=tier_colors,
           autopct='%1.0f%%', startangle=90,
           textprops={'color':'white','fontsize':10})
    ax.set_title('Events by Sampling Tier', fontsize=11, pad=8)
    plt.tight_layout()
    st.pyplot(fig)
    charts_for_excel['Sampling Tiers'] = make_chart_buf(fig)

    st.dataframe(sg2.rename(columns={'tier':'Tier','events':'Events','units':'Units',
        'sampled':'Sampled','avg':'Avg Units/Demo','conv':'Conv %'}),
        use_container_width=True, hide_index=True)

    section("Coupon Impact")
    wc = r[r['_coupons']>0]['_units'].mean() if (r['_coupons']>0).any() else 0
    nc = r[r['_coupons']==0]['_units'].mean() if (r['_coupons']==0).any() else 0
    lift = round((wc/nc-1)*100,1) if nc > 0 else 0
    coup_e = (r['_coupons']>0).sum()

    c1,c2,c3 = st.columns(3)
    kpi_card(c1, f"{coup_e}", "EVENTS WITH COUPONS")
    kpi_card(c2, f"{wc:.2f}", "AVG UNITS WITH COUPONS")
    kpi_card(c3, f"+{lift}%", "COUPON LIFT")

    fig, ax = plt.subplots(1,1,figsize=(7,4))
    fig.patch.set_facecolor(BLK); setup_ax(ax)
    bars = ax.bar(['With Coupons','Without Coupons'], [wc,nc],
                  color=[GRN,RED], edgecolor='none', width=0.5)
    ax.set_title('Coupon Impact on Avg Units/Demo', fontsize=11, pad=8)
    ax.set_ylabel('Avg Units / Demo')
    bar_labels(ax, bars, '{:.2f}', fontsize=12)
    ax.text(0.5, 0.92, f'Lift: +{lift}%', transform=ax.transAxes,
            ha='center', color=GLD, fontsize=12, fontweight='bold')
    plt.tight_layout()
    st.pyplot(fig)
    charts_for_excel['Coupon Analysis'] = make_chart_buf(fig)

# ── TAB 6: SKUs ───────────────────────────────────────────────
with tabs[5]:
    s1=k['s1']; s2=k['s2']; s3=k['s3']; tot=s1+s2+s3
    if tot > 0:
        section("SKU / Flavor Breakdown")
        c1,c2,c3 = st.columns(3)
        kpi_card(c1, f"{s1:.0f}", f"PINEAPPLE GUAVA ({s1/tot*100:.1f}%)")
        kpi_card(c2, f"{s2:.0f}", f"LEMON LIME ({s2/tot*100:.1f}%)")
        kpi_card(c3, f"{s3:.0f}", f"ORANGE MANGO ({s3/tot*100:.1f}%)")

        fig, axes = plt.subplots(1,2,figsize=(12,4))
        fig.patch.set_facecolor(BLK)
        ax = axes[0]
        ax.set_facecolor(BLK); ax.figure.patch.set_facecolor(BLK)
        sku_colors = ['#A8C4E0','#1F2260','#8B7362']
        wedges, texts, autotexts = ax.pie(
            [s1,s2,s3], labels=['Pineapple Guava','Lemon Lime','Orange Mango'],
            colors=sku_colors, autopct='%1.1f%%', startangle=90,
            textprops={'color':'white','fontsize':10})
        ax.set_title('Unit Share by SKU', fontsize=11, pad=8, color=GLD)

        ax = axes[1]; setup_ax(ax)
        bars = ax.bar(['Pineapple Guava','Lemon Lime','Orange Mango'],
                      [s1,s2,s3], color=sku_colors, edgecolor='none')
        ax.set_title('Units Sold by SKU', fontsize=11, pad=8)
        ax.set_ylabel('Total Units')
        bar_labels(ax, bars, '{:.0f}', fontsize=11)
        plt.tight_layout()
        st.pyplot(fig)
        charts_for_excel['SKU Breakdown'] = make_chart_buf(fig)
    else:
        st.info("SKU columns not mapped. Go to Edit Column Mapping to add them.")

# ── TAB 7: CAMPAIGN COMPARE ───────────────────────────────────
with tabs[6]:
    if len(sel_campaigns) > 1:
        section("Campaign Comparison")
        rows = []
        for camp in sel_campaigns:
            rc = r[r['_campaign']==camp]
            kc = compute_kpis(rc)
            rows.append({'Campaign':camp,'Events':kc['n'],'Units':kc['units'],
                'Samples':kc['samp'],'Revenue':f"${kc['rev']:,.0f}",
                'Conv %':kc['conv'],'Avg Units/Demo':kc['avg_u'],
                'Avg Sampled/Demo':kc['avg_s'],'Units/BA Hour':kc['uph']})
        comp_df = pd.DataFrame(rows)
        st.dataframe(comp_df, use_container_width=True, hide_index=True)

        fig, axes = plt.subplots(1,3,figsize=(14,4))
        fig.patch.set_facecolor(BLK)
        for ax,(col,title,color) in zip(axes,[
            ('Avg Units/Demo','Avg Units per Demo',ORG),
            ('Conv %','Conversion Rate %',GLD),
            ('Units/BA Hour','Units per BA Hour',TEAL)]):
            setup_ax(ax)
            vals = comp_df[col].tolist()
            bars = ax.bar(comp_df['Campaign'], vals, color=color, edgecolor='none')
            ax.set_title(title, fontsize=11, pad=8)
            bar_labels(ax, bars, '{:.2f}', fontsize=10)
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=20, ha='right', fontsize=9)
        plt.tight_layout()
        st.pyplot(fig)
    else:
        st.info("Upload or select more than one campaign to see the comparison.")

# ── TAB 8: DOWNLOAD ───────────────────────────────────────────
with tabs[7]:
    st.markdown("### 📥 Download Full Excel Report")
    st.markdown("The report includes all data tables AND charts embedded directly in each sheet.")

    with st.spinner("Building Excel report with charts..."):
        excel_file = build_excel(r, m, charts_for_excel)

    st.download_button(
        label="⬇️  Download Excel Report",
        data=excel_file,
        file_name="Campaign_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.markdown("**What is included:**")
    cols_info = [
        "✅ KPI Summary — all headline numbers with formulas",
        "✅ All Events — clean raw data with conversion %",
        "✅ By State — performance per state with chart",
        "✅ By BA — all ambassadors ranked by avg units with chart",
        "✅ By Day — day of week analysis with chart",
        "✅ Top 20 Stores — highest performing locations",
        "✅ Low 20 Stores — lowest performing locations",
        "✅ Sampling Tiers — impact of sample volume with chart",
        "✅ Coupon Analysis — lift from coupons with chart",
        "✅ SKU Breakdown — flavor split with chart",
    ]
    for c in cols_info:
        st.markdown(c)
