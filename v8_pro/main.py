# -*- coding: utf-8 -*-
"""车辆运营与报警月度报告 V8.4 自动识别最终版
重点修复：1.KeyError '_p'；2.Word 图表图例/数据标签/数据表重叠；3.图表统一独立区域排版。
"""
from pathlib import Path
import re, shutil
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn

BASE_DIR=Path(__file__).resolve().parent; DATA_DIR=BASE_DIR/'data'; OUTPUT_DIR=BASE_DIR/'output'; TEMPLATE=BASE_DIR/'report_template.docx'
# V8.4：自动识别输入文件与月份；保留 V8.3 已验证的图表排版代码
MONTH_FILES={}
VIDEO_FILES={}
REPORT_OLD_MONTH='上月'
REPORT_NEW_MONTH='本月'
REPORT_YEAR='2026'
WORD_REPORT=OUTPUT_DIR/'车辆运营与报警月度总结_V8.4.docx'; EXCEL_REPORT=OUTPUT_DIR/'车辆运营与报警月度总结_V8.4.xlsx'; CHART_DIR=OUTPUT_DIR/'_charts'

def _month_key(path):
    name=txt(path.name)
    # 支持：5月份、5月、2026年5月、2026-05、2026_05 等常见文件名
    patterns=[r'(20\d{2})\D*(1[0-2]|[1-9])\s*月',r'(20\d{2})[-_./](1[0-2]|0?[1-9])',r'(?<!\d)(1[0-2]|[1-9])\s*月']
    for pat in patterns:
        m=re.search(pat,name)
        if m:
            if len(m.groups())==2 and m.group(1).startswith('20'):
                return int(m.group(1))*100+int(m.group(2)), f'{int(m.group(2))}月', int(m.group(1))
            return int(m.group(1)), f'{int(m.group(1))}月', None
    return None, None, None

def _excel_sheets(path):
    try:return pd.ExcelFile(path).sheet_names
    except Exception:return []

def discover_inputs():
    """扫描 data 目录，自动识别两个月度文件和两份视频异常文件。"""
    global REPORT_OLD_MONTH, REPORT_NEW_MONTH, REPORT_YEAR
    if not DATA_DIR.exists(): raise FileNotFoundError(f'找不到数据目录：{DATA_DIR}')
    files=sorted([p for p in DATA_DIR.glob('*.xlsx') if not p.name.startswith('~$')], key=lambda p:p.stat().st_mtime)
    monthly=[]; video=[]
    for p in files:
        sheets=_excel_sheets(p)
        if '报警详情' in sheets and '行驶里程' in sheets:
            monthly.append(p); continue
        try:
            cols={txt(c) for c in pd.read_excel(p,nrows=2).columns}
        except Exception:
            continue
        if {'归属车组','设备编号','状态名称','状态类型','状态内容'}.issubset(cols): video.append(p)
    if len(monthly)<2:
        raise FileNotFoundError(f'自动识别月度数据失败：需要至少2个同时包含“报警详情”和“行驶里程”的Excel文件，实际找到{len(monthly)}个。')
    parsed=[(p,*_month_key(p)) for p in monthly]
    if all(x[1] is not None for x in parsed): parsed.sort(key=lambda x:x[1])
    else: parsed.sort(key=lambda x:x[0].stat().st_mtime)
    old_p,_,old_label,_=parsed[-2]; new_p,_,new_label,_=parsed[-1]
    REPORT_OLD_MONTH=old_label or '上月'; REPORT_NEW_MONTH=new_label or '本月'
    years=[x[3] for x in parsed if x[3] is not None]
    REPORT_YEAR=str(years[-1] if years else 2026)
    MONTH_FILES.clear(); MONTH_FILES[REPORT_OLD_MONTH]=old_p; MONTH_FILES[REPORT_NEW_MONTH]=new_p
    if len(video)<2:
        raise FileNotFoundError(f'自动识别视频异常数据失败：需要至少2个符合字段要求的Excel文件，实际找到{len(video)}个。')
    video=sorted(video,key=lambda p:p.stat().st_mtime)
    VIDEO_FILES.clear(); VIDEO_FILES['A']=video[-2]; VIDEO_FILES['B']=video[-1]
    print(f'自动识别：旧月份={REPORT_OLD_MONTH} -> {old_p.name}')
    print(f'自动识别：新月份={REPORT_NEW_MONTH} -> {new_p.name}')
    print(f'自动识别：视频A={VIDEO_FILES["A"].name}')
    print(f'自动识别：视频B={VIDEO_FILES["B"].name}')
    return old_p,new_p

ALARM_ALIASES={'车距过近':['车距过近','车距过小','前车距离过近','距离过近'],'疲劳驾驶':['疲劳驾驶','司机疲劳','驾驶员疲劳'],'驾驶员分心':['驾驶员分心','驾驶员分心驾驶','司机分心','分心驾驶'],'无驾驶员':['无驾驶员','驾驶员离岗'],'违规打电话':['违规打电话','驾驶员打电话','司机打电话'],'违规抽烟':['违规抽烟','驾驶员抽烟','司机抽烟'],'碰撞报警':['碰撞报警','前车碰撞','碰撞预警','碰撞'],'左右盲区报警':['左右盲区报警','盲区检测','左侧盲区','右侧盲区','盲区报警','盲区'],'驾驶员打哈欠':['驾驶员打哈欠','司机打哈欠','打哈欠']}
UNITS={'新华都':['新华都'],'兴万祥':['兴万祥'],'富达':['富达']}; VIDEO_PRIORITY={'DMS':1,'ADAS':2,'DSC':3,'普通摄像头':4,'未知':99}

def txt(v):
    if v is None:return ''
    try:
        if pd.isna(v):return ''
    except:pass
    return str(v).strip()
def norm(v):return re.sub(r'\s+','',txt(v)).lower()
def fi(v):
    try:return '—' if pd.isna(v) else f'{int(round(float(v))):,}'
    except:return txt(v)
def fn(v,d=2):
    try:return '—' if pd.isna(v) else f'{float(v):,.{d}f}'
    except:return txt(v)
def fp(v):
    try:return '—' if pd.isna(v) else f'{float(v)*100:.2f}%'
    except:return '—'
def trend(new,old):
    if pd.isna(new) or pd.isna(old):return '无法比较'
    d=float(new)-float(old); b=max(abs(float(old)),1e-12)
    return '基本持平' if abs(d)/b<=1e-6 else ('上升' if d>0 else '下降')
def alarm_name(v):
    x=norm(v)
    for n,als in ALARM_ALIASES.items():
        if any(norm(a) in x for a in als):return n
    return txt(v)
def unit_of(v):
    x=norm(v)
    for u,ks in UNITS.items():
        if any(norm(k) in x for k in ks):return u
    return '其他'

def read_month(path):
    if not path.exists():raise FileNotFoundError(f'找不到：{path}')
    x=pd.ExcelFile(path)
    for s in ['报警详情','行驶里程']:
        if s not in x.sheet_names:raise ValueError(f'{path.name}缺少工作表：{s}')
    a=pd.read_excel(path,sheet_name='报警详情'); m=pd.read_excel(path,sheet_name='行驶里程'); a.columns=[txt(c) for c in a.columns]; m.columns=[txt(c) for c in m.columns]; return a,m

def mileage_summary(df):
    c=next((c for c in ['行驶里程(km)','行驶里程（km）','行驶里程','运行里程','总里程'] if c in df.columns),None)
    if not c:raise ValueError(f'行驶里程表缺少里程字段：{list(df.columns)}')
    x=pd.to_numeric(df[c].astype(str).str.replace(',','',regex=False),errors='coerce').fillna(0); pc=next((c for c in ['车牌号码','车牌号','车牌','车辆编号'] if c in df.columns),None)
    return {'vehicles':int(df[pc].map(txt).ne('').sum()) if pc else len(x),'mileage':float(x.sum())}

def alarm_summary(df):
    c=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in df.columns),None)
    if not c:raise ValueError(f'报警详情缺少报警类型字段：{list(df.columns)}')
    x=df[c].map(alarm_name); counts=x.value_counts(); total=len(df)
    return pd.DataFrame([{'报警类型':n,'报警次数':int(counts.get(n,0)),'报警占比':int(counts.get(n,0))/total if total else 0} for n in ALARM_ALIASES])

def top_tables(df):
    ac=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in df.columns),None); gc=next((c for c in ['归属车组','归属分组','车组','班组'] if c in df.columns),None)
    empty=pd.DataFrame(columns=['排名','车组','报警次数']); specs={'车距过近高频违规车组TOP5':('车距过近',5),'疲劳驾驶高频违规车组TOP10':('疲劳驾驶',10),'驾驶员分心高频违规车组TOP10':('驾驶员分心',10),'前车碰撞高频违规车组TOP5':('碰撞报警',5),'左右盲区报警高频车组TOP5':('左右盲区报警',5),'打电话和抽烟违规车组TOP5':(['违规打电话','违规抽烟'],5)}
    if not ac or not gc:return {k:empty.copy() for k in specs}
    t=df.copy(); t['_type']=t[ac].map(alarm_name); t['_group']=t[gc].map(txt); t=t[t['_group'].ne('')]
    out={}
    for title,(types,n) in specs.items():
        types=types if isinstance(types,list) else [types]; q=t[t['_type'].isin(types)]
        if q.empty:out[title]=empty.copy();continue
        z=q.groupby('_group').size().reset_index(name='报警次数').sort_values(['报警次数','_group'],ascending=[False,True]).head(n).rename(columns={'_group':'车组'}); z.insert(0,'排名',range(1,len(z)+1)); out[title]=z
    return out

def month_data(path):
    a,m=read_month(path); ms=mileage_summary(m); s=alarm_summary(a); total=len(a); risk=total/ms['mileage']*100 if ms['mileage'] else np.nan
    ac=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in a.columns),None); gc=next((c for c in ['归属车组','归属分组','车组','班组'] if c in a.columns),None)
    if ac and gc:
        z=a.copy(); z['_type']=z[ac].map(alarm_name); z['_unit']=z[gc].map(unit_of); z=z[z['_type'].isin(ALARM_ALIASES)]
        units=z.groupby('_unit').size().reset_index(name='报警次数').rename(columns={'_unit':'单位'}) if not z.empty else pd.DataFrame(columns=['单位','报警次数'])
        if not units.empty:units=units.sort_values(['报警次数','单位'],ascending=[False,True]); units['占总报警比例']=units['报警次数']/total if total else 0
    else:units=pd.DataFrame(columns=['单位','报警次数','占总报警比例'])
    return {'path':path,'alarm_raw':a,'mileage_raw':m,'vehicles':ms['vehicles'],'mileage':ms['mileage'],'alarm_total':total,'risk100':risk,'alarm_summary':s,'unit_df':units,'top_tables':top_tables(a)}

def channel(v):
    s=txt(v)
    for p in [r'通道号\s*[:：#]?\s*(\d+)',r'通道\s*[:：#]?\s*(\d+)',r'channel\s*[:：#_\-]?\s*(\d+)',r'ch(?:annel)?\s*[:：#_\-]?\s*(\d+)']:
        m=re.search(p,s,re.I)
        if m:return float(m.group(1))
    nums=re.findall(r'\b\d+\b',s); return float(nums[0]) if len(nums)==1 else np.nan

def video_system(row):
    s=norm(f"{row.get('状态类型','')} {row.get('状态名称','')} {row.get('状态内容','')}")
    if any(k in s for k in ['dms','驾驶员监控','驾驶员监测','驾驶员状态监测']):return 'DMS'
    if any(k in s for k in ['adas','高级驾驶辅助','前向辅助驾驶']):return 'ADAS'
    if any(k in s for k in ['dsc','驾驶员状态','驾驶状态监测']):return 'DSC'
    if any(k in s for k in ['摄像头','camera','视频']):return '普通摄像头'
    return '未知'

def load_video(path):
    df=pd.read_excel(path); df.columns=[txt(c) for c in df.columns]; req=['归属车组','设备编号','状态名称','状态类型','状态内容']; miss=[c for c in req if c not in df.columns]
    if miss:raise ValueError(f'{path.name}缺少字段：{miss}')
    mask=df['状态名称'].astype(str).str.contains('摄像头|视频丢失|视频异常|视频丢帧|视频中断',case=False,na=False)|df['状态内容'].astype(str).str.contains('摄像头异常|视频丢失|视频异常|视频丢帧|视频中断',case=False,na=False)
    d=df.loc[mask].copy()
    if d.empty:return d,pd.DataFrame(columns=['归属车组','通道号','系统类型','异常次数'])
    d['通道号']=pd.to_numeric(d['通道号'],errors='coerce') if '通道号' in d.columns else np.nan; d['通道号']=d['通道号'].where(d['通道号'].notna(),d['状态内容'].map(channel)); d['系统类型']=d.apply(video_system,axis=1)
    s=d.groupby(['归属车组','通道号','系统类型'],dropna=False).size().reset_index(name='异常次数')
    s['_p']=s['系统类型'].map(VIDEO_PRIORITY).fillna(99).astype(int)  # 修复 KeyError '_p'
    s=s.sort_values(['_p','异常次数','归属车组'],ascending=[True,False,True],na_position='last').drop(columns='_p').reset_index(drop=True)
    return d,s

def setup_font():
    try:
        from matplotlib import font_manager
        names={f.name for f in font_manager.fontManager.ttflist}
        for n in ['Microsoft YaHei','SimHei','Noto Sans CJK SC']:
            if n in names:plt.rcParams['font.family']=n;break
    except:pass
    plt.rcParams['axes.unicode_minus']=False

def make_chart(title,headers,rows,draw,path,size=(10,6)):
    fig=plt.figure(figsize=size,dpi=180); gs=fig.add_gridspec(2,1,height_ratios=[0.70,0.22],hspace=0.38); ax=fig.add_subplot(gs[0]); tabax=fig.add_subplot(gs[1]); draw(ax)
    ax.set_title(title,fontsize=12,fontweight='bold',pad=10); ax.grid(axis='y',alpha=.25); ax.tick_params(axis='both',labelsize=8)
    handles,labels=ax.get_legend_handles_labels()
    if handles:ax.legend(handles,labels,loc='upper center',bbox_to_anchor=(.5,1.02),ncol=min(4,len(labels)),fontsize=7.5,frameon=False)
    tabax.axis('off'); tab=tabax.table(cellText=rows,colLabels=headers,loc='center',cellLoc='center',colLoc='center',bbox=[0,.02,1,.96]); tab.auto_set_font_size(False); tab.set_fontsize(7)
    for (r,c),cell in tab.get_celld().items():
        cell.set_edgecolor('#D9D9D9');cell.set_linewidth(.5)
        if r==0:cell.set_facecolor('#1F4E78');cell.get_text().set_color('white');cell.get_text().set_weight('bold')
        else:cell.set_facecolor('white')
    fig.subplots_adjust(left=.08,right=.97,top=.88,bottom=.06);fig.savefig(path,dpi=180,facecolor='white');plt.close(fig);return path

def labels(ax,bars,fmt='{:.0f}',fs=7):
    for b in bars:
        h=b.get_height()
        if np.isfinite(h) and h>0:ax.annotate(fmt.format(h),xy=(b.get_x()+b.get_width()/2,h),xytext=(0,4),textcoords='offset points',ha='center',va='bottom',fontsize=fs,clip_on=False)

def charts(old,new,trend_df,unit):
    OUTPUT_DIR.mkdir(parents=True,exist_ok=True);shutil.rmtree(CHART_DIR,ignore_errors=True);CHART_DIR.mkdir(parents=True,exist_ok=True);setup_font()
    p1=CHART_DIR/'01_车辆运营与报警对比.png'; mileage=[old['mileage']/1000,new['mileage']/1000]; alarms=[old['alarm_total'],new['alarm_total']]; risk=[old['risk100'],new['risk100']]
    def d1(ax):
        x=np.arange(2);w=.23;b1=ax.bar(x-w,mileage,w,label='总里程（千公里）');b2=ax.bar(x,alarms,w,label='总报警');b3=ax.bar(x+w,risk,w,label='百公里报警');ax.set_xticks(x);ax.set_xticklabels([REPORT_OLD_MONTH,REPORT_NEW_MONTH]);labels(ax,b1,'{:.3f}');labels(ax,b2,'{:.0f}');labels(ax,b3,'{:.2f}')
    make_chart(f'{REPORT_OLD_MONTH}—{REPORT_NEW_MONTH}车辆运营与报警对比',['指标',REPORT_OLD_MONTH,REPORT_NEW_MONTH],[['总里程（千公里）',f'{mileage[0]:.3f}',f'{mileage[1]:.3f}'],['总报警',f'{alarms[0]:.0f}',f'{alarms[1]:.0f}'],['百公里报警',f'{risk[0]:.2f}',f'{risk[1]:.2f}']],d1,p1,(9.6,6.3))
    names=['车距过近','疲劳驾驶','驾驶员分心','无驾驶员','违规打电话','违规抽烟','碰撞报警','驾驶员打哈欠'];core=trend_df.set_index('报警类型').reindex(names).reset_index().fillna(0);p2=CHART_DIR/'02_报警趋势分析_主要类型.png'
    def d2(ax):
        x=np.arange(len(core));w=.34;b1=ax.bar(x-w/2,core[REPORT_OLD_MONTH+'次数'],w,label=REPORT_OLD_MONTH);b2=ax.bar(x+w/2,core[REPORT_NEW_MONTH+'次数'],w,label=REPORT_NEW_MONTH);ax.set_xticks(x);ax.set_xticklabels(names,fontsize=7);labels(ax,b1,fs=6);labels(ax,b2,fs=6)
    make_chart('5—6月份主要报警类型对比分析',['月份']+names,[['5月']+[f"{int(v):,}" for v in core['5月次数']],['6月']+[f"{int(v):,}" for v in core['6月次数']]],d2,p2,(12.2,6.8))
    p2b=CHART_DIR/'02B_左右盲区报警趋势.png';blind=trend_df[trend_df['报警类型'].eq('左右盲区报警')];vals=[0,0] if blind.empty else [float(blind.iloc[0]['5月次数']),float(blind.iloc[0]['6月次数'])]
    def d2b(ax):b=ax.bar([REPORT_OLD_MONTH,REPORT_NEW_MONTH],vals,.42,label='左右盲区报警');labels(ax,b,fs=8)
    make_chart('左右盲区报警趋势对比',['指标','5月','6月'],[['左右盲区报警',f'{vals[0]:.0f}',f'{vals[1]:.0f}']],d2b,p2b,(9.6,5.8))
    p3=CHART_DIR/'03_各单位报警趋势.png'
    if unit.empty:p3=None
    else:
        def d3(ax):
            x=np.arange(len(unit));w=.34;b1=ax.bar(x-w/2,unit[REPORT_OLD_MONTH+'报警次数'],w,label=REPORT_OLD_MONTH);b2=ax.bar(x+w/2,unit[REPORT_NEW_MONTH+'报警次数'],w,label=REPORT_NEW_MONTH);ax.set_xticks(x);ax.set_xticklabels(unit['单位']);labels(ax,b1);labels(ax,b2)
        make_chart('各单位报警趋势对比分析',['月份']+list(unit['单位']),[[REPORT_OLD_MONTH]+[f'{int(v):,}' for v in unit[REPORT_OLD_MONTH+'报警次数']],[REPORT_NEW_MONTH]+[f'{int(v):,}' for v in unit[REPORT_NEW_MONTH+'报警次数']]],d3,p3,(9.8,6.2))
    return p1,p2,p2b,p3

def add_table(doc,df,pct=(),widths=None,fs=9):
    if df is None or df.empty:doc.add_paragraph('暂无可用数据。');return
    t=doc.add_table(rows=1,cols=len(df.columns));t.style='Table Grid';t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    for i,c in enumerate(df.columns):
        cell=t.rows[0].cells[i];cell.text=str(c);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if widths:cell.width=Cm(widths[i])
        for p in cell.paragraphs:
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(0)
            for r in p.runs:r.bold=True;r.font.name='Microsoft YaHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');r.font.size=Pt(fs)
    for _,row in df.iterrows():
        cells=t.add_row().cells
        for i,c in enumerate(df.columns):
            v=row[c];s=fp(v) if c in pct else fi(v) if isinstance(v,(int,np.integer)) else fn(v) if isinstance(v,(float,np.floating)) else txt(v);cells[i].text=s;cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths:cells[i].width=Cm(widths[i])
            for p in cells[i].paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(0)
                for r in p.runs:r.font.name='Microsoft YaHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');r.font.size=Pt(fs)

def heading(doc,text,level=1):
    p=doc.add_paragraph(style=f'Heading {level}');p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(4);p.paragraph_format.keep_with_next=True;r=p.add_run(text);r.font.name='微软雅黑';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');r.font.size=Pt(15 if level==1 else 12);r.bold=True

def picture(doc,path):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(7);p.paragraph_format.keep_together=True;p.add_run().add_picture(str(path),width=Cm(14.6))

def word_report(old,new,overview,at,unit,video,cp,risk_df,auto):
    doc=Document(str(TEMPLATE)) if TEMPLATE.exists() else Document();sec=doc.sections[0];sec.top_margin=sec.bottom_margin=Cm(2.54);sec.left_margin=sec.right_margin=Cm(3.175);body=doc._element.body
    for child in list(body):
        if not child.tag.endswith('sectPr'):body.remove(child)
    normal=doc.styles['Normal'];normal.font.name='Microsoft YaHei';normal._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');normal.font.size=Pt(11)
    for s,z in [('Heading 1',15),('Heading 2',12)]:st=doc.styles[s];st.font.name='Microsoft YaHei';st._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');st.font.size=Pt(z);st.font.bold=True
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('多宝山铜业智能终端运营情况月度总结');r.font.name='Microsoft YaHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');r.font.size=Pt(18)
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(10);r=p.add_run(f'{REPORT_YEAR}年{REPORT_OLD_MONTH}—{REPORT_NEW_MONTH}运营、报警及视频异常综合分析');r.font.name='Microsoft YaHei';r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑');r.font.size=Pt(14)
    doc.add_paragraph('目录');
    for s in ['一、安全运营分析','    1．车辆运营里程','    2．报警趋势分析','三、典型违规案例晾晒','四、重点视频异常','五、重点结论与整改建议']:doc.add_paragraph(s+' '+'.'*18)
    doc.add_page_break();heading(doc,'安全运营分析');heading(doc,'车辆运营里程',2)
    md=new['mileage']-old['mileage'];ad=new['alarm_total']-old['alarm_total'];rd=new['risk100']-old['risk100'];mw='增加' if md>0 else '减少' if md<0 else '基本持平';aw='增加' if ad>0 else '减少' if ad<0 else '基本持平';rw='上升' if rd>0 else '下降' if rd<0 else '基本持平'
    p=doc.add_paragraph(f"{REPORT_NEW_MONTH}总运行车辆数 {fi(new['vehicles'])} 辆，总运行里程 {fn(new['mileage'])} KM；相较{REPORT_OLD_MONTH}里程{mw} {fn(abs(md))} KM；总报警次数{aw} {fi(abs(ad))} 次；百公里风险事件数由 {fn(old['risk100'])} 变化为 {fn(new['risk100'])}。") ;p.paragraph_format.first_line_indent=Cm(.74);picture(doc,cp[0]);add_table(doc,overview,widths=[2.75,2.55,2.55,3.05,3.05],fs=10)
    heading(doc,'报警趋势分析',2);parts=[]
    for _,r in at.iterrows():parts.append(f"{r['报警类型']}：{fp(r[REPORT_NEW_MONTH+'占比'])}（{fi(r[REPORT_NEW_MONTH+'次数'])}次）较{REPORT_OLD_MONTH}{r['趋势']}")
    p=doc.add_paragraph('；'.join(parts)+'。详细趋势分析如图所示：');p.paragraph_format.first_line_indent=Cm(.74);picture(doc,cp[1]);picture(doc,cp[2]);add_table(doc,at,pct=[REPORT_OLD_MONTH+'占比',REPORT_NEW_MONTH+'占比'],widths=[2.45,1.55,1.55,1.55,1.55,1.55,2.25,1.95],fs=9);doc.add_paragraph('各单位报警排名：');add_table(doc,unit,pct=[REPORT_NEW_MONTH+'占比'],widths=[2.35,2.45,2.25,2.35,2.35,2.9],fs=9)
    if cp[3]:picture(doc,cp[3])
    heading(doc,'典型违规案例晾晒');doc.add_paragraph('（备注：因司机未签到所以按照车组进行数据展示）')
    for title,df in new['top_tables'].items():heading(doc,title,2);add_table(doc,df,widths=[1.8,9.35,3.5],fs=9.5)
    heading(doc,'重点风险车组 TOP10');doc.add_paragraph('根据当前月份已识别的主要报警类型，对车组进行综合风险排序；该排名用于重点关注，不改变原始报警统计口径。');add_table(doc,risk_df,widths=[0.8,2.7,1.7,1.6,1.2,1.2,1.2,1.2,1.4,1.4,2.0],fs=8.5)
    heading(doc,'重点视频异常');doc.add_paragraph('业务优先级：DMS > ADAS > DSC > 普通摄像头。');add_table(doc,video.head(20),widths=[4,2.5,4,4.15],fs=9)
    heading(doc,'重点结论与整改建议')
    for s in list(auto)+[f'百公里风险事件数{rw}。','视频异常优先核查DMS、ADAS、DSC等核心系统。','后续月度统一使用报警详情、行驶里程、视频异常三类数据生成报告。']:doc.add_paragraph(s,style='List Bullet')
    doc.add_paragraph('编制人：车辆运营与报警数据分析系统');doc.add_paragraph('日期：2026年6月');OUTPUT_DIR.mkdir(parents=True,exist_ok=True);doc.save(WORD_REPORT);return WORD_REPORT

def compare(old,new):
    ov=pd.DataFrame([['总运行车辆数',old['vehicles'],new['vehicles'],new['vehicles']-old['vehicles'],trend(new['vehicles'],old['vehicles'])],['总运行里程（KM）',old['mileage'],new['mileage'],new['mileage']-old['mileage'],trend(new['mileage'],old['mileage'])],['总报警次数',old['alarm_total'],new['alarm_total'],new['alarm_total']-old['alarm_total'],trend(new['alarm_total'],old['alarm_total'])],['百公里风险事件数',old['risk100'],new['risk100'],new['risk100']-old['risk100'],trend(new['risk100'],old['risk100'])]],columns=['指标',REPORT_OLD_MONTH,REPORT_NEW_MONTH,'变化','趋势'])
    o=old['alarm_summary'].set_index('报警类型');n=new['alarm_summary'].set_index('报警类型');rows=[]
    for name in ALARM_ALIASES:
        oc=int(o.loc[name,'报警次数']);nc=int(n.loc[name,'报警次数']);op=float(o.loc[name,'报警占比']);np_=float(n.loc[name,'报警占比']);rows.append([name,oc,nc,nc-oc,op,np_,(np_-op)*100,'上升' if np_>op else '下降' if np_<op else '基本持平'])
    at=pd.DataFrame(rows,columns=['报警类型',REPORT_OLD_MONTH+'次数',REPORT_NEW_MONTH+'次数','次数变化',REPORT_OLD_MONTH+'占比',REPORT_NEW_MONTH+'占比','占比变化（百分点）','趋势']);ou=old['unit_df'].set_index('单位')['报警次数'] if not old['unit_df'].empty else pd.Series(dtype=float);u=new['unit_df'].copy()
    if u.empty:uc=pd.DataFrame(columns=['单位','5月报警次数','6月报警次数','变化','6月占比','趋势'])
    else:
        u[REPORT_OLD_MONTH+'报警次数']=u['单位'].map(ou).fillna(0).astype(int);u[REPORT_NEW_MONTH+'报警次数']=u['报警次数'].astype(int);u['变化']=u[REPORT_NEW_MONTH+'报警次数']-u[REPORT_OLD_MONTH+'报警次数'];u[REPORT_NEW_MONTH+'占比']=u[REPORT_NEW_MONTH+'报警次数']/new['alarm_total'] if new['alarm_total'] else 0;u['趋势']=[trend(a,b) for a,b in zip(u[REPORT_NEW_MONTH+'报警次数'],u[REPORT_OLD_MONTH+'报警次数'])];uc=u[['单位',REPORT_OLD_MONTH+'报警次数',REPORT_NEW_MONTH+'报警次数','变化',REPORT_NEW_MONTH+'占比','趋势']]
    return ov,at,uc

def excel_report(old,new,ov,at,unit,video,risk_df):
    OUTPUT_DIR.mkdir(parents=True,exist_ok=True)
    if EXCEL_REPORT.exists():
        try:EXCEL_REPORT.unlink()
        except PermissionError:raise PermissionError('请先关闭WPS/Excel中已经打开的旧Excel报告。')
    with pd.ExcelWriter(EXCEL_REPORT,engine='openpyxl') as w:
        ov.to_excel(w,sheet_name='一_车辆运营里程',index=False,startrow=3);at.to_excel(w,sheet_name='二_报警趋势分析',index=False,startrow=3);unit.to_excel(w,sheet_name='二_报警趋势分析',index=False,startrow=7+len(at));video.to_excel(w,sheet_name='四_视频异常',index=False,startrow=3);old['alarm_summary'].to_excel(w,sheet_name='五_'+REPORT_OLD_MONTH+'报警统计',index=False);new['alarm_summary'].to_excel(w,sheet_name='六_'+REPORT_NEW_MONTH+'报警统计',index=False);old['mileage_raw'].to_excel(w,sheet_name='七_'+REPORT_OLD_MONTH+'里程明细',index=False);new['mileage_raw'].to_excel(w,sheet_name='八_'+REPORT_NEW_MONTH+'里程明细',index=False);old['alarm_raw'].to_excel(w,sheet_name='九_'+REPORT_OLD_MONTH+'报警明细',index=False);new['alarm_raw'].to_excel(w,sheet_name='十_'+REPORT_NEW_MONTH+'报警明细',index=False)
        risk_df.to_excel(w,sheet_name='十一_重点风险车组',index=False,startrow=3)
        ws=w.book.create_sheet('三_典型违规案例');row=1
        for title,df in new['top_tables'].items():ws.cell(row,1,title);row+=1;df.to_excel(w,sheet_name='三_典型违规案例',index=False,startrow=row-1);row+=len(df)+3
        c=w.book.create_sheet('_图表数据');c.append(['指标',REPORT_OLD_MONTH,REPORT_NEW_MONTH]);c.append(['总运行里程（千KM）',old['mileage']/1000,new['mileage']/1000]);c.append(['总报警次数',old['alarm_total'],new['alarm_total']]);c.append(['百公里风险事件数',old['risk100'],new['risk100']]);c['E1']='报警类型';c['F1']=REPORT_OLD_MONTH;c['G1']=REPORT_NEW_MONTH;o=old['alarm_summary'].set_index('报警类型');n=new['alarm_summary'].set_index('报警类型')
        for i,name in enumerate(ALARM_ALIASES,start=2):c.cell(i,5,name);c.cell(i,6,int(o.loc[name,'报警次数']));c.cell(i,7,int(n.loc[name,'报警次数']))
        c['I1']='单位';c['J1']=REPORT_OLD_MONTH;c['K1']=REPORT_NEW_MONTH
        for i,r in enumerate(unit.itertuples(index=False,name=None),start=2):c.cell(i,9,r[0]);c.cell(i,10,r[1]);c.cell(i,11,r[2])
        c.sheet_state='hidden'
    wb=load_workbook(EXCEL_REPORT);c=wb['_图表数据']
    def chart(ws,title,cat,vals,start,end,anchor,width=15,height=8):
        if end<start:return
        ch=BarChart();ch.type='col';ch.style=10;ch.title=title;ch.y_axis.title='数量';ch.height=height;ch.width=width;ch.legend.position='b';ch.gapWidth=60;ch.add_data(Reference(c,min_col=min(vals),max_col=max(vals),min_row=start-1,max_row=end),titles_from_data=True);ch.set_categories(Reference(c,min_col=cat,min_row=start,max_row=end));ws.add_chart(ch,anchor)
    chart(wb['一_车辆运营里程'],'车辆运营与报警对比',1,[2,3],2,4,'G4');chart(wb['二_报警趋势分析'],'主要报警类型趋势对比',5,[6,7],2,9,'J2',18,9);chart(wb['二_报警趋势分析'],'左右盲区报警趋势对比',5,[6,7],10,10,'J21')
    if not unit.empty:chart(wb['二_报警趋势分析'],'各单位报警趋势对比',9,[10,11],2,1+len(unit),'J38')
    fill=PatternFill('solid',fgColor='1F4E78');font=Font(name='Microsoft YaHei',bold=True,color='FFFFFF');thin=Side(style='thin',color='D9E1F2')
    for ws in wb.worksheets:
        if ws.title=='_图表数据':continue
        ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
        for cell in ws[1]:cell.fill=fill;cell.font=font;cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);cell.border=Border(bottom=thin)
        for row in ws.iter_rows(min_row=2):
            for cell in row:cell.alignment=Alignment(vertical='center',wrap_text=True)
        for i,col in enumerate(ws.columns,1):ws.column_dimensions[get_column_letter(i)].width=min(38,max(10,max([len(str(x.value)) for x in col if x.value is not None]+[10])+2))
    wb['一_车辆运营里程']['A1']=f'第一部分 车辆运营里程（{REPORT_OLD_MONTH} vs {REPORT_NEW_MONTH}）';wb['二_报警趋势分析']['A1']=f'第二部分 报警趋势分析（{REPORT_OLD_MONTH} vs {REPORT_NEW_MONTH}）';wb['四_视频异常']['A1']='第四部分 重点视频异常';wb.save(EXCEL_REPORT);wb.close();return EXCEL_REPORT

def risk_analysis(alarm_raw):
    """V8.6：基于当前月份报警详情生成重点风险车组TOP10。"""
    ac=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in alarm_raw.columns),None)
    gc=next((c for c in ['归属车组','归属分组','车组','班组'] if c in alarm_raw.columns),None)
    cols=['排名','车组','综合风险分','风险等级','车距过近','疲劳驾驶','驾驶员分心','碰撞报警','左右盲区报警','电话/抽烟','主要问题']
    if not ac or not gc or alarm_raw.empty:return pd.DataFrame(columns=cols)
    d=alarm_raw.copy();d['_type']=d[ac].map(alarm_name);d['_group']=d[gc].map(txt)
    d=d[d['_group'].ne('')&d['_type'].isin(ALARM_ALIASES)]
    weights={'车距过近':1.0,'疲劳驾驶':1.2,'驾驶员分心':1.2,'碰撞报警':1.5,'左右盲区报警':1.1,'违规打电话':1.0,'违规抽烟':0.8}
    d['_score']=d['_type'].map(weights)
    p=d.groupby(['_group','_type'])['_score'].sum().unstack(fill_value=0)
    for t in weights:
        if t not in p.columns:p[t]=0
    p['电话/抽烟']=p['违规打电话']+p['违规抽烟']
    p['综合风险分']=p['车距过近']+p['疲劳驾驶']+p['驾驶员分心']+p['碰撞报警']+p['左右盲区报警']+p['电话/抽烟']
    r=p.reset_index().rename(columns={'_group':'车组'})
    r['风险等级']=r['综合风险分'].map(lambda x:'高风险' if x>=20 else '中风险' if x>=10 else '一般')
    def problems(row):
        q=[('碰撞报警',row['碰撞报警']),('疲劳驾驶',row['疲劳驾驶']),('驾驶员分心',row['驾驶员分心']),('车距过近',row['车距过近']),('左右盲区报警',row['左右盲区报警']),('电话/抽烟',row['电话/抽烟'])]
        q=sorted([(a,b) for a,b in q if b>0],key=lambda x:x[1],reverse=True)
        return '、'.join(a for a,_ in q[:2]) if q else '—'
    r['主要问题']=r.apply(problems,axis=1)
    r=r.sort_values(['综合风险分','车组'],ascending=[False,True]).head(10).reset_index(drop=True)
    r.insert(0,'排名',range(1,len(r)+1));r['综合风险分']=r['综合风险分'].round(1)
    return r[cols]

def auto_conclusions(old,new,risk_df):
    out=[]
    diff=new['alarm_total']-old['alarm_total']
    out.append(f"{REPORT_NEW_MONTH}总报警较{REPORT_OLD_MONTH}{'增加'+fi(diff)+'次' if diff>0 else '减少'+fi(abs(diff))+'次' if diff<0 else '基本持平'}。")
    if not new['alarm_summary'].empty:
        z=new['alarm_summary'].sort_values('报警次数',ascending=False).iloc[0]
        out.append(f"{REPORT_NEW_MONTH}报警量最高的类型为“{z['报警类型']}”，共{fi(z['报警次数'])}次，占当月报警{fp(z['报警占比'])}。")
    if not risk_df.empty:
        high=risk_df[risk_df['风险等级'].eq('高风险')]
        focus=high if not high.empty else risk_df
        names='、'.join(focus['车组'].astype(str).head(3))
        out.append(f"重点风险车组{'（高风险）' if not high.empty else ''}：{names}。")
    return out

def main():
    print('='*70);print('车辆运营与报警月度报告 V8.4 - 自动识别最终版');print('='*70)
    discover_inputs()
    old=month_data(MONTH_FILES[REPORT_OLD_MONTH]);new=month_data(MONTH_FILES[REPORT_NEW_MONTH])
    va,vsa=load_video(VIDEO_FILES['A']);vb,vsb=load_video(VIDEO_FILES['B'])
    ov,at,unit=compare(old,new);risk_df=risk_analysis(new['alarm_raw']);auto=auto_conclusions(old,new,risk_df);cp=charts(old,new,at,unit);x=excel_report(old,new,ov,at,unit,vsb,risk_df);w=word_report(old,new,ov,at,unit,vsb,cp,risk_df,auto)
    print(f"{REPORT_OLD_MONTH}：车辆 {fi(old['vehicles'])}，里程 {fn(old['mileage'])} KM，报警 {fi(old['alarm_total'])}")
    print(f"{REPORT_NEW_MONTH}：车辆 {fi(new['vehicles'])}，里程 {fn(new['mileage'])} KM，报警 {fi(new['alarm_total'])}")
    print(f'A视频异常：{len(va)} 条；B视频异常：{len(vb)} 条');print('全部生成成功：');print('Word：',w);print('Excel：',x);print('图表：',CHART_DIR)

if __name__=='__main__':
    try:main()
    except Exception as e:print('\n程序运行失败：',type(e).__name__,e);raise
