# -*- coding: utf-8 -*-
"""车辆运营与报警月度报告 V8.2：模板最终美化 + Word + Excel 双输出"""
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

BASE_DIR=Path(__file__).resolve().parent
DATA_DIR=BASE_DIR/'data'; OUTPUT_DIR=BASE_DIR/'output'; TEMPLATE=BASE_DIR/'report_template.docx'
MONTH_FILES={'5月':DATA_DIR/'多宝山铜业5月份综合数据.xlsx','6月':DATA_DIR/'多宝山铜业6月份综合数据.xlsx'}
VIDEO_FILES={'A':DATA_DIR/'data_A.xlsx','B':DATA_DIR/'data_B.xlsx'}
WORD_REPORT=OUTPUT_DIR/'车辆运营与报警月度总结_5月-6月_V8.docx'
EXCEL_REPORT=OUTPUT_DIR/'车辆运营与报警月度总结_5月-6月_V8.xlsx'
CHART_DIR=OUTPUT_DIR/'_charts'

ALARM_ALIASES={
'车距过近':['车距过近','车距过小','前车距离过近','距离过近'],
'疲劳驾驶':['疲劳驾驶','司机疲劳','驾驶员疲劳'],
'驾驶员分心':['驾驶员分心','驾驶员分心驾驶','司机分心','分心驾驶'],
'无驾驶员':['无驾驶员','驾驶员离岗'],
'违规打电话':['违规打电话','驾驶员打电话','司机打电话'],
'违规抽烟':['违规抽烟','驾驶员抽烟','司机抽烟'],
'碰撞报警':['碰撞报警','前车碰撞','碰撞预警','碰撞'],
'左右盲区报警':['左右盲区报警','盲区检测','左侧盲区','右侧盲区','盲区报警','盲区'],
'驾驶员打哈欠':['驾驶员打哈欠','司机打哈欠','打哈欠']}
UNITS={'新华都':['新华都'],'兴万祥':['兴万祥'],'富达':['富达']}


def txt(v):
    if v is None:return ''
    try:
        if pd.isna(v):return ''
    except Exception:pass
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
def pctchg(new,old):
    return np.nan if pd.isna(new) or pd.isna(old) or float(old)==0 else (float(new)-float(old))/abs(float(old))
def trend(new,old):
    if pd.isna(new) or pd.isna(old):return '无法比较'
    diff=float(new)-float(old); base=max(abs(float(old)),1e-12)
    return '基本持平' if abs(diff)/base<=1e-6 else ('上升' if diff>0 else '下降')

def alarm_name(v):
    x=norm(v)
    for name,aliases in ALARM_ALIASES.items():
        if any(norm(a) in x for a in aliases):return name
    return txt(v)

def read_month(path):
    if not path.exists():raise FileNotFoundError(f'找不到：{path}')
    xls=pd.ExcelFile(path)
    if '报警详情' not in xls.sheet_names:raise ValueError(f'{path.name}缺少工作表：报警详情')
    if '行驶里程' not in xls.sheet_names:raise ValueError(f'{path.name}缺少工作表：行驶里程')
    a=pd.read_excel(path,sheet_name='报警详情'); m=pd.read_excel(path,sheet_name='行驶里程')
    a.columns=[txt(c) for c in a.columns]; m.columns=[txt(c) for c in m.columns]
    return a,m

def mileage_summary(df):
    mc=next((c for c in ['行驶里程(km)','行驶里程（km）','行驶里程','运行里程','总里程'] if c in df.columns),None)
    if not mc:raise ValueError(f'行驶里程表缺少里程字段：{list(df.columns)}')
    x=pd.to_numeric(df[mc].astype(str).str.replace(',','',regex=False),errors='coerce').fillna(0)
    pc=next((c for c in ['车牌号码','车牌号','车牌','车辆编号'] if c in df.columns),None)
    vehicles=int(df[pc].map(txt).ne('').sum()) if pc else len(x)
    return {'vehicles':vehicles,'mileage':float(x.sum())}

def alarm_summary(df):
    ac=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in df.columns),None)
    if not ac:raise ValueError(f'报警详情缺少报警类型字段：{list(df.columns)}')
    x=df[ac].map(alarm_name); counts=x.value_counts(); total=len(df)
    return pd.DataFrame([{'报警类型':n,'报警次数':int(counts.get(n,0)),'报警占比':int(counts.get(n,0))/total if total else 0} for n in ALARM_ALIASES])

def unit_of(v):
    g=norm(v)
    for u,ks in UNITS.items():
        if any(norm(k) in g for k in ks):return u
    return '其他'

def top_tables(df):
    ac=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in df.columns),None)
    gc=next((c for c in ['归属车组','归属分组','车组','班组'] if c in df.columns),None)
    empty=pd.DataFrame(columns=['排名','车组','报警次数'])
    if not ac or not gc:return {'车距过近高频违规车组TOP5':empty.copy(),'疲劳驾驶高频违规车组TOP10':empty.copy(),'驾驶员分心高频违规车组TOP10':empty.copy(),'前车碰撞高频违规车组TOP5':empty.copy(),'左右盲区报警高频车组TOP5':empty.copy(),'打电话和抽烟违规车组TOP5':empty.copy()}
    t=df.copy(); t['_type']=t[ac].map(alarm_name); t['_group']=t[gc].map(txt); t=t[t['_group'].ne('')]
    def make(types,n):
        q=t[t['_type'].isin(types if isinstance(types,list) else [types])]
        if q.empty:return empty.copy()
        z=q.groupby('_group').size().reset_index(name='报警次数').sort_values('报警次数',ascending=False).head(n).rename(columns={'_group':'车组'}); z.insert(0,'排名',range(1,len(z)+1)); return z
    return {'车距过近高频违规车组TOP5':make('车距过近',5),'疲劳驾驶高频违规车组TOP10':make('疲劳驾驶',10),'驾驶员分心高频违规车组TOP10':make('驾驶员分心',10),'前车碰撞高频违规车组TOP5':make('碰撞报警',5),'左右盲区报警高频车组TOP5':make('左右盲区报警',5),'打电话和抽烟违规车组TOP5':make(['违规打电话','违规抽烟'],5)}

def month_data(path):
    a,m=read_month(path); ms=mileage_summary(m); s=alarm_summary(a); total=len(a); risk=total/ms['mileage']*100 if ms['mileage'] else np.nan
    ac=next((c for c in ['报警类型','报警名称','事件类型','事件名称'] if c in a.columns),None); gc=next((c for c in ['归属车组','归属分组','车组','班组'] if c in a.columns),None)
    if ac and gc:
        z=a.copy(); z['_type']=z[ac].map(alarm_name); z['_unit']=z[gc].map(unit_of); z=z[z['_type'].isin(ALARM_ALIASES)]
        units=z.groupby('_unit').size().reset_index(name='报警次数').rename(columns={'_unit':'单位'}).sort_values('报警次数',ascending=False) if not z.empty else pd.DataFrame(columns=['单位','报警次数'])
        if not units.empty:units['占总报警比例']=units['报警次数']/total if total else 0
    else:units=pd.DataFrame(columns=['单位','报警次数','占总报警比例'])
    return {'path':path,'alarm_raw':a,'mileage_raw':m,'vehicles':ms['vehicles'],'mileage':ms['mileage'],'alarm_total':total,'risk100':risk,'alarm_summary':s,'unit_df':units,'top_tables':top_tables(a)}

def channel(v):
    s=txt(v)
    for p in [r'通道号\s*[:：#]?\s*(\d+)',r'通道\s*[:：#]?\s*(\d+)',r'channel\s*[:：#_\-]?\s*(\d+)',r'ch(?:annel)?\s*[:：#_\-]?\s*(\d+)']:
        m=re.search(p,s,re.I)
        if m:return float(m.group(1))
    nums=re.findall(r'\b\d+\b',s)
    return float(nums[0]) if len(nums)==1 else np.nan

def video_system(row):
    s=norm(f"{row.get('状态类型','')} {row.get('状态名称','')} {row.get('状态内容','')}")
    if any(k in s for k in ['dms','驾驶员监控','驾驶员监测','驾驶员状态监测']):return 'DMS'
    if any(k in s for k in ['adas','高级驾驶辅助','前向辅助驾驶']):return 'ADAS'
    if any(k in s for k in ['dsc','驾驶员状态','驾驶状态监测']):return 'DSC'
    if any(k in s for k in ['摄像头','camera','视频']):return '普通摄像头'
    return '未知'

def load_video(path):
    df=pd.read_excel(path); df.columns=[txt(c) for c in df.columns]
    req=['归属车组','设备编号','状态名称','状态类型','状态内容']; miss=[c for c in req if c not in df.columns]
    if miss:raise ValueError(f'{path.name}缺少字段：{miss}')
    mask=df['状态名称'].astype(str).str.contains('摄像头|视频丢失|视频异常|视频丢帧|视频中断',case=False,na=False)|df['状态内容'].astype(str).str.contains('摄像头异常|视频丢失|视频异常|视频丢帧|视频中断',case=False,na=False)
    d=df.loc[mask].copy()
    if d.empty:return d,pd.DataFrame(columns=['归属车组','通道号','系统类型','异常次数'])
    d['通道号']=pd.to_numeric(d['通道号'],errors='coerce') if '通道号' in d.columns else np.nan
    ex=d['状态内容'].map(channel); d['通道号']=d['通道号'].where(d['通道号'].notna(),ex); d['系统类型']=d.apply(video_system,axis=1)
    s=d.groupby(['归属车组','通道号','系统类型'],dropna=False).size().reset_index(name='异常次数').sort_values('异常次数',ascending=False)
    return d,s

def charts(old,new,alarm_trend,unit):
    """生成与原始模板风格接近的柱状图：蓝色主色、数据标签、底部数据表。"""
    OUTPUT_DIR.mkdir(parents=True,exist_ok=True)
    shutil.rmtree(CHART_DIR,ignore_errors=True)
    CHART_DIR.mkdir(parents=True,exist_ok=True)

    try:
        from matplotlib import font_manager
        avail={f.name for f in font_manager.fontManager.ttflist}
        for f in ['Microsoft YaHei','SimHei','Noto Sans CJK SC']:
            if f in avail:
                plt.rcParams['font.family']=f
                break
    except Exception:
        pass
    plt.rcParams['axes.unicode_minus']=False

    def add_table(ax, headers, rows, bbox=(0,-0.02,1,0.23), fontsize=7.5):
        table=ax.table(cellText=rows,colLabels=headers,loc='bottom',bbox=bbox,cellLoc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(fontsize)
        for cell in table.get_celld().values():
            cell.set_edgecolor('#D9D9D9')
            cell.set_linewidth(0.5)
            cell.set_facecolor('white')
        return table

    def label_bars(ax,bars,fmt='{:.0f}',fontsize=7):
        for b in bars:
            h=b.get_height()
            if h <= 0: continue
            ax.annotate(fmt.format(h),(b.get_x()+b.get_width()/2,h),xytext=(0,3),
                        textcoords='offset points',ha='center',va='bottom',fontsize=fontsize)

    # 1. 车辆运营里程：完全贴近模板“柱状图 + 数据表”结构。
    p1=CHART_DIR/'01_车辆运营里程.png'
    fig,ax=plt.subplots(figsize=(8.8,4.65),dpi=180)
    months=['5月','6月']
    mileage=np.array([old['mileage']/1000,new['mileage']/1000],dtype=float)
    alarms=np.array([old['alarm_total'],new['alarm_total']],dtype=float)
    risk=np.array([old['risk100'],new['risk100']],dtype=float)
    # 使用与模板一致的“单坐标轴 + 三组指标”表达方式；数值过大时保留真实值。
    x=np.arange(len(months)); w=.24
    b1=ax.bar(x-w,mileage,w,label='总里程（千公里）')
    b2=ax.bar(x,alarms,w,label='总报警')
    b3=ax.bar(x+w,risk,w,label='百公里报警')
    ax.set_title('5-6月份里程和报警对比',fontsize=12,fontweight='bold',pad=8)
    ax.set_xticks(x); ax.set_xticklabels(months,fontsize=8)
    ax.tick_params(axis='y',labelsize=8)
    ax.grid(axis='y',alpha=.25)
    ax.legend(loc='lower left',fontsize=7,frameon=False)
    label_bars(ax,b1,fmt='{:.3f}',fontsize=7)
    label_bars(ax,b2,fmt='{:.0f}',fontsize=7)
    label_bars(ax,b3,fmt='{:.2f}',fontsize=7)
    rows=[
        ['总里程（千公里）',f'{mileage[0]:.3f}',f'{mileage[1]:.3f}'],
        ['总报警',f'{alarms[0]:.0f}',f'{alarms[1]:.0f}'],
        ['百公里报警',f'{risk[0]:.2f}',f'{risk[1]:.2f}'],
    ]
    add_table(ax,['指标','5月','6月'],rows,bbox=(0,-0.02,1,0.22),fontsize=7)
    fig.subplots_adjust(left=.10,right=.98,top=.87,bottom=.30)
    fig.savefig(p1,dpi=180,bbox_inches='tight',facecolor='white')
    plt.close(fig)

    # 2. 主要报警类型：模板式“多系列柱状图 + 底部数据表”。
    p2=CHART_DIR/'02_报警趋势分析_主要类型.png'
    names=['车距过近','疲劳驾驶','驾驶员分心','无驾驶员','违规打电话','违规抽烟','碰撞报警','驾驶员打哈欠']
    core=alarm_trend[alarm_trend['报警类型'].isin(names)].set_index('报警类型').reindex(names).reset_index()
    x=np.arange(len(core)); w=.34
    fig,ax=plt.subplots(figsize=(11.2,5.0),dpi=180)
    b1=ax.bar(x-w/2,core['5月次数'],w,label='5月')
    b2=ax.bar(x+w/2,core['6月次数'],w,label='6月')
    ax.set_title('5-6月份报警数据对比分析情况',fontsize=12,fontweight='bold',pad=8)
    ax.set_xticks(x); ax.set_xticklabels(core['报警类型'],fontsize=7,rotation=0)
    ax.tick_params(axis='y',labelsize=7)
    ax.grid(axis='y',alpha=.25)
    ax.legend(loc='upper left',fontsize=7,frameon=False)
    label_bars(ax,b1,fontsize=6); label_bars(ax,b2,fontsize=6)
    rows=[
        ['5月']+[f'{int(v):,}' for v in core['5月次数']],
        ['6月']+[f'{int(v):,}' for v in core['6月次数']],
    ]
    add_table(ax,['月份']+list(core['报警类型']),rows,bbox=(0,-0.06,1,0.22),fontsize=6.5)
    fig.subplots_adjust(left=.08,right=.99,top=.87,bottom=.30)
    fig.savefig(p2,dpi=180,bbox_inches='tight',facecolor='white')
    plt.close(fig)

    # 2B. 左右盲区单独展示，避免极大值把其他类型压扁。
    p2b=CHART_DIR/'02B_左右盲区报警趋势.png'
    blind=alarm_trend[alarm_trend['报警类型'].eq('左右盲区报警')]
    vals=[0,0] if blind.empty else [float(blind.iloc[0]['5月次数']),float(blind.iloc[0]['6月次数'])]
    fig,ax=plt.subplots(figsize=(8.8,4.3),dpi=180)
    bars=ax.bar(['5月','6月'],vals,width=.42)
    ax.set_title('左右盲区报警趋势对比',fontsize=12,fontweight='bold',pad=8)
    ax.tick_params(axis='both',labelsize=8); ax.grid(axis='y',alpha=.25)
    label_bars(ax,bars,fontsize=8)
    add_table(ax,['月份','5月','6月'],[['左右盲区报警',f'{vals[0]:.0f}',f'{vals[1]:.0f}']],bbox=(0,-.05,1,.20),fontsize=7)
    fig.subplots_adjust(left=.10,right=.98,top=.86,bottom=.28)
    fig.savefig(p2b,dpi=180,bbox_inches='tight',facecolor='white')
    plt.close(fig)

    # 3. 各单位趋势：柱状图 + 底部数据表。
    p3=CHART_DIR/'03_各单位报警趋势.png'
    if not unit.empty:
        x=np.arange(len(unit)); w=.34
        fig,ax=plt.subplots(figsize=(9.0,4.8),dpi=180)
        b1=ax.bar(x-w/2,unit['5月报警次数'],w,label='5月')
        b2=ax.bar(x+w/2,unit['6月报警次数'],w,label='6月')
        ax.set_title('各单位报警趋势对比分析',fontsize=12,fontweight='bold',pad=8)
        ax.set_xticks(x); ax.set_xticklabels(unit['单位'],fontsize=8)
        ax.tick_params(axis='y',labelsize=7); ax.grid(axis='y',alpha=.25)
        ax.legend(loc='upper left',fontsize=7,frameon=False)
        label_bars(ax,b1,fontsize=7); label_bars(ax,b2,fontsize=7)
        rows=[['5月']+[f'{int(v):,}' for v in unit['5月报警次数']],['6月']+[f'{int(v):,}' for v in unit['6月报警次数']]]
        add_table(ax,['月份']+list(unit['单位']),rows,bbox=(0,-.06,1,.22),fontsize=7)
        fig.subplots_adjust(left=.09,right=.98,top=.86,bottom=.30)
        fig.savefig(p3,dpi=180,bbox_inches='tight',facecolor='white')
        plt.close(fig)
    else:
        p3=None
    return p1,p2,p2b,p3

def add_doc_table(doc,df,pct_cols=(),widths=None,font_size=10):
    if df is None or df.empty:
        doc.add_paragraph('暂无可用数据。')
        return

    t=doc.add_table(rows=1,cols=len(df.columns))
    t.style='Table Grid'
    t.alignment=WD_TABLE_ALIGNMENT.CENTER
    t.autofit=False

    if widths and len(widths)==len(df.columns):
        for row in t.rows:
            for i,w in enumerate(widths):
                row.cells[i].width=Cm(w)

    for i,c in enumerate(df.columns):
        cell=t.rows[0].cells[i]
        cell.text=str(c)
        cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if widths and i < len(widths):
            cell.width=Cm(widths[i])
        for p in cell.paragraphs:
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after=Pt(0)
            for r in p.runs:
                r.bold=True
                r.font.name='Microsoft YaHei'
                r.font.size=Pt(font_size)

    for _,row in df.iterrows():
        cells=t.add_row().cells
        for i,c in enumerate(df.columns):
            v=row[c]
            if c in pct_cols:
                s=fp(v)
            elif isinstance(v,(float,np.floating)):
                s=fn(v)
            elif isinstance(v,(int,np.integer)):
                s=fi(v)
            else:
                s=txt(v)
            cells[i].text=s
            cells[i].vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if widths and i < len(widths):
                cells[i].width=Cm(widths[i])
            for p in cells[i].paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after=Pt(0)
                for r in p.runs:
                    r.font.name='Microsoft YaHei'
                    r.font.size=Pt(font_size)
    return t

def _set_template_paragraph(paragraph, text_value, size=11, bold=False,
                            align=None, space_before=0, space_after=3,
                            first_line_cm=None):
    """按原始模板的字体、字号、间距重建段落。"""
    paragraph.clear()
    if align is not None:
        paragraph.alignment = align
    pf = paragraph.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_line_cm is not None:
        pf.first_line_indent = Cm(first_line_cm)
    run = paragraph.add_run(str(text_value))
    run.font.name = 'Microsoft YaHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(size)
    run.bold = bold
    return paragraph


def _template_heading(doc, text_value, level=1):
    # 原模板 Heading 1/2 已经自带中文自动编号，因此这里绝不能再写“一、/1．”。
    p = doc.add_paragraph(style=f'Heading {level}')
    p.paragraph_format.space_before = Pt(2.5)
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(str(text_value))
    r.font.name = '微软雅黑'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    r.font.size = Pt(15 if level == 1 else 12)
    r.bold = True
    return p


def _add_template_picture(doc, path, width_cm=14.6, after=3):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.keep_together = True
    run = p.add_run()
    run.add_picture(str(path), width=Cm(width_cm))
    return p


def _restore_template_header_footer(docx_path):
    """python-docx 保存时会丢失模板中的浮动页眉/页脚引用；这里补回模板的 header1/footer1 引用。"""
    from zipfile import ZipFile, ZIP_DEFLATED
    from lxml import etree
    tmp=docx_path.with_suffix('.patched.docx')
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    with ZipFile(docx_path,'r') as zin, ZipFile(tmp,'w',ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            if item.filename=='word/document.xml':
                root=etree.fromstring(data)
                sect=root.xpath('.//w:sectPr[last()]',namespaces=ns)[0]
                refs=[('headerReference','rId5'),('footerReference','rId6')]
                for tag,rid in refs:
                    if not sect.xpath(f'./w:{tag}',namespaces=ns):
                        el=etree.Element('{%s}%s'%(ns['w'],tag))
                        el.set('{%s}id'%ns['r'],rid)
                        el.set('{%s}type'%ns['w'],'default')
                        sect.insert(0,el)
                data=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
            zout.writestr(item,data)
    tmp.replace(docx_path)

def word_report(old,new,overview,alarm_trend,unit,video_stats,chart_paths):
    """V8.2 模板最终美化版：严格按 report_template.docx 的 A4、边距、字号、图表宽度组织。"""
    doc=Document(str(TEMPLATE)) if TEMPLATE.exists() else Document()
    sec=doc.sections[0]
    # 与用户原始模板一致：A4 + 上下2.54cm + 左右3.175cm。
    sec.top_margin=Cm(2.54)
    sec.bottom_margin=Cm(2.54)
    sec.left_margin=Cm(3.175)
    sec.right_margin=Cm(3.175)

    # 保留模板样式体系，只清空正文内容。
    body=doc._element.body
    for child in list(body):
        if not child.tag.endswith('sectPr'):
            body.remove(child)

    normal=doc.styles['Normal']
    normal.font.name='Microsoft YaHei'
    normal._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    normal.font.size=Pt(11)
    normal.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY

    for style_name, size in [('Heading 1',15),('Heading 2',14),('Heading 3',13)]:
        st=doc.styles[style_name]
        st.font.name='Microsoft YaHei'
        st._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        st.font.size=Pt(size)
        st.font.bold=True
        st.paragraph_format.space_before=Pt(2.5)
        st.paragraph_format.space_after=Pt(2.5)

    # ---------------- 标题 ----------------
    p=doc.add_paragraph(style='Normal')
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.space_after=Pt(2)
    r=p.add_run('多宝山铜业智能终端运营情况月度总结')
    r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑')
    r.font.size=Pt(18); r.bold=False

    p=doc.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(10)
    r=p.add_run('2026年5月—6月运营、报警及视频异常综合分析')
    r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑')
    r.font.size=Pt(14); r.bold=False

    # ---------------- 模板式封面 + 目录 ----------------
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(12)
    p.paragraph_format.space_after=Pt(6)
    r=p.add_run('目录'); r.font.name='微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(14); r.bold=False
    toc_items=[
        '一、安全运营分析',
        '    1．车辆运营里程',
        '    2．报警趋势分析',
        '三、典型违规案例晾晒',
        '四、重点视频异常',
        '五、重点结论与整改建议',
    ]
    for item in toc_items:
        p=doc.add_paragraph()
        p.paragraph_format.space_after=Pt(3)
        r=p.add_run(item + ' ' + '.'*18)
        r.font.name='微软雅黑'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(10.5)
    doc.add_page_break()

    # ---------------- 第一部分：车辆运营里程 ----------------
    _template_heading(doc,'安全运营分析',1)
    _template_heading(doc,'车辆运营里程',2)

    md=new['mileage']-old['mileage']
    ad=new['alarm_total']-old['alarm_total']
    rd=new['risk100']-old['risk100']
    md_word='增加' if md>0 else '减少' if md<0 else '基本持平'
    ad_word='增加' if ad>0 else '减少' if ad<0 else '基本持平'
    rd_word='上升' if rd>0 else '下降' if rd<0 else '基本持平'

    p=doc.add_paragraph(style='Normal')
    p.paragraph_format.first_line_indent=Cm(0.74)
    p.paragraph_format.space_after=Pt(4)
    text_value=(
        f"6月总运行车辆数 {fi(new['vehicles'])} 辆，总运行里程 {fn(new['mileage'])} KM；"
        f"相较5月里程{md_word} {fn(abs(md))} KM；"
        f"总报警次数{ad_word} {fi(abs(ad))} 次；"
        f"百公里风险事件数由 {fn(old['risk100'])} 变化为 {fn(new['risk100'])}。"
        "详细趋势分析如图所示："
    )
    r=p.add_run(text_value); r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(11)

    _add_template_picture(doc, chart_paths[0], 14.6, after=5)

    # 运营指标表：5列压缩到模板正文宽度14.65cm。
    add_doc_table(doc,overview,
                  widths=[2.75,2.55,2.55,3.05,3.05],
                  font_size=10)

    # ---------------- 报警趋势 ----------------
    _template_heading(doc,'报警趋势分析',2)
    p=doc.add_paragraph('报警类型分布：',style='Normal')
    p.paragraph_format.space_after=Pt(2)

    # 保留模板“一段话 + 图表”的阅读方式，但控制长度，避免跨页乱跳。
    parts=[]
    for _,r in alarm_trend.iterrows():
        direction='上升' if r['趋势']=='上升' else '下降' if r['趋势']=='下降' else '持平'
        parts.append(
            f"{r['报警类型']}：{fp(r['6月占比'])}（{fi(r['6月次数'])}次）"
            f"较5月{direction}"
        )
    p=doc.add_paragraph(style='Normal')
    p.paragraph_format.first_line_indent=Cm(0.74)
    p.paragraph_format.space_after=Pt(4)
    r=p.add_run('；'.join(parts)+'。详细趋势分析如图所示：')
    r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(11)

    # 关键修复：所有图表统一14.6cm，严格落在模板正文区域，不再向右溢出。
    _add_template_picture(doc, chart_paths[1], 14.6, after=4)
    _add_template_picture(doc, chart_paths[2], 14.6, after=5)

    add_doc_table(doc,alarm_trend,
                  pct_cols=['5月占比','6月占比'],
                  widths=[2.45,1.55,1.55,1.55,1.55,1.55,2.25,1.95],
                  font_size=9)

    p=doc.add_paragraph('各单位报警排名：',style='Normal')
    p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(2)
    add_doc_table(doc,unit,
                  pct_cols=['6月占比'],
                  widths=[2.35,2.45,2.25,2.35,2.35,2.9],
                  font_size=9)

    if chart_paths[3]:
        _add_template_picture(doc, chart_paths[3], 14.6, after=6)

    # ---------------- 典型违规案例 ----------------
    _template_heading(doc,'典型违规案例晾晒',1)
    p=doc.add_paragraph('（备注：因司机未签到所以按照车组进行数据展示）',style='Normal')
    p.paragraph_format.space_after=Pt(4)
    for title,df in new['top_tables'].items():
        _template_heading(doc,title,2)
        add_doc_table(doc,df,widths=[1.8,9.35,3.5],font_size=9.5)

    # ---------------- 视频异常 ----------------
    _template_heading(doc,'重点视频异常',1)
    p=doc.add_paragraph('业务优先级：DMS > ADAS > DSC > 普通摄像头。',style='Normal')
    p.paragraph_format.space_after=Pt(4)
    add_doc_table(doc,
                  video_stats.head(20) if not video_stats.empty else video_stats,
                  widths=[4.0,2.5,4.0,4.15],font_size=9)

    # ---------------- 结论 ----------------
    _template_heading(doc,'重点结论与整改建议',1)
    conclusions=[
        f"6月总报警较5月{ad_word} {fi(abs(ad))} 次。",
        f"百公里风险事件数{rd_word}。",
        '优先关注高频报警车组及高频报警类型。',
        '视频异常优先核查DMS、ADAS、DSC等核心系统。',
        '后续月度统一使用报警详情、行驶里程、视频异常三类数据生成报告。'
    ]
    for s in conclusions:
        p=doc.add_paragraph(s,style='List Bullet')
        p.paragraph_format.space_after=Pt(2)
        p.paragraph_format.left_indent=Cm(0.5)
        p.paragraph_format.first_line_indent=Cm(0)
        for run in p.runs:
            run.font.name='Microsoft YaHei'; run._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); run.font.size=Pt(10.5)

    # 末尾留出与模板相近的编制信息。
    p=doc.add_paragraph()
    p.paragraph_format.space_before=Pt(6)
    r=p.add_run('编制人：车辆运营与报警数据分析系统')
    r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(10.5)
    p=doc.add_paragraph()
    r=p.add_run('日期：2026年6月')
    r.font.name='Microsoft YaHei'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'微软雅黑'); r.font.size=Pt(10.5)

    doc.save(WORD_REPORT)
    _restore_template_header_footer(WORD_REPORT)
    return WORD_REPORT

def style_ws(ws):
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions;fill=PatternFill('solid',fgColor='1F4E78');font=Font(name='Microsoft YaHei',bold=True,color='FFFFFF');thin=Side(style='thin',color='D9E1F2')
    for c in ws[1]:c.fill=fill;c.font=font;c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True);c.border=Border(bottom=thin)
    for row in ws.iter_rows(min_row=2):
        for c in row:c.alignment=Alignment(vertical='center',wrap_text=True)
    for i,col in enumerate(ws.columns,1):ws.column_dimensions[get_column_letter(i)].width=min(38,max(10,max([len(str(c.value)) for c in col if c.value is not None]+[10])+2))

def excel_report(old,new,overview,alarm_trend,unit,video_stats):
    if EXCEL_REPORT.exists():
        try:
            EXCEL_REPORT.unlink()
        except PermissionError:
            raise PermissionError('请先关闭WPS/Excel中已经打开的旧Excel报告。')

    with pd.ExcelWriter(EXCEL_REPORT,engine='openpyxl') as w:
        overview.to_excel(w,sheet_name='一_车辆运营里程',index=False,startrow=3)
        alarm_trend.to_excel(w,sheet_name='二_报警趋势分析',index=False,startrow=3)
        unit.to_excel(w,sheet_name='二_报警趋势分析',index=False,startrow=7+len(alarm_trend))
        video_stats.to_excel(w,sheet_name='四_视频异常',index=False,startrow=3)

        old['alarm_summary'].to_excel(w,sheet_name='五_5月报警统计',index=False)
        new['alarm_summary'].to_excel(w,sheet_name='六_6月报警统计',index=False)
        old['mileage_raw'].to_excel(w,sheet_name='七_5月里程明细',index=False)
        new['mileage_raw'].to_excel(w,sheet_name='八_6月里程明细',index=False)
        old['alarm_raw'].to_excel(w,sheet_name='九_5月报警明细',index=False)
        new['alarm_raw'].to_excel(w,sheet_name='十_6月报警明细',index=False)

        ws3=w.book.create_sheet('三_典型违规案例')
        ws3.cell(1,1,'第三部分 典型违规案例晾晒')
        row=2
        for title,df in new['top_tables'].items():
            ws3.cell(row,1,title)
            row+=1
            df.to_excel(w,sheet_name='三_典型违规案例',index=False,startrow=row-1)
            row+=len(df)+3

        c=w.book.create_sheet('_图表数据')
        c.append(['指标','5月','6月'])
        c.append(['总运行里程（千KM）',old['mileage']/1000,new['mileage']/1000])
        c.append(['总报警次数',old['alarm_total'],new['alarm_total']])
        c.append(['百公里风险事件数',old['risk100'],new['risk100']])

        c['E1']='报警类型'
        c['F1']='5月'
        c['G1']='6月'
        for i,n in enumerate(ALARM_ALIASES,start=2):
            c.cell(i,5,n)
            c.cell(i,6,int(old['alarm_summary'].set_index('报警类型').loc[n,'报警次数']))
            c.cell(i,7,int(new['alarm_summary'].set_index('报警类型').loc[n,'报警次数']))

        c['I1']='单位'
        c['J1']='5月'
        c['K1']='6月'
        for i,rowdata in enumerate(unit.itertuples(index=False,name=None),start=2):
            c.cell(i,9,rowdata[0])
            c.cell(i,10,rowdata[1])
            c.cell(i,11,rowdata[2])

        c.sheet_state='hidden'

    wb=load_workbook(EXCEL_REPORT)

    wb['一_车辆运营里程']['A1']='第一部分 车辆运营里程'
    wb['一_车辆运营里程']['A2']=f"旧数据：{old['path'].name}    新数据：{new['path'].name}"
    wb['二_报警趋势分析']['A1']='第二部分 报警趋势分析'
    wb['二_报警趋势分析']['A2']='报警占比按当月总报警次数计算。'
    wb['四_视频异常']['A1']='第四部分 重点视频异常'
    wb['四_视频异常']['A2']='业务优先级：DMS > ADAS > DSC > 普通摄像头'

    c=wb['_图表数据']

    def chart(ws,title,cat,vcols,start,end,anchor,width=15,height=8):
        if end < start:
            return
        ch=BarChart()
        ch.type='col'
        ch.style=10
        ch.title=title
        ch.y_axis.title='数量'
        ch.height=height
        ch.width=width
        ch.legend.position='b'
        ch.gapWidth=60
        ch.add_data(
            Reference(c,min_col=min(vcols),max_col=max(vcols),min_row=start-1,max_row=end),
            titles_from_data=True
        )
        ch.set_categories(Reference(c,min_col=cat,min_row=start,max_row=end))
        ws.add_chart(ch,anchor)

    chart(
        wb['一_车辆运营里程'],
        '车辆运营与报警对比',
        1,[2,3],2,4,'G4',
        width=15,height=8
    )

    # Excel中同样拆开报警图，避免左右盲区报警压扁其他柱子
    chart(
        wb['二_报警趋势分析'],
        '主要报警类型趋势对比',
        5,[6,7],2,9,'J2',
        width=18,height=9
    )
    # 只把左右盲区单独做成一个小图，数据位于第10行
    chart(
        wb['二_报警趋势分析'],
        '左右盲区报警趋势对比',
        5,[6,7],10,10,'J21',
        width=15,height=8
    )

    if not unit.empty:
        chart(
            wb['二_报警趋势分析'],
            '各单位报警趋势对比',
            9,[10,11],2,1+len(unit),'J38',
            width=15,height=8
        )

    for ws in wb.worksheets:
        if ws.title!='_图表数据':
            style_ws(ws)

    # 对核心工作表做额外列宽控制，防止数字被拆行
    ws=wb['二_报警趋势分析']
    widths={
        'A':16,'B':12,'C':12,'D':12,'E':12,'F':12,'G':16,'H':10
    }
    for col,width in widths.items():
        ws.column_dimensions[col].width=width

    ws=wb['一_车辆运营里程']
    for col,width in {'A':18,'B':14,'C':14,'D':14,'E':12}.items():
        ws.column_dimensions[col].width=width

    wb.save(EXCEL_REPORT)
    wb.close()
    return EXCEL_REPORT

def main():
    print('='*78);print('车辆运营与报警月度报告 V8.2 - 模板最终美化 + Word + Excel 双输出');print('='*78)
    for p in [*MONTH_FILES.values(),*VIDEO_FILES.values()]:
        if not p.exists():raise FileNotFoundError(f'找不到输入文件：{p}')
    old=month_data(MONTH_FILES['5月']);new=month_data(MONTH_FILES['6月'])
    va,vs_a=load_video(VIDEO_FILES['A']);vb,vs_b=load_video(VIDEO_FILES['B']);new['video_stats']=vs_b;old['video_stats']=vs_a;new['video_detail']=vb;old['video_detail']=va
    print(f"5月：车辆 {fi(old['vehicles'])}，里程 {fn(old['mileage'])} KM，报警 {fi(old['alarm_total'])}")
    print(f"6月：车辆 {fi(new['vehicles'])}，里程 {fn(new['mileage'])} KM，报警 {fi(new['alarm_total'])}")
    print(f"A视频异常：{len(va)} 条；B视频异常：{len(vb)} 条")
    overview,alarm_trend,uc=compare(old,new)
    charts_=charts(old,new,alarm_trend,uc)
    x=excel_report(old,new,overview,alarm_trend,uc,vs_b);w=word_report(old,new,overview,alarm_trend,uc,vs_b,charts_)
    print('\n全部生成成功：');print('Word ：',w);print('Excel：',x);print('图表：',CHART_DIR)

def compare(old,new):
    overview=pd.DataFrame([['总运行车辆数',old['vehicles'],new['vehicles'],new['vehicles']-old['vehicles'],trend(new['vehicles'],old['vehicles'])],['总运行里程（KM）',old['mileage'],new['mileage'],new['mileage']-old['mileage'],trend(new['mileage'],old['mileage'])],['总报警次数',old['alarm_total'],new['alarm_total'],new['alarm_total']-old['alarm_total'],trend(new['alarm_total'],old['alarm_total'])],['百公里风险事件数',old['risk100'],new['risk100'],new['risk100']-old['risk100'],trend(new['risk100'],old['risk100'])]],columns=['指标','5月','6月','变化','趋势'])
    rows=[];o=old['alarm_summary'].set_index('报警类型');n=new['alarm_summary'].set_index('报警类型')
    for name in ALARM_ALIASES:
        oc=int(o.loc[name,'报警次数']);nc=int(n.loc[name,'报警次数']);op=float(o.loc[name,'报警占比']);np_=float(n.loc[name,'报警占比']);rows.append([name,oc,nc,nc-oc,op,np_,(np_-op)*100,'上升' if np_>op else '下降' if np_<op else '基本持平'])
    at=pd.DataFrame(rows,columns=['报警类型','5月次数','6月次数','次数变化','5月占比','6月占比','占比变化（百分点）','趋势'])
    ou=old['unit_df'].set_index('单位')['报警次数'] if not old['unit_df'].empty else pd.Series(dtype=float); u=new['unit_df'].copy();
    if u.empty:uc=pd.DataFrame(columns=['单位','5月报警次数','6月报警次数','变化','6月占比','趋势'])
    else:
        u['5月报警次数']=u['单位'].map(ou).fillna(0).astype(int);u['6月报警次数']=u['报警次数'].astype(int);u['变化']=u['6月报警次数']-u['5月报警次数'];u['6月占比']=u['6月报警次数']/new['alarm_total'] if new['alarm_total'] else 0;u['趋势']=[trend(a,b) for a,b in zip(u['6月报警次数'],u['5月报警次数'])];uc=u[['单位','5月报警次数','6月报警次数','变化','6月占比','趋势']]
    return overview,at,uc

if __name__=='__main__':
    try:main()
    except Exception as e:
        print('\n程序运行失败：',type(e).__name__,e);raise
