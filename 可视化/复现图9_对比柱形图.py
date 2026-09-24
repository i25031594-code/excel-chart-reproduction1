"""复现图9：2021与2022商品销量对比柱形图。"""
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from openpyxl import load_workbook
EXCEL_FILE="第二章 图表(前15).xlsx"; SHEET="9 对比柱形图"; OUTPUT="图9_对比柱形图.png"; BG="#1A1E43"; WHITE="#FFFFFF"; BLUE="#0070C0"; LIGHT="#82ADD7"
def base_dir(): return Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
def excel_path():
    for p in (base_dir()/EXCEL_FILE,Path.cwd()/EXCEL_FILE,Path.cwd()/"upload"/EXCEL_FILE):
        if p.exists(): return p
    raise FileNotFoundError(f"未找到{EXCEL_FILE}")
def font():
    for p in ["C:/Windows/Fonts/msyh.ttc","C:/Windows/Fonts/simhei.ttf","/System/Library/Fonts/PingFang.ttc","/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]:
        if Path(p).exists(): return FontProperties(fname=p)
    return FontProperties(family=["Microsoft YaHei","SimHei","sans-serif"])
def read_data():
    wb=load_workbook(excel_path(),data_only=True,read_only=True); ws=wb[SHEET]; n=[ws.cell(r,2).value for r in range(3,8)]; a=np.array([ws.cell(r,3).value for r in range(3,8)],float); b=np.array([ws.cell(r,4).value for r in range(3,8)],float); wb.close(); return n,a,b
def main():
    n,a,b=read_data(); d=a-b; fp=font(); x=np.arange(5); fig=plt.figure(figsize=(11.56,8.44),dpi=150,facecolor=BG); ax=fig.add_axes([.06,.14,.86,.56],facecolor=BG)
    p1=x-.15; p2=x+.15; ba=ax.bar(p1,a,.23,color=BLUE); bb=ax.bar(p2,b,.23,color=LIGHT)
    ax.set_ylim(0,5600); ax.set_xticks(x,n); ax.set_yticks([]); ax.tick_params(axis="x",length=0,colors=WHITE,pad=8)
    for t in ax.get_xticklabels(): t.set_fontproperties(fp); t.set_fontsize(9)
    for s in ax.spines.values(): s.set_visible(False)
    ax.spines["bottom"].set_visible(True); ax.spines["bottom"].set_color("#A7ABC0"); ax.spines["bottom"].set_alpha(.55)
    for bars,vals in [(ba,a),(bb,b)]:
        for bar,val in zip(bars,vals): ax.text(bar.get_x()+bar.get_width()/2,val+100,f"{int(val)}",ha="center",color=WHITE,fontsize=9,fontproperties=fp)
    for xi,hi,lo,dif in zip(p2,a,b,d):
        ax.vlines(xi,lo+130,hi-100,color=WHITE,lw=1); ax.plot([xi],[hi],marker="s",ms=5,color=WHITE); ax.text(xi,(hi+lo)/2,f"{int(dif)}",ha="center",va="center",color=WHITE,fontsize=9,fontproperties=fp)
    ax.scatter([],[],s=45,marker="s",color=BLUE,label="2021销量"); ax.scatter([],[],s=45,marker="s",color=LIGHT,label="2022销量")
    lg=ax.legend(loc="upper left",bbox_to_anchor=(0,.99),ncol=2,frameon=False,handlelength=.8,handletextpad=.4,columnspacing=1.5,prop=fp)
    for t in lg.get_texts(): t.set_color(WHITE); t.set_fontsize(9)
    fig.text(.055,.92,"2022年商品对比去年销售情况",color=WHITE,fontsize=20,fontweight="bold",fontproperties=fp,va="top")
    fig.text(.055,.82,"商品整体比去年销量有所下降，其中隔离下降最多，下降33%",color=WHITE,fontsize=14,fontproperties=fp,va="top")
    fig.text(.045,.045,"*注：数据来源于公司销售系统，统计日期截至2022.01.01",color="#D9D9D9",fontsize=8,fontproperties=fp)
    out=base_dir()/OUTPUT; fig.savefig(out,dpi=300,facecolor=BG); print(f"图片已保存：{out.resolve()}"); plt.show()
if __name__=="__main__": main()
