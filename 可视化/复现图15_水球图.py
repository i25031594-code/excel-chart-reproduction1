"""复现图15：65%目标完成率水球图。"""
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib.patches import Circle, PathPatch
from matplotlib.path import Path as MplPath
from openpyxl import load_workbook
EXCEL_FILE="第二章 图表(前15).xlsx"; SHEET="15 水球图"; OUTPUT="图15_水球图.png"; BG="#1A1E43"; WHITE="#FFFFFF"; BLUE="#0070C0"; CYAN="#00B0F0"
def base_dir(): return Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
def excel_path():
    for p in (base_dir()/EXCEL_FILE,Path.cwd()/EXCEL_FILE,Path.cwd()/"upload"/EXCEL_FILE):
        if p.exists(): return p
    raise FileNotFoundError(f"未找到{EXCEL_FILE}")
def font():
    for p in ["C:/Windows/Fonts/msyh.ttc","C:/Windows/Fonts/simhei.ttf","/System/Library/Fonts/PingFang.ttc","/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]:
        if Path(p).exists(): return FontProperties(fname=p)
    return FontProperties(family=["Microsoft YaHei","SimHei","sans-serif"])
def read_rate():
    wb=load_workbook(excel_path(),data_only=True,read_only=True); v=float(wb[SHEET]["B3"].value); wb.close(); return v
def main():
    rate=read_rate(); fp=font(); fig=plt.figure(figsize=(11.56,8.08),dpi=150,facecolor=BG); ax=fig.add_axes([.25,.14,.50,.64],facecolor=BG); ax.set_xlim(-1.15,1.15); ax.set_ylim(-1.15,1.15); ax.set_aspect("equal"); ax.axis("off")
    circle=Circle((0,0),1,facecolor="none",edgecolor=BLUE,lw=3); ax.add_patch(circle)
    level=-1+2*rate; x=np.linspace(-1.2,1.2,700); wave=level+.08*np.sin(2*np.pi*(x+1)/1.35); verts=np.column_stack([np.r_[x,x[::-1]],np.r_[wave,np.full_like(x,-1.25)]]); patch=PathPatch(MplPath(verts),facecolor=BLUE,edgecolor="none",alpha=1); patch.set_clip_path(circle); ax.add_patch(patch)
    wave2=level+.045*np.sin(2*np.pi*(x+1)/1.35+1.2); verts2=np.column_stack([np.r_[x,x[::-1]],np.r_[wave2,np.full_like(x,-1.25)]]); patch2=PathPatch(MplPath(verts2),facecolor=CYAN,edgecolor="none",alpha=.30); patch2.set_clip_path(circle); ax.add_patch(patch2)
    ax.text(0,0.02,f"{rate:.0%}",ha="center",va="center",color=WHITE,fontsize=36,fontproperties=fp,fontweight="bold"); ax.text(0,-.25,"目标完成率",ha="center",va="center",color=WHITE,fontsize=11,fontproperties=fp)
    fig.text(.06,.92,"2022年上半年目标完成率",color=WHITE,fontsize=20,fontweight="bold",fontproperties=fp,va="top"); fig.text(.06,.82,"截至6月30日销售目标总体完成率达到65%",color=WHITE,fontsize=14,fontproperties=fp,va="top"); fig.text(.045,.045,"*注：数据来源于公司销售系统",color="#D9D9D9",fontsize=8,fontproperties=fp)
    out=base_dir()/OUTPUT; fig.savefig(out,dpi=300,facecolor=BG); print(f"图片已保存：{out.resolve()}"); plt.show()
if __name__=="__main__": main()
