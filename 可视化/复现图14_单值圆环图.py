"""复现图14：85%目标完成率单值渐变圆环图。"""
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.font_manager import FontProperties
from openpyxl import load_workbook
EXCEL_FILE="第二章 图表(前15).xlsx"; SHEET="14 单值圆环图"; OUTPUT="图14_单值圆环图.png"; BG="#1A1E43"; WHITE="#FFFFFF"
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
    rate=read_rate(); fp=font(); fig=plt.figure(figsize=(11.56,8.08),dpi=150,facecolor=BG); ax=fig.add_axes([.27,.17,.46,.58],facecolor=BG); ax.set_aspect("equal"); ax.axis("off")
    wedges,_=ax.pie([rate,1-rate],startangle=300,counterclock=False,colors=["none",(1,1,1,.12)],wedgeprops=dict(width=.45,edgecolor="none"))
    cmap=LinearSegmentedColormap.from_list("purple_red",["#7030A0","#E74E69"]); grad=np.tile(np.linspace(0,1,700)[:,None],(1,700)); im=ax.imshow(grad,extent=[-1,1,-1,1],origin="upper",cmap=cmap,zorder=0); im.set_clip_path(wedges[0])
    ax.text(0,.08,f"{rate:.0%}",ha="center",va="center",color=WHITE,fontsize=34,fontproperties=fp); ax.text(0,-.33,"目标完成率",ha="center",va="center",color=WHITE,fontsize=11,fontproperties=fp)
    fig.text(.06,.92,"2022年上半年目标完成率",color=WHITE,fontsize=20,fontweight="bold",fontproperties=fp,va="top"); fig.text(.06,.82,"截至6月30日销售目标总体完成率达到85%",color=WHITE,fontsize=14,fontproperties=fp,va="top"); fig.text(.045,.045,"*注：数据来源于公司销售系统",color="#D9D9D9",fontsize=8,fontproperties=fp)
    out=base_dir()/OUTPUT; fig.savefig(out,dpi=300,facecolor=BG); print(f"图片已保存：{out.resolve()}"); plt.show()
if __name__=="__main__": main()
