"""Generate all charts (PNG) for the TIME PG/PD resource-justification document."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import os

OUT = os.path.dirname(os.path.abspath(__file__))
TEAL="#0f7d88"; TEAL_D="#0b5c65"; GOLD="#e0a83c"; GREEN="#2e9e6b"; RED="#d9645f"; GREY="#c9cdd0"; INK="#1f2a30"
plt.rcParams.update({"font.family":"DejaVu Sans","font.size":11,"axes.edgecolor":"#c9d2d4",
                     "axes.grid":True,"grid.color":"#e6eef0","axes.axisbelow":True})

def save(fig,name):
    fig.savefig(os.path.join(OUT,name),dpi=180,bbox_inches="tight",facecolor="white")
    plt.close(fig); print("wrote",name)

# 1) DONUTS: Impression share vs Revenue share (FY2025 clean)
def donuts():
    fig,axes=plt.subplots(1,2,figsize=(8.2,3.7))
    cols=[GOLD,GREEN,RED]; labels=["Direct","Programmatic","House"]
    imp=[35.7,55.5,8.8]; rev=[84.5,15.2,0.3]
    for ax,data,title in [(axes[0],imp,"Share of IMPRESSIONS"),(axes[1],rev,"Share of REVENUE")]:
        w,_,at=ax.pie(data,colors=cols,startangle=90,counterclock=False,
                      autopct=lambda p:f"{p:.0f}%" if p>3 else "",pctdistance=0.78,
                      wedgeprops=dict(width=0.42,edgecolor="white",linewidth=2),
                      textprops=dict(color="white",fontweight="bold",fontsize=12))
        ax.set_title(title,fontsize=12,fontweight="bold",color=INK,pad=8)
    fig.legend(labels,loc="lower center",ncol=3,frameon=False,bbox_to_anchor=(0.5,-0.04),fontsize=11)
    fig.suptitle("Programmatic + House = 64% of impressions, but only ~16% of revenue",
                 fontsize=12.5,fontweight="bold",color=TEAL_D,y=1.03)
    save(fig,"c1_donuts.png")

# 2) SSP revenue horizontal bar (top programmatic partners)
def ssp():
    data=[("Media.net",593477),("Ad Exchange",408188),("Amazon / A9",319257),
          ("EBDA (Open Bidding)",130530),("Media.net Bytes",36196),("Magnite",20418)]
    data=data[::-1]
    names=[d[0] for d in data]; vals=[d[1] for d in data]
    fig,ax=plt.subplots(figsize=(8.2,3.6))
    bars=ax.barh(names,vals,color=TEAL,height=0.62)
    bars[-1].set_color(TEAL_D); bars[-2].set_color(TEAL_D)
    for b,v in zip(bars,vals):
        ax.text(v+8000,b.get_y()+b.get_height()/2,f"${v/1000:,.0f}K",va="center",fontsize=10,color=INK,fontweight="bold")
    ax.set_xlim(0,660000); ax.xaxis.set_major_formatter(FuncFormatter(lambda x,_:f"${x/1000:,.0f}K"))
    ax.set_title("Top programmatic partners by revenue (US, Jan 2025 – Jun 2026)",
                 fontsize=12,fontweight="bold",color=TEAL_D,pad=8)
    ax.grid(axis="y",visible=False)
    save(fig,"c2_ssp.png")

# 3) Grouped bar: top-6 ad units Direct vs Programmatic eCPM (FY2025)
def adunits():
    units=["stickyfooter1","leaderboard1","inline1","inline2","rightrail1","inline3"]
    direct=[22.92,18.66,18.91,20.76,19.56,20.05]
    prog=[1.79,1.78,3.25,2.82,2.88,2.74]
    x=range(len(units)); w=0.38
    fig,ax=plt.subplots(figsize=(8.6,3.9))
    b1=ax.bar([i-w/2 for i in x],direct,w,label="Direct (Standard)",color=GOLD)
    b2=ax.bar([i+w/2 for i in x],prog,w,label="Programmatic",color=GREEN)
    for b in b1: ax.text(b.get_x()+b.get_width()/2,b.get_height()+0.4,f"${b.get_height():.0f}",ha="center",fontsize=9.5,fontweight="bold",color=INK)
    for b in b2: ax.text(b.get_x()+b.get_width()/2,b.get_height()+0.4,f"${b.get_height():.1f}",ha="center",fontsize=9.5,fontweight="bold",color=INK)
    ax.set_xticks(list(x)); ax.set_xticklabels(units,fontsize=10)
    ax.set_ylabel("eCPM ($)"); ax.set_ylim(0,27)
    ax.set_title("Same ad unit, two prices: Direct vs Programmatic eCPM (FY2025)",
                 fontsize=12,fontweight="bold",color=TEAL_D,pad=8)
    ax.legend(frameon=False,fontsize=10,loc="upper right"); ax.grid(axis="x",visible=False)
    save(fig,"c3_adunits.png")

# 4) Opportunity scenario grouped bar ($8 vs $10 at 10/20/30% shift)
def opp():
    scen=["10% shift\n(27.8M imp)","20% shift\n(55.6M imp)","30% shift\n(83.3M imp)"]
    at8=[159,319,478]; at10=[215,430,645]
    x=range(len(scen)); w=0.38
    fig,ax=plt.subplots(figsize=(8.2,3.9))
    b1=ax.bar([i-w/2 for i in x],at8,w,label="Target eCPM $8",color=TEAL)
    b2=ax.bar([i+w/2 for i in x],at10,w,label="Target eCPM $10",color=GOLD)
    for bars in (b1,b2):
        for b in bars: ax.text(b.get_x()+b.get_width()/2,b.get_height()+8,f"+${b.get_height():,.0f}K",ha="center",fontsize=10,fontweight="bold",color=INK)
    ax.set_xticks(list(x)); ax.set_xticklabels(scen,fontsize=10)
    ax.set_ylabel("Incremental revenue / year"); ax.set_ylim(0,760)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y,_:f"${y:,.0f}K"))
    ax.set_title("Incremental annual revenue — shifting top-6 programmatic inventory to PG/PD",
                 fontsize=11.5,fontweight="bold",color=TEAL_D,pad=8)
    ax.legend(frameon=False,fontsize=10,loc="upper left"); ax.grid(axis="x",visible=False)
    save(fig,"c4_opportunity.png")

if __name__=="__main__":
    donuts(); ssp(); adunits(); opp()
    print("all charts done")
