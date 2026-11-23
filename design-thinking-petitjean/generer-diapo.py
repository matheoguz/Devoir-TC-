import sys, json
from pptx import Presentation
from pptx.util import Inches as I, Pt, Emu
from pptx.dml.color import RGBColor as C
from pptx.enum.shapes import MSO_SHAPE as SH
from pptx.enum.text import PP_ALIGN as AL, MSO_ANCHOR as AN
from pptx.enum.dml import MSO_LINE_DASH_STYLE as DASH

NUIT=C(0x1E,0x19,0x16); NUIT2=C(0x2B,0x24,0x1F); BRAISE=C(0xE2,0x58,0x2B); OR=C(0xF2,0xA3,0x3A)
KRAFT=C(0xD8,0xB8,0x8A); KRAFT2=C(0xE9,0xD4,0xB0); CREME=C(0xF7,0xF0,0xE3); ENCRE=C(0x2A,0x20,0x1A); GRIS=C(0x8A,0x7B,0x6C)
POST=[C(0xFF,0xE3,0x8A),C(0xFF,0xC9,0xA8),C(0xF6,0xE7,0xC8),C(0xFF,0xD6,0x7A)]
SERIF="Georgia"; SANS="Arial"; MONO="Courier New"

p=Presentation(); p.slide_width=I(13.333); p.slide_height=I(7.5); W,H=p.slide_width,p.slide_height
oral=json.load(open(sys.argv[1]))

def bg(s,c): f=s.background.fill; f.solid(); f.fore_color.rgb=c
def shp(s,kind,x,y,w,h,fill=None,line=None,rot=0,lw=1.25,dash=None,grad=None):
    r=s.shapes.add_shape(kind,x,y,w,h); r.shadow.inherit=False
    if grad:
        r.fill.gradient(); r.fill.gradient_angle=grad[2]
        st=r.fill.gradient_stops; st[0].color.rgb=grad[0]; st[0].position=0; st[1].color.rgb=grad[1]; st[1].position=1
    elif fill is None: r.fill.background()
    else: r.fill.solid(); r.fill.fore_color.rgb=fill
    if line is None: r.line.fill.background()
    else:
        r.line.color.rgb=line; r.line.width=Pt(lw)
        if dash: r.line.dash_style=dash
    r.rotation=rot
    if r.has_text_frame: r.text_frame.text=""
    return r
def tx(s,x,y,w,h,runs,size=18,color=ENCRE,font=SANS,bold=False,italic=False,al=AL.LEFT,an=AN.TOP,sp=1.1,rot=0,gap=0.4):
    tb=s.shapes.add_textbox(x,y,w,h); tb.rotation=rot; tf=tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=an
    for a in ('margin_left','margin_right','margin_top','margin_bottom'): setattr(tf,a,0)
    if isinstance(runs,str): runs=[runs]
    for i,para in enumerate(runs):
        pr=tf.paragraphs[0] if i==0 else tf.add_paragraph(); pr.alignment=al; pr.line_spacing=sp
        if i: pr.space_before=Pt(size*gap)
        if isinstance(para,str): para=[(para,{})]
        for t,o in para:
            r=pr.add_run(); r.text=t; f=r.font
            f.size=Pt(o.get('size',size)); f.name=o.get('font',font); f.bold=o.get('bold',bold); f.italic=o.get('italic',italic); f.color.rgb=o.get('color',color)
    return tb
def kicker(s,t,dark=False,x=I(0.75),y=I(0.6)):
    shp(s,SH.OVAL,x,y+I(0.06),I(0.14),I(0.14),fill=BRAISE)
    tx(s,x+I(0.28),y,I(9),I(0.3),t.upper(),size=12,bold=True,color=OR if dark else BRAISE)
def big_num(s,n,dark=False):
    tx(s,I(9.3),I(-0.55),I(4.2),I(3.2),n,size=210,font=SERIF,color=NUIT2 if dark else KRAFT2,al=AL.RIGHT)
def foot(s,i,dark=False):
    c=GRIS
    tx(s,I(0.75),I(6.95),I(8),I(0.3),"LE P'TIT FEU  ·  BOULANGERIE PETITJEAN",size=9,color=c,bold=True)
    tx(s,I(11.6),I(6.95),I(1),I(0.3),f"{i:02d} / 10",size=9,color=c,bold=True,al=AL.RIGHT)
def ph(s,x,y,w,h,lab,rot=0,fill=KRAFT2,tc=GRIS):
    r=shp(s,SH.RECTANGLE,x,y,w,h,fill=fill,line=GRIS,dash=DASH.DASH,rot=rot)
    tx(s,x,y,w,h,[lab,[("glisser l'image ici",{'size':11,'italic':True})]],size=14,color=tc,al=AL.CENTER,an=AN.MIDDLE,rot=rot)
def new(c):
    s=p.slides.add_slide(p.slide_layouts[6]); bg(s,c); return s

# ---------- 1 TITRE
s=new(NUIT)
shp(s,SH.OVAL,I(7.6),I(-1.2),I(7.5),I(7.5),grad=(BRAISE,OR,315))
shp(s,SH.OVAL,I(6.9),I(4.9),I(1.3),I(1.3),fill=BRAISE)
tx(s,I(0.75),I(0.7),I(8),I(0.3),"PROJET CRÉATIF  ·  DESIGN THINKING  ·  R3.MDEE.16",size=12,bold=True,color=OR)
tx(s,I(0.75),I(1.6),I(9),I(1.3),"Le",size=60,font=SERIF,italic=True,color=CREME)
tx(s,I(0.75),I(2.35),I(10),I(2.2),"P'tit Feu",size=120,font=SERIF,color=CREME)
tx(s,I(0.8),I(4.45),I(6.5),I(1.2),"La formule midi au feu de bois qui ramène les étudiants du Creusot à la boulangerie.",size=20,color=KRAFT2,sp=1.2)
t=shp(s,SH.ROUNDED_RECTANGLE,I(0.8),I(5.9),I(3.3),I(0.6),fill=None,line=OR,lw=1.5)
tx(s,I(0.8),I(5.9),I(3.3),I(0.6),"Boulangerie-Pâtisserie Petitjean",size=13,bold=True,color=OR,al=AL.CENTER,an=AN.MIDDLE)
tx(s,I(4.4),I(5.9),I(5),I(0.6),"[Prénom 1] · [Prénom 2] · [Prénom 3]",size=13,color=KRAFT2,an=AN.MIDDLE)

# ---------- 2 CONTEXTE
s=new(CREME); kicker(s,"Le contexte"); foot(s,2)
tx(s,I(0.75),I(1.0),I(11),I(1.6),[[("Un artisan ",{}),("reconnu",{'italic':True,'color':BRAISE}),(",",{})],[("dans un marché qui change.",{})]],size=34,font=SERIF,sp=1.0,gap=0)
k=shp(s,SH.RECTANGLE,I(0.75),I(2.9),I(4.6),I(3.0),fill=NUIT,rot=-2)
tx(s,I(1.1),I(3.15),I(4.0),I(2.6),[[("105 rue Édith Cavell",{'size':12,'color':OR,'bold':True})],[("Le four à bois",{'size':30,'font':SERIF,'color':CREME})],[("Pain et galettes cuits au feu de bois",{'color':KRAFT2})],[("★ 4,9 / 5",{'size':26,'font':SERIF,'color':OR})],[("note moyenne des clients",{'color':KRAFT2,'size':12})]],size=14,rot=-2,gap=0.3)
for i,(a,b) in enumerate([("13","boulangeries au Creusot :\nla concurrence est forte"),("81 %","des boulangeries font\ndéjà du snacking"),("6,33 €","le panier moyen\nen boulangerie (2025)")]):
    y=I(2.75)+i*I(1.12)
    tx(s,I(6.1),y,I(2.6),I(1.0),a,size=50,font=SERIF,color=BRAISE,an=AN.MIDDLE)
    tx(s,I(8.85),y,I(3.8),I(1.0),b.split("\n"),size=15,an=AN.MIDDLE,gap=0)
    if i<2: shp(s,SH.RECTANGLE,I(6.1),y+I(1.04),I(6.5),Emu(9000),fill=KRAFT)
tx(s,I(0.75),I(6.4),I(4),I(0.3),"NOTRE DÉMARCHE",size=10,bold=True,color=GRIS)
for i,t in enumerate(["Comprendre","Définir","Imaginer","Prototyper","Tester"]):
    x=I(2.6)+i*I(2.05)
    shp(s,SH.OVAL,x,I(6.38),I(0.32),I(0.32),fill=BRAISE)
    tx(s,x,I(6.38),I(0.32),I(0.32),str(i+1),size=11,bold=True,color=CREME,al=AL.CENTER,an=AN.MIDDLE)
    tx(s,x+I(0.42),I(6.38),I(1.6),I(0.32),t,size=13,bold=True,an=AN.MIDDLE)
tx(s,I(6.1),I(6.0),I(6.5),I(0.25),"Sources : MyBoulange, Pages Jaunes, Epsimas, Zepros Resto",size=9,color=GRIS)

# ---------- 3 COMPRENDRE (post-it)
s=new(NUIT); big_num(s,"01",True); kicker(s,"Étape 1 · Comprendre",True); foot(s,3,True)
tx(s,I(0.75),I(1.0),I(11),I(0.9),"Ce que vivent les étudiants le midi",size=34,font=SERIF,color=CREME)
notes_=[("ILS DISENT","« C'est bon mais c'est cher pour un midi. »\n« J'ai 30 minutes de pause. »",-3),("ILS PENSENT","« Le pain artisanal c'est mieux… mais c'est pour les parents. »",2),
        ("ILS FONT","Mangent vite, en groupe.\nCommandent sur leur téléphone.",2.5),("ILS RESSENTENT","L'envie de manger vrai sans se ruiner.\nLa frustration de la file d'attente.",-2)]
for i,(t,d,r) in enumerate(notes_):
    x=I(0.85)+(i%2)*I(3.45); y=I(2.15)+(i//2)*I(2.25)
    shp(s,SH.RECTANGLE,x,y,I(3.15),I(2.0),fill=POST[i],rot=r)
    tx(s,x+I(0.25),y+I(0.2),I(2.7),I(1.7),[[(t,{'size':11,'bold':True,'color':BRAISE})]]+[[(l,{})] for l in d.split("\n")],size=14,rot=r,gap=0.35)
tx(s,I(7.9),I(2.1),I(2.3),I(0.4),"4 FREINS",size=13,bold=True,color=OR)
tx(s,I(10.3),I(2.1),I(2.5),I(0.4),"4 MOTIVATIONS",size=13,bold=True,color=OR)
for i,(a,b) in enumerate([("Le prix","Le goût, le four à bois"),("Le temps","Une formule complète"),("L'image « vieillotte »","Commander à l'avance"),("Les chaînes plus « fun »","Original, à partager")]):
    y=I(2.6)+i*I(0.95)
    tx(s,I(7.9),y,I(2.3),I(0.85),[[("✕  ",{'color':BRAISE,'bold':True}),(a,{})]],size=15,color=CREME,an=AN.MIDDLE)
    tx(s,I(10.3),y,I(2.5),I(0.85),[[("✓  ",{'color':OR,'bold':True}),(b,{})]],size=15,color=CREME,an=AN.MIDDLE)
    shp(s,SH.RECTANGLE,I(7.9),y+I(0.88),I(4.85),Emu(6000),fill=NUIT2)
tx(s,I(0.85),I(6.55),I(7),I(0.3),[[("Recueil : ",{'bold':True,'color':OR}),("échanges dans le groupe + [X] étudiants interrogés à l'IUT",{})]],size=12,color=KRAFT2)

# ---------- 4 DÉFINIR (carte étudiante + HMW)
s=new(CREME); big_num(s,"02"); kicker(s,"Étape 2 · Définir"); foot(s,4)
tx(s,I(0.75),I(1.0),I(7),I(0.9),"Léa, et notre question",size=34,font=SERIF)
c=shp(s,SH.ROUNDED_RECTANGLE,I(0.75),I(2.05),I(5.3),I(4.55),fill=C(0xFF,0xFF,0xFF),line=KRAFT,rot=-1.5); c.adjustments[0]=0.05
shp(s,SH.RECTANGLE,I(0.82),I(2.12),I(5.16),I(0.55),fill=BRAISE,rot=-1.5)
tx(s,I(1.05),I(2.13),I(4.8),I(0.5),"CARTE ÉTUDIANTE  ·  IUT DU CREUSOT",size=11,bold=True,color=CREME,an=AN.MIDDLE,rot=-1.5)
shp(s,SH.OVAL,I(1.05),I(2.95),I(1.25),I(1.25),grad=(BRAISE,OR,45),rot=-1.5)
tx(s,I(1.05),I(2.95),I(1.25),I(1.25),"L",size=44,font=SERIF,color=CREME,al=AL.CENTER,an=AN.MIDDLE)
tx(s,I(2.55),I(2.95),I(3.4),I(1.3),[[("Léa, 19 ans",{'size':24,'font':SERIF})],[("1re année de BUT · CROUS · petit budget",{'size':12,'color':GRIS})]],size=14,rot=-1.5,gap=0.2)
for i,(a,b) in enumerate([("Le midi","Fast-food ou sandwich, 3 fois par semaine"),("Son frein","Trop cher et trop lent"),("La ferait revenir","< 7 €, commandé sur son téléphone, original")]):
    y=I(4.4)+i*I(0.7)
    tx(s,I(1.05),y,I(4.8),I(0.65),[[(a.upper()+"   ",{'size':10,'bold':True,'color':BRAISE}),(b,{})]],size=14,rot=-1.5,an=AN.MIDDLE)
shp(s,SH.RECTANGLE,I(6.6),I(2.05),I(6.0),I(4.55),grad=(BRAISE,OR,300))
tx(s,I(6.85),I(1.55),I(1.5),I(1.4),"“",size=120,font=SERIF,color=NUIT)
tx(s,I(7.05),I(2.55),I(5.3),I(3.9),[[("Comment pourrait-on proposer aux étudiants du Creusot une offre de repas ",{}),("rapide, abordable et originale",{'bold':True,'color':NUIT}),(" qui valorise le ",{}),("four à bois",{'bold':True,'color':NUIT}),(" de Ludovic Petitjean, pour qu'ils le choisissent plutôt que la restauration rapide ?",{})]],size=21,font=SERIF,color=CREME,sp=1.2)

# ---------- 5 IMAGINER (outil)
s=new(KRAFT2); big_num(s,"03"); kicker(s,"Étape 3 · Imaginer"); foot(s,5)
tx(s,I(0.75),I(1.0),I(8),I(0.9),"Le brainstorming, et pourquoi lui",size=34,font=SERIF)
tx(s,I(0.75),I(2.1),I(4.4),I(1.3),[[("Notre problème est large : ",{'bold':True}),("offre, prix, rapidité, image. Il fallait d'abord un maximum de pistes.",{})]],size=16,sp=1.15)
tx(s,I(0.75),I(3.45),I(4.4),I(0.3),"LES RÈGLES",size=11,bold=True,color=BRAISE)
for i,t in enumerate(["15 minutes chronométrées","Zéro critique pendant la séance","La quantité d'abord","Tri et vote seulement après"]):
    y=I(3.85)+i*I(0.6)
    shp(s,SH.OVAL,I(0.78),y+I(0.1),I(0.28),I(0.28),fill=NUIT)
    tx(s,I(0.78),y+I(0.1),I(0.28),I(0.28),str(i+1),size=10,bold=True,color=CREME,al=AL.CENTER,an=AN.MIDDLE)
    tx(s,I(1.25),y,I(3.9),I(0.5),t,size=15,an=AN.MIDDLE)
tx(s,I(5.5),I(2.1),I(3.2),I(0.3),"ET PAS…",size=11,bold=True,color=BRAISE)
for i,(a,b) in enumerate([("SCAMPER","modifie un produit existant ; nous créons une offre nouvelle"),("Méthode 635","demande 6 participants"),("Mind mapping","structure, mais produit moins d'idées")]):
    y=I(2.5)+i*I(1.35)
    shp(s,SH.RECTANGLE,I(5.5),y,I(3.2),I(1.2),fill=CREME)
    tx(s,I(5.7),y+I(0.12),I(2.9),I(1.0),[[(a,{'bold':True,'size':15})],[(b,{'size':12,'color':GRIS})]],size=13,gap=0.2)
ph(s,I(9.05),I(2.2),I(3.55),I(4.3),"Photo de nos post-it",rot=3,fill=CREME)

# ---------- 6 TRI (ardoise)
s=new(NUIT); kicker(s,"Étape 3 · Imaginer",True); foot(s,6,True)
tx(s,I(0.75),I(1.0),I(10),I(0.9),"Nos idées passées au tamis de 3 critères",size=32,font=SERIF,color=CREME)
crit=["Frein de Léa","Four à bois","Faisable seul"]
ideas=[("Formule midi à prix fixe",[1,1,1]),("Galette garnie au feu de bois",[1,1,1]),("Précommande QR + retrait sans file",[1,0,1]),("Carte de fidélité numérique",[1,0,1]),("Recette du mois votée sur Instagram",[0,1,1]),("Packaging kraft « cuit au feu de bois »",[0,1,1]),("Borne de commande en boutique",[1,0,0]),("Gâteaux trompe-l'œil",[0,0,1]),("Coin goûter / révisions",[0,0,0])]
x0=I(0.75); cx=[I(7.0),I(8.75),I(10.5)]
for j,c_ in enumerate(crit): tx(s,cx[j],I(1.95),I(1.75),I(0.4),c_.upper(),size=11,bold=True,color=OR,al=AL.CENTER)
for i,(idea,v) in enumerate(ideas):
    y=I(2.3)+i*I(0.44); keep=sum(v)>=2
    tx(s,x0,y,I(6.2),I(0.42),[[("●  " if keep else "○  ",{'color':BRAISE if keep else GRIS}),(idea,{})]],size=15,font=MONO,color=CREME if keep else GRIS,bold=keep,an=AN.MIDDLE)
    for j,ok in enumerate(v): tx(s,cx[j],y,I(1.75),I(0.42),"✓" if ok else "·",size=18,color=OR if ok else GRIS,bold=True,al=AL.CENTER,an=AN.MIDDLE)
    if not keep: shp(s,SH.RECTANGLE,x0+I(0.35),y+I(0.21),Emu(int(len(idea)*I(0.125))),Emu(12000),fill=GRIS)
r=shp(s,SH.ROUNDED_RECTANGLE,I(0.75),I(6.48),I(7.2),I(0.4),fill=BRAISE); 
tx(s,I(0.75),I(6.48),I(7.2),I(0.4),"Au moins 2 critères sur 3 = on garde. Le reste est barré.",size=13,bold=True,color=CREME,al=AL.CENTER,an=AN.MIDDLE)

# ---------- 7 CONCEPT (ticket de caisse) + parcours
s=new(CREME); kicker(s,"Le concept retenu"); foot(s,7)
tk=shp(s,SH.RECTANGLE,I(0.9),I(0.95),I(4.3),I(5.75),fill=C(0xFF,0xFF,0xFF),line=KRAFT,rot=-2.5)
lines=[[("BOULANGERIE PETITJEAN",{'bold':True})],[("105 rue Édith Cavell",{'size':11})],[("- - - - - - - - - - - - - - -",{'color':GRIS})],
       [("FORMULE LE P'TIT FEU",{'bold':True,'color':BRAISE})],[("1 galette feu de bois",{})],[("   (3 recettes dont",{'size':11})],[("    1 votée sur Insta)",{'size':11})],
       [("1 boisson",{})],[("1 mini-dessert maison",{})],[("- - - - - - - - - - - - - - -",{'color':GRIS})],[("TOTAL         6,50 €",{'bold':True,'size':20})],
       [("Précommandé sur mobile",{'size':11})],[("Fidélité : 9 / 10 ●●●●●●●●●○",{'size':11,'color':BRAISE})]]
tx(s,I(1.2),I(1.25),I(3.7),I(5.3),lines,size=14,font=MONO,rot=-2.5,sp=1.0,gap=0.15,al=AL.LEFT)
tx(s,I(5.9),I(0.95),I(7),I(1.3),[[("Le ",{}),("P'tit Feu",{'italic':True,'color':BRAISE})]],size=54,font=SERIF)
tx(s,I(5.95),I(2.1),I(6.6),I(0.8),"Une formule midi au feu de bois, précommandée sur téléphone, récupérée sans file.",size=17,color=GRIS,sp=1.15)
tx(s,I(5.95),I(3.05),I(6),I(0.3),"LE PARCOURS DE LÉA",size=11,bold=True,color=BRAISE)
shp(s,SH.RECTANGLE,I(6.15),I(3.6),Emu(15000),I(3.05),fill=KRAFT)
for i,(h_,d) in enumerate([("10h00","Voit l'affiche à l'IUT, scanne le QR code"),("10h02","Choisit et paie : 6,50 €"),("12h05","Comptoir P'tit Feu : zéro file"),("12h15","Le sac kraft part en story"),("Ensuite","Sa carte de fidélité avance")]):
    y=I(3.5)+i*I(0.64)
    shp(s,SH.OVAL,I(6.04),y+I(0.12),I(0.24),I(0.24),fill=BRAISE if i in (0,2) else NUIT)
    tx(s,I(6.5),y,I(1.2),I(0.5),h_,size=16,font=SERIF,color=BRAISE,an=AN.MIDDLE)
    tx(s,I(7.75),y,I(4.9),I(0.5),d,size=15,an=AN.MIDDLE)

# ---------- 8 PROTOTYPE (polaroids + phone)
s=new(NUIT); big_num(s,"04",True); kicker(s,"Étape 4 · Prototyper",True); foot(s,8,True)
tx(s,I(0.75),I(1.0),I(10),I(0.9),"Trois éléments, trois freins levés",size=34,font=SERIF,color=CREME)
for k,(x,rot,lab,t,d) in enumerate([(I(0.8),-3,"Affiche / visuel Instagram","L'affiche & Instagram","contre l'image « vieillotte »"),(I(4.75),2,"Packaging kraft","Le sac kraft","contre le manque de visibilité")]):
    shp(s,SH.RECTANGLE,x,I(2.1),I(3.6),I(4.2),fill=CREME,rot=rot)
    ph(s,x+I(0.2),I(2.3),I(3.2),I(3.0),lab,rot=rot,fill=KRAFT2)
    tx(s,x+I(0.2),I(5.4),I(3.2),I(0.8),[[(t,{'font':SERIF,'size':17})],[(d,{'size':12,'color':BRAISE,'bold':True})]],size=14,rot=rot,al=AL.CENTER,gap=0.1)
ph_=shp(s,SH.ROUNDED_RECTANGLE,I(9.1),I(1.85),I(2.6),I(4.75),fill=C(0x0E,0x0B,0x0A),line=GRIS,lw=2); ph_.adjustments[0]=0.12
shp(s,SH.ROUNDED_RECTANGLE,I(10.05),I(1.98),I(0.7),I(0.12),fill=NUIT2)
ph(s,I(9.25),I(2.25),I(2.3),I(4.15),"Écran de précommande",fill=NUIT2,tc=KRAFT)
tx(s,I(8.6),I(6.62),I(3.6),I(0.3),"L'écran de précommande · contre la file d'attente",size=11,bold=True,color=OR,al=AL.CENTER)

# ---------- 9 TESTER (bulles)
s=new(CREME); big_num(s,"05"); kicker(s,"Étape 5 · Tester"); foot(s,9)
tx(s,I(0.75),I(1.0),I(9),I(0.9),"Ludovic Petitjean nous a dit…",size=34,font=SERIF)
ob=[("« C'est rentable, à 6,50 € ? »","Proche du panier moyen (6,33 €). Four déjà chaud, 3 recettes seulement : peu de pertes."),
    ("« Je n'ai pas le temps de gérer une appli. »","→ On a supprimé la borne et l'appli : un formulaire en ligne + un QR code."),
    ("« Ça ne va pas abîmer mon image ? »","Le four à bois et ses classiques revisités sont au centre de tout."),
    ("« Comment les jeunes vont le savoir ? »","Affiches avec QR code au lycée et à l'IUT, Instagram, recette du mois votée.")]
for i,(a,b) in enumerate(ob):
    x=I(0.75)+(i%2)*I(6.05); y=I(2.0)+(i//2)*I(2.3)
    bub=shp(s,SH.ROUNDED_RECTANGLE,x,y,I(5.75),I(0.85),fill=NUIT); bub.adjustments[0]=0.3
    tri=shp(s,SH.ISOSCELES_TRIANGLE,x+I(0.5),y+I(0.8),I(0.35),I(0.25),fill=NUIT,rot=180)
    tx(s,x+I(0.3),y,I(5.2),I(0.85),a,size=16,font=SERIF,italic=True,color=CREME,an=AN.MIDDLE)
    tx(s,x+I(0.3),y+I(1.15),I(5.3),I(0.95),b,size=15,sp=1.1)
r=shp(s,SH.RECTANGLE,I(0.75),I(6.45),I(11.85),I(0.42),fill=KRAFT2)
tx(s,I(0.95),I(6.45),I(11.5),I(0.42),[[("Ce que le test a changé : ",{'bold':True,'color':BRAISE}),("la borne a disparu du prototype, remplacée par la précommande sur téléphone.",{})]],size=13,an=AN.MIDDLE)

# ---------- 10 CONCLUSION
s=new(NUIT); foot(s,10,True)
shp(s,SH.OVAL,I(-2.5),I(3.2),I(7),I(7),grad=(BRAISE,OR,45))
kicker(s,"Ce que ça change",True)
tx(s,I(0.75),I(1.0),I(11.8),I(1.6),[[("On lève les trois freins de Léa avec ce que Ludovic Petitjean a de plus fort : ",{}),("son four à bois.",{'italic':True,'color':OR})]],size=30,font=SERIF,color=CREME,sp=1.15)
for col,(t,items) in enumerate([("Pour Léa",["Un vrai repas à 6,50 €","Zéro file d'attente","Un produit à partager","Une 10e formule offerte"]),("Pour la boulangerie",["Le public du midi, parti au fast-food","Le four à bois au cœur de l'offre","De la visibilité gratuite","Une différence face aux 12 autres"])]):
    x=I(5.0)+col*I(4.0)
    tx(s,x,I(2.95),I(3.7),I(0.5),t,size=22,font=SERIF,color=OR)
    shp(s,SH.RECTANGLE,x,I(3.5),I(3.6),Emu(12000),fill=BRAISE)
    tx(s,x,I(3.7),I(3.7),I(2.3),items,size=15,color=CREME,gap=0.45)
tx(s,I(5.0),I(6.2),I(7.6),I(0.5),"Merci ! Vos questions ?",size=24,font=SERIF,italic=True,color=CREME)

for sl,t in zip(p.slides,oral): sl.notes_slide.notes_text_frame.text=t
p.save(sys.argv[2]); print(len(p.slides._sldIdLst))
