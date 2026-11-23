import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CREME=RGBColor(0xFA,0xF5,0xEC); BRUN=RGBColor(0x3A,0x27,0x1B); FEU=RGBColor(0xD4,0x5D,0x24)
KRAFT=RGBColor(0xE9,0xDC,0xC6); GRIS=RGBColor(0x7A,0x6A,0x5C); BLANC=RGBColor(0xFF,0xFF,0xFF)
TITRE="Georgia"; TXT="Calibri"

p=Presentation(); p.slide_width=Inches(13.333); p.slide_height=Inches(7.5)
W,H=p.slide_width,p.slide_height
blank=p.slide_layouts[6]

def bg(s,c=CREME):
    f=s.background.fill; f.solid(); f.fore_color.rgb=c
def box(s,x,y,w,h,fill=None,line=None,shape=MSO_SHAPE.RECTANGLE):
    r=s.shapes.add_shape(shape,x,y,w,h)
    if fill is None: r.fill.background()
    else: r.fill.solid(); r.fill.fore_color.rgb=fill
    if line is None: r.line.fill.background()
    else: r.line.color.rgb=line; r.line.width=Pt(1.25)
    r.shadow.inherit=False
    return r
def text(s,x,y,w,h,runs,size=18,color=BRUN,font=TXT,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,spacing=1.15):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    tf.vertical_anchor=anchor
    for a in ('margin_left','margin_right','margin_top','margin_bottom'): setattr(tf,a,0)
    if isinstance(runs,str): runs=[runs]
    for i,para in enumerate(runs):
        par=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        par.alignment=align; par.line_spacing=spacing
        if i: par.space_before=Pt(size*0.5)
        if isinstance(para,str): para=[(para,{})]
        for t,o in para:
            r=par.add_run(); r.text=t
            r.font.size=Pt(o.get('size',size)); r.font.name=o.get('font',font)
            r.font.bold=o.get('bold',bold); r.font.italic=o.get('italic',False)
            r.font.color.rgb=o.get('color',color)
    return tb
def header(s,num,kicker,titre):
    bg(s)
    box(s,0,0,Inches(0.18),H,fill=FEU)
    text(s,Inches(0.8),Inches(0.55),Inches(9),Inches(0.4),kicker.upper(),size=13,color=FEU,bold=True)
    text(s,Inches(0.8),Inches(0.9),Inches(11.5),Inches(1),titre,size=36,font=TITRE)
    text(s,Inches(11.8),Inches(6.85),Inches(1),Inches(0.3),f"{num:02d}",size=12,color=GRIS,align=PP_ALIGN.RIGHT)
    text(s,Inches(0.8),Inches(6.85),Inches(8),Inches(0.3),"Le P'tit Feu · Boulangerie-Pâtisserie Petitjean",size=11,color=GRIS)
def notes(s,t): s.notes_slide.notes_text_frame.text=t

N=[]  # notes per slide, filled from oral
import json
oral=json.load(open(sys.argv[1]))

def card(s,x,y,w,h,t,d,tsize=19,dsize=15,top=True,fill=BLANC,tc=BRUN,dc=GRIS):
    box(s,x,y,w,h,fill=fill)
    if top: box(s,x,y,w,Inches(0.07),fill=FEU)
    text(s,x+Inches(0.25),y+Inches(0.25),w-Inches(0.5),Inches(0.6),t,size=tsize,bold=True,color=tc)
    text(s,x+Inches(0.25),y+Inches(0.25)+Inches(tsize/72*1.6),w-Inches(0.5),h-Inches(0.9),d.split("\n"),size=dsize,color=dc,spacing=1.05)
def placeholder(s,x,y,w,h,lab):
    r=box(s,x,y,w,h,fill=KRAFT,line=GRIS); r.line.dash_style=7
    text(s,x,y,w,h,[lab,"(clic droit › Modifier l'image, ou glisser l'image)"],size=15,color=GRIS,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
n=[0]
def new(kicker=None,titre=None,dark=False):
    n[0]+=1; s=p.slides.add_slide(blank)
    if dark: bg(s,BRUN)
    else: header(s,n[0],kicker,titre)
    return s

from PIL import Image as _I
import os as _os
VIS=_os.path.join(_os.path.dirname(_os.path.abspath(sys.argv[2])),'visuels')
def pic(s,name,x,y,w,h,focus=0.5):
    im=_I.open(_os.path.join(VIS,name)); W0,H0=im.size; r=w/h
    if W0/H0>r: nw=int(H0*r); l=int((W0-nw)*focus); im=im.crop((l,0,l+nw,H0))
    else: nh=int(W0/r); t=int((H0-nh)*focus); im=im.crop((0,t,W0,t+nh))
    tmp='/tmp/_crop_%s'%name; im.save(tmp,quality=92); s.shapes.add_picture(tmp,x,y,w,h)
def lbl(s,x,y,w,t): text(s,x,y,w,Inches(0.35),t.upper(),size=13,color=FEU,bold=True)

# 1 Titre
s=new(dark=True)
box(s,Inches(0.8),Inches(2.2),Inches(0.9),Inches(0.08),fill=FEU)
text(s,Inches(0.8),Inches(1.5),Inches(10),Inches(0.5),"PROJET CRÉATIF · DESIGN THINKING · R3.MDEE.16",size=14,color=FEU,bold=True)
text(s,Inches(0.8),Inches(2.5),Inches(11),Inches(1.6),"Le P'tit Feu",size=80,color=CREME,font=TITRE)
text(s,Inches(0.8),Inches(4.0),Inches(11),Inches(1),"La formule midi au feu de bois qui ramène les étudiants du Creusot à la boulangerie",size=24,color=KRAFT)
text(s,Inches(0.8),Inches(6.2),Inches(11),Inches(0.5),"Boulangerie-Pâtisserie Petitjean  ·  BUT Techniques de commercialisation  ·  [Prénoms du groupe]",size=14,color=KRAFT)

# 2 Contexte + démarche
s=new("Le contexte","Un artisan reconnu, un marché qui change")
card(s,Inches(0.8),Inches(2.0),Inches(4.6),Inches(3.0),"Boulangerie Petitjean","105 rue Édith Cavell, Le Creusot\nPain et galettes au four à bois\nNote moyenne de 4,9/5",dsize=16)
for i,(a,b) in enumerate([("13","boulangeries au Creusot"),("81 %","des boulangeries font du snacking"),("6,33 €","panier moyen en 2025")]):
    x=Inches(5.6)+i*Inches(2.33)
    box(s,x,Inches(2.0),Inches(2.2),Inches(3.0),fill=BLANC)
    text(s,x+Inches(0.2),Inches(2.3),Inches(1.9),Inches(1),a,size=34,color=FEU,font=TITRE)
    text(s,x+Inches(0.2),Inches(3.4),Inches(1.85),Inches(1.5),b,size=15)
lbl(s,Inches(0.8),Inches(5.3),Inches(6),"Notre démarche")
for i,t in enumerate(["1 Comprendre","2 Définir","3 Imaginer","4 Prototyper","5 Tester"]):
    x=Inches(0.8)+i*Inches(2.36)
    box(s,x,Inches(5.7),Inches(2.24),Inches(0.6),fill=BRUN)
    text(s,x,Inches(5.7),Inches(2.24),Inches(0.6),t,size=15,color=CREME,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(0.8),Inches(6.45),Inches(11.7),Inches(0.3),"Sources : MyBoulange, Pages Jaunes, Epsimas 2025-2026, Zepros Resto 2025",size=10,color=GRIS)

# 3 Comprendre
s=new("Étape 1 · Comprendre","Ce que vivent les étudiants le midi")
q=[("Ils disent","« C'est bon mais c'est cher pour un midi. »\n« J'ai 30 minutes de pause. »"),("Ils pensent","Le pain artisanal, c'est mieux… mais c'est pour les parents."),
   ("Ils font","Mangent vite, en groupe, commandent sur leur téléphone."),("Ils ressentent","Envie de manger « vrai » sans se ruiner ; frustration dans la file.")]
bw=Inches(3.3); bh=Inches(1.95)
for i,(t,d) in enumerate(q):
    x=Inches(0.8)+(i%2)*(bw+Inches(0.15)); y=Inches(2.0)+(i//2)*(bh+Inches(0.15))
    box(s,x,y,bw,bh,fill=BLANC)
    lbl(s,x+Inches(0.2),y+Inches(0.15),bw,t)
    text(s,x+Inches(0.2),y+Inches(0.55),bw-Inches(0.4),bh-Inches(0.6),d.split("\n"),size=14,spacing=1.0)
fr=["Prix : 5 à 7 € max le midi","Temps : pause courte, file au rush","Image « traditionnelle »","Chaînes de snacking plus « fun »"]
mo=["Goût du fait maison, four à bois","Formule complète à bon prix","Commander avant, sans attendre","Produit original, fidélité"]
for col,(t,items,c) in enumerate([("Freins",fr,BRUN),("Motivations",mo,FEU)]):
    x=Inches(7.9)+col*Inches(2.35)
    text(s,x,Inches(2.0),Inches(2.2),Inches(0.4),t,size=18,bold=True,color=c)
    for i,it in enumerate(items):
        y=Inches(2.5)+i*Inches(0.95)
        box(s,x,y,Inches(2.2),Inches(0.85),fill=BLANC); box(s,x,y,Inches(0.07),Inches(0.85),fill=c)
        text(s,x+Inches(0.2),y,Inches(1.9),Inches(0.85),it,size=13,anchor=MSO_ANCHOR.MIDDLE,spacing=1.0)
text(s,Inches(0.8),Inches(6.35),Inches(11.7),Inches(0.4),[[("Recueil : ",{'bold':True,'color':FEU}),("échanges au sein de notre groupe de l'IUT",{})]],size=13)

# 4 Définir
s=new("Étape 2 · Définir","Léa, 19 ans, et notre question")
box(s,Inches(0.8),Inches(2.0),Inches(5.4),Inches(4.4),fill=BLANC)
box(s,Inches(1.1),Inches(2.3),Inches(1.1),Inches(1.1),fill=FEU,shape=MSO_SHAPE.OVAL)
text(s,Inches(1.1),Inches(2.3),Inches(1.1),Inches(1.1),"L",size=40,color=CREME,font=TITRE,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(2.45),Inches(2.35),Inches(3.6),Inches(1.1),[[("Léa, 19 ans",{'size':22,'bold':True})],[("BUT à l'IUT du Creusot, CROUS, petit budget",{'color':GRIS})]],size=14,spacing=1.0)
rows=[("Habitudes","Fast-food ou sandwich 3 midis par semaine."),("Frein principal","Trop cher et trop lent le midi."),("La ferait revenir","Formule < 7 €, commandée sur son téléphone, originale.")]
for i,(t,d) in enumerate(rows):
    y=Inches(3.55)+i*Inches(0.95)
    lbl(s,Inches(1.1),y,Inches(4.8),t)
    text(s,Inches(1.1),y+Inches(0.33),Inches(4.9),Inches(0.6),d,size=14,spacing=1.0)
box(s,Inches(6.5),Inches(2.0),Inches(6.0),Inches(4.4),fill=BRUN)
lbl(s,Inches(6.85),Inches(2.25),Inches(5),"Comment pourrait-on…")
text(s,Inches(6.85),Inches(2.7),Inches(5.35),Inches(3.6),[[("…proposer aux étudiants du Creusot une offre de repas ",{}),("rapide, abordable et originale",{'color':FEU,'bold':True}),(" qui valorise le ",{}),("four à bois",{'color':FEU,'bold':True}),(" de Ludovic Petitjean, pour qu'ils choisissent la boulangerie plutôt que la restauration rapide ?",{})]],size=21,color=CREME,font=TITRE,spacing=1.2)

# 5 Outil
s=new("Étape 3 · Imaginer","Le brainstorming, et pourquoi lui")
lbl(s,Inches(0.8),Inches(2.0),Inches(4),"Pourquoi ce choix")
text(s,Inches(0.8),Inches(2.4),Inches(3.9),Inches(1.3),"Notre problème est large (offre, prix, rapidité, image) : il fallait d'abord un maximum de pistes.",size=15)
lbl(s,Inches(0.8),Inches(3.75),Inches(4),"Les règles respectées")
text(s,Inches(0.8),Inches(4.15),Inches(3.9),Inches(2.2),["15 minutes chronométrées","Aucune critique pendant la séance","La quantité d'abord","Tri et vote seulement après"],size=14,spacing=0.95)
lbl(s,Inches(4.9),Inches(2.0),Inches(3.5),"Pourquoi pas un autre outil")
for i,(a,b) in enumerate([("SCAMPER","fait évoluer un produit existant ; nous créons une offre nouvelle"),("Méthode 635","demande 6 participants"),("Mind mapping","structure, mais produit moins d'idées")]):
    y=Inches(2.4)+i*Inches(1.33)
    box(s,Inches(4.9),y,Inches(3.4),Inches(1.2),fill=BLANC)
    text(s,Inches(5.1),y+Inches(0.12),Inches(3.0),Inches(1.0),[[(a,{'bold':True,'size':15})],[(b,{'color':GRIS})]],size=13,spacing=1.0)
box(s,Inches(8.55),Inches(2.0),Inches(3.95),Inches(4.4),fill=BRUN)
lbl(s,Inches(8.8),Inches(2.2),Inches(3.5),"Les idées de la séance")
text(s,Inches(8.8),Inches(2.65),Inches(3.5),Inches(3.7),["Formule midi à prix fixe","Galettes street-food","Mini-pizzas à partager","Gâteaux trompe-l'œil","Recette du mois sur Instagram","Borne de commande","Précommande QR code","Créneau étudiants 11h45-13h30","Carte de fidélité numérique","Packaging kraft","Instagram + affiches","Coin goûter / révisions"],size=12,color=CREME,spacing=1.0)

# 6 Idées & tri
s=new("Étape 3 · Imaginer","Nos idées, triées avec 3 critères")
crit=["Répond au frein de Léa","Valorise le four à bois","Faisable par un artisan seul"]
ideas=[("Formule midi à prix fixe",[1,1,1]),("Galette garnie au feu de bois",[1,1,1]),("Borne de commande en boutique",[1,0,1]),("Précommande en ligne (QR code)",[1,0,1]),("Carte de fidélité numérique",[1,0,1]),("Recette du mois votée sur Instagram",[0,1,1]),("Packaging kraft « cuit au feu de bois »",[0,1,1]),("Gâteaux trompe-l'œil",[0,0,1]),("Coin goûter / révisions",[0,0,0])]
x0=Inches(0.8); cwid=[Inches(5.0),Inches(2.25),Inches(2.25),Inches(2.2)]
xs=[x0]
for w in cwid[:-1]: xs.append(xs[-1]+w)
y=Inches(1.95); tw=sum(cwid,Emu(0))
box(s,x0,y,tw,Inches(0.6),fill=BRUN)
text(s,xs[0]+Inches(0.2),y,cwid[0],Inches(0.6),"Idée du brainstorming",size=14,color=CREME,bold=True,anchor=MSO_ANCHOR.MIDDLE)
for j,c in enumerate(crit): text(s,xs[j+1]+Inches(0.1),y,cwid[j+1]-Inches(0.2),Inches(0.6),c,size=12,color=CREME,bold=True,anchor=MSO_ANCHOR.MIDDLE,align=PP_ALIGN.CENTER,spacing=1.0)
for i,(idea,v) in enumerate(ideas):
    yy=y+Inches(0.6)+i*Inches(0.42); keep=sum(v)>=2
    box(s,x0,yy,tw,Inches(0.4),fill=BLANC if keep else KRAFT)
    text(s,xs[0]+Inches(0.2),yy,cwid[0],Inches(0.4),idea,size=14,bold=keep,color=BRUN if keep else GRIS,anchor=MSO_ANCHOR.MIDDLE)
    for j,ok in enumerate(v): text(s,xs[j+1],yy,cwid[j+1],Inches(0.4),"✓" if ok else "–",size=16,color=FEU if ok else GRIS,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(0.8),Inches(6.4),Inches(11.7),Inches(0.35),"En gras : au moins 2 critères sur 3, ce sont les idées qui forment le concept.",size=13,color=GRIS)

# 7 Concept + parcours
s=new(dark=True)
text(s,Inches(0.8),Inches(0.5),Inches(9),Inches(0.4),"LE CONCEPT RETENU",size=13,color=FEU,bold=True)
text(s,Inches(0.8),Inches(0.85),Inches(7),Inches(1.0),"Le P'tit Feu",size=48,color=CREME,font=TITRE)
text(s,Inches(0.8),Inches(1.75),Inches(8.2),Inches(0.9),"Galette au feu de bois (3 recettes dont 1 votée sur Instagram) + boisson + mini-dessert maison. Commande sur la borne en boutique ou à l'avance en ligne. La 10e est offerte.",size=16,color=KRAFT)
r=box(s,Inches(9.5),Inches(0.6),Inches(3.0),Inches(1.6),fill=FEU,shape=MSO_SHAPE.ROUNDED_RECTANGLE); r.adjustments[0]=0.1
text(s,Inches(9.5),Inches(0.65),Inches(3.0),Inches(1.1),"6,50 €",size=46,color=BLANC,font=TITRE,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(9.5),Inches(1.7),Inches(3.0),Inches(0.4),"la formule complète",size=13,color=BLANC,align=PP_ALIGN.CENTER)
lbl(s,Inches(0.8),Inches(3.05),Inches(8),"Le parcours de Léa")
st=[("10h","Elle voit l'affiche à l'IUT et scanne le QR code"),("10h02","Elle choisit et paie : 6,50 €"),("12h05","Comptoir P'tit Feu : pas de file"),("12h15","Le sac kraft finit en story"),("Ensuite","Sa carte de fidélité avance")]
cw=Inches(2.24); gap=Inches(0.125)
for i,(h_,d) in enumerate(st):
    x=Inches(0.8)+i*(cw+gap); y=Inches(3.5)
    box(s,x,y,cw,Inches(2.9),fill=RGBColor(0x4A,0x35,0x27))
    text(s,x+Inches(0.25),y+Inches(0.3),cw,Inches(0.6),h_,size=26,color=FEU,font=TITRE)
    text(s,x+Inches(0.25),y+Inches(1.15),cw-Inches(0.5),Inches(1.6),d,size=16,color=CREME)

# 8 Prototype
s=new("Étape 4 · Prototyper","Un prototype en 3 éléments, un frein chacun")
for k,(img,f,t,d) in enumerate([("instagram.jpg",1.0,"Instagram & visuels","Contre l'image « traditionnelle »"),("packaging-kraft.jpg",0.75,"Packaging kraft","Contre le manque de visibilité"),("borne-boutique.jpg",0.45,"Borne + commande en ligne","Contre la file et le manque de temps")]):
    x=Inches(0.8)+k*Inches(3.98)
    pic(s,img,x,Inches(2.0),Inches(3.75),Inches(3.3),f)
    text(s,x,Inches(5.45),Inches(3.75),Inches(0.4),t,size=18,bold=True)
    text(s,x,Inches(5.85),Inches(3.75),Inches(0.5),d,size=14,color=FEU,bold=True)

# 9 Tester
s=new("Étape 5 · Tester","Les objections de L. Petitjean")
ob=[("« C'est rentable, à 6,50 € ? »","Proche du panier moyen (6,33 €). Four déjà chaud, 3 recettes seulement. La borne est un investissement courant dans le métier."),
    ("« Je n'ai pas le temps de gérer une appli. »","Pas d'appli à créer : une borne et une page de commande prêtes à l'emploi, avec le même menu de 3 recettes."),
    ("« Ça ne va pas abîmer mon image ? »","Le four à bois et les classiques revisités sont au centre de tout."),
    ("« Comment les jeunes vont le savoir ? »","Affiches avec QR code au lycée et à l'IUT, Instagram, vote de la recette du mois.")]
for i,(a,b) in enumerate(ob):
    x=Inches(0.8)+(i%2)*Inches(5.95); y=Inches(2.0)+(i//2)*Inches(2.2)
    box(s,x,y,Inches(5.75),Inches(2.05),fill=BLANC); box(s,x,y,Inches(0.08),Inches(2.05),fill=FEU)
    text(s,x+Inches(0.3),y+Inches(0.2),Inches(5.25),Inches(0.5),a,size=16,bold=True,font=TITRE)
    text(s,x+Inches(0.3),y+Inches(0.95),Inches(5.25),Inches(1.05),b,size=15,color=GRIS)
text(s,Inches(0.8),Inches(6.4),Inches(11.7),Inches(0.35),"Ce que le test a changé : pas d'application dédiée, un menu réduit à 3 recettes, et la commande en ligne ajoutée à la borne.",size=13,color=GRIS)

# 10 Bénéfices + conclusion
s=new(dark=True)
text(s,Inches(0.8),Inches(0.6),Inches(9),Inches(0.4),"CE QUE ÇA CHANGE",size=13,color=FEU,bold=True)
text(s,Inches(0.8),Inches(1.0),Inches(11.7),Inches(1.2),"On lève les trois freins de Léa avec ce que Ludovic Petitjean a de plus fort : son four à bois.",size=28,color=CREME,font=TITRE,spacing=1.15)
for col,(t,items) in enumerate([("Pour Léa",["Un vrai repas à 6,50 €","Zéro file d'attente","Un produit à partager, une récompense"]),("Pour la boulangerie",["Le public du midi, parti au fast-food","Le four à bois au cœur de l'offre","Visibilité gratuite, différence face aux 12 autres"])]):
    x=Inches(0.8)+col*Inches(6.0)
    box(s,x,Inches(2.65),Inches(5.7),Inches(3.0),fill=RGBColor(0x4A,0x35,0x27))
    text(s,x+Inches(0.35),Inches(2.9),Inches(5),Inches(0.5),t,size=22,bold=True,color=FEU)
    text(s,x+Inches(0.35),Inches(3.55),Inches(5.0),Inches(2.0),items,size=18,color=CREME,spacing=1.05)
text(s,Inches(0.8),Inches(6.2),Inches(11.5),Inches(0.5),"Merci de votre attention. Avez-vous des questions ?",size=20,color=CREME,bold=True)

for sl,t in zip(p.slides,oral): notes(sl,t)
p.save(sys.argv[2]); print(len(p.slides._sldIdLst))
