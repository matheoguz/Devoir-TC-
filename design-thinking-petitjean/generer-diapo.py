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

# 1 Titre
s=new(dark=True)
box(s,Inches(0.8),Inches(2.2),Inches(0.9),Inches(0.08),fill=FEU)
text(s,Inches(0.8),Inches(1.5),Inches(10),Inches(0.5),"PROJET CRÉATIF · DESIGN THINKING · R3.MDEE.16",size=14,color=FEU,bold=True)
text(s,Inches(0.8),Inches(2.5),Inches(11),Inches(1.6),"Le P'tit Feu",size=80,color=CREME,font=TITRE)
text(s,Inches(0.8),Inches(4.0),Inches(11),Inches(1),"La formule midi au feu de bois qui ramène les étudiants du Creusot à la boulangerie",size=24,color=KRAFT)
text(s,Inches(0.8),Inches(6.2),Inches(11),Inches(0.5),"Boulangerie-Pâtisserie Petitjean  ·  BUT Techniques de commercialisation  ·  [Prénoms du groupe]",size=14,color=KRAFT)

# 2 Contexte
s=new("Le contexte","Un artisan reconnu, un marché qui change")
card(s,Inches(0.8),Inches(2.05),Inches(5.6),Inches(4.4),"Boulangerie Petitjean","105 rue Édith Cavell, Le Creusot\nPain et galettes cuits au four à bois\nNote moyenne de 4,9/5\nRevisite les classiques avec une touche perso")
k=[("13","boulangeries au Creusot : un marché de proximité très concurrentiel"),("81 %","des boulangeries font déjà du snacking ou de la restauration rapide"),("6,33 €","panier moyen en boulangerie en 2025")]
for i,(a,b) in enumerate(k):
    y=Inches(2.05)+i*Inches(1.5)
    box(s,Inches(6.8),y,Inches(5.7),Inches(1.3),fill=BLANC)
    text(s,Inches(7.05),y,Inches(1.9),Inches(1.3),a,size=32,color=FEU,font=TITRE,anchor=MSO_ANCHOR.MIDDLE)
    text(s,Inches(9.0),y,Inches(3.3),Inches(1.3),b,size=14,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(0.8),Inches(6.55),Inches(11.7),Inches(0.3),"Sources : MyBoulange, Pages Jaunes, Epsimas 2025-2026, Zepros Resto 2025",size=10,color=GRIS)

# 3 Démarche
s=new("Notre démarche","Les 5 étapes du Design Thinking")
ph=[("01","Comprendre","Carte d'empathie, freins, motivations"),("02","Définir","Persona et question HMW"),("03","Imaginer","Brainstorming, tri par critères, concept"),("04","Prototyper","Parcours client, visuels, packaging, écran"),("05","Tester","Objections de L. Petitjean, prototype corrigé")]
cw=Inches(2.24); gap=Inches(0.125)
for i,(a,t,d) in enumerate(ph):
    x=Inches(0.8)+i*(cw+gap); y=Inches(2.5)
    box(s,x,y,cw,Inches(3.4),fill=BLANC); box(s,x,y,cw,Inches(0.07),fill=FEU)
    text(s,x+Inches(0.25),y+Inches(0.35),cw,Inches(0.8),a,size=36,color=FEU,font=TITRE)
    text(s,x+Inches(0.25),y+Inches(1.3),cw-Inches(0.5),Inches(0.5),t,size=21,bold=True)
    text(s,x+Inches(0.25),y+Inches(1.9),cw-Inches(0.5),Inches(1.4),d,size=15,color=GRIS)

# 4 Empathie
s=new("Étape 1 · Comprendre","La carte d'empathie")
q=[("Ils disent","« C'est bon mais c'est cher pour un midi. »\n« J'ai 30 minutes de pause. »\n« Je prends un sandwich au Carrefour ou un tacos. »"),
   ("Ils pensent","Le pain artisanal, c'est mieux… mais c'est pour les parents.\nIl faut que ce soit rentable et que ça nourrisse.\nSi c'est beau, ça part en story."),
   ("Ils font","Mangent vite entre midi et deux, souvent en groupe.\nCommandent depuis leur téléphone.\nViennent surtout pour une viennoiserie ou un goûter."),
   ("Ils ressentent","L'envie de manger « vrai » sans se ruiner.\nDe la frustration dans les files d'attente.\nLe plaisir de partager une bonne pâtisserie.")]
bw=Inches(5.75); bh=Inches(2.0)
for i,(t,d) in enumerate(q):
    x=Inches(0.8)+(i%2)*(bw+Inches(0.2)); y=Inches(2.0)+(i//2)*(bh+Inches(0.15))
    box(s,x,y,bw,bh,fill=BLANC)
    text(s,x+Inches(0.3),y+Inches(0.18),bw,Inches(0.4),t.upper(),size=14,color=FEU,bold=True)
    text(s,x+Inches(0.3),y+Inches(0.58),bw-Inches(0.6),bh-Inches(0.7),d.split("\n"),size=15,spacing=1.0)
text(s,Inches(0.8),Inches(6.25),Inches(11.7),Inches(0.4),[[("Recueil : ",{'bold':True,'color':FEU}),("échanges dans le groupe + [X] étudiants interrogés à l'IUT (à compléter)",{})]],size=14)

# 5 Freins/motivations
s=new("Étape 1 · Comprendre","Ce qui freine, ce qui motive")
fr=["Le prix : un midi étudiant, c'est 5 à 7 €","Le temps : pause courte, file d'attente au rush","Une image « traditionnelle », absente des réseaux","Le snacking des chaînes, plus visible et plus « fun »"]
mo=["Le goût et le fait maison : le four à bois","Une formule complète, nourrissante, à bon prix","Commander avant, récupérer sans attendre","Un produit original, « instagrammable », la fidélité"]
for col,(t,items,c) in enumerate([("4 freins",fr,BRUN),("4 motivations",mo,FEU)]):
    x=Inches(0.8)+col*Inches(6.0)
    text(s,x,Inches(2.1),Inches(5.6),Inches(0.5),t,size=24,bold=True,color=c)
    for i,it in enumerate(items):
        y=Inches(2.85)+i*Inches(0.92)
        box(s,x,y,Inches(5.6),Inches(0.78),fill=BLANC); box(s,x,y,Inches(0.08),Inches(0.78),fill=c)
        text(s,x+Inches(0.3),y,Inches(5.2),Inches(0.78),it,size=17,anchor=MSO_ANCHOR.MIDDLE)

# 6 Persona
s=new("Étape 2 · Définir","Notre persona : Léa, 19 ans")
box(s,Inches(0.8),Inches(2.1),Inches(3.6),Inches(4.4),fill=BRUN)
box(s,Inches(1.75),Inches(2.45),Inches(1.7),Inches(1.7),fill=FEU,shape=MSO_SHAPE.OVAL)
text(s,Inches(1.75),Inches(2.45),Inches(1.7),Inches(1.7),"L",size=60,color=CREME,font=TITRE,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(1.0),Inches(4.4),Inches(3.2),Inches(2),[[("Léa, 19 ans",{'size':24,'bold':True,'color':CREME})],[("1re année de BUT, IUT du Creusot",{'color':KRAFT})],[("Appartement CROUS, petit budget",{'color':KRAFT})]],size=15,align=PP_ALIGN.CENTER)
rows=[("Ses habitudes","Mange dehors 3 midis par semaine : sandwich de supermarché ou fast-food entre amis. Boulangerie le week-end, pour les viennoiseries."),
      ("Son frein principal","Le rapport prix / rapidité : pour elle, la boulangerie est trop chère et trop lente le midi."),
      ("Ce qui la ferait revenir","Une formule midi à moins de 7 €, commandable sur son téléphone, avec un produit original.")]
for i,(t,d) in enumerate(rows):
    y=Inches(2.1)+i*Inches(1.5)
    text(s,Inches(4.9),y,Inches(7.5),Inches(0.4),t.upper(),size=14,color=FEU,bold=True)
    text(s,Inches(4.9),y+Inches(0.42),Inches(7.6),Inches(1),d,size=18)

# 7 HMW
s=new("Étape 2 · Définir","Notre question « Comment pourrait-on… ? »")
box(s,Inches(0.8),Inches(2.1),Inches(11.7),Inches(3.2),fill=BRUN)
text(s,Inches(1.2),Inches(1.95),Inches(1),Inches(1),"“",size=96,color=FEU,font=TITRE)
text(s,Inches(1.6),Inches(2.75),Inches(10.4),Inches(2.4),[[("Comment pourrait-on proposer aux étudiants du Creusot une offre de repas ",{}),("rapide, abordable et originale",{'color':FEU,'bold':True}),(" qui valorise le ",{}),("four à bois",{'color':FEU,'bold':True}),(" de Ludovic Petitjean, pour qu'ils choisissent la boulangerie plutôt que la restauration rapide ?",{})]],size=24,color=CREME,font=TITRE,spacing=1.2)
card(s,Inches(0.8),Inches(5.5),Inches(5.75),Inches(1.15),"Le besoin de Léa","rapide · abordable · original",tsize=16,dsize=15)
card(s,Inches(6.75),Inches(5.5),Inches(5.75),Inches(1.15),"L'identité de l'artisan","le four à bois · le savoir-faire",tsize=16,dsize=15)

# 8 Outil
s=new("Étape 3 · Imaginer","Pourquoi le brainstorming ?")
text(s,Inches(0.8),Inches(2.05),Inches(5.6),Inches(0.4),"NOTRE PROBLÈME EST LARGE",size=14,color=FEU,bold=True)
text(s,Inches(0.8),Inches(2.5),Inches(5.6),Inches(1.6),"Offre, prix, rapidité, image : il fallait d'abord ouvrir un maximum de pistes, sans se censurer.",size=19)
text(s,Inches(0.8),Inches(3.85),Inches(5.6),Inches(0.4),"LES RÈGLES RESPECTÉES",size=14,color=FEU,bold=True)
text(s,Inches(0.8),Inches(4.3),Inches(5.6),Inches(2.2),["15 minutes chronométrées","Aucune critique pendant la production","La quantité avant la qualité","On rebondit sur les idées des autres","Tri et vote seulement après"],size=16,spacing=0.95)
alt=[("SCAMPER","Fait évoluer un produit qui existe. Or nous voulions créer une offre midi qui n'existe pas encore."),("Méthode 635","Demande 6 participants : notre groupe n'avait pas cette taille."),("Mind mapping","Utile pour structurer, moins pour produire beaucoup d'idées vite.")]
text(s,Inches(6.9),Inches(2.05),Inches(5.6),Inches(0.4),"POURQUOI PAS UN AUTRE OUTIL",size=14,color=FEU,bold=True)
for i,(a,b) in enumerate(alt):
    y=Inches(2.55)+i*Inches(1.33)
    box(s,Inches(6.9),y,Inches(5.6),Inches(1.18),fill=BLANC)
    text(s,Inches(7.15),y+Inches(0.15),Inches(5.2),Inches(0.4),a,size=17,bold=True)
    text(s,Inches(7.15),y+Inches(0.52),Inches(5.2),Inches(0.7),b,size=14,color=GRIS)

# 9 Idées + trace
s=new("Étape 3 · Imaginer","Nos idées, regroupées par thèmes")
th=[("Offre","Formule midi à prix fixe\nGalettes street-food\nMini-pizzas à partager"),("Produits","Gâteaux trompe-l'œil\nRecette du mois votée sur Instagram"),("Rapidité","Borne de commande\nPrécommande QR code\nCréneau étudiants 11h45-13h30"),("Fidélité & image","Carte de fidélité numérique\nPackaging kraft\nInstagram + affiches"),("Lieu","Coin goûter / révisions le mercredi")]
cw=Inches(2.0); gap=Inches(0.1)
for i,(t,d) in enumerate(th):
    r,c=divmod(i,3)
    x=Inches(0.8)+c*(cw+gap); y=Inches(2.05)+r*Inches(2.3)
    card(s,x,y,cw,Inches(2.15),t,d,tsize=16,dsize=13)
placeholder(s,Inches(7.3),Inches(2.05),Inches(5.2),Inches(4.45),"Photo de nos post-it / feuilles\n(trace du brainstorming)")

# 10 Tri
s=new("Étape 3 · Imaginer","Le tri : 3 critères explicites")
crit=["Répond au frein de Léa (prix, temps)","Met en valeur le four à bois","Réalisable par un artisan seul"]
ideas=[("Formule midi à prix fixe",[1,1,1]),("Précommande QR code + retrait",[1,0,1]),("Galette garnie au feu de bois",[1,1,1]),("Recette du mois votée sur Instagram",[0,1,1]),("Carte de fidélité numérique",[1,0,1]),("Borne de commande en boutique",[1,0,0]),("Coin goûter / révisions",[0,0,0])]
x0=Inches(0.8); cwid=[Inches(4.6),Inches(2.4),Inches(2.4),Inches(2.3)]
xs=[x0]; 
for w in cwid[:-1]: xs.append(xs[-1]+w)
y=Inches(2.05)
box(s,x0,y,sum(cwid,Emu(0)),Inches(0.75),fill=BRUN)
text(s,xs[0]+Inches(0.2),y,cwid[0],Inches(0.75),"Idée",size=14,color=CREME,bold=True,anchor=MSO_ANCHOR.MIDDLE)
for j,c in enumerate(crit): text(s,xs[j+1]+Inches(0.1),y,cwid[j+1]-Inches(0.2),Inches(0.75),c,size=12,color=CREME,bold=True,anchor=MSO_ANCHOR.MIDDLE,align=PP_ALIGN.CENTER,spacing=1.0)
for i,(idea,v) in enumerate(ideas):
    yy=y+Inches(0.75)+i*Inches(0.48); keep=sum(v)>=2
    box(s,x0,yy,sum(cwid,Emu(0)),Inches(0.45),fill=BLANC if keep else KRAFT)
    text(s,xs[0]+Inches(0.2),yy,cwid[0],Inches(0.45),idea,size=15,bold=keep,color=BRUN if keep else GRIS,anchor=MSO_ANCHOR.MIDDLE)
    for j,ok in enumerate(v): text(s,xs[j+1],yy,cwid[j+1],Inches(0.45),"✓" if ok else "–",size=18,color=FEU if ok else GRIS,bold=True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(0.8),Inches(6.3),Inches(11.7),Inches(0.4),"On garde les idées qui remplissent au moins 2 critères (en gras) : elles forment le concept.",size=14,color=GRIS)

# 11 Concept
s=new(dark=True)
text(s,Inches(0.8),Inches(0.55),Inches(9),Inches(0.4),"LE CONCEPT RETENU",size=13,color=FEU,bold=True)
text(s,Inches(0.8),Inches(0.9),Inches(7),Inches(1.2),"Le P'tit Feu",size=54,color=CREME,font=TITRE)
text(s,Inches(0.8),Inches(1.95),Inches(7.5),Inches(0.5),"La formule midi des lycéens et étudiants",size=20,color=KRAFT)
r=box(s,Inches(9.3),Inches(0.7),Inches(3.2),Inches(1.7),fill=FEU,shape=MSO_SHAPE.ROUNDED_RECTANGLE); r.adjustments[0]=0.1
text(s,Inches(9.3),Inches(0.75),Inches(3.2),Inches(1.2),"6,50 €",size=50,color=BLANC,font=TITRE,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
text(s,Inches(9.3),Inches(1.85),Inches(3.2),Inches(0.4),"la formule complète",size=14,color=BLANC,align=PP_ALIGN.CENTER)
it=[("La galette au feu de bois","3 recettes au choix, dont une « recette du mois » votée sur Instagram"),("+ boisson + mini-dessert","Un classique revisité : mini-éclair, cookie maison"),("Précommande sur téléphone","QR code → choix → paiement, retrait sans file au comptoir P'tit Feu"),("Fidélité numérique","La 10e formule offerte")]
for i,(a,b) in enumerate(it):
    x=Inches(0.8)+(i%2)*Inches(6.0); y=Inches(3.0)+(i//2)*Inches(1.75)
    card(s,x,y,Inches(5.7),Inches(1.55),a,b,tsize=20,dsize=15,top=False,fill=RGBColor(0x4A,0x35,0x27),tc=FEU,dc=CREME)

# 12 Parcours client
s=new("Étape 4 · Prototyper","Le parcours de Léa, de la classe au comptoir")
st=[("10h","En cours","Elle voit l'affiche P'tit Feu à l'IUT et scanne le QR code."),("10h02","Elle commande","Galette du mois + boisson + cookie. Paiement en ligne : 6,50 €."),("12h05","Elle récupère","Comptoir P'tit Feu : pas de file, sa commande l'attend."),("12h15","Elle partage","Le sac kraft « cuit au feu de bois » finit en story."),("Plus tard","Elle revient","Sa carte de fidélité avance : la 10e est offerte.")]
cw=Inches(2.24); gap=Inches(0.125)
for i,(h_,t,d) in enumerate(st):
    x=Inches(0.8)+i*(cw+gap); y=Inches(2.3)
    box(s,x,y,cw,Inches(3.9),fill=BLANC)
    text(s,x+Inches(0.25),y+Inches(0.3),cw,Inches(0.6),h_,size=26,color=FEU,font=TITRE)
    text(s,x+Inches(0.25),y+Inches(1.05),cw-Inches(0.5),Inches(0.5),t,size=19,bold=True)
    text(s,x+Inches(0.25),y+Inches(1.6),cw-Inches(0.5),Inches(2.2),d,size=15,color=GRIS)
    if i<4:
        a=s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,x+cw-Inches(0.05),y+Inches(1.8),Inches(0.25),Inches(0.3)); a.fill.solid(); a.fill.fore_color.rgb=FEU; a.line.fill.background()

# 13-15 Prototypes
for k,(t,lab,leg) in enumerate([("L'affiche et les visuels Instagram","Votre visuel produits / affiche","Rendre la boulangerie « instagrammable » : la galette et les desserts trompe-l'œil mis en scène, avec le QR code de précommande.\n\nRépond au frein : l'image « traditionnelle »."),
                            ("Le packaging kraft","Votre visuel du packaging","Logo Petitjean + « cuit au feu de bois ». Le sac circule à l'IUT et en story : c'est une publicité gratuite.\n\nRépond au frein : la visibilité face aux chaînes."),
                            ("L'écran de précommande","Votre maquette d'écran","Sur le téléphone, après le QR code : choisir sa galette, sa boisson, son dessert, payer, choisir son heure de retrait.\n\nRépond au frein : le temps et la file d'attente.")]):
    s=new(f"Étape 4 · Prototype {k+1}/3",t)
    placeholder(s,Inches(0.8),Inches(2.05),Inches(7.4),Inches(4.5),lab)
    text(s,Inches(8.7),Inches(2.15),Inches(3.8),Inches(0.4),"CE QU'IL MONTRE",size=14,color=FEU,bold=True)
    text(s,Inches(8.7),Inches(2.6),Inches(3.8),Inches(3.9),leg.split("\n"),size=17,spacing=1.05)

# 16 Test
s=new("Étape 5 · Tester","Tester : les objections de L. Petitjean")
ob=[("« C'est rentable, à 6,50 € ? »","Prix proche du panier moyen (6,33 €). Galettes cuites dans un four déjà chaud, mini-dessert en petit format, 3 recettes seulement pour limiter les pertes."),
    ("« Je n'ai pas le temps de gérer une appli. »","On a abandonné la borne et l'appli : un simple formulaire en ligne + QR code, sans développement."),
    ("« Ça ne va pas abîmer mon image ? »","Le four à bois et les classiques revisités sont au centre : sur le sac, l'affiche et dans la recette du mois."),
    ("« Comment les jeunes vont le savoir ? »","Affiches avec QR code au lycée et à l'IUT, compte Instagram, vote de la recette du mois.")]
for i,(a,b) in enumerate(ob):
    x=Inches(0.8)+(i%2)*Inches(5.95); y=Inches(2.05)+(i//2)*Inches(2.2)
    box(s,x,y,Inches(5.75),Inches(2.05),fill=BLANC); box(s,x,y,Inches(0.08),Inches(2.05),fill=FEU)
    text(s,x+Inches(0.3),y+Inches(0.2),Inches(5.25),Inches(0.5),a,size=16,bold=True,font=TITRE)
    text(s,x+Inches(0.3),y+Inches(0.95),Inches(5.25),Inches(1.05),b,size=14,color=GRIS)

# 17 Bénéfices
s=new("Les bénéfices","Gagnant pour Léa et pour la boulangerie")
for col,(t,items) in enumerate([("Pour Léa",["Un vrai repas à 6,50 €, sous son budget de 7 €","Zéro file : elle commande en cours, récupère en 2 minutes","Un produit original à partager","Une récompense : la 10e formule offerte"]),
                                 ("Pour la boulangerie Petitjean",["Un nouveau public le midi, aujourd'hui parti au fast-food","Le four à bois mieux utilisé, au cœur de l'offre","De la visibilité gratuite sur Instagram","Une différence nette face aux 12 autres boulangeries"])]):
    x=Inches(0.8)+col*Inches(6.0)
    box(s,x,Inches(2.05),Inches(5.7),Inches(4.45),fill=BRUN if col else BLANC)
    text(s,x+Inches(0.35),Inches(2.3),Inches(5),Inches(0.5),t,size=22,bold=True,color=FEU)
    text(s,x+Inches(0.35),Inches(3.0),Inches(5.0),Inches(3.4),items,size=17,color=CREME if col else BRUN,spacing=1.05)

# 18 Conclusion
s=new(dark=True)
text(s,Inches(0.8),Inches(1.2),Inches(9),Inches(0.4),"EN RÉSUMÉ",size=13,color=FEU,bold=True)
text(s,Inches(0.8),Inches(1.7),Inches(11.5),Inches(2.2),"Le prix, le temps et l'image : on lève les trois freins de Léa avec ce que Ludovic Petitjean a de plus fort, son four à bois.",size=32,color=CREME,font=TITRE,spacing=1.2)
for i,(a,b) in enumerate([("6,50 €","sous le budget de Léa"),("0 file","précommande et retrait au comptoir"),("1 histoire","le feu de bois, à partager en story")]):
    x=Inches(0.8)+i*Inches(4.0)
    text(s,x,Inches(4.4),Inches(3.7),Inches(0.8),a,size=36,color=FEU,font=TITRE)
    text(s,x,Inches(5.2),Inches(3.7),Inches(0.6),b,size=16,color=KRAFT)
text(s,Inches(0.8),Inches(6.4),Inches(11.5),Inches(0.5),"Merci de votre attention. Avez-vous des questions ?",size=18,color=CREME,bold=True)

for sl,t in zip(p.slides,oral): notes(sl,t)
p.save(sys.argv[2]); print(len(p.slides._sldIdLst))
