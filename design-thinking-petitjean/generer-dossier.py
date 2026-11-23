from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
BRUN=RGBColor(0x3A,0x27,0x1B); FEU=RGBColor(0xD4,0x5D,0x24); GRIS=RGBColor(0x7A,0x6A,0x5C)
d=Document()
for s in d.sections: s.left_margin=s.right_margin=Cm(2.3); s.top_margin=s.bottom_margin=Cm(2)
st=d.styles['Normal']; st.font.name='Calibri'; st.font.size=Pt(11.5); st.paragraph_format.space_after=Pt(6); st.paragraph_format.line_spacing=1.15
for n,sz,c in [('Heading 1',18,BRUN),('Heading 2',13.5,FEU),('Title',30,BRUN)]:
    h=d.styles[n]; h.font.name='Georgia'; h.font.size=Pt(sz); h.font.color.rgb=c; h.font.bold=(n!='Title')
    h.element.rPr.rFonts.set(qn('w:asciiTheme'),'') if False else None
def shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),hexc); tcPr.append(sh)
def P(t="",bold_lead=None,italic=False,al=None,size=None,color=None):
    p=d.add_paragraph()
    if bold_lead: r=p.add_run(bold_lead); r.bold=True
    r=p.add_run(t); r.italic=italic
    if size: r.font.size=Pt(size)
    if color: r.font.color.rgb=color
    if al: p.alignment=al
    return p
def B(t,lead=None):
    p=d.add_paragraph(style='List Bullet')
    if lead: r=p.add_run(lead); r.bold=True
    p.add_run(t); return p
def table(rows,widths=None,head=True):
    t=d.add_table(rows=len(rows),cols=len(rows[0])); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.autofit=False
    if widths:
        for j,w in enumerate(widths): t.columns[j].width=Cm(w)
    for i,row in enumerate(rows):
        for j,v in enumerate(row):
            c=t.cell(i,j); c.text=""
            par=c.paragraphs[0]; r=par.add_run(v); r.font.size=Pt(10.5)
            if (head and i==0) or v in ('Ce qu\'ils FONT','Ce qu\'ils RESSENTENT'): r.bold=True; r.font.color.rgb=RGBColor(0xFF,0xFF,0xFF); shade(c,'3A271B')
            if widths: c.width=Cm(widths[j])
    d.add_paragraph(); return t
def cadre(t):
    tb=d.add_table(rows=1,cols=1); tb.style='Table Grid'; c=tb.cell(0,0); shade(c,'F3E9DA')
    c.text=""; r=c.paragraphs[0].add_run(t); r.italic=True; r.font.color.rgb=GRIS; c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(5): c.add_paragraph()
    d.add_paragraph()
def img(name,w=14,leg=None):
    d.add_picture('visuels/'+name,width=Cm(w)); d.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
    if leg: P(leg,italic=True,al=WD_ALIGN_PARAGRAPH.CENTER,size=10,color=GRIS)
def pb(): d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

# --- Page de garde
for _ in range(4): d.add_paragraph()
P("R3.MDEE.16  ·  BUT Techniques de commercialisation  ·  Travail de groupe",al=WD_ALIGN_PARAGRAPH.CENTER,size=11,color=GRIS)
P("Projet créatif — Design Thinking",al=WD_ALIGN_PARAGRAPH.CENTER,size=16,color=FEU)
t=d.add_paragraph("Le P'tit Feu",style='Title'); t.alignment=WD_ALIGN_PARAGRAPH.CENTER
P("Une formule midi au feu de bois pour ramener les étudiants du Creusot à la Boulangerie-Pâtisserie Petitjean",al=WD_ALIGN_PARAGRAPH.CENTER,size=13,italic=True)
for _ in range(6): d.add_paragraph()
P("Dossier de groupe",al=WD_ALIGN_PARAGRAPH.CENTER,size=12,color=BRUN).runs[0].bold=True
P("[Prénom NOM]  ·  [Prénom NOM]  ·  [Prénom NOM]",al=WD_ALIGN_PARAGRAPH.CENTER)
P("Enseignant(e) : [Nom]  ·  Année universitaire 2026-2027",al=WD_ALIGN_PARAGRAPH.CENTER,color=GRIS)
pb()
d.add_heading("Sommaire",1)
for x in ["Introduction","1. Le contexte : la Boulangerie-Pâtisserie Petitjean","2. Phase 1 — Comprendre la cible","3. Phase 2 — Définir le problème","4. Phase 3 — Imaginer des solutions","5. Phase 4 — Le prototype","6. Phase 5 — Tester et préparer le pitch","Conclusion","Sources","Annexes"]:
    P(x)
pb()

# --- Introduction
d.add_heading("Introduction",1)
P("Quand on nous a présenté le cas de la Boulangerie-Pâtisserie Petitjean, une chose nous a tout de suite parlé : nous faisons nous-mêmes partie de la cible. Comme beaucoup d'étudiants, nous passons devant des boulangeries tous les jours, nous trouvons que c'est bon, et pourtant, le midi, nous finissons souvent au fast-food ou au rayon sandwichs du supermarché.")
P("Ludovic Petitjean tient sa boulangerie au 105 rue Édith Cavell, au Creusot, depuis plusieurs années. Il est connu pour son pain cuit au four à bois et pour sa manière de revisiter les classiques. Mais comme la plupart des artisans, il doit faire face à un marché qui bouge : le snacking prend de plus en plus de place, les clients commandent de plus en plus en ligne, et les jeunes sont sans arrêt sollicités par les chaînes de restauration rapide.")
P("Sa demande était claire : imaginer un produit ou un service nouveau qui donne envie aux jeunes de venir chez lui, sans que la boulangerie perde ce qui fait sa force, son savoir-faire d'artisan. Pour y répondre, nous avons suivi les cinq étapes du Design Thinking vues en cours. Plutôt que de partir directement d'une idée, nous avons d'abord cherché à comprendre pourquoi les étudiants ne viennent pas, avant de chercher comment les faire venir. Ce dossier retrace ce chemin, étape par étape, jusqu'à la solution que nous avons retenue : « Le P'tit Feu ».")

# --- Contexte
d.add_heading("1. Le contexte : la Boulangerie-Pâtisserie Petitjean",1)
d.add_heading("1.1 L'entreprise",2)
P("La boulangerie de Ludovic Petitjean a un atout que peu de ses concurrents peuvent mettre en avant : un four à bois. Le pain et les galettes y sont cuits au feu de bois, et c'est d'ailleurs ce que les clients citent le plus souvent dans leurs avis. Sa note moyenne de 4,9 sur 5 sur les plateformes locales montre qu'il a une clientèle fidèle et satisfaite, et beaucoup le considèrent comme une institution du quartier.")
P("Son positionnement est celui d'un artisan qui reprend les grands classiques de la boulangerie-pâtisserie en y ajoutant sa touche personnelle. C'est un point important pour nous : la solution que nous allions proposer ne devait pas aller contre cette image, mais au contraire s'appuyer dessus.")
P("Il faut quand même rappeler que la concurrence est forte. On compte 13 boulangeries-pâtisseries au Creusot, ce qui est beaucoup pour une ville de cette taille. Dans un marché de proximité comme celui-là, les clients ont le choix, et il faut leur donner une vraie raison de venir chez un artisan plutôt que chez un autre.")
d.add_heading("1.2 Le marché de la boulangerie",2)
P("À l'échelle de la France, le secteur se porte plutôt bien. On compte environ 43 777 boulangeries, pour un chiffre d'affaires d'environ 17,2 milliards d'euros en 2025. Mais derrière ces chiffres, le métier change vite. Aujourd'hui, 81 % des boulangeries proposent du snacking, de la restauration rapide ou un service traiteur. Le panier moyen a aussi augmenté, passant de 5,92 € en 2024 à 6,33 € en 2025, et les artisans investissent de plus en plus dans le numérique, comme la commande en ligne ou les caisses connectées.")
P("Ce qui nous a le plus intéressés, ce sont les jeunes. Les moins de 35 ans s'intéressent de plus en plus aux boulangeries locales, pour leur côté authentique et convivial. Mais cette même génération est aussi celle que les chaînes de restauration rapide visent en priorité. Les jeunes ont donc envie d'aller chez l'artisan, mais au moment de manger, ce n'est pas lui qu'ils choisissent. C'est ce décalage que nous avons voulu comprendre.")

# --- Phase 1
d.add_heading("2. Phase 1 — Comprendre la cible",1)
d.add_heading("2.1 Comment nous avons recueilli les avis",2)
P("Notre cible, ce sont les lycéens et les étudiants du Creusot. Comme nous sommes nous-mêmes étudiants, nous avons commencé par parler entre nous, assez librement, de nos habitudes : ce que nous achetons en boulangerie, à quel moment de la journée, combien nous dépensons le midi, et pourquoi nous allons parfois ailleurs.")
P("Nous nous sommes vite rendu compte que nos propres réponses ne suffisaient pas : nous risquions de prendre nos habitudes pour celles de tout le monde. Nous avons donc posé les mêmes questions à [X] étudiants de l'IUT du Creusot [à compléter : où et quand]. Les questions sont en annexe 2. Nous avons noté leurs réponses le plus fidèlement possible, avec leurs propres mots, pour pouvoir ensuite séparer ce que les gens disent de ce qu'ils pensent vraiment.")
d.add_heading("2.2 La carte d'empathie",2)
P("Nous avons rassemblé ce que nous avions entendu dans une carte d'empathie, qui classe les réponses en quatre parties : ce que la cible dit, ce qu'elle pense, ce qu'elle fait et ce qu'elle ressent.")
table([["Ce qu'ils DISENT","Ce qu'ils PENSENT"],
       ["« La boulangerie c'est bon, mais c'est cher pour un midi. »\n« J'ai pas le temps, j'ai 30 minutes de pause. »\n« Je prends un sandwich au Carrefour ou un tacos. »","Le pain artisanal, c'est mieux, mais c'est plutôt « pour les parents ».\nIls veulent que ce soit rentable pour eux et que ça les nourrisse bien.\nSi c'est beau, ils le postent en story."],
       ["Ce qu'ils FONT","Ce qu'ils RESSENTENT"],
       ["Ils mangent vite entre midi et deux, souvent en groupe.\nIls commandent souvent depuis leur téléphone.\nIls vont à la boulangerie surtout pour une viennoiserie ou un goûter.","L'envie de manger « vrai » et bon sans se ruiner.\nDe la frustration dans les files d'attente.\nLe plaisir de partager une bonne pâtisserie avec leurs amis."]],[8,8])
P("En relisant cette carte, un point nous a marqués. Quand les étudiants parlent de la boulangerie, ils ne critiquent jamais la qualité. Au contraire, ils pensent que le pain artisanal est meilleur. Ce qu'ils reprochent, c'est le prix, le temps, et une image un peu « vieille », comme si la boulangerie était un endroit pour leurs parents et pas pour eux.")
P("Autrement dit, le problème n'est pas le produit, c'est la façon dont il leur est proposé, surtout le midi. Cette idée a guidé toute la suite de notre travail.")
d.add_heading("2.3 Les freins",2)
P("Le premier frein est le prix. Un étudiant dépense en général entre 5 et 7 € pour son repas du midi, et beaucoup ont l'impression qu'à la boulangerie, ils vont dépasser ce budget. Que ce soit vrai ou non, c'est cette impression qui compte, car c'est elle qui les fait passer leur chemin.")
P("Le deuxième frein est le temps. Avec une pause d'une demi-heure, faire la queue au moment du rush n'est pas possible. Les fast-foods et les supermarchés ont un avantage ici : on sait qu'on sera servi vite, ou on peut commander à l'avance.")
P("Le troisième frein est l'image. La boulangerie est vue comme un lieu traditionnel, et elle est peu présente sur les réseaux sociaux, là où les jeunes passent beaucoup de temps. À l'inverse, les chaînes de restauration rapide, qui représentent le quatrième frein, sont partout : en publicité, sur Instagram, sur les applications de livraison. Leur offre paraît plus moderne et plus « fun ».")
d.add_heading("2.4 Les motivations",2)
P("Heureusement, nous avons aussi trouvé de vraies raisons d'espérer. La première, c'est le goût : les étudiants aiment le fait maison et l'authenticité, et le four à bois est un argument qui leur parle. La deuxième, c'est l'envie d'une formule complète, qui nourrit vraiment et qui reste à un prix raisonnable.")
P("Ils sont aussi très sensibles à la rapidité, et l'idée de commander à l'avance pour récupérer sa commande sans attendre leur plaît. Enfin, ils aiment les produits originaux, ceux qu'on a envie de prendre en photo et de montrer, et ils sont attirés par les programmes de fidélité.")

# --- Phase 2
d.add_heading("3. Phase 2 — Définir le problème",1)
d.add_heading("3.1 Le persona : Léa",2)
P("Pour ne pas perdre de vue à qui nous nous adressions, nous avons réuni tout ce que nous avions appris dans un seul personnage, représentatif de notre cible.")
table([["Champ","Notre persona"],["Prénom, âge","Léa, 19 ans"],["Situation","Étudiante en 1re année de BUT à l'IUT du Creusot, vit en appartement CROUS, petit budget"],["Habitudes de consommation","Mange dehors 3 midis par semaine : un sandwich de supermarché ou un fast-food avec ses amis. Va à la boulangerie le week-end pour des viennoiseries."],["Frein principal","Le rapport prix / rapidité : elle pense que la boulangerie est trop chère et trop lente pour le midi."],["Ce qui la ferait revenir","Une formule midi rapide à moins de 7 €, commandable sur son téléphone, avec un produit original."]],[4.5,11.5])
P("Ce qui est intéressant avec Léa, c'est qu'elle n'a rien contre la boulangerie. Elle y va déjà le week-end, pour se faire plaisir avec des viennoiseries. Le problème, c'est qu'elle n'y pense pas au moment où elle mange le plus souvent dehors, c'est-à-dire le midi, en semaine. C'est donc ce moment précis que nous avons décidé de cibler.")
P("Nous avons fait en sorte que Léa soit crédible : son budget, ses horaires et ses habitudes ressemblent à ceux des étudiants que nous avons interrogés. Nous voulions éviter le piège d'un persona caricatural, qui n'aurait servi qu'à justifier une idée déjà en tête.")
d.add_heading("3.2 Notre question « Comment pourrait-on… ? »",2)
P("À partir de Léa, nous avons formulé notre problème sous la forme d'une question :")
p=P(""); r=p.add_run("« Comment pourrait-on proposer aux étudiants du Creusot une offre de repas rapide, abordable et originale qui valorise le four à bois de Ludovic Petitjean, pour qu'ils choisissent la boulangerie plutôt que la restauration rapide ? »"); r.italic=True; r.bold=True; r.font.color.rgb=BRUN
P("Nous avons mis du temps à trouver cette formulation. Nos premières versions étaient soit trop vagues, du genre « comment attirer les jeunes ? », soit elles contenaient déjà la solution. Celle que nous avons gardée a l'avantage de relier deux choses : d'un côté, ce dont Léa a besoin, un repas rapide, pas cher et original, et de l'autre, ce qui fait l'identité de l'artisan, son four à bois et son savoir-faire. La question dit ce que nous cherchons, mais elle laisse encore ouverte la façon d'y arriver.")

# --- Phase 3
d.add_heading("4. Phase 3 — Imaginer des solutions",1)
d.add_heading("4.1 Pourquoi le brainstorming",2)
P("Pour trouver des idées, nous devions choisir un outil de créativité parmi ceux vus en cours. Nous avons choisi le brainstorming direct, qui consiste à produire le plus d'idées possible en groupe, sans les juger, puis à les trier seulement à la fin.")
P("Ce choix vient directement de notre problème. Notre question touche à plusieurs sujets à la fois : ce qu'on vend, le prix, la rapidité et l'image de la boulangerie. Avant de choisir une piste, il fallait donc en ouvrir le plus possible, et le brainstorming est l'outil le plus adapté pour ça.")
P("Nous avons aussi réfléchi aux autres outils. Le SCAMPER nous a semblé moins adapté, car il sert à faire évoluer un produit qui existe déjà à partir de sept verbes d'action. Or nous ne voulions pas modifier un produit de la boulangerie, mais créer une offre du midi qui n'existait pas encore. La méthode 635 demande six participants, alors que notre groupe n'en compte que trois. Quant au mind mapping, il est très utile pour organiser un sujet, mais il pousse moins à produire beaucoup d'idées en peu de temps.")
d.add_heading("4.2 Le déroulement de la séance",2)
P("Nous nous sommes donné 15 minutes, chronomètre en main. Pendant ce temps, la règle était simple : aucune critique. Chaque idée était notée au fur et à mesure, même celles qui nous paraissaient un peu folles, et nous essayions de rebondir sur les idées des autres plutôt que de les discuter. Ce n'est pas toujours facile de ne pas commenter, mais c'est justement ce qui permet de faire sortir des idées auxquelles on n'aurait pas pensé seul.")
P("Une fois le temps écoulé, nous avons regroupé les idées par thèmes, puis nous les avons triées avec des critères précis. La liste complète des idées notées pendant la séance se trouve en annexe 1.")
d.add_heading("4.3 Les idées générées",2)
P("Nos idées se sont réparties en cinq grands thèmes. Côté offre, nous avons imaginé une formule midi à prix fixe, avec une galette cuite au feu de bois, une boisson et un dessert. Sont aussi venues des galettes garnies façon street-food, et des mini-pizzas au feu de bois à partager entre amis.")
P("Côté produits, nous avons pensé à des viennoiseries et des gâteaux en trompe-l'œil, en forme de pomme, de citron, de basket ou même de sac à main, qui ont tout pour être pris en photo. Une autre idée était de faire voter les clients sur Instagram pour choisir une recette du mois.")
P("Pour la rapidité, il y avait une borne de commande comme dans les fast-foods, une précommande par QR code avec un retrait sans faire la queue, et un créneau « Spécial étudiants » de 11h45 à 13h30. Pour la fidélité et l'image, nous avons proposé une carte de fidélité numérique avec la dixième formule offerte, un emballage en kraft avec la mention « cuit au feu de bois », ainsi qu'un compte Instagram et des affiches au lycée et à l'IUT. Enfin, une idée concernait le lieu : aménager un coin goûter et révisions le mercredi après-midi.")
d.add_heading("4.4 Comment nous avons trié",2)
P("Nous ne voulions pas choisir au hasard ni simplement garder les idées qui nous plaisaient le plus. Nous avons donc fixé trois critères, tirés de notre question et de la demande de Ludovic Petitjean. Est-ce que l'idée lève le frein principal de Léa, le prix ou la rapidité ? Est-ce qu'elle met en valeur le four à bois et le savoir-faire de l'artisan ? Et est-ce qu'un artisan seul peut la mettre en place sans investissement trop lourd ?")
table([["Idée","Frein de Léa","Four à bois","Faisable seul","Retenue ?"],
       ["Formule midi à prix fixe","✓","✓","✓","Oui"],["Galette garnie au feu de bois","✓","✓","✓","Oui"],["Borne de commande en boutique","✓","–","✓","Oui"],["Précommande en ligne (QR code)","✓","–","✓","Oui"],
       ["Carte de fidélité numérique","✓","–","✓","Oui"],["Recette du mois votée sur Instagram","–","✓","✓","Oui"],["Packaging kraft « cuit au feu de bois »","–","✓","✓","Oui"],
       ["Gâteaux trompe-l'œil","–","–","✓","Non"],["Coin goûter / révisions","–","–","–","Non"]],[6.5,2.3,2.3,2.3,2.3])
P("Nous avons gardé les idées qui remplissaient au moins deux critères sur trois. Certains choix ont été faciles, d'autres moins. La borne de commande, par exemple, nous a fait hésiter. Elle répond très bien au problème du temps, mais elle n'a pas de lien avec le savoir-faire de l'artisan, et elle représente un coût. Nous l'avons finalement gardée, car c'est un investissement que beaucoup d'artisans font déjà : en 2025, ils ont investi en moyenne 17 623 €, en grande partie dans la digitalisation. Les gâteaux trompe-l'œil étaient une belle idée visuelle, mais ils ne répondaient ni au prix ni au temps, qui sont les vrais freins de Léa.")
d.add_heading("4.5 Le concept retenu : « Le P'tit Feu »",2)
P("En rassemblant les idées gardées, nous avons construit un seul concept cohérent, que nous avons appelé « Le P'tit Feu ». Le nom fait directement référence au four à bois, tout en restant simple et sympathique pour des jeunes.")
P("Le P'tit Feu est une formule midi à 6,50 €, pensée pour les lycéens et les étudiants. Elle comprend une galette cuite au four à bois, avec trois garnitures au choix, dont une recette du mois choisie par les clients grâce à un vote sur Instagram. La galette est accompagnée d'une boisson et d'un mini-dessert « classique revisité », comme un mini-éclair ou un cookie maison, ce qui rappelle le style de Ludovic Petitjean.")
P("Pour commander, il y a deux possibilités. En boutique, le client passe par une borne de commande, comme dans les fast-foods : il choisit sa formule, paie, et n'a plus qu'à la récupérer au comptoir « P'tit Feu ». Il peut aussi commander à l'avance en ligne, grâce à un QR code présent sur les affiches et sur Instagram, et passer la chercher sans attendre. Enfin, une carte de fidélité numérique offre la dixième formule.")
P("Le prix de 6,50 € n'a pas été choisi au hasard. Il reste sous le budget maximum de Léa, qui est de 7 €, et il est très proche du panier moyen en boulangerie, qui était de 6,33 € en 2025. Il se situe aussi au niveau d'un menu étudiant dans la restauration rapide, ce qui permet à la boulangerie de se battre à armes égales.")

# --- Phase 4
d.add_heading("5. Phase 4 — Le prototype",1)
P("Une idée, aussi bonne soit-elle, reste difficile à imaginer tant qu'on ne la voit pas. Nous avons donc réalisé un prototype de moyenne fidélité, en trois éléments : une présence sur Instagram, un packaging kraft et une borne de commande. Nous avons choisi chaque élément pour qu'il réponde à un frein précis de Léa. Pour les visuels, nous nous sommes aidés d'un outil d'intelligence artificielle, mais seulement une fois nos idées trouvées en groupe : l'IA nous a servi à mettre en image ce que nous avions imaginé, pas à l'imaginer à notre place.")
d.add_heading("5.1 Le parcours de Léa",2)
P("Avant de dessiner quoi que ce soit, nous avons imaginé une journée de Léa avec Le P'tit Feu. Il est 10 heures, elle est en cours, et elle voit l'affiche du P'tit Feu à l'IUT. Elle scanne le QR code avec son téléphone. Deux minutes plus tard, elle a choisi la galette du mois, une boisson et un cookie, et elle a payé ses 6,50 € en ligne.")
P("À midi cinq, elle passe à la boulangerie. Il y a du monde, mais elle va directement au comptoir P'tit Feu, où sa commande l'attend. Un quart d'heure plus tard, assise avec ses amis, elle prend en photo son sac kraft « cuit au feu de bois » et le poste en story. Sa carte de fidélité avance d'une case. Ses amis, qui n'avaient pas commandé à l'avance, passent par la borne en boutique : en quelques clics, ils ont eux aussi leur formule, sans attendre au comptoir.")
P("Ce petit scénario nous a beaucoup aidés, parce qu'il montre que chaque frein est levé à un moment précis : le prix au moment de payer, le temps au moment de récupérer la commande, et l'image au moment de partager.")
d.add_heading("5.2 Premier élément : les visuels et le support Instagram",2)
P("Le premier élément de notre prototype est la page Instagram de la boulangerie, « boulangerie.petitjean ». La présentation du compte annonce le pain cuit au feu de bois et la formule « Le P'tit Feu » à 6,50 €, avec un bouton pour commander en ligne. Les publications mettent en scène les viennoiseries, le café et le nouveau packaging, dans une ambiance chaleureuse qui donne envie de venir. L'objectif est de casser l'image « traditionnelle » de la boulangerie : si les produits sont beaux et donnent envie d'être photographiés, ce sont les étudiants eux-mêmes qui feront la publicité de la boulangerie auprès de leurs amis.")
img("instagram.jpg",16,"Le P'tit Feu sur Instagram : la page de la boulangerie et la publication du nouveau packaging")
d.add_heading("5.3 Deuxième élément : le packaging kraft",2)
P("Le deuxième élément est un emballage en kraft, sur lequel on retrouve le logo de la boulangerie et la mention « cuit au feu de bois ». Ce sac répond au manque de visibilité de la boulangerie face aux chaînes. Il va circuler dans la rue, dans les couloirs de l'IUT et sur les réseaux sociaux. Chaque client devient une petite publicité ambulante, qui met en avant l'argument le plus fort de l'artisan.")
img("packaging-kraft.jpg",15,"Le packaging kraft « cuit au feu de bois » : sac, boîte et gobelet à emporter")
d.add_heading("5.4 Troisième élément : la borne de commande",2)
P("Le troisième élément est une borne de commande, installée dans la boutique. Nous l'avons d'abord dessinée sous forme de croquis, en choisissant des matériaux qui rappellent l'univers de l'artisan : du bois de chêne, un socle en acier noir et un logo en forme de flamme en cuivre, pour faire le lien avec le four à bois. Elle comporte un écran tactile, un terminal de paiement sans contact et une sortie de ticket.")
img("borne-croquis.jpg",15,"Croquis de la borne : matériaux, dimensions et éléments")
P("Nous l'avons ensuite mise en situation dans la boulangerie. L'écran affiche le menu du P'tit Feu, et le client compose sa formule en quelques secondes, paie, puis récupère sa commande au comptoir. Le même menu est disponible en ligne, ce qui permet aux étudiants pressés de commander avant même d'arriver. Cet élément répond directement au problème du temps et de la file d'attente, et il reprend une habitude que la cible a déjà dans les fast-foods.")
img("borne-boutique.jpg",15,"La borne « Le P'tit Feu » en situation dans la boulangerie")

# --- Phase 5
d.add_heading("6. Phase 5 — Tester et préparer le pitch",1)
d.add_heading("6.1 Les objections de Ludovic Petitjean",2)
P("Pour tester notre concept, nous nous sommes mis à la place de Ludovic Petitjean. Nous nous sommes demandé ce qu'il pourrait nous répondre, en tant qu'artisan qui a déjà beaucoup de travail et qui tient à sa réputation. Cet exercice a été utile, car il nous a obligés à revoir certains points de notre prototype.")
P("Sa première question serait sûrement : « Est-ce que c'est rentable à 6,50 € ? ». Nous pensons que oui. Ce prix est proche du panier moyen en boulangerie, et les galettes cuisent dans un four qui est déjà chaud pour le pain, ce qui limite les coûts. Mais cette objection nous a quand même fait revoir l'offre : nous avons limité le choix à trois recettes et réduit le dessert à un petit format, pour éviter de gâcher des produits. Quant à la borne, c'est un vrai coût, mais il faut le comparer au temps gagné au moment du rush, quand un seul vendeur ne peut pas servir tout le monde. Les chiffres du secteur montrent d'ailleurs que les artisans investissent déjà beaucoup dans ce type d'équipement.")
P("Il pourrait aussi nous dire : « Je n'ai pas le temps de gérer une application. ». C'est l'objection qui a le plus fait évoluer notre projet. Au départ, nous avions pensé à créer une application propre à la boulangerie. Mais une application demande du temps, de l'argent et un suivi que Ludovic Petitjean ne peut pas assurer seul. Nous avons donc abandonné cette idée. À la place, la borne et la commande en ligne reposent sur des solutions déjà prêtes, comme celles qu'utilisent les restaurants, et elles affichent exactement le même menu. Ce menu est volontairement court, avec trois recettes, pour que tout reste simple à gérer.")
P("Une autre inquiétude pourrait être : « Est-ce que ça ne va pas abîmer mon image ? ». Nous comprenons cette crainte, car un artisan n'a pas envie de ressembler à un fast-food. Mais notre concept fait l'inverse : il met en avant le four à bois et les classiques revisités, c'est-à-dire exactement ce qui fait sa réputation. Pour que ce soit bien visible, la mention « cuit au feu de bois » apparaît sur l'emballage et sur Instagram, et la borne reprend les matériaux et la flamme qui rappellent le four.")
P("Enfin, il nous demanderait sans doute : « Comment les jeunes vont-ils savoir que ça existe ? ». Nous prévoyons des affiches au lycée et à l'IUT, un compte Instagram et le vote de la recette du mois, qui fait participer les clients. Cette question nous a donné une idée simple : mettre le QR code directement sur l'affiche, pour qu'un étudiant puisse commander au moment même où il la voit.")
P("[À compléter si vous l'avez fait : nous avons aussi montré notre prototype à [X] étudiants. Ils nous ont dit… Nous avons donc modifié…]",italic=True,color=GRIS)
d.add_heading("6.2 Ce que le projet apporte",2)
P("Pour Léa, et plus largement pour les étudiants, Le P'tit Feu permet de manger un vrai repas, fait maison, pour 6,50 €, sans dépasser son budget. Grâce à la borne et à la commande en ligne, elle ne perd plus de temps dans la file d'attente. Elle découvre un produit original, qu'elle a envie de partager, et sa fidélité est récompensée.")
P("Pour la boulangerie, le bénéfice est tout aussi important. Le P'tit Feu lui permet d'attirer, le midi, un public qui part aujourd'hui vers la restauration rapide. Il donne encore plus de place au four à bois, qui devient le cœur d'une nouvelle offre. Les photos partagées par les clients lui apportent une visibilité gratuite sur les réseaux sociaux. Et surtout, il lui donne une vraie différence par rapport aux douze autres boulangeries du Creusot.")
d.add_heading("6.3 La trame de notre pitch",2)
P("Notre pitch dure 10 minutes et suit la trame conseillée. Nous nous sommes réparti la parole pour que chacun de nous trois parle environ trois minutes. [Prénom 1] commence par le contexte, présente la cible et le problème à partir de la carte d'empathie. [Prénom 2] présente Léa, notre question « Comment pourrait-on… ? », l'outil de créativité et la façon dont nous avons trié les idées. [Prénom 3] présente ensuite le concept, le parcours de Léa, le prototype et les objections de l'artisan. Nous terminons tous les trois en présentant les bénéfices pour la cible et pour la boulangerie, avant de répondre aux questions.")

# --- Conclusion
d.add_heading("Conclusion",1)
P("Ce projet nous a appris à ne pas partir trop vite vers une solution. Au début, nous avions plein d'idées en tête, et si nous les avions suivies directement, nous serions sans doute passés à côté du vrai problème. En prenant le temps d'écouter la cible, nous avons compris que les étudiants ne boudent pas la boulangerie parce qu'ils ne l'aiment pas, mais parce qu'elle ne correspond pas à leur rythme et à leur budget le midi.")
P("Le P'tit Feu essaie de répondre à ces trois freins à la fois. Le prix, avec une formule à 6,50 €. Le temps, avec une borne en boutique et une commande en ligne, sans attente. L'image, avec des produits et un emballage qu'on a envie de montrer. Et tout cela en s'appuyant sur ce que Ludovic Petitjean a de plus précieux, son four à bois, sans lui demander de gros investissements.")
P("Notre projet a bien sûr ses limites. Nous n'avons pas pu tester la formule en conditions réelles, et nos chiffres sur la rentabilité restent des estimations. La suite logique serait de la lancer pendant quelques semaines, par exemple autour de l'IUT, et de suivre le nombre de formules vendues, de précommandes et d'abonnés sur Instagram. Cela permettrait de voir si Léa, et tous les étudiants qui lui ressemblent, prennent vraiment le chemin de la boulangerie le midi.")

# --- Sources
d.add_heading("Sources",1)
for s_ in ["MyBoulange, fiche et avis clients de la Boulangerie Pâtisserie L. Petitjean — myboulange.fr","Pages Jaunes, boulangeries-pâtisseries artisans au Creusot — pagesjaunes.fr","Boulangeries-Patisseries.fr, recensement des boulangeries du Creusot","Epsimas, Étude de marché de la boulangerie 2025-2026 — epsimas.com","Zepros Resto, « Chiffres d'affaires, rentabilité, fréquentation... en 2025 » — resto.zepros.fr","AffaireDidées, « SIRHA 2025 : La boulangerie française à l'aube d'une nouvelle ère »","Act Développement, « Les inquiétudes et perspectives économiques du secteur en 2025 »"]:
    B(s_)
d.add_heading("Annexes",1)
d.add_heading("Annexe 1 — Les idées notées pendant le brainstorming",2)
P("Voici la liste des idées produites pendant les 15 minutes de la séance, regroupées ensuite par thèmes.")
for th,ids in [("Offre",["Formule midi à prix fixe : galette au feu de bois + boisson + dessert","Galettes garnies façon street-food","Mini-pizzas au feu de bois à partager"]),("Produits",["Viennoiseries et gâteaux trompe-l'œil (pomme, citron, basket, sac à main)","Recette du mois votée sur Instagram"]),("Rapidité",["Borne ou écran de commande comme dans les fast-foods","Précommande par QR code et retrait sans file","Créneau « Spécial étudiants » de 11h45 à 13h30"]),("Fidélité et image",["Carte de fidélité numérique, 10e formule offerte","Packaging kraft « cuit au feu de bois »","Compte Instagram","Affiches au lycée et à l'IUT"]),("Lieu",["Coin goûter / révisions le mercredi après-midi"])]:
    p=P(""); r=p.add_run(th); r.bold=True; r.font.color.rgb=FEU
    for x in ids: B(x)
d.add_heading("Annexe 2 — Questions posées aux étudiants",2)
for q in ["Où mangez-vous le midi, et combien de fois par semaine ?","Combien dépensez-vous en moyenne pour un repas du midi ?","Qu'achetez-vous en boulangerie, et à quel moment ?","Qu'est-ce qui vous empêche d'aller à la boulangerie le midi ?","Qu'est-ce qui vous ferait venir plus souvent ?"]:
    d.add_paragraph(q,style='List Number')
d.save('Dossier-Le-Ptit-Feu.docx')
