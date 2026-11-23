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

# --- Sommaire
d.add_heading("Sommaire",1)
for x in ["Introduction","1. Le contexte : la Boulangerie-Pâtisserie Petitjean","2. Phase 1 — Comprendre la cible","3. Phase 2 — Définir le problème","4. Phase 3 — Imaginer des solutions","5. Phase 4 — Le prototype","6. Phase 5 — Tester et préparer le pitch","Conclusion","Sources","Annexes"]:
    P(x)
pb()

# --- Intro
d.add_heading("Introduction",1)
P("Ludovic Petitjean tient depuis plusieurs années sa boulangerie-pâtisserie au 105 rue Édith Cavell, au Creusot. Reconnu pour son pain cuit au four à bois et pour sa façon de revisiter les grands classiques, il fait face, comme beaucoup d'artisans, à un marché qui change : montée du snacking, commandes de plus en plus digitales et clientèle jeune très sollicitée par les chaînes de restauration rapide.")
P("Il nous a confié une mission : imaginer un produit ou un service inédit capable de séduire une clientèle jeune, sans trahir son savoir-faire artisanal. Pour y répondre, nous avons suivi les cinq étapes de la démarche Design Thinking : comprendre la cible, définir le problème, imaginer des solutions, prototyper, puis tester notre concept face aux objections de l'artisan. Ce dossier présente chacune de ces étapes et la solution à laquelle elles nous ont menés : « Le P'tit Feu ».")

# --- Contexte
d.add_heading("1. Le contexte : la Boulangerie-Pâtisserie Petitjean",1)
d.add_heading("1.1 L'entreprise",2)
table([["Élément","Donnée"],["Adresse","105 rue Édith Cavell, 71200 Le Creusot"],["Spécificité","Four à bois : pain et galettes cuits au feu de bois"],["Réputation","Note moyenne de 4,9/5 sur les plateformes locales"],["Positionnement","Artisan qui revisite les classiques avec une touche personnelle"],["Concurrence locale","13 boulangeries-pâtisseries recensées au Creusot"]],[4.5,11.5])
d.add_heading("1.2 Le marché",2)
P("Le secteur est dynamique mais se transforme : en France, on compte environ 43 777 boulangeries pour un chiffre d'affaires d'environ 17,2 milliards d'euros en 2025. Surtout, 81 % des boulangeries proposent désormais du snacking, de la restauration rapide ou un service traiteur, et le panier moyen est passé de 5,92 € en 2024 à 6,33 € en 2025. Les artisans investissent aussi davantage, notamment dans la digitalisation (commandes en ligne, caisses connectées).")
P("Les moins de 35 ans montrent un intérêt croissant pour les boulangeries locales, pour leur authenticité et leur convivialité, mais cette génération reste très sollicitée par la restauration rapide. C'est tout l'enjeu de notre projet : faire revenir les jeunes chez un artisan, au moment où ils lui échappent le plus, le midi.")

# --- Phase 1
d.add_heading("2. Phase 1 — Comprendre la cible",1)
d.add_heading("2.1 Notre méthode de recueil",2)
P("Avant d'imaginer une solution, nous avons cherché à comprendre notre cible, les lycéens et étudiants du Creusot. Nous avons d'abord échangé au sein du groupe sur nos propres habitudes en boulangerie : ce que nous achetons, ce qui nous freine et ce qui nous ferait revenir plus souvent. Nous avons ensuite interrogé [X] étudiants de l'IUT du Creusot, en leur posant les mêmes questions [à compléter : nombre de personnes, lieu, date]. Nous avons noté leurs phrases telles quelles, afin de bien distinguer ce que la cible dit de ce qu'elle pense ou ressent réellement.")
d.add_heading("2.2 La carte d'empathie",2)
table([["Ce qu'ils DISENT","Ce qu'ils PENSENT"],
       ["« La boulangerie c'est bon, mais c'est cher pour un midi. »\n« J'ai pas le temps, j'ai 30 minutes de pause. »\n« Je prends un sandwich au Carrefour ou un tacos. »","Le pain artisanal, c'est mieux, mais c'est plutôt « pour les parents ».\nIls veulent que ce soit rentable pour eux et que ça les nourrisse bien.\nSi c'est beau, ils le postent en story."],
       ["Ce qu'ils FONT","Ce qu'ils RESSENTENT"],
       ["Ils mangent vite entre midi et deux, souvent en groupe.\nIls commandent souvent depuis leur téléphone.\nIls vont à la boulangerie surtout pour une viennoiserie ou un goûter.","L'envie de manger « vrai » et bon sans se ruiner.\nDe la frustration dans les files d'attente.\nLe plaisir de partager une bonne pâtisserie avec leurs amis."]],[8,8])
P("Cette carte fait apparaître un décalage intéressant : ce que les étudiants disent porte surtout sur le prix et le temps, alors que ce qu'ils pensent montre qu'ils reconnaissent la qualité de l'artisanal. Le problème n'est donc pas le produit, mais la façon dont il leur est proposé le midi.")
d.add_heading("2.3 Les freins identifiés",2)
B(" un repas du midi coûte environ 5 à 7 € pour un étudiant, et la boulangerie est perçue comme chère.","Le prix :")
B(" la pause est courte et la file d'attente est longue au moment du rush.","Le manque de temps :")
B(" la boulangerie est vue comme « traditionnelle » et est peu présente sur les réseaux sociaux.","L'image :")
B(" l'offre snacking des chaînes de restauration rapide est très visible et jugée plus « fun ».","La concurrence :")
d.add_heading("2.4 Les motivations identifiées",2)
B(" le fait maison et l'authenticité ; le four à bois est un vrai argument.","Le goût :")
B(" une formule complète, nourrissante et à un prix convenable.","La formule :")
B(" pouvoir commander à l'avance et récupérer sans attendre, par exemple sur internet.","La rapidité :")
B(" un produit original et « instagrammable », ainsi qu'un programme de fidélité.","L'originalité :")

# --- Phase 2
d.add_heading("3. Phase 2 — Définir le problème",1)
d.add_heading("3.1 Le persona",2)
P("À partir de la carte d'empathie, nous avons construit un persona qui représente notre cible.")
table([["Champ","Notre persona"],["Prénom, âge","Léa, 19 ans"],["Situation","Étudiante en 1re année de BUT à l'IUT du Creusot, vit en appartement CROUS, petit budget"],["Habitudes de consommation","Mange dehors 3 midis par semaine : un sandwich de supermarché ou un fast-food avec ses amis. Va à la boulangerie le week-end pour des viennoiseries."],["Frein principal","Le rapport prix / rapidité : elle pense que la boulangerie est trop chère et trop lente pour le midi."],["Ce qui la ferait revenir","Une formule midi rapide à moins de 7 €, commandable sur son téléphone, avec un produit original."]],[4.5,11.5])
P("Léa n'est pas opposée à la boulangerie : elle y va déjà le week-end. Elle n'y va simplement pas au moment où elle mange dehors le plus souvent, le midi. C'est ce moment que nous avons choisi de cibler.")
d.add_heading("3.2 Notre question « Comment pourrait-on… ? »",2)
p=P(""); r=p.add_run("« Comment pourrait-on proposer aux étudiants du Creusot une offre de repas rapide, abordable et originale qui valorise le four à bois de Ludovic Petitjean, pour qu'ils choisissent la boulangerie plutôt que la restauration rapide ? »"); r.italic=True; r.bold=True; r.font.color.rgb=BRUN
P("Cette question relie deux éléments : le besoin de la cible (rapide, abordable, original), issu directement des freins et du persona, et l'identité de l'artisan (le four à bois et son savoir-faire), qui doit rester au cœur de la solution. Elle est centrée sur l'utilisateur et ne contient pas encore la solution.")

# --- Phase 3
d.add_heading("4. Phase 3 — Imaginer des solutions",1)
d.add_heading("4.1 L'outil choisi : le brainstorming",2)
P("Parmi les outils de créativité vus en cours, nous avons choisi le brainstorming (direct), qui consiste à produire un maximum d'idées en groupe, sans jugement, avant de les trier.")
P("Ce choix se justifie par notre problématique, qui est large : elle touche à la fois l'offre, le prix, la rapidité et l'image de la boulangerie auprès des jeunes. Il nous fallait d'abord ouvrir un grand nombre de pistes, dans toutes ces directions, avant de choisir. Nous avons écarté les autres outils pour les raisons suivantes :")
B(" cet outil fait évoluer un produit existant à partir de 7 verbes. Or nous ne voulions pas modifier un produit de la boulangerie, mais créer une offre midi qui n'existe pas encore.","SCAMPER :")
B(" elle demande 6 participants, et notre groupe n'en compte que 3.","La méthode 635 :")
B(" il aide surtout à structurer un problème, mais produit moins d'idées en peu de temps.","Le mind mapping :")
d.add_heading("4.2 Le déroulement de la séance",2)
P("Nous avons respecté les règles du brainstorming vues en cours :")
B(" 15 minutes chronométrées de production d'idées en groupe ;","Durée :")
B(" aucune critique pendant la production, toutes les idées sont notées ;","Pas de jugement :")
B(" chaque idée est écrite sur papier, et on rebondit sur les idées des autres ;","Quantité :")
B(" ce n'est qu'après la séance que nous avons regroupé les idées par thèmes, puis que nous les avons triées.","Tri après :")
P("La trace de cette séance (photo de nos feuilles) se trouve en annexe 1.",italic=True,color=GRIS)
d.add_heading("4.3 Les idées générées",2)
table([["Thème","Idées"],["Offre","Formule midi à prix fixe (galette au feu de bois + boisson + dessert) ; galettes garnies façon street-food ; mini-pizzas au feu de bois à partager"],["Produits","Viennoiseries et gâteaux trompe-l'œil (pomme, citron, basket, sac à main) très « instagrammables » ; recette du mois votée sur Instagram"],["Rapidité","Borne ou écran de commande comme dans les fast-foods ; précommande par QR code et retrait sans file ; créneau « Spécial étudiants » de 11h45 à 13h30"],["Fidélité et image","Carte de fidélité numérique (10e formule offerte) ; packaging kraft « cuit au feu de bois » ; compte Instagram et affiches au lycée et à l'IUT"],["Lieu","Coin goûter / révisions le mercredi après-midi"]],[3.5,12.5])
d.add_heading("4.4 Le tri des idées",2)
P("Pour trier nos idées, nous avons utilisé trois critères explicites, tirés de notre question HMW et de la mission de l'artisan :")
B(" l'idée lève-t-elle le frein principal de Léa (prix ou rapidité) ?","Critère 1 —")
B(" l'idée met-elle en valeur le four à bois et le savoir-faire ?","Critère 2 —")
B(" un artisan seul peut-il la mettre en place sans gros investissement ?","Critère 3 —")
table([["Idée","Frein de Léa","Four à bois","Faisable seul","Retenue ?"],
       ["Formule midi à prix fixe","✓","✓","✓","Oui"],["Galette garnie au feu de bois","✓","✓","✓","Oui"],["Précommande QR code + retrait sans file","✓","–","✓","Oui"],
       ["Carte de fidélité numérique","✓","–","✓","Oui"],["Recette du mois votée sur Instagram","–","✓","✓","Oui"],["Packaging kraft « cuit au feu de bois »","–","✓","✓","Oui"],
       ["Borne de commande en boutique","✓","–","–","Non"],["Gâteaux trompe-l'œil","–","–","✓","Non"],["Coin goûter / révisions","–","–","–","Non"]],[6.5,2.3,2.3,2.3,2.3])
P("Nous avons retenu les idées qui remplissent au moins deux critères sur trois. La borne de commande, par exemple, répond bien au manque de temps, mais elle coûte cher et ne valorise pas le four à bois : nous l'avons écartée.")
d.add_heading("4.5 Le concept retenu : « Le P'tit Feu »",2)
P("Les idées retenues forment un concept unique : « Le P'tit Feu », une formule midi à 6,50 € pensée pour les lycéens et les étudiants.")
B(" une galette cuite au four à bois, garnie au choix parmi 3 recettes, dont une « recette du mois » créée avec les clients grâce à un vote sur Instagram ;","Le plat :")
B(" une boisson et un mini-dessert « classique revisité » (mini-éclair, cookie maison) ;","L'accompagnement :")
B(" précommande via un QR code (lien Instagram / click & collect) et retrait sans file au comptoir « P'tit Feu » ;","La commande :")
B(" une carte de fidélité numérique, la 10e formule est offerte.","La fidélité :")
P("Le prix de 6,50 € reste sous le budget maximum de Léa (7 €) et proche du panier moyen en boulangerie (6,33 € en 2025).")

# --- Phase 4
d.add_heading("5. Phase 4 — Le prototype",1)
P("Pour donner une forme concrète à notre concept, nous avons réalisé un prototype de moyenne fidélité, en trois éléments. Chacun répond à un frein précis de notre persona. [Si l'IA a été utilisée : nous avons utilisé un outil d'intelligence artificielle uniquement pour visualiser nos idées, après les avoir trouvées en groupe, et nous avons corrigé les visuels proposés.]")
d.add_heading("5.1 Le parcours client",2)
table([["Moment","Ce que fait Léa"],["10h00, en cours","Elle voit l'affiche « Le P'tit Feu » à l'IUT et scanne le QR code."],["10h02","Elle choisit la galette du mois, une boisson et un cookie, et paie 6,50 € en ligne."],["12h05","Elle passe au comptoir « P'tit Feu » : pas de file, sa commande l'attend."],["12h15","Elle partage une photo de son sac kraft « cuit au feu de bois » en story."],["Ensuite","Sa carte de fidélité avance : la 10e formule sera offerte."]],[4,12])
d.add_heading("5.2 Élément 1 : les visuels et le support Instagram",2)
P("Nous avons amélioré la présentation des produits avec une affiche et des visuels pensés pour Instagram, mettant en scène la galette au feu de bois et les desserts, avec le QR code de précommande. Cet élément répond au frein de l'image « traditionnelle » : un produit beau et partagé sur les réseaux fait connaître la boulangerie auprès des jeunes.")
cadre("[Insérer ici le visuel produits / l'affiche]")
d.add_heading("5.3 Élément 2 : le packaging kraft « à emporter »",2)
P("Le packaging kraft porte le logo de la boulangerie et la mention « cuit au feu de bois ». Il répond au manque de visibilité face aux chaînes : le sac circule dans l'IUT et sur les réseaux, et met en avant l'argument du four à bois partout où il passe.")
cadre("[Insérer ici le visuel du packaging]")
d.add_heading("5.4 Élément 3 : l'écran de précommande",2)
P("La maquette d'écran montre la précommande sur téléphone : après avoir scanné le QR code, le client choisit sa galette, sa boisson et son dessert, paie et choisit son heure de retrait. Cet élément répond au frein du temps et de la file d'attente, et correspond à l'habitude de la cible de commander depuis son téléphone.")
cadre("[Insérer ici la maquette d'écran de téléphone]")

# --- Phase 5
d.add_heading("6. Phase 5 — Tester et préparer le pitch",1)
d.add_heading("6.1 Les objections anticipées de Ludovic Petitjean",2)
P("Nous nous sommes mis à la place de l'artisan pour anticiper ses objections. Ce test nous a amenés à faire évoluer notre prototype.")
table([["Objection","Notre réponse","Ce que ça a changé"],
       ["« Est-ce rentable à 6,50 € ? »","Le prix est proche du panier moyen (6,33 €). Les galettes cuisent dans un four déjà chaud.","Nous limitons l'offre à 3 recettes et à un mini-dessert, pour éviter les pertes."],
       ["« Je n'ai pas le temps de gérer une appli. »","Un simple formulaire en ligne et un QR code suffisent, sans développement coûteux.","Nous avons abandonné l'idée de la borne et de l'application : l'écran du prototype est devenu un écran de précommande sur téléphone."],
       ["« Est-ce que ça ne dénature pas mon image ? »","Le concept met en avant le four à bois et les classiques revisités, au cœur de son savoir-faire.","La mention « cuit au feu de bois » est présente sur le packaging, l'affiche et l'écran."],
       ["« Comment les jeunes vont-ils le savoir ? »","Affiches au lycée et à l'IUT, compte Instagram, vote de la recette du mois.","Le QR code de précommande a été ajouté directement sur l'affiche."]],[4.2,5.9,5.9])
P("[À compléter si vous l'avez fait : nous avons montré le prototype à [X] étudiants ; ils nous ont dit… ; nous avons donc modifié…]",italic=True,color=GRIS)
d.add_heading("6.2 Les bénéfices attendus",2)
table([["Pour la cible (Léa)","Pour la boulangerie Petitjean"],["Un vrai repas à 6,50 €, sous son budget","Un nouveau public le midi, aujourd'hui parti au fast-food"],["Aucune file d'attente grâce à la précommande","Le four à bois mieux utilisé et mis au cœur de l'offre"],["Un produit original, à partager sur les réseaux","De la visibilité gratuite sur Instagram"],["Une récompense : la 10e formule offerte","Une différence nette face aux 12 autres boulangeries du Creusot"]],[8,8])
d.add_heading("6.3 La trame de notre pitch (10 minutes)",2)
table([["Partie","Contenu","Qui","Durée"],["1","Rappel du contexte, de la cible et du problème (carte d'empathie)","[Prénom 1]","≈ 3 min"],["2","Persona, question HMW, outil de créativité et sa justification, tri des idées","[Prénom 2]","≈ 3 min"],["3","Concept, parcours client, prototype, objections de l'artisan","[Prénom 3]","≈ 3 min"],["4","Bénéfices pour la cible et pour la boulangerie, conclusion (à trois)","Tous","≈ 1 min"],["—","Questions du jury","Tous","après"]],[1.3,9.5,3,2.2])

# --- Conclusion
d.add_heading("Conclusion",1)
P("La démarche Design Thinking nous a permis de partir des besoins réels de notre cible plutôt que de nos propres idées. En écoutant les étudiants, nous avons compris que le problème n'était pas la qualité de la boulangerie, qu'ils reconnaissent, mais son prix perçu, son manque de rapidité et son image le midi.")
P("Avec « Le P'tit Feu », nous proposons une réponse à ces trois freins : une formule à 6,50 €, précommandée sur téléphone et récupérée sans file, avec des produits et un packaging qui donnent envie d'être partagés. Cette solution s'appuie sur le plus grand atout de Ludovic Petitjean, son four à bois, et reste réalisable par un artisan, sans gros investissement. Une prochaine étape serait de tester la formule en conditions réelles, sur quelques semaines, en suivant le nombre de formules vendues et de précommandes.")

# --- Sources
d.add_heading("Sources",1)
for s_ in ["MyBoulange, fiche et avis clients de la Boulangerie Pâtisserie L. Petitjean — myboulange.fr","Pages Jaunes, boulangeries-pâtisseries artisans au Creusot — pagesjaunes.fr","Boulangeries-Patisseries.fr, recensement des boulangeries du Creusot","Epsimas, Étude de marché de la boulangerie 2025-2026 — epsimas.com","Zepros Resto, « Chiffres d'affaires, rentabilité, fréquentation... en 2025 » — resto.zepros.fr","AffaireDidées, « SIRHA 2025 : La boulangerie française à l'aube d'une nouvelle ère »","Act Développement, « Les inquiétudes et perspectives économiques du secteur en 2025 »"]:
    B(s_)

# --- Annexes
d.add_heading("Annexes",1)
d.add_heading("Annexe 1 — Trace du brainstorming",2)
cadre("[Insérer ici la photo de vos feuilles / post-it]")
d.add_heading("Annexe 2 — Questions posées aux étudiants",2)
for q in ["Où mangez-vous le midi, et combien de fois par semaine ?","Combien dépensez-vous en moyenne pour un repas du midi ?","Qu'achetez-vous en boulangerie, et à quel moment ?","Qu'est-ce qui vous empêche d'aller à la boulangerie le midi ?","Qu'est-ce qui vous ferait venir plus souvent ?"]:
    d.add_paragraph(q,style='List Number')
d.save('Dossier-Le-Ptit-Feu.docx')
