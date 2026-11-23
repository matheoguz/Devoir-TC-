import json,re
from docx import Document
from docx.shared import Pt, RGBColor
n=json.load(open('oral.json'))
titres=["Titre","Contexte et démarche","Comprendre : carte d'empathie, freins, motivations","Définir : Léa et notre question","Imaginer : le brainstorming","Nos idées triées","Le concept et le parcours de Léa","Le prototype en 3 éléments","Tester : les objections","Bénéfices et conclusion (à trois)"]
d=Document(); d.styles['Normal'].font.name='Calibri'; d.styles['Normal'].font.size=Pt(12)
d.add_heading("Oral 10 min — Le P'tit Feu (3 personnes, 10 diapos)",0)
d.add_paragraph("≈ 1 100 mots : 8 à 9 minutes lus calmement + la présentation des visuels (diapo 8) = 10 minutes. Chacun parle environ 3 minutes. Remplacez [Personne 1], [Personne 2], [Personne 3] par vos prénoms.")
t=d.add_table(rows=4,cols=3); t.style='Light Grid Accent 1'
for r,row in enumerate([("Qui","Diapos","Partie"),("Personne 1","1 à 3 + début de la 10","Contexte, comprendre"),("Personne 2","4 à 6 + milieu de la 10","Définir, imaginer"),("Personne 3","7 à 9 + fin de la 10","Concept, prototype, tester")]):
    for c,v in enumerate(row): t.cell(r,c).text=v
for i,(t_,x) in enumerate(zip(titres,n),1):
    d.add_heading(f"Diapo {i} — {t_}",1)
    for who,txt in re.findall(r"\[(PERSONNE \d)\] (.*?)(?=\s*\[PERSONNE \d\]|$)",x,re.S):
        p=d.add_paragraph(); r=p.add_run(who.title()); r.bold=True; r.font.color.rgb=RGBColor(0xD4,0x5D,0x24)
        d.add_paragraph(txt.strip())
d.add_heading("Questions probables du jury",1)
for q,a in [("Pourquoi 6,50 € ?","Proche du panier moyen (6,33 €), sous les 7 € de Léa, au niveau d'un menu étudiant de fast-food."),("Une borne, ça coûte cher ?","Oui, c'est un investissement, mais les artisans investissent déjà en moyenne 17 623 € (2025), surtout dans le numérique. Et elle fait gagner du temps au rush."),("Comment savoir si ça marche ?","Formules vendues par jour, commandes en ligne et sur la borne, abonnés Instagram, cartes de fidélité remplies."),("Les autres boulangeries font déjà du snacking ?","Oui (81 %), mais pas une formule au feu de bois avec recette votée par les clients et commande sans attente."),("Vous avez utilisé l'IA ?","Oui, seulement pour mettre en image nos idées (borne, packaging, Instagram), après les avoir trouvées en groupe."),("Comment avez-vous noté vos idées ?","Répondez honnêtement : dites comment vous les avez vraiment notées (sur un doc, un téléphone, à l'oral…).")]:
    p=d.add_paragraph(style='List Bullet'); r=p.add_run(q+" "); r.bold=True; p.add_run(a)
d.save('Oral-Le-Ptit-Feu.docx')
