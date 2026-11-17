from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


OUT = Path("Dossier_projet_Chatbot_Conseil_Cosmetique_Pharmacie.docx")
NAVY = "17365D"
BLUE = "2F75B5"
PALE_BLUE = "D9EAF7"
LIGHT = "F3F6F9"
GRAY = "666666"

doc = Document()
section = doc.sections[0]
section.top_margin = Inches(0.75)
section.bottom_margin = Inches(0.7)
section.left_margin = Inches(0.8)
section.right_margin = Inches(0.8)

styles = doc.styles
styles["Normal"].font.name = "Aptos"
styles["Normal"].font.size = Pt(10)
styles["Normal"].font.color.rgb = RGBColor.from_string("222222")
styles["Normal"].paragraph_format.space_after = Pt(6)
for name, size, color in [
    ("Title", 28, NAVY),
    ("Heading 1", 19, NAVY),
    ("Heading 2", 13, BLUE),
    ("Heading 3", 11, NAVY),
]:
    st = styles[name]
    st.font.name = "Aptos Display" if name != "Normal" else "Aptos"
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = RGBColor.from_string(color)
    st.paragraph_format.keep_with_next = True
    if name == "Heading 1":
        st.paragraph_format.space_before = Pt(16)
        st.paragraph_format.space_after = Pt(7)
    elif name == "Heading 2":
        st.paragraph_format.space_before = Pt(10)
        st.paragraph_format.space_after = Pt(4)


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, color=None, size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.autofit = True
    for i, h in enumerate(headers):
        set_cell_text(t.rows[0].cells[i], h, True, "FFFFFF", 9)
        shade(t.rows[0].cells[i], NAVY)
    tr_pr = t.rows[0]._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, value in enumerate(row):
            set_cell_text(cells[i], value, size=8.5)
            if ri % 2 == 1:
                shade(cells[i], LIGHT)
    if widths:
        for row in t.rows:
            for idx, width in enumerate(widths):
                row.cells[idx].width = Inches(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t


def para(text="", style=None, bold_prefix=None, italic=False):
    p = doc.add_paragraph(style=style)
    if bold_prefix and text.startswith(bold_prefix):
        r = p.add_run(bold_prefix)
        r.bold = True
        p.add_run(text[len(bold_prefix):])
    else:
        r = p.add_run(text)
        r.italic = italic
    return p


def bullet(text, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    p.paragraph_format.space_after = Pt(2)
    p.add_run(text)
    return p


def numbered(text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(2)
    p.add_run(text)
    return p


def heading(text, level=1):
    doc.add_heading(text, level=level)


def callout(label, text, fill=PALE_BLUE):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    cell = t.cell(0, 0)
    shade(cell, fill)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(label + "  ")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(NAVY)
    p.add_run(text)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Page ")
    run.font.size = Pt(8)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    paragraph._p.append(fld)


# Header and footer
header = section.header.paragraphs[0]
header.text = "PROJET DE FIN D'ANNÉE  |  MASTER 1 INTELLIGENCE ARTIFICIELLE"
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for r in header.runs:
    r.font.name = "Aptos"
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor.from_string(GRAY)
page_field(section.footer.paragraphs[0])

# Cover page
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(65)
r = p.add_run("DOSSIER DE PROJET")
r.bold = True
r.font.size = Pt(13)
r.font.color.rgb = RGBColor.from_string(BLUE)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(22)
r = p.add_run("Assistant conversationnel pour le conseil\n de produits cosmétiques en pharmacie")
r.bold = True
r.font.name = "Aptos Display"
r.font.size = Pt(27)
r.font.color.rgb = RGBColor.from_string(NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(10)
r = p.add_run("Aide à la sélection de produits selon le besoin, le type de peau,\nles ingrédients et les précautions à vérifier")
r.font.size = Pt(13)
r.font.color.rgb = RGBColor.from_string(GRAY)

doc.add_paragraph()
table(
    ["Élément", "À compléter / information"],
    [
        ("Formation", "Master 1 Intelligence Artificielle"),
        ("Étudiant(e)(s)", "[Nom(s) et prénom(s)]"),
        ("Encadrant(e)", "[Nom et fonction]"),
        ("Établissement", "[Université / école]"),
        ("Année universitaire", "2025–2026"),
        ("Version du dossier", f"Version de travail — {date.today().strftime('%d/%m/%Y')}"),
    ],
    [2.0, 4.7],
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(22)
r = p.add_run("Document de cadrage, spécifications, architecture, protocole d’évaluation\net trame de mémoire et de soutenance")
r.italic = True
r.font.size = Pt(10)
doc.add_page_break()

heading("Avant-propos et statut du document", 1)
para(
    "Ce dossier rassemble une proposition complète pour cadrer et réaliser le projet. "
    "Les informations entre crochets sont à personnaliser. Les choix d’architecture sont des "
    "propositions de MVP, à confirmer avec l’encadrant et les contraintes de l’établissement."
)
callout(
    "État réel du dépôt",
    "Le dépôt fourni contient un script de préparation Open Beauty Facts et des données exportées. "
    "Il ne contient pas encore d’application conversationnelle, de moteur de recommandation validé, "
    "ni d’évaluation utilisateur. Les sections décrivant le chatbot constituent donc une spécification "
    "et un protocole de réalisation, pas des résultats expérimentaux déjà obtenus.",
    "FFF2CC",
)
heading("Sommaire", 1)
for item in [
    "1. Résumé exécutif",
    "2. Contexte, problématique et objectifs",
    "3. Périmètre et utilisateurs",
    "4. Besoins et exigences",
    "5. Données et qualité",
    "6. Solution proposée et architecture",
    "7. Parcours conversationnel et règles de recommandation",
    "8. Sécurité, éthique et cadre réglementaire",
    "9. Méthodologie et plan de réalisation",
    "10. Vérification et évaluation",
    "11. Risques et mesures de maîtrise",
    "12. Prototype attendu et scénario de démonstration",
    "13. Trame du mémoire",
    "14. Plan de soutenance",
    "15. Conclusion et perspectives",
    "Annexes : exigences, cas de test et sources",
]:
    bullet(item)
para("Dans Word, mettre à jour ce sommaire manuellement si les titres ou la pagination évoluent.", italic=True)
doc.add_page_break()

heading("1. Résumé exécutif", 1)
para(
    "Le projet vise à concevoir un assistant conversationnel destiné aux pharmaciens et à leurs équipes. "
    "À partir d’un échange guidé, l’outil recueille uniquement les éléments utiles au conseil cosmétique "
    "(besoin exprimé, préférences, informations de peau pertinentes et précautions déclarées), puis "
    "présente des produits candidats et les éléments de leur composition qui méritent vérification."
)
para(
    "L’assistant est une aide au travail du professionnel, non un outil de diagnostic, de prescription "
    "ou de remplacement du jugement pharmaceutique. Une approche hybride est proposée : dialogue "
    "structuré, recherche dans un catalogue contrôlé, filtres déterministes de précaution et génération "
    "de réponse limitée aux informations sourcées. En cas d’incertitude ou de signal d’alerte, le système "
    "s’abstient de recommander et invite à un examen par le pharmacien."
)
para(
    "Le dépôt contient actuellement un pipeline de préparation fondé sur Open Beauty Facts (OBF). "
    "Son rapport local indique 80 331 lignes lues, 18 891 produits conservés et 44 249 ingrédients "
    "distincts. Ces chiffres décrivent le jeu préparé par le script, pas un catalogue validé pour le conseil. "
    "L’échantillon observé révèle des notices multilingues ou bruitées, des produits non cosmétiques et "
    "des séparations d’ingrédients peu fiables. La première étape du travail est donc un audit et un "
    "nettoyage pilotés par un pharmacien."
)
table(
    ["Élément", "Proposition"],
    [
        ("Public cible", "Pharmaciens et préparateurs, sous supervision du pharmacien"),
        ("Valeur attendue", "Structurer le recueil du besoin et accélérer une présélection explicable"),
        ("MVP", "Application locale simple, catalogue contrôlé, règles explicites et recherche textuelle"),
        ("Données initiales", "Open Beauty Facts ; provenance et licence à respecter, qualité à auditer"),
        ("Principe de sécurité", "Aucune décision médicale autonome ; abstention et escalade humaine"),
        ("Critère de réussite", "Réponses traçables, pertinentes, sûres et évaluées avec des professionnels"),
    ],
)

heading("2. Contexte, problématique et objectifs", 1)
heading("2.1 Contexte", 2)
para(
    "Le conseil en dermocosmétique exige de relier le besoin formulé par la personne, ses préférences, "
    "la composition et les précautions d’emploi d’un produit. Les catalogues sont nombreux et les "
    "informations produits peuvent être incomplètes ou hétérogènes. Un assistant numérique peut "
    "structurer la recherche et rendre les critères de sélection plus visibles, à condition de conserver "
    "un professionnel comme décideur."
)
heading("2.2 Problématique", 2)
callout(
    "Question de recherche",
    "Comment concevoir et évaluer un assistant conversationnel explicable qui aide un pharmacien à "
    "présélectionner des produits cosmétiques à partir d’un besoin utilisateur et d’informations "
    "d’ingrédients, tout en maîtrisant les données incomplètes et le risque de conseil inapproprié ?",
)
heading("2.3 Objectif général", 2)
para(
    "Concevoir, réaliser et évaluer un prototype d’aide au conseil cosmétique en pharmacie, capable de "
    "recueillir un besoin, de retrouver des produits pertinents dans un catalogue contrôlé, d’expliquer "
    "les critères retenus et de signaler les limites ou précautions nécessitant une validation humaine."
)
heading("2.4 Objectifs spécifiques", 2)
for s in [
    "Analyser les besoins des pharmaciens et définir les scénarios d’usage du MVP.",
    "Auditer la qualité, la couverture, la provenance et les limites du catalogue disponible.",
    "Construire un modèle de données produit et un référentiel d’ingrédients documenté.",
    "Implémenter un parcours conversationnel court avec recherche et règles transparentes.",
    "Prévoir les situations d’abstention, de demande de clarification et d’orientation vers le professionnel.",
    "Évaluer la pertinence, la traçabilité, l’utilisabilité et les erreurs de sécurité avec un protocole explicite.",
]:
    bullet(s)
heading("2.5 Hypothèse de travail", 2)
para(
    "Un assistant hybride, dont les réponses sont contraintes par un catalogue et des règles explicables, "
    "peut réduire le temps de recherche de produits par rapport à une recherche libre, sans augmenter "
    "les recommandations inadaptées, si la qualité des données et la supervision professionnelle sont "
    "suffisantes. Cette hypothèse doit être testée ; elle n’est pas un résultat établi."
)

heading("3. Périmètre et utilisateurs", 1)
heading("3.1 Utilisateurs et rôles", 2)
table(
    ["Rôle", "Besoin", "Responsabilité dans le système"],
    [
        ("Pharmacien", "Comparer rapidement des options et vérifier les précautions",
         "Valide, corrige ou écarte les suggestions ; reste responsable du conseil"),
        ("Préparateur / équipe", "Structurer le recueil d’informations",
         "Utilise l’outil selon les procédures de l’officine et sous supervision"),
        ("Administrateur catalogue", "Maintenir les fiches et les sources",
         "Contrôle la provenance, la fraîcheur et les champs obligatoires"),
        ("Personne conseillée", "Exprimer un besoin et recevoir un conseil compréhensible",
         "Ne reçoit pas de diagnostic ni de recommandation automatique sans validation"),
    ],
)
heading("3.2 Inclus dans le MVP", 2)
for s in [
    "Dialogue guidé pour préciser le besoin cosmétique et les préférences.",
    "Recherche de candidats dans une liste de produits dont les champs essentiels ont été contrôlés.",
    "Affichage de la composition disponible, de la source et de la date de vérification.",
    "Explication des critères de correspondance et des informations manquantes.",
    "Signalement des précautions déclarées et orientation vers une vérification par le pharmacien.",
    "Journalisation technique minimale, sans identité de la personne conseillée.",
]:
    bullet(s)
heading("3.3 Hors périmètre explicite", 2)
for s in [
    "Diagnostic d’une maladie de peau, interprétation de symptômes ou prescription.",
    "Remplacement du pharmacien, de la notice fabricant ou d’un avis médical.",
    "Affirmations d’efficacité clinique déduites uniquement d’une liste d’ingrédients.",
    "Identification automatique certaine d’allergies ou de contre-indications à partir de texte bruité.",
    "Personnalisation publicitaire, profilage commercial et conservation de dossiers individuels.",
    "Déploiement réel en officine avant validation métier, sécurité et conformité.",
]:
    bullet(s)
heading("3.4 Scénarios d’usage principaux", 2)
numbered("La personne exprime un objectif cosmétique ; le pharmacien complète les préférences utiles.")
numbered("L’assistant demande une précision si le besoin est trop vague pour filtrer le catalogue.")
numbered("Le système présente une courte liste de candidats avec critères, sources et limites.")
numbered("Le pharmacien examine la composition et les précautions avant de formuler son conseil.")
numbered("En présence d’un signal d’alerte, d’une donnée critique manquante ou d’un résultat vide, l’assistant s’abstient et propose une vérification humaine.")

heading("4. Besoins et exigences", 1)
heading("4.1 Exigences fonctionnelles", 2)
table(
    ["ID", "Exigence", "Priorité", "Critère d’acceptation"],
    [
        ("F-01", "Recueillir le besoin en langage simple et permettre sa reformulation.",
         "Must", "Le pharmacien peut confirmer ou corriger le résumé du besoin."),
        ("F-02", "Demander uniquement les précisions nécessaires au filtrage.",
         "Must", "Les champs non nécessaires ne sont pas requis pour continuer."),
        ("F-03", "Rechercher dans les produits actifs du catalogue contrôlé.",
         "Must", "Chaque résultat renvoie à un identifiant produit et à sa source."),
        ("F-04", "Afficher les ingrédients et avertissements comme éléments à vérifier.",
         "Must", "La fiche indique clairement données manquantes et date de révision."),
        ("F-05", "Expliquer pourquoi un candidat apparaît.",
         "Must", "Au moins un critère observable est affiché pour chaque candidat."),
        ("F-06", "S’abstenir si un critère critique est inconnu ou si aucun résultat fiable n’existe.",
         "Must", "Aucune recommandation n’est fabriquée en réponse à un catalogue vide."),
        ("F-07", "Permettre au professionnel de corriger les filtres et de rejeter un candidat.",
         "Should", "Les contrôles sont disponibles avant la présentation à la personne."),
        ("F-08", "Afficher une mention de limites et le rôle du pharmacien.",
         "Must", "La mention reste visible dans l’écran de résultat et la démonstration."),
    ],
)
heading("4.2 Exigences non fonctionnelles", 2)
table(
    ["ID", "Exigence"],
    [
        ("NF-01", "Traçabilité : associer chaque information produit à une source et une date."),
        ("NF-02", "Confidentialité : minimiser les données personnelles et ne pas conserver l’échange par défaut."),
        ("NF-03", "Explicabilité : distinguer faits catalogués, règles et texte généré."),
        ("NF-04", "Robustesse : gérer champs absents, fautes de frappe et recherche sans résultat."),
        ("NF-05", "Accessibilité : interface lisible, libellés simples et contraste suffisant."),
        ("NF-06", "Maintenabilité : séparer interface, catalogue, règles, sources et journalisation."),
        ("NF-07", "Sécurité : contrôle d’accès si déployé, gestion sûre des secrets, mises à jour et sauvegardes."),
        ("NF-08", "Performance MVP : afficher un résultat de recherche local en moins de 2 secondes sur le poste cible, objectif à mesurer."),
    ],
)

heading("5. Données et qualité", 1)
heading("5.1 État des données dans le dépôt", 2)
para(
    "Le script prepare_obf.py télécharge ou lit un export Open Beauty Facts, sélectionne des champs, "
    "nettoie certains caractères et exporte un CSV/JSONL ainsi qu’un rapport. Le rapport fourni indique "
    "80 331 lignes brutes lues, 18 891 produits après nettoyage, 44 249 ingrédients distincts, "
    "12 218 produits avec un marqueur d’allergène et 10 824 avec un marqueur sensible. Ces valeurs "
    "sont des statistiques produites par le script ; elles ne mesurent ni l’exactitude ni la couverture "
    "des produits cosmétiques pertinents."
)
table(
    ["Constat du dépôt", "Conséquence", "Action recommandée"],
    [
        ("Champs ingrédients extraits par découpage sur virgules/points-virgules",
         "Les noms composés et les textes hors liste INCI peuvent être mal segmentés.",
         "Auditer un échantillon annoté ; utiliser un parseur et une validation adaptés aux listes INCI."),
        ("Notice parfois multilingue, bruitée ou avec erreur d’encodage",
         "Faux positifs de recherche et règles d’ingrédients peu fiables.",
         "Détecter langue/qualité, isoler la section INCI et exclure les fiches non fiables."),
        ("Exemples de produits non cosmétiques ou de catégories non pertinentes",
         "La recherche peut proposer un aliment, un dentifrice ou un produit hors périmètre.",
         "Filtrer par catégorie contrôlée et faire valider le périmètre par un pharmacien."),
        ("Noms de produits/ingrédients parfois mal transcrits",
         "Une absence de correspondance ne garantit pas l’absence d’un ingrédient.",
         "Traiter l’information manquante comme inconnue, pas comme négative."),
        ("Marqueurs codés en listes dans prepare_obf.py",
         "Un marqueur indicatif ne constitue pas un diagnostic ni une contre-indication complète.",
         "Référentiel versionné et vérifié par un professionnel ; afficher source et date."),
        ("Pays/catégories incomplets dans certains enregistrements",
         "La disponibilité locale et l’adéquation au marché cible sont incertaines.",
         "Limiter les résultats au catalogue officinal local validé."),
    ],
)
heading("5.2 Modèle de données minimal", 2)
table(
    ["Entité", "Champs proposés"],
    [
        ("Produit", "product_id, nom, marque, catégorie, pays, statut, URL source, date de mise à jour"),
        ("Composition", "product_id, INCI original, ingrédients normalisés, méthode et confiance d’extraction"),
        ("Précaution", "type, texte, source, version, statut de validation, validateur, date de revue"),
        ("Besoin de conseil", "catégories contrôlées et préférences temporaires ; aucune identité nécessaire"),
        ("Résultat", "produits affichés, critères satisfaits/non vérifiés, version du moteur et du catalogue"),
    ],
)
heading("5.3 Pipeline de qualité proposé", 2)
for s in [
    "Conserver la provenance et la licence de la source dès l’ingestion.",
    "Dédupliquer par identifiant stable, puis vérifier les fiches et catégories admissibles.",
    "Séparer la liste INCI des autres textes de l’emballage ; conserver l’original.",
    "Normaliser sans supprimer le texte source ; documenter chaque transformation.",
    "Marquer les champs manquants et les extractions incertaines comme inconnus.",
    "Faire valider un échantillon stratifié par un pharmacien ; mesurer précision, rappel et accord inter-évaluateurs.",
    "Versionner les données, référentiels et dates de revue ; retirer ou signaler les fiches périmées.",
]:
    numbered(s)
callout(
    "Décision recommandée",
    "Ne pas utiliser les marqueurs actuels du script comme filtres de sécurité en production. Ils sont "
    "présentés comme indicatifs par le code lui-même et nécessitent une révision métier et réglementaire.",
    "FCE4D6",
)

heading("6. Solution proposée et architecture", 1)
heading("6.1 Architecture MVP", 2)
para("Architecture simple, locale et contrôlable pour le prototype de Master :")
table(
    ["Couche", "Rôle", "Technologie indicative"],
    [
        ("Interface", "Dialogue guidé, filtres, fiches et avertissements", "Streamlit (prototype)"),
        ("Orchestration", "État de session, validation des entrées, étapes du dialogue", "Python"),
        ("Moteur de recherche", "Filtrage catégorie/pays, texte et critères explicités", "Pandas + SQLite/CSV"),
        ("Moteur de règles", "Règles documentées, abstention et signaux à vérifier", "Python, règles versionnées"),
        ("Données", "Catalogue validé, dictionnaire et sources", "CSV/JSONL au début ; SQLite ensuite"),
        ("Génération éventuelle", "Reformuler uniquement à partir de contenus récupérés", "Optionnelle ; ajout après baseline"),
    ],
)
para(
    "Pour une évolution vers un usage multi-postes, séparer l’interface d’une API (par exemple FastAPI), "
    "ajouter authentification et contrôle d’accès, gestion des versions du catalogue, surveillance, "
    "journalisation minimisée et revue de sécurité. Ce déploiement n’est pas nécessaire au MVP."
)
heading("6.2 Flux de traitement", 2)
for s in [
    "Le professionnel ouvre une nouvelle session sans saisir l’identité de la personne.",
    "L’interface collecte le besoin et les préférences utiles ; l’utilisateur confirme le résumé.",
    "Le validateur vérifie les valeurs attendues et les critères de prudence.",
    "Le moteur applique d’abord les règles d’abstention, puis recherche les candidats admissibles.",
    "Le moteur renvoie les fiches, sources et raisons de sélection ; aucune donnée inconnue n’est inférée.",
    "L’interface présente des options et leurs limites ; le pharmacien décide de la suite.",
]:
    numbered(s)
heading("6.3 Usage éventuel d’un modèle de langage", 2)
para(
    "Une baseline sans modèle génératif est recommandée. Si un LLM est ajouté, son rôle doit rester "
    "limité à la compréhension de formulations ou à la reformulation de résultats déjà récupérés. "
    "Il ne doit ni inventer des propriétés, ni conclure à l’absence d’allergène, ni produire une contre-indication "
    "non étayée. Chaque affirmation produit doit être reliée à une source consultable. Les réponses "
    "hors catalogue ou sans preuve doivent déclencher une clarification ou une abstention."
)

heading("7. Parcours conversationnel et règles de recommandation", 1)
heading("7.1 Questions utiles", 2)
for s in [
    "Quel est l’objectif cosmétique recherché ?",
    "Quelles préférences pratiques sont importantes (texture, parfum, budget indicatif, format) ?",
    "Y a-t-il un ingrédient que la personne sait déjà devoir éviter ? Le système ne confirme pas l’allergie.",
    "Le pharmacien doit-il examiner une situation particulière avant de poursuivre ?",
]:
    bullet(s)
para(
    "Les questions doivent être minimales et contextualisées. Les informations de santé ne doivent pas "
    "être collectées « au cas où ». Si le parcours fait apparaître une affection diagnostiquée, des symptômes "
    "importants, une réaction en cours, un traitement ou une situation particulière, le prototype doit "
    "orienter vers le pharmacien et suspendre toute suggestion automatique."
)
heading("7.2 Logique de sélection", 2)
table(
    ["Étape", "Comportement"],
    [
        ("1. Vérifier", "Contrôler l’intégrité de la session et la présence des informations strictement nécessaires."),
        ("2. Mettre en sécurité", "Détecter un cas nécessitant une vérification professionnelle ; interrompre la suggestion automatique."),
        ("3. Filtrer", "Retenir uniquement les produits actifs, dans le périmètre et dont les champs critiques sont suffisamment fiables."),
        ("4. Classer", "Trier selon des critères explicites : besoin/catégorie, préférence et disponibilité des données."),
        ("5. Expliquer", "Présenter correspondances, critères non vérifiés, source et fraîcheur des informations."),
        ("6. S’abstenir", "Ne pas proposer de produit si les données sont insuffisantes ou le résultat non fiable."),
    ],
)
heading("7.3 Exemple de dialogue de démonstration", 2)
table(
    ["Intervenant", "Exemple"],
    [
        ("Pharmacien", "La personne cherche un soin hydratant sans parfum, pour le visage."),
        ("Assistant", "Je peux rechercher des produits du catalogue vérifié. Souhaitez-vous préciser une texture ou un budget ?"),
        ("Pharmacien", "Texture légère ; aucune autre préférence connue."),
        ("Assistant", "J’ai trouvé deux fiches dans le catalogue. Le premier produit est indiqué comme soin visage et sa fiche ne liste pas de parfum ; cette information doit être vérifiée sur l’emballage actuel. Voici sa source et sa date de contrôle. Le second contient une donnée de composition incomplète : je le signale et ne le classe pas comme option confirmée."),
        ("Pharmacien", "Je vérifie la composition et le besoin avant de conseiller."),
    ],
)
callout(
    "Important",
    "Dialogue fictif destiné à illustrer l’interface. Il ne valide aucun produit, ingrédient, bénéfice ou conseil clinique.",
    "FFF2CC",
)

heading("8. Sécurité, éthique et cadre réglementaire", 1)
heading("8.1 Principes de prudence", 2)
for s in [
    "Le pharmacien garde la décision finale ; l’assistant n’établit aucun diagnostic.",
    "La composition incomplète ou inconnue est explicitement distinguée d’une absence d’ingrédient.",
    "Les termes « hypoallergénique », « non comédogène » ou équivalents ne sont pas inférés automatiquement.",
    "Aucune donnée de santé n’est stockée par défaut ; toute collecte doit être justifiée, minimisée et validée.",
    "Le système expose ses sources, sa date de revue, ses limites et sa version.",
    "Les règles de précaution sont revues et approuvées par un pharmacien compétent avant tout usage.",
    "Toute évolution vers un usage réel nécessite une analyse juridique, sécurité et protection des données.",
]:
    bullet(s)
heading("8.2 Données personnelles", 2)
para(
    "Les informations de peau ou de santé peuvent révéler des données sensibles selon le contexte. "
    "Pour le prototype, privilégier des données fictives ou anonymisées, une session éphémère, l’absence "
    "de compte utilisateur et la non-conservation des conversations. Avant un déploiement, déterminer "
    "les rôles de responsable de traitement et de sous-traitant, la base légale, les durées, les droits, "
    "les mesures de sécurité et la nécessité d’une analyse d’impact avec le référent compétent."
)
heading("8.3 Réglementation et responsabilité", 2)
para(
    "Le règlement européen relatif aux produits cosmétiques et les règles applicables aux données "
    "personnelles sont à examiner avec l’encadrant et les interlocuteurs juridiques de l’établissement. "
    "La qualification réglementaire d’un logiciel dépend notamment de sa finalité et de ses fonctions : "
    "ne pas conclure automatiquement qu’un outil est ou n’est pas un dispositif médical. Vérifier les "
    "textes consolidés et les obligations en vigueur à la date du projet."
)
heading("8.4 Biais et équité", 2)
para(
    "Un catalogue dominé par certaines marques, langues, pays ou types de peau peut conduire à des "
    "résultats incomplets. Documenter les populations et catégories absentes, ne pas présenter la liste "
    "comme exhaustive et faire tester l’interface par des profils variés."
)

heading("9. Méthodologie et plan de réalisation", 1)
heading("9.1 Démarche", 2)
for s in [
    "Cadrer avec l’encadrant et un pharmacien référent les cas d’usage, exclusions et critères d’acceptation.",
    "Auditer les données et définir un petit catalogue pilote validé avant de développer les règles.",
    "Implémenter une recherche de référence sans génération, puis une interface simple.",
    "Ajouter les règles d’abstention et tester les situations limites avant toute démonstration.",
    "Conduire une évaluation contrôlée et consigner résultats, limites et décisions de révision.",
    "Rédiger le mémoire à partir des travaux réellement réalisés et distinguer propositions et résultats.",
]:
    numbered(s)
heading("9.2 Calendrier indicatif sur 16 semaines", 2)
table(
    ["Semaines", "Travaux", "Livrables / jalon"],
    [
        ("1–2", "Cadrage, revue documentaire, entretiens métier", "Problématique, périmètre, scénarios validés"),
        ("3–4", "Audit source et qualité des données", "Rapport qualité, schéma produit, critères d’exclusion"),
        ("5–6", "Catalogue pilote et référentiel validé", "Jeu de test versionné et traçable"),
        ("7–9", "Recherche déterministe et règles de prudence", "Baseline testable, cas d’abstention"),
        ("10–11", "Interface conversationnelle du MVP", "Prototype local démontrable"),
        ("12", "Tests fonctionnels, données et sécurité", "Compte rendu de tests et corrections"),
        ("13–14", "Évaluation avec scénarios et utilisateurs experts", "Résultats mesurés et retours structurés"),
        ("15", "Analyse, rédaction du mémoire et limites", "Version complète du mémoire"),
        ("16", "Révision et répétition de soutenance", "Rapport final et présentation"),
    ],
)
heading("9.3 Livrables attendus", 2)
for s in [
    "Dossier de cadrage et cahier des charges.",
    "Rapport d’audit des données, provenance et règles validées.",
    "Prototype local avec instructions de lancement.",
    "Jeu de scénarios de test et résultats reproductibles.",
    "Mémoire avec méthode, résultats observés, limites et perspectives.",
    "Présentation de soutenance et démonstration sur données fictives ou contrôlées.",
]:
    bullet(s)

heading("10. Vérification et évaluation", 1)
heading("10.1 Tests fonctionnels minimaux", 2)
table(
    ["Cas", "Entrée/situation", "Résultat attendu"],
    [
        ("T-01", "Besoin clair et catalogue comportant des fiches valides", "Candidats sourcés et critères visibles."),
        ("T-02", "Aucune fiche correspondante", "Abstention, aucun produit inventé."),
        ("T-03", "Composition manquante", "Champ signalé inconnu ; aucune affirmation d’absence."),
        ("T-04", "Texte non pertinent ou produit hors catégorie", "Fiche exclue ou explicitement non recommandable."),
        ("T-05", "Allergie déclarée à un ingrédient", "Vérification humaine ; pas de garantie basée sur parsing imparfait."),
        ("T-06", "Symptôme préoccupant ou demande de diagnostic", "Arrêt du conseil automatique et orientation vers professionnel."),
        ("T-07", "Formulation ambiguë", "Question de clarification plutôt qu’inférence silencieuse."),
        ("T-08", "Source/date absente", "Fiche non présentée comme vérifiée."),
    ],
)
heading("10.2 Indicateurs", 2)
table(
    ["Dimension", "Indicateur proposé", "Méthode"],
    [
        ("Qualité données", "Précision des champs clés et taux de champs inconnus", "Annotation manuelle d’un échantillon stratifié"),
        ("Recherche", "Précision@k et rappel@k sur requêtes jugées par des pharmaciens", "Jeu de requêtes de référence"),
        ("Sécurité", "Taux de scénarios à risque conduisant correctement à l’abstention", "Tests adversariaux contrôlés, revus par expert"),
        ("Traçabilité", "Part des assertions produit avec source et date", "Audit des réponses"),
        ("Utilisabilité", "Temps de réalisation, taux de réussite et questionnaire bref", "Sessions supervisées"),
        ("Robustesse", "Comportement sur données manquantes, bruitées et formulations ambiguës", "Tests de cas limites"),
    ],
)
para(
    "Définir les seuils d’acceptation avec l’encadrant et le pharmacien avant l’évaluation. Rapporter "
    "l’effectif, la méthode d’échantillonnage et les intervalles ou limites. Ne pas présenter un score "
    "global comme preuve de sécurité clinique."
)
heading("10.3 Comparaison expérimentale possible", 2)
para(
    "Comparer la recherche manuelle dans le catalogue à la recherche assistée sur les mêmes scénarios, "
    "en mesurant le temps, la pertinence jugée et les erreurs de précaution. Un petit nombre de "
    "participants ne permet pas de généraliser ; présenter les résultats comme exploratoires."
)

heading("11. Risques et mesures de maîtrise", 1)
table(
    ["Risque", "Niveau", "Mesure de réduction"],
    [
        ("Composition mal extraite ou obsolète", "Élevé", "Catalogue pilote vérifié, date de contrôle, abstention si incertitude."),
        ("Recommandation interprétée comme diagnostic", "Élevé", "Limites visibles, validation humaine, pas de diagnostic ni de prescription."),
        ("Fausse assurance d’absence d’allergène", "Élevé", "Données manquantes = inconnues ; contrôle étiquette et pharmacien."),
        ("Biais de couverture du catalogue", "Moyen/élevé", "Documenter couverture et exclusions ; sources complémentaires contrôlées."),
        ("Collecte excessive de données de santé", "Élevé", "Minimisation, session éphémère, revue protection des données."),
        ("Hallucination d’un modèle génératif", "Élevé si LLM", "Baseline sans LLM, génération contrainte, citations et tests d’abstention."),
        ("Dépendance à une source communautaire", "Moyen", "Vérification indépendante et procédure de mise à jour."),
        ("Projet trop ambitieux pour le calendrier", "Moyen", "MVP sans LLM, catalogue limité, jalons et exclusions claires."),
    ],
)

heading("12. Prototype attendu et scénario de démonstration", 1)
heading("12.1 Écrans du MVP", 2)
table(
    ["Écran", "Contenu"],
    [
        ("Accueil", "Finalité de l’outil, limites, rôle du professionnel et bouton de nouvelle session."),
        ("Recueil", "Besoin, préférences facultatives et question de clarification."),
        ("Résultats", "Liste courte, raisons de sélection, données inconnues, sources et dates."),
        ("Fiche produit", "Nom, marque, catégorie, composition source, indicateurs à vérifier."),
        ("Abstention", "Raison de non-recommandation et prochaine étape humaine."),
    ],
)
heading("12.2 Démonstration en 3 minutes", 2)
numbered("Présenter la finalité et rappeler qu’il s’agit d’une aide au pharmacien, pas d’un diagnostic.")
numbered("Lancer une recherche sur un scénario fictif sans donnée personnelle.")
numbered("Montrer les critères retenus, la source, la date et un champ inconnu.")
numbered("Introduire un cas ambigu ou une information de sécurité manquante et montrer l’abstention.")
numbered("Conclure sur les limites du catalogue et les prochaines étapes de validation.")
heading("12.3 Critères de démonstration réussie", 2)
for s in [
    "Aucun conseil clinique non étayé n’est formulé.",
    "Chaque candidat peut être relié à une fiche source.",
    "Les informations manquantes sont visibles.",
    "Le cas de prudence entraîne effectivement une abstention ou une escalade.",
    "Le professionnel comprend pourquoi le système a proposé ou écarté un candidat.",
]:
    bullet(s)

heading("13. Trame du mémoire", 1)
table(
    ["Chapitre", "Contenu à rédiger"],
    [
        ("Introduction", "Contexte, enjeu, problématique, objectifs, hypothèse, contributions."),
        ("1. État de l’art", "Conseil dermocosmétique, systèmes de recommandation, dialogue, IA explicable et sécurité."),
        ("2. Analyse du besoin", "Parties prenantes, scénarios, exigences, périmètre et critères d’acceptation."),
        ("3. Données", "Source, licence, collecte, nettoyage, audit, limites et gouvernance."),
        ("4. Conception", "Architecture, modèle de données, dialogue, règles, abstention et choix techniques."),
        ("5. Réalisation", "Implémentation réelle, interfaces, modules, versions et difficultés."),
        ("6. Évaluation", "Protocole, jeux de test, participants, indicateurs, résultats et analyse."),
        ("7. Discussion", "Interprétation, limites, biais, risques, validité et améliorations."),
        ("Conclusion", "Réponse à la problématique, contributions, perspectives et recommandations."),
        ("Annexes", "Cahier des charges, scénarios, questionnaires, règles, sources et instructions."),
    ],
)
callout(
    "Consigne de rédaction",
    "N’écrire les résultats qu’après avoir exécuté les tests et l’évaluation. Séparer clairement les "
    "résultats mesurés, les propositions de conception et les travaux futurs.",
    "FFF2CC",
)

heading("14. Plan de soutenance", 1)
table(
    ["Diapositive", "Message clé"],
    [
        ("1. Titre", "Sujet, étudiant, encadrant et établissement."),
        ("2. Contexte", "Pourquoi le conseil cosmétique mobilise information et expertise."),
        ("3. Problématique", "Question de recherche et limites d’un conseil automatisé."),
        ("4. Objectifs et périmètre", "Ce que le prototype fait et ne fait pas."),
        ("5. Données", "Source, statistiques locales et principaux défauts détectés."),
        ("6. Architecture", "Parcours du besoin à la fiche sourcée et à la validation humaine."),
        ("7. Démonstration", "Recherche normale puis situation où le système s’abstient."),
        ("8. Méthode d’évaluation", "Scénarios, métriques, participants et précautions."),
        ("9. Résultats", "Uniquement mesures réellement obtenues et limites."),
        ("10. Discussion", "Risques, qualité des données, biais et valeur pour le pharmacien."),
        ("11. Conclusion", "Réponse, contributions et prochaines étapes."),
        ("12. Questions", "Références principales et ouverture à la discussion."),
    ],
)
para("Pour une présentation de 10 minutes, viser environ 45 secondes par diapositive hors démonstration et questions.")

heading("15. Conclusion et perspectives", 1)
para(
    "Le projet a un intérêt pédagogique et pratique s’il est conçu comme un outil d’aide à la recherche "
    "et à la vérification, non comme un conseiller de santé autonome. Le dépôt fournit un point de départ "
    "pour explorer Open Beauty Facts, mais les défauts visibles dans les données rendent indispensable "
    "un catalogue pilote validé avant toute recommandation. Le socle recommandé est un moteur déterministe "
    "explicable, un parcours court, une traçabilité systématique et une abstention sûre."
)
para(
    "Les perspectives incluent l’enrichissement contrôlé du catalogue, la validation par plusieurs "
    "professionnels, l’analyse comparative avec les pratiques actuelles et, seulement après une baseline "
    "satisfaisante, l’étude d’un modèle de langage soumis à des garde-fous et à une évaluation dédiée."
)

heading("Annexe A — Questions à valider avec l’encadrant", 1)
for s in [
    "Qui est l’utilisateur principal et qui valide les règles métier ?",
    "Quel pays et quel cadre officinal sont visés ?",
    "Quelles catégories de produits sont incluses et lesquelles sont exclues ?",
    "Quel catalogue peut être vérifié par un pharmacien dans le temps imparti ?",
    "Le prototype doit-il être uniquement local ou déployé sur un environnement de démonstration ?",
    "Quelles règles institutionnelles encadrent les données, les tests utilisateurs et les outils d’IA ?",
    "Quels seuils d’acceptation et quelles métriques seront exigés pour le mémoire ?",
    "Le projet doit-il intégrer un LLM, ou une baseline déterministe suffit-elle ?",
]:
    bullet(s)

heading("Annexe B — Liste de contrôle avant démonstration", 1)
for s in [
    "Utiliser uniquement des données fictives ou des données dont l’usage est autorisé.",
    "Vérifier que le catalogue de démo comporte source, date et statut de validation.",
    "Tester un résultat positif, un résultat vide, une donnée manquante et un cas d’abstention.",
    "Vérifier que l’interface ne prétend pas diagnostiquer ou exclure avec certitude un allergène.",
    "Ne pas afficher de données personnelles dans les captures ou le journal.",
    "Préparer une démonstration hors ligne et une sauvegarde du jeu de données contrôlé.",
]:
    bullet(s)

heading("Annexe C — Références et ressources", 1)
references = [
    ("Dépôt du projet", "prepare_obf.py et rapport local data_Open Beauty Facts/report.txt ; statistiques rapportées dans ce dossier."),
    ("Open Beauty Facts — données", "https://world.openbeautyfacts.org/data"),
    ("Open Data Commons — ODbL 1.0", "https://opendatacommons.org/licenses/odbl/1-0/"),
    ("Règlement (CE) no 1223/2009 relatif aux produits cosmétiques", "https://eur-lex.europa.eu/eli/reg/2009/1223/oj"),
    ("Règlement (UE) 2016/679 (RGPD)", "https://eur-lex.europa.eu/eli/reg/2016/679/oj"),
    ("Règlement (UE) 2023/1545, modification concernant l’étiquetage des allergènes de parfum", "https://eur-lex.europa.eu/eli/reg/2023/1545/oj"),
    ("CNIL — comprendre le RGPD", "https://www.cnil.fr/fr/rgpd-de-quoi-parle-t-on"),
]
for title, url in references:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(title + " — ")
    r.bold = True
    p.add_run(url)
para(
    "Les références réglementaires doivent être consultées dans leur version consolidée et vérifiées "
    "au moment de la rédaction finale. Pour Open Beauty Facts, conserver l’attribution et vérifier les "
    "conditions ODbL applicables à la réutilisation et à la redistribution des données.",
    italic=True,
)

doc.core_properties.title = "Assistant conversationnel pour le conseil de produits cosmétiques en pharmacie"
doc.core_properties.subject = "Dossier de projet de Master 1 en intelligence artificielle"
doc.core_properties.author = "[Nom de l'étudiant(e)]"
doc.core_properties.keywords = "pharmacie, cosmétique, chatbot, recommandation, IA"
doc.save(OUT)
print(f"Document créé : {OUT.resolve()}")
