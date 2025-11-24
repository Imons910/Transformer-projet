from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os

def create_word_report():
    doc = Document()

    # --- STYLE DU TITRE ---
    title = doc.add_heading('Rapport de Projet : Fine-Tuning de Depth Anything V2 avec LoRA', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- INFOS BINOME ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Membres du binôme : Alban Haton & Simon Guillet\n')
    run.bold = True
    run.font.size = Pt(12)
    p.add_run('Master Vision & Robotique\nDate : Janvier 2025')

    doc.add_page_break()

    # ===================================================================

    # --- 1. INTRODUCTION ---
    doc.add_heading('1. Introduction et Objectifs', level=1)
    
    p = doc.add_paragraph()
    p.add_run("L'estimation de la profondeur à partir d'une seule image (Monocular Depth Estimation) est un défi majeur. Si des modèles comme ")
    p.add_run("Depth Anything V2").bold = True
    p.add_run(" offrent d'excellentes performances généralistes, ils manquent de précision sur des capteurs industriels spécifiques.")
    
    p = doc.add_paragraph("L'objectif est d'adapter ce modèle sur un jeu de données acquis avec une caméra Zivid (RGB + Depth). Pour répondre aux contraintes matérielles, nous avons utilisé la technique LoRA (Low-Rank Adaptation).")

    # --- 2. DONNÉES ---
    doc.add_heading('2. Données et Prétraitement', level=1)
    
    doc.add_heading('2.1. Extraction et Conversion d\'Échelle', level=2)
    p = doc.add_paragraph("Les données Zivid sont initialement en millimètres. Une Loss initiale très élevée (2.6e6) a mis en évidence ce problème. Nous avons appliqué une conversion :")
    
    # Formule simulée
    equation = doc.add_paragraph("Si Z > 100 alors Z = Z / 1000.0")
    equation.alignment = WD_ALIGN_PARAGRAPH.CENTER
    equation.runs[0].italic = True

    doc.add_heading('2.2. Gestion des Données Manquantes', level=2)
    doc.add_paragraph("Nous générons un masque binaire M pour ignorer les pixels invalides (NaN ou infinis) : M = 1 si valide, 0 sinon.")

    # --- 3. METHODOLOGIE ---
    doc.add_heading('3. Méthodologie : Implémentation de LoRA', level=1)
    
    doc.add_heading('3.1. Justification Théorique', level=2)
    p = doc.add_paragraph("Le Fine-Tuning complet est coûteux. LoRA suppose que l'adaptation a un rang intrinsèque faible. Au lieu de modifier la matrice de poids W0, nous apprenons deux matrices de rang r (ici r=16) : B et A.")
    
    eq = doc.add_paragraph("W = W0 + ΔW = W0 + B A")
    eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    eq.runs[0].bold = True

    p = doc.add_paragraph("Nous avons ciblé les matrices Query et Value des couches d'attention. Cela représente seulement 1.18% des paramètres totaux (294 912 paramètres), rendant le projet faisable sur une seule carte graphique.")

    # --- 4. PROTOCOLE ---
    doc.add_heading('4. Protocole d\'Entraînement', level=1)
    
    items = [
        ("Fonction de Coût :", "MSE Masquée (Masked Mean Squared Error)"),
        ("Optimiseur :", "AdamW (Learning Rate = 1e-4)"),
        ("Durée :", "10 Époques"),
        ("Batch Size :", "4")
    ]
    for label, val in items:
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(label).bold = True
        p.add_run(f" {val}")

    # --- 5. RESULTATS ---
    doc.add_heading('5. Résultats et Analyse', level=1)

    doc.add_heading('5.1. Choix des Métriques', level=2)
    p = doc.add_paragraph("Le sujet suggérait des métriques de classification. S'agissant d'une régression dense, nous utilisons les standards de l'état de l'art (Eigen et al.) :")
    metrics = [
        "RMSE (Root Mean Squared Error) : Pénalise les fortes erreurs.",
        "AbsRel : Erreur relative normalisée.",
        "Précision (δ < 1.25) : Pourcentage de pixels fiables."
    ]
    for m in metrics:
        doc.add_paragraph(m, style='List Bullet')

    doc.add_heading('5.2. Comparaison Quantitative', level=2)
    doc.add_paragraph("Comparaison entre le modèle de base (Zero-Shot) et après 10 époques de Fine-Tuning :")

    # --- TABLEAU ---
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Medium Shading 1 Accent 1' # Style bleu pro
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Métrique'
    hdr_cells[1].text = 'Zero-Shot'
    hdr_cells[2].text = 'Fine-Tuned (10 Ep)'
    hdr_cells[3].text = 'Amélioration'

    data = [
        ('Précision (δ < 1.25)', '16.64%', '75.26%', '+58.6 pts'),
        ('RMSE (m)', '1.795 m', '0.319 m', '-82%'),
        ('AbsRel', '1.289', '0.150', '-88%'),
    ]

    for metric, zero, fine, imp in data:
        row_cells = table.add_row().cells
        row_cells[0].text = metric
        row_cells[1].text = zero
        row_cells[2].text = fine
        row_cells[3].text = imp

    # --- ANALYSE ---
    doc.add_heading('5.3. Analyse des Courbes', level=2)
    doc.add_paragraph("[INSÉRER ICI L'IMAGE : courbes_entrainement.png]")
    p = doc.add_paragraph("Figure 1 : Évolution de la Loss et de la Précision.")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True
    
    doc.add_paragraph("La convergence est rapide. La précision atteint 75.26%, confirmant que le modèle apprend la structure 3D spécifique aux données Zivid.")

    doc.add_heading('5.4. Analyse Qualitative', level=2)
    doc.add_paragraph("[INSÉRER ICI L'IMAGE : resultat_epoch_10.png]")
    p = doc.add_paragraph("Figure 2 : Gauche: RGB | Milieu: Vérité | Droite: Prédiction.")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].italic = True
    
    doc.add_paragraph("Visuellement, le modèle restaure la géométrie des objets avec des contours nets, corrigeant les défauts du modèle initial.")

    # --- 6. DEFIS ---
    doc.add_heading('6. Défis Techniques', level=1)
    doc.add_paragraph("1. Compatibilité RTX 5070 : Erreur 'NoKernelImage' due à l'architecture trop récente. Résolu via un repli CPU pour garantir la reproductibilité.", style='List Number')
    doc.add_paragraph("2. Échelle : L'écart mm/m empêchait la convergence. La normalisation dynamique a été la clé.", style='List Number')

    # --- 7. CONCLUSION ---
    doc.add_heading('7. Conclusion', level=1)
    p = doc.add_paragraph("Ce projet démontre la puissance de LoRA. En entraînant moins de ")
    p.add_run("1.2%").bold = True
    p.add_run(" des paramètres, nous avons transformé un modèle générique (16% de précision) en un modèle spécialisé fiable (")
    p.add_run("75.3%").bold = True
    p.add_run("), validant l'approche pour des applications industrielles.")

    # SAUVEGARDE
    output_filename = 'Rapport_Projet_Final.docx'
    doc.save(output_filename)
    print(f"✅ Fichier '{output_filename}' généré avec succès !")

if __name__ == "__main__":
    create_word_report()