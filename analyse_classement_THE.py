import os
import sqlite3
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.label import DataLabelList

##############################################################
#### Auteurs:
    # Selma Bennani
    # Adam Bouhaben
    # Fabrice Huret
    # Billy Lemoine

##############################################################

##############################################################
## Fonctions utilitaires pour dessiner
# un diagramme à barres et circulaire
## À NE PAS MODIFIER
##############################################################

from openpyxl.chart import LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties

def generer_line_chart(ws, titre, titre_x, titre_y):
    """
    Génère un graphique en courbes (LineChart) et l'insère dans une feuille Excel.
    - Les données sont extraites à partir des colonnes :Année, Enseignement,
    Recherche et Score THE. Trois séries sont affichées sur le même graphique.

    Args:
    - ws : Worksheet, feuille Excel contenant les données
    - titre : str, titre du graphique
    - titre_x : str, titre de l’axe X
    - titre_y : str, titre de l’axe Y

    Returns:
    None
    """

    chart = LineChart()
    chart.title = titre
    chart.x_axis.title = titre_x
    chart.y_axis.title = titre_y

    # Données : colonnes B à D, avec en-têtes
    data = Reference(ws, min_col=2, max_col=4, min_row=1, max_row=ws.max_row)

    # Catégories : années (colonne A, sans header)
    categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # Taille du graphique
    chart.height = 12
    chart.width = 20

    # Couleurs des courbes
    couleurs = ["4F81BD", "C0504D", "9BBB59"]  # bleu, rouge, vert

    for serie, couleur in zip(chart.series, couleurs):
        serie.graphicalProperties = GraphicalProperties(
            solidFill=couleur
        )
        serie.graphicalProperties.line.width = 25000  # ligne plus visible

    ws.add_chart(chart, "G5")


def generer_bar_chart(ws, titre, titre_x, titre_y):
    """
    Génère un graphique à barres et l'insère dans une feuille Excel.

    Args:
    -ws : Worksheet, Feuille Excel où les données et le graphique
    seront ajoutés.
    -titre, titre_x, titre_y : str, Titre du graphique à barres, et
    titres des axes X et Y

    Returns :
    None

    """

    chart = BarChart()
    chart.type = "col"
    chart.title = titre
    chart.y_axis.title = titre_y
    chart.x_axis.title = titre_x
    # Données = colonne 2, à partir de la ligne 1 (inclut le header)
    data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
    # Catégories = colonne 1, à partir de la ligne 2 (sans header)
    categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    chart.legend = None  # supprime la légende
    chart.height = 12
    chart.width = 20
    #exemples de couleurs : "4F81BD" → bleu, "C0504D" → rouge,
    # "9BBB59" → vert, "8064A2" → violet, "F79646" → orange
    serie = chart.series[0]
    serie.graphicalProperties = GraphicalProperties(
        solidFill="8064A2")

    serie.graphicalProperties.line.noFill = True
    ws.add_chart(chart, "G5")


def generer_pie_chart(ws, titre):
    """
    Génère un graphique circulaire et l'insère dans une feuille Excel.

    Args:
    -ws : Worksheet, Feuille Excel où les données et le graphique seront ajoutés.
    -titre : str, Titre du graphique circulaire.

    Returns:
    None

    """
    chart = PieChart()
    chart.title = titre

    data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row, max_col=2)
    categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    # Labels: afficher uniquement le %
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showSerName = False   # <- enlève "Nombre d'universités"
    chart.dataLabels.showVal = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showLegendKey = False

    chart.height = 12
    chart.width = 20
    ws.add_chart(chart, "G5")

##############################################################
#### Fonctions à compléter pour explorer la BDD
##############################################################

def ouvrir_connexion_bdd(chemin_vers_bdd): 
    """
    Établit une connexion à une base de données SQLite.
    et paramètre les résultats sous forme de dictionnaire

    Args:
    chemin_bdd : str, Chemin du fichier de la base de données SQLite.

    Returns:
    sqlite3.Connection, Objet de connexion à la base de données
    si la connexion est réussie, ou None si la connexion échoue (fichier de BDD
    n'existe pas)

    """
    # On vérifie si le chemin existe pour exécuter la fonction
    if os.path.exists(chemin_vers_bdd):
        conn = sqlite3.connect(chemin_vers_bdd)
        conn.row_factory = sqlite3.Row 
        print("connexion reussi")
        return conn
    print("Le chemin n'existe pas")
    return None

def fermer_connexion_bdd(connexion):
    """
    Ferme la connexion à la base de données SQLite et affiche un message
    si la connexion est fermée, ou affiche un
    message d'erreur si aucune connexion active à fermer

    Args:
    connexion : sqlite3.Connection, Objet de connexion à fermer.

    Returns:
    None
    """
    if isinstance(connexion, sqlite3.Connection):
        connexion.close()
        print("La connexion a été fermée.")
    else:
        print("L'argument n'est pas une connexion sqlite3.")

def kpis_dashboard(connexion):
    """
    Calcule des indicateurs globaux (KPIs) à partir de la base SQLite.
    Retourne un dictionnaire contenant :
    - le nombre d'universités, de pays, de régions
    - des moyennes (PIB/hab, taux d’alphabétisation)
    - des moyennes des scores (enseignement, recherche, international, THE, ratios H/F)

    Args:
    connexion : sqlite3.Connection, connexion ouverte vers la base.

    Returns:
    dict : dictionnaire {nom_kpi: valeur} (valeurs numériques arrondies).
    """
    cur = connexion.cursor()
    req = """
        SELECT (SELECT count(DISTINCT nom_univ) FROM universites) as Nbre_univ, 
        (SELECT count(DISTINCT nom_pays) FROM pays) as Nbre_pays,
        (SELECT count(DISTINCT nom_reg) FROM regions) as Nbre_reg,
        (SELECT ROUND(AVG(pib_par_hab), 3) FROM pays) as Moy_pib_hab, 
        (SELECT ROUND(AVG(taux_alphabet), 3) FROM pays) as Moy_taux_alphabet,
        (SELECT ROUND(AVG(score_enseig), 3) FROM classements) as Moy_enseig,
        (SELECT ROUND(AVG(score_qualite_rech), 3) FROM classements) as Moy_rech,
        (SELECT ROUND(AVG(pct_etud_internat), 3) FROM classements) as Moy_etud_intern,
        (SELECT ROUND(AVG(score_the), 3) FROM classements) as Moy_score_the,
        (SELECT ROUND(AVG(ratio_fem), 3) FROM classements) as Moy_ratio_fem,
        (SELECT ROUND(AVG(ratio_hom), 3) FROM classements) as Moy_ratio_hom;

        """
    cur.execute(req)
    return dict(cur.fetchone())

def top5_pays_par_critere(connexion, annee=2025, critere="score_enseig"):
    """
    Renvoie les 5 pays ayant la meilleure moyenne pour un critère donné
    (enseignement, recherche ou score THE) pour une année donnée.

    Args:
    connexion : sqlite3.Connection
    annee : int
    critere : str

    Returns:
    list[tuple] : [(nom_pays, moyenne), ...]
    """
    cur = connexion.cursor()
    req = """
        SELECT p.nom_pays, round(AVG(c.""" + critere + """), 2) AS moy_score
        FROM classements as c
        JOIN universites as u
        ON c.id_univ = u.id_univ
        JOIN pays as p
        ON u.id_pays = p.id_pays
        WHERE c.annee = ?
        GROUP BY p.nom_pays
        ORDER BY moy_score DESC
        LIMIT 5;
          """
    cur.execute(req, (annee, ))
    # Création d'un dictionnaire pour récuperer les paire clés-valeurs
    L = list(dict(ligne) for ligne in cur.fetchall())
    # Initialisation à vide d'une liste de tuple
    L_t = []
    # On ajoute chaque tuple avec nom_pays et moy_score dans la liste
    for k in range(len(L)):
        L_t.append((L[k]["nom_pays"], L[k]['moy_score']))
    return L_t

def top10_pays_multi_annees(connexion, annee_deb=2020, annee_fin=2025, critere="score_enseig"):
    """
    Classe et renvoie les 10 meilleurs pays ayant plus que 20 universités sur plusieurs années
    selon un critère donné, en tenant compte du nombre d’universités par pays.

    Args:
    connexion : sqlite3.Connection
    annee_deb : int
    annee_fin : int
    critere : str

    Returns: 10 premiers pays
    list[tuple] : [(nom_pays, score_final, nb_univ), ...]
    """
    curs = connexion.cursor()
    req = """
        SELECT p.nom_pays, ROUND((AVG(c."""+ critere +""")) * ((COUNT(DISTINCT(u.id_univ))) * 1.0 / (COUNT(DISTINCT(u.id_univ)) + 20) * 1.0), 3) AS score_final,
        COUNT(DISTINCT(u.id_univ)) AS nbre_univ
        FROM classements AS c
        JOIN universites AS u ON u.id_univ = c.id_univ
        JOIN pays AS p ON p.id_pays = u.id_pays
        WHERE c.annee BETWEEN ? AND ? AND c."""+ critere +""" IS NOT NULL
        GROUP BY p.nom_pays
        HAVING COUNT(DISTINCT(u.id_univ)) > 20
        ORDER BY score_final DESC
        LIMIT 10;
        """
    curs.execute(req, (annee_deb, annee_fin,))
    # Conversion en liste de tuples les résultats de la requête SQL
    res = list(tuple(ligne) for ligne in curs.fetchall())
    return res

def repartition_universites_par_region(connexion, annee=2025):
    """
    Renvoie la répartition des universités par région pour une année donnée.

    Args:
    conn : sqlite3.Connection
    annee : int

    Returns:
    list[tuple] : [(region, nombre_universites), ...]
    """
    curseur = connexion.cursor()
    requete = """
    SELECT count(nom_univ) as univ, nom_reg
    FROM universites as u
    INNER JOIN pays AS p 
    on u.id_pays = p.id_pays
    INNER JOIN regions as r
    on r.id_reg = p.id_reg
	INNER JOIN classements AS c
	on c.id_univ = u.id_univ
	WHERE annee = ?
    group by nom_reg;
    """
    curseur.execute(requete, (annee,))
    # Conversion en liste de tuples les résultats de la requête SQL, et on ordonne les résultats avec les clés
    list_repartition = [(ligne['nom_reg'], ligne['univ'] ) for ligne in curseur.fetchall()]
    return list_repartition

def top20_universites(connexion, annee=2025):
    """
    Renvoie les 20 meilleures universités selon le rang pour une année donnée.

    Args:
    connexion : sqlite3.Connection
    annee : int

    Returns:
    list[tuple] :
    (université, région, pays, rang, score_the, score_enseig, score_recherche)
    """
    cur = connexion.cursor()
    req = """
       SELECT c.rang, u.nom_univ, p.nom_pays, r.nom_reg, c.score_enseig, c.score_qualite_rech, c.score_the
       FROM classements as c
       JOIN universites as u ON c.id_univ = u.id_univ
       JOIN pays as p ON u.id_pays = p.id_pays
       JOIN regions as r ON p.id_reg = r.id_reg
       WHERE annee = ?   
       AND c.score_enseig IS NOT NULL
       AND c.score_qualite_rech IS NOT NULL
       AND c.score_the IS NOT NULL                     
       ORDER BY rang
       LIMIT 20;
       """
    cur.execute(req, (annee,))
    # Conversion en liste de tuples les résultats de la requête SQL, et on ordonne les résultats avec les clés
    list_top20 = [(ligne['rang'], ligne['nom_univ'],
                         ligne['nom_reg'], ligne['nom_pays'],
                         ligne['score_the'], ligne['score_enseig'],
                         ligne['score_qualite_rech']) for ligne in cur.fetchall()]
    return list_top20
    

def n_universites_par_indicateur(connexion, annee=2025, indicateur="score_enseig", n=5, sens="top"):
    """
    Renvoie les n meilleures ou pires universités selon un indicateur donné.

    Args:
    connexion : sqlite3.Connection
    annee : int
    indicateur : str
    n : int
    sens : str ("top" ou "bottom")

    Returns:
    list[tuple] : [(valeur, université, pays, rang), ...]
    """
    indicateurs= {
        "score_enseig": "score_enseig",
        "score_qualite_rech": "score_qualite_rech"
    }
    if sens == "top":
        ordre = "DESC"
    else :
        ordre = "ASC"
    var = indicateurs[indicateur]
    curseur = connexion.cursor()
    requete = f"""
    select nom_univ, {var} AS valeur, rang, nom_pays
    FROM classements AS c
    JOIN universites AS u 
    ON u.id_univ = c.id_univ
    JOIN pays AS p 
    ON u.id_pays = p.id_pays
    where annee = ? And {var} is not null
    order by {var} {ordre}
    LIMIT ?; 
    """
    curseur.execute(requete,(annee, n))
    #Conversion en liste de tuples les résultats de la requête SQL, et on ordonne les résultats avec les clés
    list_n_universites = [(ligne['valeur'], ligne['nom_univ'], ligne['nom_pays'], ligne['rang']) for ligne in curseur.fetchall()]
    return list_n_universites

def rechercher_universites(conn, pays=None,  annee= None, min_enseig=None, min_rech=None, limit=50):
    """
    Recherche des universités selon des critères optionnels
    et renvoie une liste ordonnée par rang.

    Args:
    conn : sqlite3.Connection
    pays : str ou None
    annee : int ou None
    min_enseig : float ou None
    min_rech : float ou None
    limit : int

    Returns:
    list[tuple] :
    (rang, université, pays, score_enseig, score_rech, score_the)
    """

    curseur = conn.cursor()
    requete = """
    SELECT rang, score_enseig as min_enseig, score_qualite_rech as min_rech, nom_univ, score_the, nom_pays
    FROM classements AS c
    INNER JOIN universites AS u 
    on u.id_univ = c.id_univ
    INNER JOIN pays AS p 
    on u.id_pays = p.id_pays
    where (? IS NULL OR nom_pays = ?)
        AND (? IS NULL OR annee = ?)
        AND (? IS NULL OR score_enseig >= ?)
        AND (? IS NULL OR score_qualite_rech >= ?) 
    limit ?; 
    """
    curseur.execute(requete,(pays, pays, annee, annee, min_enseig, min_enseig,
                            min_rech, min_rech, limit))
    # Conversion en liste de tuples les résultats de la requête SQL, et on ordonne les résultats avec les clés
    resultat =[(ligne['rang'], ligne['nom_univ'], ligne['nom_pays'], ligne['min_enseig'], ligne['min_rech'], ligne['score_the']) for ligne in curseur.fetchall()]
    return resultat

def evolution_universite(connexion, nom_univ):
    """
    Récupère l'évolution annuelle d'une université donnée.

    Args:
    - connexion : connexion sqlite3
    - nom_univ : str, nom (ou partie du nom) de l'université

    Returns:
    Liste de tuples :
    [(annee, score_enseig, score_qualite_rech, score_the), ...]
    """

    cur = connexion.cursor()
    req = """
        SELECT annee, c.score_enseig, c.score_qualite_rech, c.score_the
        FROM classements as c
        JOIN universites as u on u.id_univ = c.id_univ
        WHERE u.nom_univ = ?
        AND c.score_enseig IS NOT NULL
        AND c.score_qualite_rech IS NOT NULL
        AND c.score_the IS NOT NULL
        ORDER BY annee ASC;
        """
    cur.execute(req, (nom_univ,))
    # Conversion en liste de tuples les résultats de la requête SQL, et on ordonne les résultats avec les clés
    list_evolution = [(ligne['annee'], ligne['score_enseig'],
                       ligne['score_qualite_rech'], ligne['score_the'])
                       for ligne in cur.fetchall()]
    return list_evolution
    # On crée une liste de dictionnaires pour récupérer les paires clés-valeurs


def stats_internationalisation_vs_recherche (connexion,  annee=2025):
    """
    Analyse le lien entre internationalisation des étudiants
    et performances en recherche et score THE.

    Args:
    connexion : sqlite3.Connection
    annee : int

    Returns:
    dict : {classe: {"moy_rech": x, "moy_the": y}}
    """
    cur = connexion.cursor()
    req = """
        SELECT
	        CASE 
		        WHEN c.pct_etud_internat between 0 and 9 THEN '0-10'
		        WHEN c.pct_etud_internat between 10 and 19 THEN '10-20'
		        WHEN c.pct_etud_internat between 20 and 29 THEN '20-30'
		        ELSE '30+'
	        END AS classe,
        ROUND(avg(c.score_qualite_rech), 4) as moy_rech,
        ROUND(avg(c.score_the), 4) as moy_the
        FROM classements as c
        WHERE annee = ?
        AND c.score_qualite_rech IS NOT NULL
        AND c.score_the IS NOT NULL
        AND c.pct_etud_internat IS NOT NULL
        GROUP BY classe;
        """
    cur.execute(req, (annee, ))
    # On récupère dans un dictionnaire les résultats de la requête SQL
    dico_moy = dict((ligne[0], {'moy_rech' : ligne[1], 'moy_the' : ligne[2]}) for ligne in cur.fetchall())
    return dico_moy

def stats_pib_vs_perf (connexion,  annee=2025):
    """
    Compare les performances universitaires selon le niveau de PIB par habitant.

    Args:
    connexion : sqlite3.Connection
    annee : int

    Returns:
    dict : {classe_pib: {"moy_enseig": x, "moy_rech": y, "moy_the": z}}
    """
    cur = connexion.cursor()
    req = """
        SELECT
	        CASE 
		        WHEN pib_par_hab <= 1135 THEN 'Faible'
		        WHEN pib_par_hab between 1136 AND 4495 THEN 'Intermédiaire'
		        WHEN pib_par_hab >= 4496 THEN 'Elevé'
	        END AS classe,
        round(avg(c.score_enseig), 4) as moy_enseig, 
        round(avg(c.score_qualite_rech), 4) as moy_rech, 
        round(avg(c.score_the), 4) as moy_the
        FROM classements as c
        JOIN universites as u ON u.id_univ = c.id_univ
        JOIN pays as p on p.id_pays = u.id_pays
        WHERE annee = ?
        and pib_par_hab IS NOT NULL
        AND score_enseig IS NOT NULL
        AND score_qualite_rech IS NOT NULL
        AND score_the IS NOT NULL
        GROUP BY classe
        ORDER BY pib_par_hab;
        """
    # On exécute la requête de SQL via Python
    cur.execute(req, (annee, ))
    # Création d'un dictionnaire, chaque classe de PIB ayant les moyennes associées
    dico_moy = dict((ligne[0], {'moy_enseig' : ligne[1], 'moy_rech' : ligne[2], 'moy_the' : ligne[3]}) 
                       for ligne in cur.fetchall())
    return dico_moy

def repartition_par_alphabetisation(connexion, annee=2025):
    """
    Calcule la répartition des universités selon le taux d’alphabétisation
    du pays pour une année donnée.

    Args:
    connexion : sqlite3.Connection
    annee : int

    Returns:
    dict : {classe: nombre_universites}
    """
    curs = connexion.cursor()
    req = """
        SELECT
            CASE
	            WHEN p.taux_alphabet BETWEEN 80 AND 90 THEN "[80-90]"
                WHEN p.taux_alphabet < 80 THEN "<80"
                WHEN p.taux_alphabet > 90 THEN ">90"
            END AS classe,
            COUNT(DISTINCT u.id_univ) AS nbre_univ
        FROM classements AS c
        JOIN universites AS u ON u.id_univ = c.id_univ
        JOIN pays AS p ON p.id_pays = u.id_pays
        WHERE c.annee = ? AND p.taux_alphabet IS NOT NULL
        GROUP BY classe;
        """
    curs.execute(req, (annee, ))
    # On récupère dans un dictionnaire les résultats de la requête SQL
    res = dict(ligne for ligne in curs.fetchall())
    return res

def stats_mixite_vs_perf(connexion, annee=2025):
    """
    Analyse le lien entre la mixité (ratio femmes) et les performances universitaires.

    Args:
    connexion : sqlite3.Connection
    annee : int

    Returns:
    dict : {classe: {"moy_ens": x, "moy_rech": y}}
    """
    curs = connexion.cursor()
    req = """
        SELECT
	        CASE
		        WHEN c.ratio_fem BETWEEN 40 AND 60 THEN "[40-60]"
		        WHEN c.ratio_fem < 40 THEN "<40"
		        WHEN c.ratio_fem > 60 THEN ">60"
            END AS classe,
            ROUND(AVG(score_enseig), 4) AS moy_enseig,
	        ROUND(AVG(score_qualite_rech), 4) AS moy_rech
        FROM classements AS c
        JOIN universites AS u ON u.id_univ = c.id_univ
        JOIN pays AS p ON p.id_pays = u.id_pays
        WHERE c.annee = ? AND c.ratio_fem IS NOT NULL
        GROUP BY classe;
        """
    curs.execute(req, (annee,))
    # On récupère dans un dictionnaire les résultats de la requête SQL
    res = dict((ligne[0], {"moy_enseig" : ligne[1], "moy_rech" : ligne[2]}) for ligne in curs.fetchall())
    return res

def stats_migration_vs_perf(connexion, annee=2025):
    """
    Analyse le lien entre la migration nette du pays
    et les performances universitaires.

    Args:
    connexion : sqlite3.Connection
    annee : int

    Returns:
    dict : {classe: {"moy_ens": x, "moy_rech": y}}
    """
    curs = connexion.cursor()
    req = """
        SELECT
            CASE
                WHEN p.migration_nette < 0 THEN 'negative'
                WHEN p.migration_nette = 0 THEN 'nulle'
                WHEN p.migration_nette > 0 THEN 'positive'
            END AS classe,
        ROUND(AVG(c.score_enseig), 1) AS moy_enseig,
        ROUND(AVG(c.score_qualite_rech), 1) AS moy_rech
        FROM classements AS c
        JOIN universites AS u ON u.id_univ = c.id_univ
        JOIN pays AS p ON p.id_pays = u.id_pays
        WHERE c.annee = ?
        AND p.migration_nette IS NOT NULL
        AND c.score_enseig IS NOT NULL
        AND c.score_qualite_rech IS NOT NULL
        GROUP BY classe;
        """
    curs.execute(req, (annee, ))
    # On récupère dans un dictionnaire les résultats de la requête SQL
    res = dict((ligne[0], {"moy_enseig" : ligne[1], "moy_rech" : ligne[2]}) for ligne in curs.fetchall())
    return res


############################################################
###### Programme principal pour tester vos fonctions
###### et insérer dans le fichier excel
############################################################
if __name__ == "__main__":

    chemin_vers_bdd = "classement_universites_par_THE.db"
    # 1. Ouvrir la connexion
    connexion = ouvrir_connexion_bdd(chemin_vers_bdd)
    if connexion:
        
        # Création du classeur
        wb = Workbook()
        wb.remove(wb.active)

        
        # KPIS dashboard
        D_kpis = kpis_dashboard(connexion) #KPIS
        ws_KPI = wb.create_sheet("KPIs – Vue Globale")
        titre_col_Kpis = list(D_kpis.keys())
        ws_KPI.append(["Indicateur", "Valeur"])
        for col in titre_col_Kpis:
            ws_KPI.append([col, D_kpis[col]])

        
        # Top 5 pays enseig 2025
        critere="score_enseig" 
        Ltop_5_enseig = top5_pays_par_critere(connexion, annee=2025, critere=critere)
        ws_top5_enseig  = wb.create_sheet("Top5 Pays – Enseignement (2025)")
        ws_top5_enseig .append(["Pays", "Score moyen enseignement"])
        for k in range(len(Ltop_5_enseig)):
            ws_top5_enseig .append(Ltop_5_enseig[k])
        generer_bar_chart(ws_top5_enseig , "Top 5 pays en 2025 en fonction du score d'enseignement", "Pays", "Score moyen")

        
        # Top 5 pays recherche 2025
        critere="score_qualite_rech"
        Ltop_5_rech = top5_pays_par_critere(connexion, annee=2025, critere=critere)
        ws_top5_rech = wb.create_sheet("Top5 Pays – Recherche (2025)")
        ws_top5_rech.append(["Pays", "Score moyen recherche"])
        for k in range(len(Ltop_5_rech)):
            ws_top5_rech.append(Ltop_5_rech[k])
        generer_bar_chart(ws_top5_rech, "Top 5 pays en 2025 en fonction du score de recherche", "Pays", "Score moyen")

        
        # Top 5 pays score_the 2025
        critere="score_the"
        Ltop_5_score_the = top5_pays_par_critere(connexion, annee=2025, critere=critere)
        ws_top5_score_the = wb.create_sheet("Top5 Pays – THE (2025)")
        ws_top5_score_the.append(["Pays", "Score THE moyen"])
        for k in range(len(Ltop_5_score_the)):
            ws_top5_score_the.append(Ltop_5_score_the[k])
        generer_bar_chart(ws_top5_score_the, "Top 5 pays en 2025 en fonction du score de The World", "Pays", "Score moyen")


         # Top 10 pays enseignement 2020-2025
        top10e = top10_pays_multi_annees(connexion, 2020, 2025, "score_enseig")
        ws_top10e = wb.create_sheet("Top10 Pays–Ens(2020–2025)") # Création d'une nouvelle feuille
        ws_top10e.append(("Pays", "Score moyen enseignement (20-25)", "Nombre universités")) # Ajout des noms des colonnes
        for ligne in top10e:  # Ajout du résultat ligne par ligne
            ws_top10e.append(tuple(ligne)) 
        generer_bar_chart(ws_top10e, "Top10 Pays–Ens(2020–2025)", "Pays", "Score moyen") # Génération du diagramme à barres

        
        # Top 10 pays recherche 2020-2025
        critere = "score_qualite_rech"
        top10r = top10_pays_multi_annees(connexion, 2020, 2025, "score_qualite_rech")
        ws_top10r = wb.create_sheet("Top10 Pays–Rech(2020–2025)")
        ws_top10r.append(("Pays", "Score moyen recherche (20-25)", "Nombre universités"))
        for ligne in top10r:
            ws_top10r.append(tuple(ligne))
        generer_bar_chart(ws_top10r, "Top10 Pays–Rech(2020–2025)", "Pays", "Score moyen")

        
        # Répartition des universités par région pour une année donnée
        repartition_univ_region = repartition_universites_par_region(connexion, annee=2025)
        # Création de la feuille
        ws_repartition = wb.create_sheet("Univ par Région")
        ws_repartition.append(["Région", "Nombre d’universités"])
        for region, nb in repartition_univ_region:
           ws_repartition.append([region, nb])

        generer_pie_chart(ws_repartition,
        titre="Répartition des universités par région")

        
        # Création tableau score du top 20 des universités
        top20 = top20_universites(connexion)
        ws_top20 = wb.create_sheet("Top 20 Univ (Index)")
        ws_top20.append(("Rang", "Nom Université", "Nom région", "Nom pays", "Score THE", "Score enseig", "Score recherche"))
        for ligne in top20:
            ws_top20.append(tuple(ligne))

        
        # Top 5 des universités selon le Score d'enseignement
        top_5_univ = n_universites_par_indicateur(connexion)
        # Création de la feuille
        ws_repartition = wb.create_sheet("Top 5 Univ – Enseignement")
        # probleme
        ws_repartition.append(["Score enseignement", "Université", "Pays", "Rang"])
        # probleme
        for rang , univ, pays, score_enseig in top_5_univ:
           ws_repartition.append([rang , univ, pays, score_enseig])

        # Classement des 5 pires universités selon le Score d'enseignement
        bottom_5_univ = n_universites_par_indicateur(connexion, sens="bottom")
        # Création de la feuille
        ws_repartition = wb.create_sheet("Bottom 5 Univ – Enseignement")
        # probleme
        ws_repartition.append(["Score enseignement", "Université", "Pays", "Rang"])
        # probleme
        for rang , univ, pays, score_enseig in bottom_5_univ:
           ws_repartition.append([rang , univ, pays, score_enseig])

        # Top 5 des universités selon le Score de la qualité de leur recherche 
        top_5_univ_rech = n_universites_par_indicateur(connexion, indicateur="score_qualite_rech")
        # Création de la feuille
        ws_repartition = wb.create_sheet("Top 5 Univ – Recherche")
        # probleme
        ws_repartition.append(["Score recherche", "Université", "Pays", "Rang"])
        # probleme
        for rang , univ, pays, score_rech in top_5_univ_rech:
           ws_repartition.append([rang , univ, pays, score_rech])
        
        # Classement des 5 pires universités selon le Score de la qualité de leur recherche 
        bottom_5_univ_rech = n_universites_par_indicateur(connexion, sens="bottom", indicateur="score_qualite_rech")
        # Création de la feuille
        ws_repartition = wb.create_sheet(" Bottom 5 Univ – Recherche")
        # probleme
        ws_repartition.append(["Score recherche", "Université", "Pays", "Rang"])
        # probleme
        for rang , univ, pays, score_rech in bottom_5_univ_rech:
           ws_repartition.append([rang , univ, pays, score_rech])

        
        # Recherche des universités avec un score d’enseignement d’au moins 40 
        bottom_5_univ_rech = rechercher_universites(connexion, pays="France", annee=2020, min_enseig=40)
        # Création de la feuille
        ws_repartition = wb.create_sheet("Recherche Univ")
        # probleme
        ws_repartition.append(["Rang", "Université", "Pays", "Enseig", "Recherche", "Score THE"])
        # probleme
        for rang , univ, pays,score_enseig, score_rech, score_THE in bottom_5_univ_rech:
           ws_repartition.append([rang , univ, pays,score_enseig, score_rech, score_THE])
        

        # Création line chart évolution d'une université au fil des années
        evol_univ = evolution_universite(connexion, 'Panthéon-Sorbonne University – Paris 1')
        ws_evol_univ = wb.create_sheet("Evolution-Univ")
        ws_evol_univ.append(("Année", "Enseig", "Recherche", "Score THE"))
        for ligne in evol_univ:
            ws_evol_univ.append(tuple(ligne))
        generer_line_chart(ws_evol_univ, "Evolution d'une université au fil des années", "Année", "Niveau du score")


        # Création tableau internationalisation vs recherche
        intern_rech = stats_internationalisation_vs_recherche(connexion)
        ws_intern_rech = wb.create_sheet("Internat vs Recherche")
        ws_intern_rech.append(("Classe %", "Moy recherche", "Moy score THE"))
        for classe, moyennes in intern_rech.items():
            ws_intern_rech.append((classe, moyennes["moy_rech"], moyennes["moy_the"]))

        
        # Création tableau PIB vs perf
        pib_perf = stats_pib_vs_perf(connexion)
        ws_pib_perf = wb.create_sheet("PIB Par Hab vs Perf")
        ws_pib_perf.append(("Catégorie PIB", "Moy enseignement", "Moy recherche", "Moy score THE"))
        for classe, moyennes in pib_perf.items():
            ws_pib_perf.append((classe, moyennes["moy_enseig"], moyennes["moy_rech"], moyennes["moy_the"]))

        
        # Répartition par taux d'alphabétisation
        alpha = repartition_par_alphabetisation(connexion)
        ws_alpha = wb.create_sheet("Top & Alphabétisation")
        ws_alpha.append(("Classe alphabétisation", "Nombre univ"))
        for classe, nbre_univ in alpha.items():
            ws_alpha.append((classe, nbre_univ))
        generer_pie_chart(ws_alpha, "Répartition du nombre d'universités selon la classe d'alphabétisation")

        
        # Lien entre mixité et performances universitaires
        mix = stats_mixite_vs_perf(connexion)
        ws_mix = wb.create_sheet("Mixité vs Perf")
        ws_mix.append(("Classe ratio femmes", "Moy enseignement", "Moy Recherche"))
        for classe, moyennes in mix.items():
            ws_mix.append((classe, moyennes["moy_enseig"], moyennes["moy_rech"]))

        
        # Classement des pays selon la migration nette
        migr = stats_migration_vs_perf(connexion)
        ws_migr = wb.create_sheet("Migration vs Perf")
        ws_migr.append(("Classe migration", "Moy enseignement", "Moy recherche"))
        for classe, moyennes in migr.items():
            ws_migr.append((classe, moyennes["moy_enseig"], moyennes["moy_rech"]))

        
        # Sauvegarde et génération de fichier Excel
        wb.save('resultats_analyse.xlsx')
        print("Exportation et visualisation terminées.")
        fermer_connexion_bdd(connexion)


