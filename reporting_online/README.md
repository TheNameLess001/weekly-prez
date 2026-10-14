# Yassir Performance Studio — Streamlit Online

Interface en français pour importer des exports Excel/CSV, vérifier les colonnes et doublons, comparer des restaurants et générer les présentations éditables de **20 slides** validées : performance, meilleures ventes, catégories, Top 3, hausses/Drop, plan d’action et annexes.

## Déploiement Streamlit Community Cloud

Dans le dépôt `TheNameLess001/weekly-prez`, cette version est ajoutée dans `reporting_online/`. L’ancienne application à la racine est conservée.

1. Sur [Streamlit Community Cloud](https://share.streamlit.io), créer une application.
2. Dépôt : `TheNameLess001/weekly-prez`.
3. Branche : `codex/streamlit-online-reporting` (ou `main` après intégration de la PR).
4. Point d’entrée : `reporting_online/streamlit_app.py`.
5. Dans Advanced settings, sélectionner Python **3.12**, puis ajouter un secret personnel `APP_PASSWORD` dans les Secrets de l’application. `secrets.example.toml` montre la syntaxe, avec une valeur à remplacer.
6. Déployer. Choisir l’accès privé dans les réglages de partage pour le portefeuille métier. Les secrets restent dans Streamlit, jamais dans GitHub.

Le code démarre sans secret en **mode démonstration** : le jeu fictif Belhaj est utilisable, les imports réels sont désactivés. Dès qu’APP_PASSWORD est configuré, la connexion précède les imports. Les vrais chiffres et logos ne sont pas inclus dans le dépôt.

Les dépendances Python se trouvent dans `reporting_online/requirements.txt`. Le fichier `packages.txt` placé **à la racine du dépôt** installe LibreOffice Impress pour la conversion PDF, Liberation pour les polices et libcairo pour SVG. La configuration Streamlit se trouve à la racine dans `.streamlit/config.toml` pour le déploiement depuis un sous-dossier. Un déploiement autonome du dossier de cette application utilise ses propres fichiers `packages.txt` et `.streamlit/config.toml`.

## Parcours utilisateur

**Importer → vérifier → sélectionner → analyser → générer.**

- Importer plusieurs `.xlsx` ou `.csv` ; chaque feuille est inspectée.
- Sélectionner le fichier et la feuille à configurer, vérifier le grain et les colonnes, puis certifier l’exhaustivité uniquement sur les dates réellement couvertes.
- Consulter le contrôle qualité et corriger les identifiants absents. Restaurant ID reste la clé principale.
- Filtrer la ville, l’enseigne et le nom ; sélectionner un, plusieurs ou tous les restaurants du filtre.
- Choisir WoW, MoM, périodes courantes de même durée ou deux intervalles libres. Le dashboard affiche les dates, KPI, tendances, articles, catégories, Top 3 et baisses.
- Générer les PPTX éditables, PDF si LibreOffice est présent, puis télécharger le ZIP du lot. Les erreurs d’un restaurant n’interrompent pas les autres.
- Télécharger les préférences JSON pour réutiliser mapping, définitions, couleurs et périodes. Un changement de contenu du fichier invalide la certification des anciennes dates. Les logos se réimportent séparément.

## Fonctionnement en ligne

Les fichiers sont maintenant traités **sur le serveur Streamlit**, pas sur votre PC. Chaque session possède son propre répertoire temporaire et ses propres objets de calcul. Aucun cache partagé de données métier, aucune écriture dans le dépôt et aucun accès au dossier d’un autre utilisateur ne sont utilisés.

La reconnexion complète, la fermeture de session ou le redémarrage du serveur peut effacer fichiers, données et livrables. Télécharger les résultats et préférences avant de quitter. Cette version ne comporte pas de stockage métier permanent ni de synchronisation Drive. Un fichier est limité à 100 Mo et les sources d’une session à 400 Mo. Pour des imports très volumineux, un hébergement Python dimensionné et une base durable seront nécessaires.

## Règles de calcul conservées

- Aucun cumul des lignes produits avec les volumes/GMV des commandes.
- Commande : Restaurant ID, Order ID et date. Ligne article : Restaurant ID, Order ID, Line ID, date, produit et quantité. Catégorie et montant ligne complètent le reporting.
- Doublon identique exclu ; champs complémentaires fusionnés ; conflits exclus et couverture invalidée.
- Ratios recalculés depuis leurs numérateurs/dénominateurs. Moyennes opérationnelles pondérées par l’effectif de chaque durée. Médianes et utilisateurs uniques multi-périodes ne sont pas fabriqués depuis des agrégats.
- N/D est une donnée inconnue. Un restaurant/produit absent ne devient zéro que sur une source exhaustive. Les variations sont calculées lorsque les deux valeurs sont connues, avec un libellé observé ou certifié. Une valeur absente reste N/D. Pour une référence négative, seul l’écart absolu est interprétable ; le pourcentage reste N/A.
- L’agrégat hebdomadaire ne peut pas être découpé en dates fines. Les évolutions produit/Drop comparent les unités par jour sur les périodes certifiées.
- CA, commissions, promotions, contribution et GMV article exigent une définition et une convention HT/TTC validées dans les préférences. Le GMV article ne recopie jamais le montant complet de commande sur chaque article.
- Les constats et recommandations sont déterministes, sans IA distante ni conclusion causale automatique.

## Logos et thème

Importer PNG, JPG ou SVG dans Préférences. Les logos sont convertis en PNG localement sur le serveur, avec proportions conservées. Les SVG actifs ou contenant des ressources externes sont refusés. Palette et devise configurables. Arial est le choix PPTX par défaut, substitué par Liberation Sans pour le rendu Linux si nécessaire. Les vraies ressources de marque doivent être importées par l’utilisateur.

## Exécution locale / serveur Python

```bash
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

Pour un serveur Linux, installer également les paquets du `packages.txt`. Le Dockerfile est fourni pour un déploiement autonome : `docker build -t yassir-studio .`, puis fournir `APP_PASSWORD` via l’environnement du conteneur et exposer le port 8501 derrière HTTPS.

## Tests

```bash
python -m unittest discover -s tests -v
```

Les tests vérifient les calculs historiques, les articles, l’isolation de deux sessions, l’import/restauration de mappings, l’invalidation des dates sur fichier modifié, la protection par connexion et le parcours **réel Streamlit AppTest** : démo → analyse → export. Le déploiement Community Cloud et son accès doivent être vérifiés séparément après connexion au compte de publication.

Références officielles : [déploiement](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy), [dépendances](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/app-dependencies), [AppTest](https://docs.streamlit.io/develop/api-reference/app-testing/st.testing.v1.apptest).


### Calculs Admin Earnings et comparaison

- Préparation estimée = `picked up time − Accepted at`, attente éventuelle incluse. Préparation déclarée = `Ready by Restaurant − Accepted at`. Attente après prêt = `picked up time − Ready by Restaurant` ; le fichier ne fournit pas l’arrivée réelle du livreur au restaurant.
- Les heures sans date passent au lendemain seulement lorsque nécessaire ; les dates explicites incohérentes sont rejetées. Les durées invalides ne suppriment pas la commande commerciale.
- Moyennes et médianes sur les livraisons avec une durée valide, avec effectifs actuels et de référence. CTD : durée source en priorité, sinon livraison − création.
- Taux : écart en points et variation relative distincts. Totaux quotidiens divisés par les jours calendaires sélectionnés, même sur données observées ; pas d’imputation de jours absents à zéro.
- Montants affichés avec deux décimales, calculs effectués avant arrondi. Référence zéro : nouvelle activité pour les volumes commerciaux, sinon N/A et écart absolu.
- Quantités des produits extraites des descriptions `[Produit{quantité}]`. Les lignes articles explicites ont priorité pour éviter un double comptage. Aucun GMV article n’est inventé depuis le panier.
- Une catégorie ambiguë peut être récupérée lorsqu’un même produit du restaurant a une seule catégorie connue dans une commande à catégorie unique. Sinon elle reste non attribuable.
- Premiers clients observés = première livraison du restaurant dans l’historique importé. Ce KPI reste distinct des nouveaux clients globaux, qui exigent un référentiel de première commande.
- SLA : indicateur source ou seuil de click-to-door défini dans les préférences (0 = non défini). Disponibilité : heures disponibles / heures prévues, uniquement si fournies.
