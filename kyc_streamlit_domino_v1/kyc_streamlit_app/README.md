# Recertification KYC — Streamlit sur Domino

Application **1.0.0**, dérivée du notebook joint `kyc_qwen38_production_kyc_pipeline_name_similarity_v2(1).ipynb`. Elle reçoit `kyc_documents_test.zip`, exécute les fonctions validées et propose le téléchargement du rapport original du pipeline.

**État de livraison :** code complet et tests locaux sans GPU. Le déploiement Domino, le navigateur Streamlit et l'inférence réelle Qwen/H100 doivent être validés dans votre environnement. Aucun résultat de précision OCR réelle n'a été mesuré ici.

## Chemins contractuels

| Élément | Chemin exact |
|---|---|
| Code de l'application | `/mnt/kyc_streamlit_app/` |
| Entrée sauvegardée | `/mnt/kyc_app/input/kyc_documents_test.zip` |
| Extraction sûre | `/mnt/kyc_app/extracted/` |
| Référence existante, jamais demandée dans l'UI | `/mnt/recertification_lissage_updated.csv` |
| Rapport téléchargé | `/mnt/output_qwen38/reports/kyc_recertification_control.xlsx` |
| Journaux applicatifs | `/mnt/output_qwen38/logs/` |
| Métadonnées | `/mnt/output_qwen38/run_metadata/<run_id>.json` |
| Reçu du rapport courant | `/mnt/output_qwen38/run_metadata/current_report.json` |
| Verrou de traitement | `/mnt/output_qwen38/.kyc_processing.lock` |
| Verrou de durée de vie du serveur | `/mnt/output_qwen38/.kyc_server.lock` |
| Sorties et caches par archive | `/mnt/output_qwen38/datasets/<sha256_archive>/` |
| Ancien rapport conservé au début d'un nouveau traitement | `/mnt/output_qwen38/report_history/<nouveau_run_id>/kyc_recertification_control.xlsx` |

La séparation par empreinte du ZIP évite de confondre des populations différentes qui utilisent les mêmes identifiants clients. Le rapport final reste au chemin imposé. Réutiliser **le même ZIP binaire** retrouve les mêmes checkpoints. Recompresser un ZIP change son empreinte et crée un nouvel espace de travail. Les caches de notebooks historiques situés ailleurs ne sont pas migrés automatiquement.

Dans chaque `datasets/<sha256_archive>/`, le moteur conserve les sous-répertoires d'origine : `inventory`, `raw_ocr`, `structured`, `reports`, `performance`, `checkpoints`, `logs`, `images`, `customers`, `global`. `reports` est réservé à l'initialisation ; la publication est redirigée vers le répertoire contractuel global. `projection_history/<run_id>/` conserve les JSON matérialisés des exécutions précédentes avant leur reconstruction. Les caches OCR/structure ne sont pas déplacés.

## Ce qui est conservé et ce qui est adapté

Les **230 définitions de fonctions/classes** du notebook ont un AST identique dans `validated_core.py` et `notebook_config.py`. `reference/NOTEBOOK_MAPPING.md` fournit la correspondance exhaustive, et le test vérifie cette identité. Le notebook de référence fourni dans `reference/` conserve ses sources ; ses sorties historiques sont retirées pour ne pas redistribuer d'éventuelles données KYC.

Le gros moteur demeure un module unique : ses nombreuses fonctions partagent un état global. Le diviser artificiellement en OCR/matching/validation augmenterait le risque de modifier les dépendances ou leur ordre. Les responsabilités de l'application sont, elles, séparées.

Adaptations techniques explicites :

- `MATCH_ONLY` devient `FULL_PIPELINE` pour les nouveaux ZIP ; aucune limite diagnostic de population ou d'appels.
- Les chemins relatifs deviennent ceux ci-dessus, et le nom CSV configuré `recertication_lissage_updated.csv` devient le nom demandé `recertification_lissage_updated.csv`.
- Les appels automatiques de cellules sont retirés de l'import et orchestrés explicitement.
- L'état du traitement est réinitialisé entre deux archives ; les poids Qwen et le processor sont conservés.
- `print` et `display` du moteur passent par le journal local ; aucune preuve OCR n'est affichée dans Streamlit.
- La sécurité d'entrée vérifie toutes les entrées ZIP et les CRC avant l'extraction originale des seuls PDF cibles.
- Des callbacks enveloppent les fonctions de traitement sans modifier leurs arguments, retours ou règles.
- La reprise des caches RAW, structurés et benchmark est conservée. Après reconstruction des JSON documentaires, le matching Python est recalculé pour y réappliquer ses décisions. Les fonctions et écritures de checkpoints de matching restent présentes ; seule leur lecture est court-circuitée en exécution App. Cela évite qu'un reçu client ancien laisse les JSON documentaires reconstruits au statut `NOT_RUN`. Aucun appel Qwen n'est ajouté par ce recalcul.
- Les anciennes projections JSON sont archivées avant la structuration : une erreur actuelle ne peut pas réutiliser silencieusement un ancien résultat client. Les preuves/cache restent disponibles pour la reprise.
- La présence d'un unique H100 devient bloquante dans l'adaptateur. Le notebook se limitait à un avertissement si le GPU n'était pas un H100.

Les prompts, la déquantification FP8 vers BF16, `device_map="auto"`, les générations, le contrôle de thinking, le benchmark froid/chaud, la sélection des doublons famille + taille, les règles MRZ/expiration/domicile/Convention/FATCA/carton, les scores et les rapports ne sont pas réécrits.

Le seuil FATCA/carton reste **75,0**, inclusif sur le score non arrondi. Le carton conserve **nom fiable correspondant OU compte fiable exactement correspondant**, avec les gardes de reconnaissance/qualité d'origine. Un score de similarité seul ne supprime pas les indicateurs d'incertitude. Le contenu CSV n'est lu qu'après la passe structurée et n'entre jamais dans les prompts OCR.

## Architecture et exécution

`app.py` → `JobController` → verrou et thread unique → `pipeline.execute_locked` → fonctions originales de `validated_core`.

Un seul processus Streamlit et un seul thread de traitement lourd. `st.cache_resource` conserve le contrôleur. `model_manager` importe le moteur à la première demande et conserve son singleton. Le chargeur original ne charge les poids qu'au premier appel nécessitant une inférence. Aucun modèle au simple affichage de la page. Le deuxième travail réutilise le modèle ; un dépôt modifié ou un chargement partiellement échoué impose un redémarrage avant reprise.

Ordre effectif : validation ZIP → extraction et inventaire → provenance modèle → benchmark si nécessaire → `process_raw_dataset` → barrière RAW → `initialize_structure_stage` → `process_structured_dataset` → lecture CSV → `process_all_customers` → `export_reports` → `final_audit` → vérification du fichier et publication de son reçu SHA-256.

Les erreurs terminales de PDF sont conservées dans les résultats suivant les règles originales. Les erreurs fatales GPU, les échecs de persistance et les garde-fous RAW interrompent le traitement. Un rapport exploitable avec anomalies peut être livré ; la présence d'un Excel ne signifie pas que tous les documents sont conformes.

Les compteurs « conformes » concernent **la cohérence avec la base**, comme les statuts source. Les anomalies documentaires et la demande de relecture sont distinctes et peuvent recouper ces compteurs. `Clients traités` inclut les résultats client en erreur ; le compteur d'erreurs est affiché séparément. Les documents absents ne gonflent pas le nombre de PDF traités.

## Capture Domino : lecture exacte

La photo montre une **page de consignes “App development”**, pas un formulaire de déploiement avec des valeurs d'environnement renseignées.

| Élément effectivement lisible | Valeur mise en œuvre | Explication |
|---|---|---|
| Ne pas mettre `requirements.txt` à la racine du projet | Fichier uniquement dans `/mnt/kyc_streamlit_app/requirements.txt` | Respecte la consigne de la capture |
| Ajouter `app.sh` à la racine globale du projet | `app.sh` fourni | Délègue au lanceur applicatif |
| `/opt/conda/envs/[MY VENV NAME]/bin/python -m streamlit run` | `KYC_PYTHON_BIN` puis `-m streamlit run app.py` | Le nom du venv n'est pas renseigné sur la photo |
| Exemple `streamlit_app.py` | `/mnt/kyc_streamlit_app/app.py` | Nom du fichier livré |
| `--theme.base dark` | `--theme.base=dark` | Même option |
| `--server.fileWatcherType none` | `--server.fileWatcherType=none` | Même option |
| `--browser.gatherUsageStats false` | `--browser.gatherUsageStats=false` | Télémétrie Streamlit désactivée |
| `--server.address 0.0.0.0` | `0.0.0.0` | Écoute sur toutes les interfaces du conteneur |
| `--server.port 8888` | `${PORT:-8888}` | 8888 par défaut, exactement comme l'exemple |

Non visibles : tier matériel exact, nom/version de l'environnement, nom du projet, chemin physique de sa racine, variable Domino officielle de port, options de partage, persistance des montages et intitulés des champs de publication. Ces éléments ne sont pas déduits de la photo. `PORT` est un paramètre de notre script, pas une variable que la capture prouverait être fournie par Domino.

**Recommandation :** sélectionner l'environnement réellement utilisé par le notebook et un tier avec un seul H100. Garder une seule instance de l'App et un accès Domino restreint aux opérateurs KYC autorisés. Ne pas activer de publication publique.

## Environnement

Utiliser Python ≥ 3.10 et le runtime CUDA/PyTorch/Transformers du notebook validé. `runtime_prerequisites.txt` est un inventaire, pas un fichier `pip install -r`. Le modèle FP8 est chargé avec `FineGrainedFP8Config(dequantize=True)` en BF16 ; le nom du checkpoint ne signifie pas que les poids restent en FP8 en mémoire. La vérification de placement interdit l'offload CPU/disque/méta comme dans le notebook.

Seule dépendance applicative déclarée à installer : `streamlit>=1.44,<2`, avec ses dépendances si absentes. Le minimum couvre notamment `st.fragment` et `download_button(on_click="ignore")`. Une version minimale n'est pas une certification de toute combinaison de packages : consigner les versions qui passent le smoke test dans votre image Domino.

Installation **dans l'environnement de build Domino**, depuis le miroir ou wheelhouse interne autorisé, avec le même interpréteur que le notebook. Exemple si votre administrateur a provisionné `/mnt/wheelhouse` :

```bash
"${KYC_PYTHON_BIN:-python}" -m pip install --no-index --find-links=/mnt/wheelhouse -r /mnt/kyc_streamlit_app/requirements.txt
```

Ce répertoire de wheels est un exemple à adapter, pas un chemin observé. Ne pas exécuter de mise à jour globale de torch/CUDA. Ne pas installer le package tiers `fitz` : le module `fitz` utilisé est fourni par **PyMuPDF**. Ne pas installer flash-attn pour cette conversion. Aucune installation n'est faite au lancement.

## Déploiement Domino, pas à pas

1. Ouvrir le projet Domino où le notebook fonctionne et démarrer un Workspace. Ouvrir son terminal JupyterLab/IDE.
2. Vérifier la racine réelle du projet avec `pwd` et, dans le panneau des fichiers, l'emplacement où les fichiers seront versionnés/publiés. Dans un projet à fichiers, elle peut être `/mnt`. Dans un projet Git, le checkout peut être ailleurs. La capture ne tranche pas. **Le `app.sh` doit appartenir à la racine du projet publiée**, même si le code applicatif est monté sous `/mnt`.
3. Créer les dossiers demandés :

```bash
mkdir -p /mnt/kyc_streamlit_app/scripts /mnt/kyc_streamlit_app/tests /mnt/kyc_streamlit_app/reference
mkdir -p /mnt/kyc_app/input /mnt/kyc_app/extracted
mkdir -p /mnt/output_qwen38/reports /mnt/output_qwen38/logs /mnt/output_qwen38/run_metadata /mnt/output_qwen38/datasets
```

4. Télécharger `kyc_streamlit_domino_v1.zip` et le téléverser dans `/mnt` via le navigateur de fichiers du Workspace (bouton d'upload JupyterLab, s'il est disponible). Le ZIP de livraison contient `app.sh` et `kyc_streamlit_app/`. Il ne contient ni modèle, ni CSV de référence, ni dossiers KYC réels.
5. Pour une **première installation**, décompresser sans écraser les fichiers déjà présents :

```bash
unzip -n /mnt/kyc_streamlit_domino_v1.zip -d /mnt
```

Si une ancienne installation existe, arrêter son App, la sauvegarder dans un répertoire de version, puis installer le nouveau dossier de code. Ne pas utiliser une extraction partielle mélangeant deux versions. Ne pas toucher aux sorties et checkpoints lors d'une mise à jour du code.
6. Si la racine publiée est `/mnt`, `/mnt/app.sh` est déjà au bon endroit. Sinon, copier **ce seul lanceur** vers la racine réelle du checkout du projet. Publier également les sources nécessaires pour reconstruire `/mnt/kyc_streamlit_app` dans le runtime App ou monter ce dossier. Un fichier uniquement présent dans votre Workspace n'est pas automatiquement présent dans le conteneur App. Faire vérifier le montage et le mécanisme de copie par l'administrateur si le projet Git et `/mnt` sont séparés.
7. Vérifier le CSV existant :

```bash
test -s /mnt/recertification_lissage_updated.csv
```

Colonnes attendues : `Id tiers`, `Date de naissance`, `Nom abrege tiers`, `Date expiration document`, `Numero de compte`. Le lecteur conserve les chaînes et zéros initiaux, accepte UTF-8 avec ou sans BOM, et détecte un séparateur compatible avec le schéma. Les dates de référence du matching restent au format `yyyy-mm-dd`, comme dans le notebook.
8. Vérifier le modèle local dans le runtime qui servira l'App :

```bash
test -f /domino/edv/modelhub/ModelHub-model-huggingface-Qwen/Qwen3.8-27B-FP8/main/config.json
```

Si le montage réel diffère, renseigner `KYC_MODEL_PATH` dans l'environnement App. Ne pas télécharger un autre modèle.
9. Identifier l'interpréteur du notebook en y exécutant `import sys; print(sys.executable)`. Mettre **ce chemin réel** dans `KYC_PYTHON_BIN` au niveau de la configuration d'environnement Domino utilisée par l'App. Le placeholder `[MY VENV NAME]` de la photo ne doit pas être copié littéralement. Un `export` dans votre terminal ne configure que ce terminal ; reporter aussi la variable dans l'environnement App.
10. Vérifier le tier H100 avec `nvidia-smi`. Le runtime App doit disposer du même accès GPU, du même dépôt modèle et des mêmes données que le Workspace. Les poids déquantifiés demandent plus de VRAM que leur taille FP8 sur disque. Ne pas lancer le notebook et l'App simultanément sur le même GPU.
11. Donner les permissions et vérifier :

```bash
chmod +x /mnt/app.sh
chmod +x /mnt/kyc_streamlit_app/scripts/run_app.sh
chmod +x /mnt/kyc_streamlit_app/scripts/validate_environment.sh
bash /mnt/kyc_streamlit_app/scripts/validate_environment.sh
```

Si le lanceur a été copié dans une autre racine de projet, appliquer aussi `chmod +x` à cette copie. Le validateur ne charge aucun poids.
12. Si Streamlit manque, ajouter uniquement les dépendances manquantes dans l'image de calcul selon la procédure interne, reconstruire l'environnement, puis répéter le validateur. Une installation interactive dans le Workspace peut ne pas être reproduite au déploiement.
13. Vérifier la persistance et les permissions. **`/mnt` n'est pas une garantie de persistance après redéploiement.** La documentation Domino indique que les changements d'une App ne sont pas automatiquement sauvegardés dans le dépôt du projet. Les répertoires `/mnt/kyc_app` et `/mnt/output_qwen38` doivent être adossés à un stockage Domino persistant autorisé et accessible à l'App. Vérifier aussi le support effectif de `flock` sur ce stockage.

```bash
findmnt -T /mnt/output_qwen38
findmnt -T /mnt/kyc_app
```

Faire un essai de fichier marqueur puis arrêt/redémarrage de l'App avant la recette. Si le stockage est éphémère, faire provisionner les montages persistants aux chemins demandés ; un simple `mkdir` ne suffit pas. La reprise sur disque fonctionne si ces données sont conservées. Aucun appel API de stockage externe n'est ajouté.
14. Faire le test de lancement manuel dans le terminal :

```bash
cd /mnt/kyc_streamlit_app
bash scripts/run_app.sh
```

Attendre l'annonce Streamlit. Dans un second terminal, si `curl` est disponible :

```bash
curl --fail http://127.0.0.1:8888/_stcore/health
```

Résultat attendu : `ok`. Utiliser l'URL de preview/proxy proposée par votre Workspace pour le test navigateur ; son format dépend du déploiement Domino. Ne pas inventer une URL publique. Après le smoke test manuel, arrêter ce serveur avec Ctrl+C avant de publier, pour libérer son verrou et la VRAM.
15. Ouvrir la rubrique Apps du projet et préparer la publication/version de l'App. Le nom exact des boutons dépend de votre version Domino et n'est pas visible sur la photo. Synchroniser les fichiers/committer les sources si nécessaire avant de publier cette version.
16. Utiliser le mécanisme `app.sh` à la racine montré sur la capture. Si l'interface demande une commande de lancement, utiliser `bash app.sh` depuis cette racine. Le lanceur délègue exactement à `bash /mnt/kyc_streamlit_app/scripts/run_app.sh`. Si elle demande la commande absolue et `/mnt` est la racine, utiliser `bash /mnt/app.sh`.
17. Sélectionner l'environnement validé et un tier avec **un H100**, **une instance** et aucun worker supplémentaire. Configurer le port attendu à `8888` si un champ existe ; sinon le script le fixe. Vérifier que `PORT` n'est pas défini avec une autre valeur par votre image. Le serveur écoute `0.0.0.0`. CORS et XSRF restent aux protections Streamlit par défaut ; ne pas les désactiver aveuglément pour résoudre un problème de proxy.
18. Publier/démarrer l'App, consulter les logs Domino et ouvrir son URL authentifiée. La page doit afficher « Application prête » sans charger les poids. Les droits Domino déterminent les personnes autorisées ; cette V1 est un outil partagé pour opérateurs, pas une plateforme multi-tenant cloisonnant des données client par utilisateur.
19. Charger `kyc_documents_test.zip`, vérifier son nom, sa taille et sa validation, puis cliquer **Lancer le traitement**. La progression est actualisée toutes les deux secondes pendant le travail, sans appels Streamlit depuis le thread. Un rafraîchissement conserve le suivi grâce au `run_id` dans l'URL. Un autre utilisateur ne peut pas démarrer un second travail.
20. À la fin, lire le résumé, vérifier puis télécharger le rapport :

```bash
test -s /mnt/output_qwen38/reports/kyc_recertification_control.xlsx
```

Le bouton lit les octets de ce fichier exact. Il exige un reçu réussi de la même exécution et du même ZIP, et vérifie le SHA-256 sous verrou. Un ancien résultat est refusé s'il n'est plus le rapport courant. Un bouton déjà rendu contient les octets de l'exécution explicitement affichée ; il ne se transforme pas silencieusement en rapport d'une autre archive.

## Comment placer chaque fichier

La table de la section suivante donne toutes les destinations. Option conseillée : upload du ZIP de livraison puis extraction, étapes 4–6. Tous les `.py`, `.sh` et fichiers auxiliaires sont alors placés ensemble.

**Option éditeur Workspace :** pour chaque ligne, créer les dossiers parents indiqués, créer un fichier portant exactement le nom de la colonne « Fichier », coller son contenu intégral depuis le guide complet, enregistrer en UTF-8 à sa destination. Les scripts utilisent des fins de ligne LF. Les `.py` ne nécessitent pas de bit exécutable ; ils sont importés ou lancés par Python. Pour les `.sh`, appliquer `chmod +x` comme ci-dessus. Les données de référence restent séparées du code.

**Option terminal :** utiliser un heredoc littéral pour chaque fichier, avec les contenus fournis dans les sections `FILE:` du guide. Exemple de création complète du lanceur, si `/mnt` est la racine projet :

```bash
cat > /mnt/app.sh <<'SH'
#!/usr/bin/env bash
set -euo pipefail
exec bash /mnt/kyc_streamlit_app/scripts/run_app.sh
SH
chmod +x /mnt/app.sh
```

Pour les autres fichiers, préférer le ZIP complet afin d'éviter une erreur de copier-coller dans les milliers de lignes du moteur validé. La capture ne prouve pas l'existence d'un éditeur web Domino particulier ; le terminal et le navigateur de fichiers du Workspace sont les mécanismes décrits ici.

| Fichier local | Destination exacte | Rôle | Exécutable |
|---|---|---|---|
| `app.sh` (racine du ZIP) | `/mnt/app.sh` si racine projet = `/mnt`, sinon racine publiée réelle | Lanceur exigé par la capture | Oui |
| `README.md` | `/mnt/kyc_streamlit_app/README.md` | Guide de déploiement et exploitation | Non |
| `VALIDATION.md` | `/mnt/kyc_streamlit_app/VALIDATION.md` | Vérifications réalisées et limites | Non |
| `app.py` | `/mnt/kyc_streamlit_app/app.py` | Interface Streamlit française, upload, suivi, téléchargement | Non |
| `archive_paths.py` | `/mnt/kyc_streamlit_app/archive_paths.py` | Validation des chemins ZIP copiée du notebook | Non |
| `config.py` | `/mnt/kyc_streamlit_app/config.py` | Chemins applicatifs et paramètres centralisés | Non |
| `exceptions.py` | `/mnt/kyc_streamlit_app/exceptions.py` | Erreurs applicatives françaises | Non |
| `file_manager.py` | `/mnt/kyc_streamlit_app/file_manager.py` | Validation ZIP/CRC, upload atomique et verrous | Non |
| `jobs.py` | `/mnt/kyc_streamlit_app/jobs.py` | Thread unique et métadonnées durables | Non |
| `logging_config.py` | `/mnt/kyc_streamlit_app/logging_config.py` | Journaux locaux avec rotation | Non |
| `model_manager.py` | `/mnt/kyc_streamlit_app/model_manager.py` | Singleton du moteur, contrôle H100 et protection du chargement | Non |
| `notebook_config.py` | `/mnt/kyc_streamlit_app/notebook_config.py` | Classes de configuration originales | Non |
| `pipeline.py` | `/mnt/kyc_streamlit_app/pipeline.py` | Orchestration des étapes originales et API synchrone | Non |
| `progress.py` | `/mnt/kyc_streamlit_app/progress.py` | Événements neutres de progression | Non |
| `reference/NOTEBOOK_MAPPING.md` | `/mnt/kyc_streamlit_app/reference/NOTEBOOK_MAPPING.md` | Table lisible des 230 correspondances | Non |
| `reference/function_mapping.json` | `/mnt/kyc_streamlit_app/reference/function_mapping.json` | Correspondance et empreinte AST de chaque définition | Non |
| `reference/source_provenance.json` | `/mnt/kyc_streamlit_app/reference/source_provenance.json` | Empreinte SHA-256 de la pièce jointe originale | Non |
| `reference/validated_notebook.ipynb` | `/mnt/kyc_streamlit_app/reference/validated_notebook.ipynb` | Sources du notebook de référence, sorties historiques retirées | Non |
| `reporting.py` | `/mnt/kyc_streamlit_app/reporting.py` | Résultat structuré, contrôle de fraîcheur et téléchargement exact | Non |
| `requirements.txt` | `/mnt/kyc_streamlit_app/requirements.txt` | Seule dépendance applicative à installer : Streamlit | Non |
| `runtime_prerequisites.txt` | `/mnt/kyc_streamlit_app/runtime_prerequisites.txt` | Inventaire du runtime existant, ne pas installer en bloc | Non |
| `scripts/compare_regression.py` | `/mnt/kyc_streamlit_app/scripts/compare_regression.py` | Comparaison des sorties notebook/App sans GPU | Non |
| `scripts/run_app.sh` | `/mnt/kyc_streamlit_app/scripts/run_app.sh` | Lanceur Streamlit, port 8888 et verrou serveur | Oui |
| `scripts/validate_environment.py` | `/mnt/kyc_streamlit_app/scripts/validate_environment.py` | Contrôles sans charger les poids | Non |
| `scripts/validate_environment.sh` | `/mnt/kyc_streamlit_app/scripts/validate_environment.sh` | Lance le validateur avec le Python configuré | Oui |
| `tests/test_application.py` | `/mnt/kyc_streamlit_app/tests/test_application.py` | 19 tests applicatifs et 80 assertions source | Non |
| `validated_core.py` | `/mnt/kyc_streamlit_app/validated_core.py` | 230 définitions originales avec notebook_config ; fonctions métier et inférence | Non |


## Paramètres configurables

| Paramètre | Défaut | Usage |
|---|---|---|
| `KYC_PYTHON_BIN` | `python` du PATH | Chemin réel du Python validé ; définir pour l'App |
| `KYC_MODEL_PATH` | Dépôt Qwen3.8 indiqué plus haut | Montage local du même checkpoint |
| `PORT` | `8888` | Port du script ; garder 8888 dans la configuration montrée |
| `KYC_MAX_UPLOAD_MB` | `1024` | Taille limite UI + écriture ; mémoire hôte adaptée |
| `KYC_ZIP_CUSTOMER_ROOT` | non défini | Préfixe interne explicite, si le ZIP contient un dossier enveloppe différent de `kyc_documents_test/` |

`config.py` centralise les chemins contractuels. `notebook_config.py` conserve toutes les valeurs originales. Les modifications métier ne se font pas depuis des widgets. Ne pas définir `KYC_CORE_CPU_TEST=1` dans l'environnement App : ce mode est réservé aux tests livrés et le lanceur le refuse.

Le ZIP doit contenir directement des dossiers clients, ou un dossier enveloppe `kyc_documents_test/` contenant ces dossiers. Pour une autre enveloppe, configurer `KYC_ZIP_CUSTOMER_ROOT`. Les cinq cibles et l'alias carton sont ceux du notebook. La sélection famille + taille peut ignorer un doublon de même taille même si son contenu diffère ; l'audit `SAME_SIZE_DIFFERENT_HASH` reste inchangé.

## Reprise, état et verrou

Même ZIP + même configuration modèle/extraction : les caches réutilisables sont retrouvés. Les pages défaillantes peuvent être retentées selon le comportement original. La reprise demande de cliquer à nouveau après un arrêt de processus ; elle n'est pas déclenchée automatiquement à l'ouverture du navigateur. Le thread continue après une déconnexion navigateur, mais pas après l'arrêt du conteneur. Le processus suivant détecte un état `RUNNING` dont le verrou n'est plus pris et le marque `INTERRUPTED`.

Le verrou de job contient `run_id`, PID, hostname et date de début. Il repose sur `flock` et un verrou mémoire commun aux threads. Le système libère `flock` après un crash. **Ne jamais supprimer le fichier de verrou ni casser un verrou sur la seule base de son âge** : supprimer un inode encore verrouillé autoriserait deux écrivains. Un fichier de verrou restant après crash est normal et ne bloque pas les traitements s'il n'est plus verrouillé. Le verrou serveur bloque un second lanceur tant que le premier serveur est vivant, même sans job actif.

Les métadonnées et reçus sont atomiques. La protection de fraîcheur s'appuie sur l'absence du rapport courant au démarrage (ancien fichier archivé), sa nouvelle publication atomique, sa date avec tolérance de granularité de filesystem, l'audit population/feuilles Excel et son SHA-256.

Les sauvegardes et checkpoints peuvent contenir des données KYC intégrales. La conservation, sauvegarde et suppression doivent suivre votre politique interne. Aucun nettoyage automatique destructif n'est livré.

## Smoke test et non-régression

Tests locaux sans GPU :

```bash
cd /mnt/kyc_streamlit_app
"${KYC_PYTHON_BIN:-python}" -m unittest discover -s tests -v
```

Le mode CPU est activé seulement dans le processus de test. Vérifications : identité des 230 définitions, 80 assertions de similarité originales, ZIP dangereux/CRC/collisions, verrous entre processus et après crash, sortie Excel avec PDF invalide, deuxième exécution/reprise, rejet d'un rapport ancien ou modifié, récupération de métadonnées interrompues. Ces tests ne sont pas des inférences Qwen.

Recette réelle : valider l'environnement ; lancer Streamlit ; ouvrir l'URL App ; téléverser le ZIP ; lancer ; observer le premier chargement seulement ; suivre les phases ; contrôler tous les clients ; contrôler le fichier Excel ; télécharger et comparer le SHA-256 ; relancer le même ZIP et vérifier la baisse des appels Qwen ; rafraîchir en cours de traitement ; tester un second navigateur ; redémarrer l'App et vérifier les checkpoints persistants.

Pour un test CLI avant publication, arrêter l'App et les kernels utilisant ce H100 :

```bash
cd /mnt/kyc_streamlit_app
flock --nonblock --no-fork /mnt/output_qwen38/.kyc_server.lock \
  "${KYC_PYTHON_BIN:-python}" pipeline.py --zip /mnt/kyc_app/input/kyc_documents_test.zip
```

Pour prouver l'équivalence métier, exécuter le notebook original en `FULL_PIPELINE` avec le même ZIP, CSV, versions, paramètres et date de contrôle, mais **un autre OUTPUT_DIR** (par exemple `/mnt/kyc_notebook_regression`). Ne jamais écrire simultanément dans les sorties de l'App. Conserver son Excel. Exécuter ensuite l'App et récupérer le SHA du ZIP :

```bash
sha256sum /mnt/kyc_app/input/kyc_documents_test.zip
```

Lancer `scripts/compare_regression.py` avec les chemins réels :

```bash
"${KYC_PYTHON_BIN:-python}" scripts/compare_regression.py \
  --notebook-output /mnt/kyc_notebook_regression \
  --app-output /mnt/output_qwen38/datasets/EMPREINTE_SHA256_REELLE \
  --notebook-excel /mnt/kyc_notebook_regression/reports/kyc_recertification_control.xlsx \
  --app-excel /mnt/output_qwen38/reports/kyc_recertification_control.xlsx
```

`EMPREINTE_SHA256_REELLE` est à remplacer par le résultat de `sha256sum`, pas un dossier livré. La comparaison contrôle les populations, textes RAW/statuts par page, champs structurés, formes FATCA, scores de similarité, décisions documentaires/client et les feuilles métier Excel. Les timestamps, temps de calcul, chemins absolus de sorties et la feuille de performance ne doivent pas être considérés comme des résultats métier identiques. Les noms de feuilles sont contrôlés ; les lignes et valeurs des principales feuilles métier aussi. Le script ne montre pas les valeurs personnelles divergentes, seulement des empreintes de chemins de divergence. Un écart de génération ou de probabilité entre deux inférences réelles doit être examiné ; les kernels GPU ne garantissent pas une identité bit à bit inter-environnements.

## Dépannage

| Symptôme | Cause probable | Diagnostic | Correction |
|---|---|---|---|
| Streamlit introuvable | Mauvais venv ou package absent | `"${KYC_PYTHON_BIN:-python}" -m pip show streamlit` | Utiliser le Python du notebook ; installer Streamlit dans l'image App via le miroir interne |
| L'URL ne répond pas | Port/proxy incorrect | `curl --fail http://127.0.0.1:8888/_stcore/health` ; logs Domino | Garder `PORT=8888`, `0.0.0.0` ; vérifier proxy et montage du code |
| Chargement de page bloqué avant analyse | Dossier/CSV non accessible | `bash scripts/validate_environment.sh --paths-only` | Monter les chemins dans le runtime App, corriger les droits |
| Modèle absent | Dépôt non monté dans l'App | `test -f "$KYC_MODEL_PATH/config.json"` si variable définie | Monter le dépôt local ; sinon utiliser le chemin par défaut documenté |
| CUDA indisponible | Tier CPU, pilote ou runtime incorrect | `nvidia-smi` puis validateur | Reprendre l'environnement GPU du notebook |
| H100 indisponible | Mauvais tier ou plusieurs GPU visibles | Validateur ; `nvidia-smi -L` | Sélectionner un seul H100 visible |
| CSV absent | Mauvais nom, ancien `recertication` ou mauvais montage | `ls -l /mnt/recertification_lissage_updated.csv` | Utiliser exactement le nom avec `fi` et le rendre accessible à l'App |
| Permission denied sous `/mnt` | UID/montage non inscriptible | `id` ; `ls -ld /mnt/kyc_app /mnt/output_qwen38` | Faire attribuer les permissions nécessaires ; ne pas rendre le stockage public |
| ZIP invalide | CRC, traversée, lien, chiffrement, doublon de chemin ou absence de dossiers | Message de validation ; `unzip -t` pour CRC | Recréer l'archive avec dossiers clients et chemins relatifs sûrs |
| Mauvais nombre de clients | Dossier enveloppe non reconnu | Examiner l'arborescence avec `unzip -l` | Définir le bon `KYC_ZIP_CUSTOMER_ROOT` avant relance |
| CUDA out of memory / offload | VRAM occupée ou insuffisante pour BF16 | `nvidia-smi`, logs du modèle | Arrêter les autres kernels, choisir la mémoire H100 suffisante, redémarrer l'App ; ne pas changer la quantification validée à l'aveugle |
| Modèle chargé à chaque visite | Plusieurs serveurs ou redémarrages du conteneur | Logs `MODEL LOADING`, PID et événements Domino | Une instance, lancer via `app.sh`, ne pas vider le cache du contrôleur en cours de job |
| Verrou apparemment périmé | Fichier JSON ancien ou processus toujours vivant | `python -c 'from config import SETTINGS; from file_manager import is_busy; print(is_busy(SETTINGS.lock_path))'` depuis le dossier app | Si False, relancer normalement ; si True, identifier le processus via Domino. Ne pas supprimer le lock |
| Rapport non généré | Benchmark, données, écriture ou export échoués | `tail -n 100 /mnt/output_qwen38/logs/application.log` ; métadonnées du run | Corriger la cause puis reprendre le même ZIP ; ancien Excel jamais offert comme résultat nouveau |
| Rapport non courant | Une autre archive a été traitée | Vérifier `run_metadata/current_report.json` | Relancer votre archive pour republier le rapport ; anciens rapports conservés pour audit |
| Rafraîchissement navigateur | Session UI réinitialisée | Conserver l'URL avec `?run=...` | Revenir à cette URL ; le thread reste actif tant que le serveur fonctionne |
| Pipeline interrompu | Arrêt App/conteneur, erreur fatale ou stockage | Métadonnées et logs, disponibilité des caches | Après correction, charger le même ZIP et cliquer à nouveau |
| Checkpoints disparus après redéploiement | Stockage App éphémère | `findmnt -T /mnt/output_qwen38` et essai marqueur après redémarrage | Provisionner un montage durable ; `/mnt` seul ne garantit pas la persistance |
| Changement du code non pris en compte | File watcher désactivé, comme la capture | Version et fichiers publiés | Arrêter puis republier/redémarrer l'App après modification |

## Sécurité et limites d'exploitation

Aucun service IA externe, aucune API externe et aucune télémétrie applicative n'est ajouté. Hugging Face est en mode local/offline ; Streamlit a `gatherUsageStats=false`. Les protections CORS/XSRF restent activées par défaut. Les preuves OCR, images, CSV et Excel restent dans les chemins Domino, hormis le téléchargement explicitement demandé par l'opérateur dans son navigateur autorisé.

Le lancement utilise `umask 077` pour les nouveaux fichiers. Les données restent accessibles au compte de service de l'App et aux administrateurs autorisés ; cette V1 ne fait pas d'isolation fine entre plusieurs opérateurs ayant accès au même outil. Ne pas rendre le répertoire de sorties statiquement public. Le CPU/RAM/disque doit être dimensionné pour le ZIP téléversé, le modèle déquantifié, les images persistées et le classeur Excel. Le moteur original construit certaines tables de rapport en mémoire ; cette conversion ne prétend pas supprimer cette limite.

Les verrous supposent un filesystem fournissant des verrous POSIX cohérents et une seule instance déployée. Ne pas lancer deux replicas sur des volumes disjoints avec le même GPU. La persistance à travers remplacement de conteneur dépend du montage Domino ; elle fait partie de la recette obligatoire.

## Références documentaires

La capture jointe est l'autorité pour ses options de lancement. Les documents officiels consultés pour compléter les comportements de plateforme sont :

- [Domino — publication Streamlit](https://docs.domino.ai/6.3/platform-capabilities/features/apps/common-app-frameworks/streamlit-app-llm) : mécanisme `app.sh` et port 8888. Les exemples de modèles distants de cette page ne sont pas utilisés.
- [Domino — persistance des données App](https://docs.domino.ai/6.3/platform-capabilities/features/apps/persist-data-in-apps) : les sorties d'une App ne sont pas automatiquement sauvegardées au dépôt ; stockage durable explicite et protection des écritures concurrentes.
- [Streamlit — cache_resource](https://docs.streamlit.io/develop/api-reference/caching-and-state/st.cache_resource) : ressources partagées entre sessions, dont l'accès doit être protégé.
- [Streamlit 1.44 — download_button](https://docs.streamlit.io/1.44.0/develop/api-reference/widgets/st.download_button) : octets gardés en mémoire pour le téléchargement et option `on_click="ignore"`.
- [Streamlit — file_uploader](https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader) : limite d'upload configurable ; le filtre de type ne remplace pas une validation de contenu.
