# Validation de livraison — 26 octobre 2026

## Vérifié dans l'environnement de création

- Lecture de toutes les cellules de code du notebook fourni et de la capture Domino.
- Comparaison AST : **230 fonctions/classes originales identiques**, configuration incluse.
- **80 assertions de similarité** de la cellule 89 originale : PASS.
- Suite applicative **19 tests unittest** : PASS (le test des 80 assertions est inclus parmi ces 19).
- Compilation Python de tous les modules : PASS.
- Syntaxe Bash des trois lanceurs/scripts : PASS.
- Parcours réel des fonctions d'inventaire, extraction sûre, passes RAW/structure, matching et Excel sur une archive synthétique avec un PDF invalide et des documents manquants ; deux clients, une référence absente, zéro appel Qwen : PASS.
- Seconde exécution avec réutilisation de l'extraction et des checkpoints applicables : PASS.
- Rejet de chemins dangereux, liens symboliques, doublons, collisions fichier/dossier, ZIP vide et CRC invalide : PASS.
- Exclusion de deux threads/processus concurrents et libération après arrêt brutal : PASS.
- Rejet de rapport ancien, modifié et provenant d'un traitement échoué : PASS.
- Reprise d'un état RUNNING sans verrou actif vers INTERRUPTED : PASS.

## Non exécuté ici

- Chargement des poids Qwen3.8-27B-FP8 : dépôt local Domino non accessible.
- CUDA/H100, performances et qualité OCR : matériel non disponible.
- Serveur et interface Streamlit dans un navigateur : Streamlit non installé dans cet environnement de création.
- Publication Domino, proxy, authentification, montages persistants et leur comportement après remplacement du conteneur : accès non fourni.
- Comparaison métier sur le véritable `kyc_documents_test.zip` : archive et CSV réels non joints.

Le code est livré pour la recette Domino décrite dans README.md. La comparaison syntaxique des fonctions ne remplace pas cette recette d'intégration. Le scénario synthétique exerce le traitement des absences/erreurs et la génération du rapport ; il ne simule pas une précision OCR réelle ni une validation du GPU.

## Adaptations surveillées

Le mode devient FULL_PIPELINE, les chemins sont centralisés, les caches sont isolés par ZIP, les vues JSON antérieures sont conservées dans projection_history, et le matching Python est recalculé après chaque reconstruction de documents. Les fonctions métier et leurs paramètres de décision restent identiques. `scripts/compare_regression.py` permet de comparer les véritables résultats du notebook et de l'application.
