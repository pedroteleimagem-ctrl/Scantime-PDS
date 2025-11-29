# ScanTime V5 - Guide d'utilisation
Dernière mise à jour : 28/09/2025

## 1. Démarrage rapide
- `ScanTime.exe` : double-cliquez pour ouvrir la version prête à l'emploi, aucune installation de Python n'est nécessaire.
- Mode développeur : ouvrez `Full_GUI.py` dans VS Code puis lancez-le avec `F5` ou la commande `python Full_GUI.py`.

## 2. Vue d'ensemble de l'interface
- Bannière supérieure : le titre et le libellé `Semaine du ...`. Double-cliquez sur ce libellé pour saisir la date de la semaine.
- Onglets : chaque semaine possède son onglet ; le bandeau juste en dessous contient `Ajouter semaine` et `Supprimer semaine`.
- Ruban d'actions : à droite du planning se trouvent les boutons principaux décrits à la section 4.
- Planning central : tableau matin / après-midi par poste pour chaque jour.
- Tableau de décompte (ShiftCount) : colonne à droite affichant le nombre de vacations par personne, mise en surbrillance lors d'une sélection.
- Tableau des contraintes : zone inférieure regroupant quotas, absences, PDS et commentaires utilisés par l'assignation automatique.

### Codes couleur
| Situation | Affichage |
| --- | --- |
| Créneau libre | Fond blanc |
| Créneau attribué | Bleu clair |
| Créneau fermé | Gris foncé |
| Conflit ou doublon | Rouge |
| Vérification - case vide | Bleu |
| Vérification - doublon | Rouge |
| Case contenant la personne sélectionnée | Jaune |
| Remplaçants proposés | Vert clair |
| Tableau de décompte (ligne active / remplaçants) | Jaune / Vert clair |

## 3. Gérer les semaines
- `Ajouter semaine` duplique la semaine active ; si plusieurs semaines existent vous choisissez celle à recopier.
- `Supprimer semaine` retire l'onglet courant (il doit toujours rester au moins une semaine).
- Renommez une semaine en double-cliquant sur le libellé `Semaine du ...` du planning.
- Les imports d'absences peuvent ajouter ou supprimer automatiquement des semaines pour s'aligner sur le mois choisi.

## 4. Ruban d'actions et gestion des postes
- `Vérification` : cochez pour mettre en évidence les cases vides (bleu) et les incohérences (rouge). Décochez pour retrouver les couleurs normales.
- `Annuler assignation` : restaure le planning enregistré juste avant la dernière assignation automatique (`Ctrl+Shift+Z` effectue la même action).
- `Fermer postes` : sélectionnez un ou plusieurs postes à rendre indisponibles (toutes les cases deviennent grises).
- `Effacer tout` : vide chaque créneau de la semaine active (annulable via `Ctrl+Z`).
- `Assignation` : lance l'affectation automatique en respectant les quotas, absences et options (voir section 8).
- `Ajouter poste` : insère un poste supplémentaire en bas du planning.
- Bouton `+` en bout de ligne poste : ouvre une fenêtre pour dupliquer ou supprimer ce poste.
- Clic droit sur le nom d'un poste : change sa couleur. Double-clic : renomme le poste (le nom doit être unique).
- Menu `File > Effacer layout` : remet la semaine active sur la mise en page par défaut (5 postes avec horaires standards) après confirmation.

## 5. Remplir le planning
- Saisie manuelle : cliquez dans une case puis tapez les initiales (les ajustements sont pris en compte instantanément dans le tableau de décompte).
- Double-clic sur une case : ferme ou rouvre le créneau (fond gris = indisponible).
- Clic droit sur un horaire (ex. `08h-13h`) : modifiez le texte et, si besoin, excluez ce créneau du calcul du décompte.
- `Ctrl` + molette : zoom avant/arrière sur le planning.
- `Ctrl+clic` sur une cellule puis sur une autre : copie la valeur (l'original reste inchangé).
- `Shift+clic` sur deux cellules : échange les valeurs entre les deux cases.
- `Ctrl+Z` annule la dernière modification manuelle ; la pile d'annulation conserve aussi les échanges et les effacements globaux.
- `Tab` et `Shift+Tab` passent respectivement au créneau suivant ou précédent.
- Lorsque vous sélectionnez une cellule, toutes les occurrences de la même personne sont surlignées en jaune et les personnes disponibles pour ce créneau sont surlignées en vert clair.
- Les créneaux fermés restent visibles mais bloqués pour l'assignation automatique.

## 6. Tableau de décompte (ShiftCount)
- Comptabilise pour chaque personne les vacations du matin (`M`), de l'après-midi (`A`) et le total hebdomadaire.
- Se met à jour dès qu'un créneau est modifié ou qu'une absence est ajoutée.
- Lorsqu'une cellule du planning est sélectionnée, la personne correspondante apparaît en jaune et les remplaçants éligibles en vert.
- Utilisez ces surbrillances pour équilibrer rapidement les vacations avant l'export.

## 7. Tableau des contraintes
- Cliquez sur `Ajouter` pour créer une nouvelle ligne ; `X` supprime la ligne correspondante.
- Colonnes :
  - `Initiales` : identifiant affiché dans le planning.
  - `Vacations/semaine` : quota maximum à respecter.
  - `Postes préférés (2 max)` : bouton de sélection multiple (les choix sont prioritaires lors de l'assignation).
  - `Postes non-assurés` : bouton qui liste tous les postes ; cochez ceux que la personne ne peut pas occuper.
  - Colonnes `Lundi` à `Dimanche` : bouton qui cycle entre Aucun / Matin / AP Midi / Journée, plus une case à cocher `PDS` pour signaler une garde la nuit précédente (la case devient rouge lorsqu'elle est active).
  - `Commentaire` : note libre visible par l'équipe.
- Les boutons conservent les sélections existantes ; un simple clic rouvre la liste pour ajuster.
- Fermez temporairement le tableau des contraintes via le bouton `X` dans son en-tête si vous avez besoin de plus d'espace.

## 8. Assignation automatique
- Préparez le tableau des contraintes : quotas renseignés, absences/PDS cochés, postes non-assurés et préférés à jour.
- Configurez les options via `Setup` :
  - `Empêcher même poste toute la journée` : interdit d'attribuer matin + après-midi au même poste.
  - `Activer limitation d'affectation` et `Configurer nombre maximum` : plafonds par poste dans la semaine.
  - `Activer repos de sécurité` : bloque automatiquement le lendemain matin après une garde (`PDS`) cochée.
- Cliquez sur `Assignation` : les cases vides sont remplies en respectant toutes les contraintes, les préférences (2 maximum par semaine) et en évitant les doublons dans un même créneau.
- Utilisez ensuite `Vérification` pour repérer les éventuels créneaux restés vides ou les conflits, puis ajustez manuellement si nécessaire.

## 9. Imports
### 9.1 Importer une mise en page (`Imports > Import Layout`)
- Choisissez un fichier `.pkl` existant : seuls les postes, couleurs, horaires et disponibilités sont repris.
- Les initiales et les contraintes de votre session actuelle restent inchangées.

### 9.2 Importer des absences Excel (`Imports > Import Absences`)
- Sélectionnez le fichier Excel (format `mois x jours x personnes`).
- Choisissez le mois/onglet puis, si besoin, mappez les semaines de votre planning avec celles du document.
- Le programme propose automatiquement les correspondances de noms ; validez chaque suggestion incertaine ou ignorez-la.
- Couleurs interprétées : jaune = absence journée, rouge = repos de garde journée, vert = formation (Matin/AP/Journée), violet = astreinte (AP), autres couleurs ignorées.
- Les absences remplissent le tableau des contraintes, les jours fermés sont gris et les onglets/semaine sont renommés selon les dates.

### 9.3 Importer des conflits depuis un planning (`Imports > Import Conflits (.pkl)`)
- Choisissez un fichier `.pkl` d'un autre planning (par exemple celui des internes).
- Associez chaque semaine de votre planning avec la semaine correspondante du fichier.
- Les demi-journées occupées dans le planning importé deviennent des absences dans votre tableau de contraintes (Matin, AP ou Journée) sans ajouter de nouvelles lignes.

### 9.4 Vérifier des conflits inter-plannings (`Imports > Vérifier Conflits inter-plannings (.pkl)`)
- Sélectionnez un fichier `.pkl` à comparer.
- Carte semaine par semaine ; le rapport liste les chevauchements (même personne, même jour, même créneau).
- Les cellules concernées reçoivent un petit badge visuel pour faciliter les corrections, sans modifier vos couleurs.

## 10. Exports
- `Export > Export to Excel` : crée un classeur avec une feuille par semaine, le planning coloré, les absences, le tableau de décompte, les statistiques individuelles (par poste exact, double vacations, présence scanner) et un graphique circulaire par poste.
- `Export > Export combiné (.pkl)` : demande un autre fichier `.pkl` puis ajoute, pour chaque créneau, une deuxième ligne avec les initiales de ce planning (utile pour superposer résidents et internes). L'export inclut également une ligne d'absences pour le second planning.

## 11. Sauvegardes
- `File > Enregistrer` sauvegarde dans le fichier courant, `File > Enregistrer sous` permet de choisir un nouveau fichier `.pkl`.
- Nom de fichier conseillé par semaine : `Planning_Semaine_DD-MM-YYYY.pkl` pour un suivi clair.
- Sauvegarde automatique : toutes les 3 minutes et avant les actions importantes, un fichier `sauvegarde_auto.pkl` est écrit dans votre dossier utilisateur.
- `File > Localiser sauvegarde automatique` ouvre l'emplacement de ce fichier pour récupérer rapidement un travail en cours.

## 12. Raccourcis et gestes utiles
- `Ctrl+Z` : annuler la dernière saisie ; `Ctrl+Shift+Z` : annuler la dernière assignation automatique.
- `Ctrl+clic` : copier une initiale ; `Shift+clic` : échanger deux créneaux.
- `Ctrl` + molette : zoomer ou dézoomer.
- `Tab` / `Shift+Tab` : naviguer entre les cellules.
- Double-clic sur une cellule : ouvrir/fermer le créneau.
- Clic droit sur un horaire : modifier le texte et/ou exclure du décompte.

## 13. Bonnes pratiques
- Avant de lancer l'assignation, enregistrez et activez `Vérification` pour repérer les points à corriger.
- Utilisez la colonne `Commentaire` pour documenter les demandes spécifiques (formations, souhaits de repos, etc.).
- Après chaque modification importante, exportez ou sauvegardez rapidement pour conserver un historique.
- En fin de préparation, activez la vérification, parcourez chaque semaine et vérifiez le tableau de décompte pour équilibrer les charges avant l'impression ou l'envoi.
