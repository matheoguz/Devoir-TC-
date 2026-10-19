# Devoir-TC-

Devoirs et dossiers du BUT Techniques de commercialisation de Matheo. Écrire en français.
Fiche du projet dans le cerveau : `../Repos-/cerveau/Projets/Études BUT TC.md`.


## Le cerveau de Matheo (à faire dans chaque session)
Tout ce qu'on fait doit finir dans **le Cerveau** : coffre privé `matheoguz/Repos-` (dossier `cerveau/`) + page privée https://claude.ai/artifact/UPP4jo5X6LqiWXkCpd3mme (outil `ArtifactData`).
- **Au début** : lire `../Repos-/cerveau/MEMOIRE.md` (le hook `rappel` le fait si `Repos-` est dans la session), puis les collections `taches` et `idees` de la page.
- **Pendant** : le projet en cours a sa fiche dans `../Repos-/cerveau/Projets/` (frontmatter `statut`, `prochaine`, `chemins`). Une décision → `Décisions.md` + `MEMOIRE.md`.
- **Commandes qui évitent d'oublier** : `/debut` en ouvrant une session, `/fin` avant de dire « c'est fait », `/idee` dès que Matheo lance une idée.
- **À la fin de chaque tâche** (c'est ce que fait `/fin`) :
  1. mettre à jour la fiche du projet (statut, prochaine étape) ; nouveau projet = nouvelle fiche, avec ses `chemins` ;
  2. mettre à jour **le tableau de bord du projet** dans la page : `dashboards/<id-de-la-fiche>` = `{liens: [{label, url}], kpis: [{label, valeur, detail, icone}], sections: [{titre, contenu (markdown)}], maj}`. Nouveau projet (un site, un devoir, une app…) = nouvelle fiche **et** son tableau de bord, tout de suite. Ne mettre que des faits (pas de chiffres inventés) ;
  3. cocher les tâches faites (`taches/<id>` : `statut: "fait"`, `fait_le`) et créer celles qui restent (`{titre, projet, echeance, statut: "a_faire", date, source: "Claude"}`) ;
  4. si le hook du journal ne tourne pas (pas de `Journal/<aujourd'hui>.md`), le lancer à la main :
     `echo '{"session_id":"<id>","transcript_path":"<~/.claude/projects/…/<id>.jsonl>","cwd":"'$PWD'"}' | node ../Repos-/cerveau/.outils/journal.mjs` ;
  5. synchroniser la page : `node ../Repos-/cerveau/.outils/vers-cerveau.mjs <scratchpad>/vault`, `ArtifactData list` pour les `version`, puis `ArtifactData batch` avec les fichiers du `manifeste.json` (50 au plus par batch, `if_version` pour les documents existants).
- Brancher un nouveau dépôt : `bash ../Repos-/cerveau/.outils/brancher.sh <dossier-du-dépôt>`.
