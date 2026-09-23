# Prompt mensuel — Analyse des fiches d'intervention SAV

À copier-coller tel quel dans une session Claude, chaque mois, avec en pièces jointes :
- le zip des fiches d'intervention scannées du mois (PDF)
- les 3 images de signatures de référence (voir plus bas)

---

## Prompt à donner à Claude

```
Voici un zip contenant des fiches d'intervention SAV remplies à la main par des
techniciens, à analyser pour le mois de [MOIS ANNÉE — ex: Septembre 2026].

Le but : pour chaque fiche, déterminer la date, le technicien, le nombre d'heures
travaillées en intervention, puis produire un total par équipe :
- EBC = Yann BAREGE + Christophe BIZEUL + Joël DANVIN
- SAV = Daniel GRONDIN + Nicolas ALBANY

Je joins aussi 3 images de signatures de référence :
- Image 1 = signature de Daniel GRONDIN
- Image 2 = signature de Nicolas ALBANY
- Image 3 = signature de Yann BAREGE
(Christophe BIZEUL est le seul technicien qui signe avec son prénom écrit en toutes
lettres — pas besoin d'image de référence pour lui.)

### Ce qu'il faut lire sur chaque fiche

1. **Documents à analyser** : uniquement les bons d'intervention / installation.
   Ignore les devis et les bons de livraison qui pourraient traîner dans le zip.
   Attention : le champ imprimé en haut de la fiche "DEVIS N°" est juste un numéro
   de référence interne du bon d'intervention — ce n'est PAS un devis à ignorer,
   c'est le document à lire normalement.

2. **Technicien** : lire le champ "Technicien". S'il y a deux noms, ce sont
   forcément un des deux duos suivants (jamais un mélange d'équipe) :
   - Christophe BIZEUL + Yann BAREGE (EBC)
   - Nicolas ALBANY + Daniel GRONDIN (SAV)
   Si le champ est vide ou illisible, identifie le technicien via la signature en
   bas à gauche de la fiche, en comparant aux 3 images de référence jointes.

3. **Heures travaillées** : cherche une mention "MO = XXh" (ou similaire),
   généralement écrite en gris au milieu de la fiche. Si elle est présente,
   utilise directement cette valeur. Sinon, calcule la différence entre l'heure
   de départ et l'heure d'arrivée indiquées sur la fiche.

4. **Date** : généralement en bas à droite, près de la signature du client
   ("le JJ/MM/AAAA"). Si elle est illisible, utilise la date de scan encodée dans
   le nom du fichier PDF (format docNNNNNAAAAMMJJhhmmss — les 8 chiffres après le
   numéro de doc sont AAAAMMJJ) pour déterminer le mois, et note-le explicitement
   ("date déduite du nom de fichier").

5. **Fichiers PDF multi-pages** : vérifie bien chaque page, un même fichier peut
   contenir plusieurs bons d'intervention distincts.

### Règles spéciales (validées, à appliquer telles quelles)

- **Deux bons liés entre eux** dans un même fichier (mention manuscrite du type
  "Va avec BI ####" / "Suit BI ####") : ce sont **2 interventions distinctes** à
  compter séparément (ne pas les fusionner en une seule). Si les heures ne sont
  écrites que sur UN SEUL des deux bons (total commun aux deux), n'attribue ces
  heures qu'à ce bon-là — n'invente pas de valeur pour l'autre et ne double pas
  les heures.
- **Un seul bon signé par deux techniciens de la même équipe** (et non deux bons
  liés) : ça compte pour **une seule intervention partagée** pour cette équipe
  (pas une intervention par technicien), avec les heures comptées une seule fois.
- **Doublons de scan** : si une fiche identique apparaît deux fois dans le zip
  (même scan), ne la compte qu'une seule fois.

### Format de sortie attendu

Un fichier Excel nommé `Productivite_SAV_<Mois><Année>.xlsx`
(ex : `Productivite_SAV_Septembre2026.xlsx`), avec 3 feuilles :

1. **Synthèse** — une ligne par équipe :
   `Equipe | Nb intervention | Nb heures en intervention pour le mois`

2. **Detail_<Mois>** (ex : `Detail_Septembre`) — une ligne par intervention comptée :
   `Fichier | Page | Date signature | Technicien | Equipe | Heures | Méthode | Notes`
   (Méthode = MO_ecrit / calcul_horaire / déduit_filename ; Notes = tout élément
   incertain à vérifier)

3. **A_verifier** — les cas limites méritant une relecture humaine (lecture
   incertaine, doublon écarté, technicien identifié seulement par signature/code
   sans nom écrit, etc.), avec le nom du fichier et le numéro de page pour
   pouvoir vérifier facilement.

Ne calcule pas les heures travaillées théoriques ni la productivité — ça,
c'est fait automatiquement par l'application en aval à partir de ce fichier.
```

---

## Notes pour l'opérateur

- Les 3 images de signatures de référence (Daniel, Nicolas, Yann) sont à joindre
  au message avec le zip du mois — sans elles, Claude ne peut pas identifier un
  technicien dont le nom n'est pas écrit sur la fiche.
- Le fichier de sortie (`Productivite_SAV_<Mois><Année>.xlsx`) est ensuite à
  uploader directement dans l'onglet SAV de l'application Bilan Mensuel, qui
  s'occupe du reste (calcul des heures travaillées théoriques à partir des
  absences, du nombre de jours ouvrés du mois, et de la productivité finale).
