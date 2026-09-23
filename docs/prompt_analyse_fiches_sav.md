# Prompt mensuel — Analyse des fiches d'intervention SAV

À copier-coller tel quel dans une session Claude, chaque mois, avec en pièces jointes :
- le zip des fiches d'intervention scannées du mois (PDF)
- les 3 images de signatures de référence (voir plus bas)

---

## Prompt à donner à Claude

```
Voici un zip contenant des fiches d'intervention SAV remplies à la main par des
techniciens, à analyser pour le mois de [MOIS ANNÉE — ex: Septembre 2026].

Le but : pour chaque fiche, déterminer le n° de BI, la date, le technicien, le
nombre d'heures travaillées en intervention, puis produire un total par équipe :
- EBC = Yann BAREGE + Christophe BIZEUL + Joël DANVIN
- SAV = Daniel GRONDIN + Nicolas ALBANY

Je joins aussi 3 images de signatures de référence :
- Image 1 = signature de Daniel GRONDIN
- Image 2 = signature de Nicolas ALBANY
- Image 3 = signature de Yann BAREGE
(Christophe BIZEUL est le seul technicien qui signe avec son prénom écrit en toutes
lettres — pas besoin d'image de référence pour lui.)

### Organisation du travail (autorisée)

- Rends chaque page de chaque PDF en image (il y a souvent 200+ pages).
- Tu es autorisé à découper le zip en 4 lots et à lancer 4 sous-agents en
  parallèle. Chaque sous-agent renvoie un tableau (fichier, page, n° BI, date,
  technicien, heures, méthode, notes) avec les images de signature de référence
  et TOUTES les règles ci-dessous. Tu consolides ensuite toi-même.
- Contrôle final obligatoire : génère une planche de contrôle (n° BI imprimé en
  haut à droite + zone date, recadrés, avec le nom du fichier à côté) et vérifie
  que chaque ligne du tableau correspond bien au bon fichier. Un décalage d'une
  ligne entre fichiers et lectures est l'erreur la plus coûteuse.
- Zoome (recadrage haute résolution) sur toute date, MO ou signature douteuse
  avant de trancher.

### Ce qu'il faut lire sur chaque page

1. **Documents à compter** :
   - les bons d'intervention / installation (formulaire "INTERVENTION").
     Attention : le champ imprimé en haut "DEVIS" + n° (ex : 0001379) est juste le
     numéro de référence interne du bon d'intervention (n° BI) — ce n'est PAS un
     devis, c'est le document à lire normalement ;
   - les BONS DE LIVRAISON (formulaire Hygierun BC/BL/Devis/Retour, case BL cochée
     ou non) : ils COMPTENT comme intervention. Avec installation et MO écrite →
     heures = MO. Livraison seule sans MO → 1 intervention, 0 h. Sur ce formulaire
     il n'y a pas de champ Technicien : utilise le champ "Commercial" (ex :
     "Nicolas") et la signature ; si aucun nom, déduis-le de l'écriture et
     signale-le dans A_verifier.
   À ignorer : les vrais devis imprimés ("Devis N°47xxx"), les e-mails, les bons de
   commande clients, les instructions d'expédition.

2. **Technicien** : lire le champ "Technicien".
   - S'il y a deux noms, ce sont forcément un des deux duos (jamais un mélange) :
     Christophe BIZEUL + Yann BAREGE (EBC) ou Nicolas ALBANY + Daniel GRONDIN (SAV).
   - Le champ contient souvent un code "1" ou "2" au lieu d'un nom (en général
     1 = Christophe, 2 = Yann), MAIS il faut TOUJOURS vérifier la signature en bas
     à gauche : c'est la signature qui fait foi (ex : une fiche codée "1" signée
     Yann = Yann).
   - Champ vide ou illisible : identifie via la signature (images de référence).
   - Joël DANVIN : ne compter QUE les fiches où il est lui-même le technicien
     (typiquement les interventions à Mayotte). Son nom apparaît aussi comme
     destinataire des e-mails et "Suivi par" sur les devis — ça ne compte pas.

3. **Heures travaillées** :
   - D'abord la mention "MO = XXh" (ou "MO Total", "maintenance = 1H", etc.),
     généralement écrite en gris au milieu/bas de la fiche. Si elle est présente,
     elle prime TOUJOURS, même si elle est très différente de l'horaire
     arrivée/départ (le signaler en note, sans changer la valeur).
   - Unités : "MO = 30" ou "30 min" = 30 minutes (0,5 h). Toute autre valeur
     (0,5 ; 1 ; 1H30 ; 2H…) est en heures.
   - Sinon (pas de MO écrite), calcule départ − arrivée.
   - Livraison de pièces seule sans MO écrite → 1 intervention, 0 h (ne pas
     calculer l'horaire).

4. **Date** :
   - D'abord la date ÉCRITE sur la fiche, en bas à droite ("le JJ/MM/AAAA"). C'est
     elle qui détermine le mois, même si elle semble incohérente avec la date de
     scan (le signaler en note).
   - Seulement si elle est illisible/masquée : utilise la date de scan encodée dans
     le nom du fichier (docNNNNNAAAAMMJJhhmmss — les 8 chiffres après le n° de doc
     sont AAAAMMJJ) et note "date déduite du nom de fichier".
   - Seules les fiches datées du mois analysé comptent dans la Synthèse. Les
     fiches d'un autre mois (souvent le mois précédent, scannées en retard) sont
     listées dans Detail avec "Mois compté = hors mois" et ne sont pas totalisées.

5. **Fichiers PDF multi-pages** : vérifie chaque page, un même fichier peut
   contenir plusieurs bons d'intervention distincts. Pages pivotées ou tronquées :
   fais pivoter l'image ; si une info est coupée, dis-le.

### Règles spéciales (validées, à appliquer telles quelles)

- **Deux bons liés** ("Va avec BI ####" / "Suit BI ####") : 2 interventions
  distinctes. Si des heures sont écrites sur CHACUN des deux bons, on compte les
  deux (même si l'un dit "MO Total"). Si les heures ne sont écrites que sur UN
  SEUL des deux bons, n'attribue ces heures qu'à ce bon-là (0 h pour l'autre, ne
  rien inventer, ne pas doubler).
- **Un seul bon signé par deux techniciens de la même équipe** : 1 seule
  intervention partagée pour l'équipe, heures comptées une seule fois.
- **Une intervention sur plusieurs jours sur un seul bon** : 1 intervention, avec
  la MO totale écrite.
- **Doublons de scan** : même n° BI et même contenu présents deux fois → compter
  une seule fois (signaler dans A_verifier). Deux n° BI différents ne sont jamais
  des doublons, même client/même jour.

### Format de sortie attendu

Un fichier Excel nommé `Productivite_SAV_<Mois><Année>.xlsx`
(ex : `Productivite_SAV_Septembre2026.xlsx`), avec 3 feuilles :

1. **Synthèse** — une ligne par équipe (uniquement les fiches du mois) :
   `Equipe | Nb intervention | Nb heures en intervention pour le mois`

2. **Detail_<Mois>** (ex : `Detail_Septembre`) — une ligne par intervention :
   `Fichier | Page | N° BI | Date signature | Mois compté | Technicien | Equipe | Heures | Méthode | Notes`
   (Méthode = MO_ecrit / calcul_horaire / déduit_filename / livraison_sans_MO ;
   Mois compté = le mois, ou "hors mois, non compté" ; lignes hors mois en gris)

3. **A_verifier** — les cas limites à relire, triés par impact en heures :
   `Fichier | Page | N° BI | Motif | Impact (h) | Détail`
   (lecture incertaine, date masquée, doublon écarté, technicien identifié par
   code/signature, BL sans nom, MO très différente de l'horaire, bons liés, hors
   mois…)

Ne calcule pas les heures travaillées théoriques ni la productivité — ça,
c'est fait automatiquement par l'application en aval à partir de ce fichier.
```

---

## Notes pour l'opérateur

- Les 3 images de signatures de référence (Daniel, Nicolas, Yann) sont à joindre
  au message avec le zip du mois — sans elles, Claude ne peut pas identifier un
  technicien dont le nom n'est pas écrit sur la fiche.
- Le fichier de sortie (`Productivite_SAV_<Mois><Année>.xlsx`) contient les
  chiffres à recopier dans la feuille `Bilan_Fiches` (colonnes `Equipe`,
  `Nb interventions`, `Nb heures en intervention`) de l'Input SAV — que ce soit
  l'onglet SAV seul ou l'Input Mensuel consolidé de l'onglet Bilan total.
  L'application s'occupe du reste (heures travaillées théoriques à partir des
  absences et des jours ouvrés du mois, puis productivité finale).
