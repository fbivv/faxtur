# Faxtur

Version 1.1.0 – nouvelle interface avec menu vertical, tableau de bord et logo réduit. Moteur Factur-X inchangé.

Faxtur est un outil Windows de conversion de factures PDF en **Factur-X / PDF/A-3b**, avec validation locale via **veraPDF**.

## État du projet

Version de pré-déploiement avant publication GitHub.

- Profil Factur-X : `BASIC WL`
- Validation PDF/A : `PDF/A-3b`
- Licence : Mozilla Public License 2.0
- Plateforme cible : Windows

## Installation utilisateur

1. Télécharger l'archive de release.
2. Dézipper.
3. Lancer `INSTALLER_FAXTUR.bat`.

L'installation se fait sans droits administrateur dans :

```text
%LOCALAPPDATA%\Faxtur
```

Les dossiers métier sont proposés par défaut sur le Bureau de l'utilisateur courant :

```text
Bureau\Factures
Bureau\Facture-X
Bureau\A traiter
```

## Lancement manuel développeur

```bat
python -m pip install -r requirements.txt
python main.py
```

Ou avec le lanceur inclus :

```bat
Faxtur.bat
```

## Moteur Factur-X

Le moteur validé est dans :

```text
engine/engine_v1.py
```

Il ne doit pas être modifié sans test de conformité complet.

## Validation

Faxtur vérifie :

- XML embarqué ;
- métadonnées XMP Factur-X ;
- attachement `/AF` ;
- profil `BASIC WL` ;
- conformité PDF/A via veraPDF.

## Licence

Copyright © 2026 Frédéric Brouard

This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0. See `LICENSE`.
