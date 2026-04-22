# 🗂️ Inventaire Logiciels Windows

Script PowerShell qui génère automatiquement un inventaire complet des logiciels installés sur une machine Windows, avec export en **CSV** et rapport **HTML interactif**.

---

## ✨ Fonctionnalités

- Détection des logiciels depuis les 3 sources du registre Windows (64 bits, 32 bits, Utilisateur)
- Détection des applications **Microsoft Store**
- Export **CSV** (délimiteur `;`, encodage UTF-8)
- Rapport **HTML** interactif avec :
  - Statistiques globales (nombre de logiciels, taille totale)
  - Recherche en temps réel
  - Filtrage par source
  - Tri par colonne
  - Badges colorés par source
- Nommage automatique avec horodatage (`inventaire_AAAA-MM-JJ_HH-mm`)

---

## 🚀 Utilisation

### Exécution directe

```powershell
.\inventaire-logiciels.ps1
```

> Si la politique d'exécution bloque le script :
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Résultat

Les fichiers générés sont enregistrés dans :
```
%USERPROFILE%\Documents\InventaireLogiciels\
```

Le rapport HTML s'ouvre automatiquement dans le navigateur par défaut à la fin de l'exécution.

---

## 📁 Structure des fichiers générés

```
Documents\
└── InventaireLogiciels\
    ├── inventaire_2026-04-22_14-30.csv
    └── inventaire_2026-04-22_14-30.html
```

---

## 📋 Colonnes exportées

| Colonne       | Description                              |
|---------------|------------------------------------------|
| Nom           | Nom du logiciel                          |
| Version       | Version installée                        |
| Editeur       | Éditeur / développeur                    |
| DateInstall   | Date d'installation (format AAAA-MM-JJ)  |
| TailleMo      | Taille estimée en Mo                     |
| Source        | Origine (registre 64/32 bits, Utilisateur, Store) |
| Emplacement   | Répertoire d'installation                |
| Desinstaller  | Commande de désinstallation              |

---

## ⚙️ Prérequis

- Windows 10 / 11
- PowerShell 5.1 ou supérieur
- Aucune dépendance externe

---

## 📄 Licence

Ce projet est distribué sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus d'informations.
