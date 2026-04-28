# 📊 Automatisation Excel Bancaire — Groq API / Llama 3

Script Python qui automatise l'extraction des données financières depuis un **PDF d'états financiers tunisiens** et remplit automatiquement le modèle Excel de décomposition bancaire, en utilisant l'IA **Llama 3.3 70B** via l'API Groq.

---

## 🎯 Objectif

Remplacer la saisie manuelle des données financières dans le modèle Excel par un processus 100% automatisé :

```
PDF États Financiers (DT)
        ↓
   Groq API / Llama 3
        ↓
Modèle Excel Saisie (KTND)
```

---

## ⚙️ Fonctionnement

Le script fonctionne en 3 étapes :

**Étape 1 — Extraction PDF**
Lecture du PDF (jusqu'à 30 pages) avec `pdfplumber` et extraction du texte brut.

**Étape 2 — Analyse par l'IA**
Le texte est envoyé à Llama 3.3 70B via l'API Groq. L'IA identifie et extrait les 28 postes financiers clés (actif, passif, résultat) et les retourne au format JSON en dinars tunisiens (DT).

**Étape 3 — Remplissage Excel**
Le script localise automatiquement la colonne `31/12/2024` dans la feuille `Saisie`, puis remplit chaque ligne via le mapping codes liasse → JSON. Les valeurs sont converties de DT en KTND (÷ 1000).

---

## 📁 Structure du projet

```
projet/
├── script_excel.py          ← Script principal
├── .env                     ← Clé API Groq (à créer, ne pas partager)
├── .env.example             ← Exemple de fichier .env
├── .gitignore
├── requirements.txt
├── SOTIPAPIER_311224.xlsx   ← Fichier de mapping Excel (construit manuellement)
├── Modele décompo.xlsx      ← Modèle Excel bancaire à remplir
└── README.md
```

---

## 🚀 Installation

### 1. Cloner le repo

```bash
git clone https://github.com/ton-username/ton-repo.git
cd ton-repo
```

### 2. Installer les dépendances

```bash
pip install groq pdfplumber openpyxl python-dotenv
```

### 3. Créer ta clé API Groq (gratuit)

1. Va sur [https://console.groq.com](https://console.groq.com)
2. Crée un compte gratuit
3. Va dans **API Keys** → **Create API Key**
4. Copie la clé générée

### 4. Configurer le fichier `.env`

Crée un fichier `.env` à la racine du projet :

```
GROQ_API_KEY=gsk_xxxxxxxxxxxxxxxxxxxxxxxx
```

> ⚠️ Ne partage jamais ce fichier. Il est dans `.gitignore` pour ne pas être publié.

---

## ▶️ Utilisation

1. Place le **PDF des états financiers** dans le dossier du projet
2. Place le **modèle Excel** (`Modele décompo.xlsx`) dans le même dossier
3. Modifie les chemins dans le script si nécessaire :
   ```python
   CHEMIN_PDF   = "Sotipapier EF indiv 31.12.2024.pdf"
   CHEMIN_EXCEL = "Modele décompo.xlsx"
   ```
4. Lance le script :
   ```bash
   python script_excel.py
   ```

Le fichier Excel rempli sera sauvegardé sous `REMPLI_NomSociété_31122024.xlsx`.

---

## 📋 Postes financiers extraits (28)

| Catégorie | Postes |
|-----------|--------|
| **Actif immobilisé** | Immo. incorporelles brutes/amort., Immo. corporelles brutes/amort., Immo. financières brutes/prov. |
| **Actif courant** | Stocks bruts/prov., Clients bruts/prov., Autres actifs courants, Placements, Liquidités |
| **Capitaux propres** | Capital social, Prime d'émission, Réserves légales, Résultats reportés, Résultat net |
| **Passif** | Prov. risques & charges, Emprunts LT, Fournisseurs, Autres passifs courants, Concours bancaires |
| **Compte de résultat** | Chiffre d'affaires, Charges personnel, Autres charges exploitation, Charges financières nettes, Impôt bénéfices |

---

## 🔧 Limites connues

- Le texte PDF est limité à **20 000 caractères** (contrainte du free tier Groq)
- L'IA peut retourner `null` si un poste est absent ou illisible dans le PDF
- Le mapping codes liasse est basé sur la structure de `SOTIPAPIER_311224.xlsx`

---

## 📦 Dépendances

```
groq
pdfplumber
openpyxl
python-dotenv
```

---

## 📄 Contexte

Projet réalisé dans le cadre d'un stage d'analyse financière. L'objectif est d'automatiser la décomposition des états financiers tunisiens (normes NCT) pour alimenter un modèle bancaire de scoring.
