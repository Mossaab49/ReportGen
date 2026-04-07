# 📄 Générateur de Rapport Interactif
**Developed by Nigo Namikaze**

Outil en ligne de commande pour générer des rapports Word professionnels avec watermark, templates personnalisés et génération automatique de texte via IA (Groq).

---

## 📁 Structure du projet

```
ton_projet/
├── report_tool.py       ← script principal
├── templates.json       ← configuration des templates
├── .env                 ← ta clé API (à créer, voir ci-dessous)
├── .gitignore           ← protège ton .env
├── README.md            ← ce fichier
└── templates/
    ├── mossaab_intro.png
    ├── mossaab_page.png
    ├── mossaab_outro.png
    └── ...
```

---

## ⚙️ Installation

### 1. Prérequis — Python 3.8+
Vérifie que Python est installé :
```powershell
python --version
```
Si non installé → [python.org/downloads](https://www.python.org/downloads)

### 2. Installer les dépendances
```powershell
pip install python-docx Pillow lxml requests
```

---

## 🔑 Obtenir une clé API Groq (gratuit)

L'outil utilise l'IA **Groq** pour générer automatiquement les paragraphes du rapport.
C'est **100% gratuit**, aucune carte bancaire requise.

### Étapes :

**1.** Aller sur → [console.groq.com](https://console.groq.com)

**2.** Cliquer sur **Sign Up** → s'inscrire avec Google ou Email

**3.** Une fois connecté → aller dans **API Keys** (menu à gauche)

**4.** Cliquer sur **Create API Key** → donner un nom → **Submit**

**5.** Copier la clé générée — elle ressemble à :
```
gsk_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```
> ⚠️ Copie-la maintenant, elle ne sera plus visible après fermeture.

---

## 🔧 Configurer la clé API — créer le fichier `.env`

Dans le dossier du projet, créer un fichier nommé **`.env`** (sans extension) :

### Windows — PowerShell :
```powershell
New-Item .env -ItemType File
notepad .env
```

### Windows — CMD :
```cmd
copy nul .env
notepad .env
```

### Contenu du fichier `.env` :
```ini
GROQ_API_KEY=gsk_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```
> ⚠️ Pas de guillemets, pas d'espaces autour du `=`  
> ✅ Remplace `gsk_xxx...` par ta vraie clé

**Sauvegarder et fermer.**

La clé est maintenant permanente — plus besoin de la ressaisir à chaque lancement.

---

## 🚀 Lancer l'outil

```powershell
cd chemin\vers\ton_projet
python report_tool.py
```

### Déroulement :

```
1. Choisir un template
2. Renseigner les infos du rapport (titre, auteur, encadrant, date)
3. Écrire ou générer l'introduction via IA
4. Ajouter les étapes (nom, images, paragraphes IA)
5. Écrire ou générer la conclusion via IA
6. Le fichier .docx est généré dans rapports_generes/
```

---

## 🖼️ Ajouter un nouveau template

Un template est composé de **3 images PNG** (format A4 recommandé : 794×1123 px) :

| Fichier | Rôle |
|---|---|
| `monid_intro.png` | Fond de la page de couverture |
| `monid_page.png` | Fond des pages de contenu |
| `monid_outro.png` | Fond de la page de conclusion |

### Étape 1 — Placer les images
Copier les 3 images dans le dossier `templates/` :
```
templates/
├── monid_intro.png
├── monid_page.png
└── monid_outro.png
```

### Étape 2 — Ajouter un bloc dans `templates.json`
Ouvrir `templates.json` et ajouter un bloc dans le tableau `"templates"` :

```json
{
  "id": "monid",
  "name": "Mon Template Personnalisé",
  "intro_bg": "monid_intro.png",
  "outro_bg": "monid_outro.png",
  "page_bg":  "monid_page.png",

  "accent_color":    [0, 112, 192],
  "secondary_color": [68, 114, 196],
  "text_dark":       [30, 30, 30],
  "subtitle_color":  [255, 145, 77],

  "title_font_size":    50,
  "report_font_size":   80,
  "heading_font_size":  30,
  "body_font_size":     14,
  "subtitle_font_size": 20,
  "caption_font_size":  13,

  "cover_blank_lines":      22,
  "page_left_margin_cm":   2.5,
  "outro_left_margin_cm":  2.5
}
```

### Explication des champs :

| Champ | Description | Exemple |
|---|---|---|
| `id` | Identifiant unique (même préfixe que les images) | `"monid"` |
| `name` | Nom affiché dans le menu | `"Mon Template"` |
| `accent_color` | Couleur des titres [R, G, B] | `[0, 112, 192]` |
| `secondary_color` | Couleur secondaire [R, G, B] | `[68, 114, 196]` |
| `text_dark` | Couleur du corps de texte [R, G, B] | `[30, 30, 30]` |
| `subtitle_color` | Couleur des sous-titres [R, G, B] | `[255, 145, 77]` |
| `cover_blank_lines` | Lignes vides avant le titre en couverture | `22` |
| `page_left_margin_cm` | Marge gauche des pages de contenu | `2.5` |
| `outro_left_margin_cm` | Marge gauche de la page conclusion | `2.5` |

> 💡 Pour trouver les valeurs RGB d'une couleur → [rgbcolorcode.com](https://rgbcolorcode.com)

### Étape 3 — Relancer le script
```powershell
python report_tool.py
```
Le nouveau template apparaît dans le menu de sélection ✅

---

## ❓ Problèmes fréquents

**La clé API n'est pas reconnue**
→ Vérifier que le fichier `.env` est bien dans le même dossier que `report_tool.py`
→ Vérifier qu'il n'y a pas de guillemets ni d'espaces : `GROQ_API_KEY=gsk_xxx`

**Erreur `model_decommissioned`**
→ Ouvrir `report_tool.py` et changer la ligne `MODEL` :
```python
MODEL = "llama-3.3-70b-versatile"
```

**Erreur `pip` introuvable**
→ Essayer `pip3` à la place de `pip`

**Le template n'apparaît pas dans le menu**
→ Vérifier que les 3 images existent bien dans `templates/`
→ Vérifier que l'`id` dans `templates.json` correspond au préfixe des noms de fichiers

**Image non trouvée lors de la génération**
→ Entrer le chemin complet : `C:\Users\toi\Images\capture.png`

---

## 📬 Contact

Développé par **Nigo Namikaze**
