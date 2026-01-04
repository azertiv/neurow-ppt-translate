# SlideTranslate AI (PowerPoint Add-in)

Add-in PowerPoint (Office.js) pour traduire **la slide actuelle** ou **tout un deck** via l'API OpenAI, en conservant **au maximum** la mise en forme (gras/italique/couleurs/sauts de ligne / bullets).

> ⚠️ Sécurité : cette version est **100% côté client** (GitHub Pages). La clé OpenAI est saisie et utilisée dans le panneau. C’est pratique pour un POC, mais pas sûr pour un usage pro. Pour la prod, ajoute un petit proxy serveur (Cloudflare Worker, Azure Function, etc.).

## Fonctionnalités

- Traduction slide actuelle ou toutes les slides
- Conservation du style :
  - Text boxes : on reconstruit le texte puis on réapplique les styles par plages (runs)
  - Tables : on utilise `TableCell.textRuns` (format conservé)
- Option "Adapter la longueur" : demande une traduction plus courte/plus longue pour limiter les débordements
- Prévisualisation (ne modifie pas le deck)
- Glossaire `Terme=Traduction`
- Exclusion via regex sur `shape.name`

## Prérequis

- Node.js 20+
- PowerPoint Desktop ou PowerPoint Online avec Office.js

## Développement local

```bash
npm install
npm run dev
```

> Pour sideload en local, il faut **HTTPS** (cert dev) — la doc Microsoft Office Add-ins explique comment faire. Sinon, teste via GitHub Pages.

## Déploiement (GitHub Pages)

1. Active GitHub Pages: **Settings → Pages → Build and deployment: GitHub Actions**
2. Push sur `main` → l'action `Deploy GitHub Pages` build et déploie.

⚙️ Vite est configuré avec `BASE_PATH="/${repo}/"` pour que les assets marchent sur Pages.

## Sideload du manifest

1. Déploie sur GitHub Pages.
2. Modifie `manifest.xml` et remplace:
   - `YOUR_GITHUB_USER`
   - `YOUR_REPO`
3. Dans PowerPoint: **Insérer → Mes compléments → Gérer mes compléments → Télécharger mon complément** (ou via le catalogue de compléments de ton org), puis sélectionne `manifest.xml`.

## Notes techniques (important)

- **TextRange ne fournit pas directement des “runs”** (rich text) hors tableaux. Pour garder le style, on détecte les changements de style en scannant le texte caractère par caractère (optimisé) puis on réapplique les attributs via `TextRange.getSubstring(...).font`.
- Pour éviter les bugs connus, l’alignement de paragraphe est appliqué prudemment.

## Structure

- `taskpane.html` / `src/taskpane/*` : UI
- `src/services/openai.ts` : appel OpenAI Responses API (`/v1/responses`)
- `src/services/ppt.ts` : extraction / traduction / application
- `src/services/formatting.ts` : extraction & restauration de styles
- `manifest.xml` : add-in + bouton ribbon

