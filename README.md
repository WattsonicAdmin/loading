<<<<<<< HEAD
# Wattsonic Container Loading V1

A local 3D container loading web tool.

## Files
- `index.html` - app entry
- `styles.css` - UI styles (light, clean)
- `app.js` - 3D scene, drag/drop, quantity fill, auto fill, weight constraints
- `container-models.json` - container data + 3D model parameters (20GP / 40GP / 40HQ)

## Run locally
Use any local HTTP server from this folder:

```bash
python3 -m http.server 8080
```

Open:
- `http://127.0.0.1:8080`

## Publish for customer access
This project is now runtime-CDN based (Three.js + jsPDF), so deployment is static-host friendly.

### Option 0: GitHub Pages (recommended for your request)
1. Create a GitHub repository and push this project to branch `main`.
2. This repo includes workflow: `.github/workflows/deploy-pages.yml`.
3. In GitHub repo settings:
   - `Settings -> Pages -> Build and deployment -> Source: GitHub Actions`.
4. Push to `main` and wait for `Deploy Static Site to GitHub Pages` workflow.
5. Open your published URL:
   - `https://<your-github-username>.github.io/<repo-name>/`

### Option A: Vercel (recommended)
1. Push this folder to a GitHub repository.
2. Import the repo in Vercel.
3. Framework preset: `Other`.
4. Build command: empty.
5. Output directory: `.` (project root).
6. Deploy and share the generated HTTPS URL.

### Option B: Netlify Drop (fastest)
1. Open Netlify.
2. Drag this folder directly into Netlify Drop.
3. Wait for deployment and share the generated URL.

## Core features
- 3D container view (20GP / 40GP / 40HQ)
- Color-coded pallet models
- Drag pallets into container
- Fill by quantity
- Auto fill by algorithm under max weight limit
- Live stats: pallet count, total weight, utilization, footprint
- Click a placed pallet to remove it

## Standards-based constraints
- CTU Code: load distribution, stack stability, and securing-risk checks
- ISO 668 / ISO 1496: dimensional and rated payload/gross checks
- CSC Convention: plate and periodic examination confirmations
=======
# loading
Wattsonic Products Loading into containers
>>>>>>> d6bf20a4329111067943e270196a6ad376bdb39c
