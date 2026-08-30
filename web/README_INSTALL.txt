Model Tracer v1.6.4 — GitHub Pages install pack

Repository: cavankohys/Excel-Tracer
Hosted web folder: /web
Expected base URL: https://cavankohys.github.io/Excel-Tracer/web

1. In GitHub: Settings > Pages.
2. Under Build and deployment, choose Deploy from a branch.
3. Branch: main. Folder: /(root). Save.
4. Wait for GitHub Pages to show the site is deployed.
5. Verify in a browser:
   https://cavankohys.github.io/Excel-Tracer/web/health.txt
   https://cavankohys.github.io/Excel-Tracer/web/taskpane.html
6. For quickest testing, sideload manifests/model-tracer.xml into Excel.

Mac fallback sideload path:
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
Copy model-tracer.xml into that folder, then restart Excel.

The manifest is already configured for your GitHub Pages URL. No PowerShell, Node.js, npm, localhost server, or certificate setup is required.
