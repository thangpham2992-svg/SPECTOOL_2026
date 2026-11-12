# SpecWebTool v5.8.44 – Link Icon and Share Preview

- Added assets/specweb-icon.png from user-provided icon as browser favicon and Apple touch icon.
- Added assets/specweb-share.jpg (1200 x 630) for link preview cards.
- Added Open Graph and Twitter card metadata to index.html.
- Retained all technical logic, database, models and exports from v5.8.43.
- Deployment note: GitHub Pages crawlers (Zalo/Facebook/Messenger) may require an **absolute public URL** in og:image. When the final published repository URL is known, replace `assets/specweb-share.jpg` in og:image with `https://<account>.github.io/<repository>/assets/specweb-share.jpg` (or root equivalent); no JavaScript-based replacement for social crawlers.
- Icon updates may require clearing website and social-network caches.
