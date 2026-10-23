# Bundled dependencies

- `xlsx-0.20.3.full.min.js`: SheetJS CE 0.20.3, downloaded from https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js. License: `SHEETJS-LICENSE.txt` (Apache 2.0).
- `fonts/inter-latin.woff2`: Inter variable Latin font from https://fonts.gstatic.com/s/inter/v20/UcC73FwrK3iLTeHuS_nVMrMxCp50SjIa1ZL7W0Q5nw.woff2. License: `fonts/INTER-LICENSE.txt` (SIL OFL).
- `fonts/jetbrains-mono-latin.woff2`: JetBrains Mono Latin font from https://fonts.gstatic.com/s/jetbrainsmono/v24/tDbY2o-flEEny0FZhsfKu5WU4zr3E_BX0PnT8RD8yKxTOlOVk6OThhvA.woff2. License: `fonts/JETBRAINS-MONO-LICENSE.txt` (SIL OFL).

These files are served locally. Do not replace them with unpinned CDN references. Refresh integrity hashes with `node scripts/update-integrity.cjs` after updating scripts or styles. Review the SheetJS security advisories at https://cdn.sheetjs.com/advisories/ when upgrading.
