# Bundled PDF dependencies

PDF viewing, workers, CJK character maps, standard fonts, annotation saving and page rearrangement load locally. No CDN fallback. Dependencies are loaded only when PDF operations need them.

- `pdfjs/`: pdfjs-dist 4.4.168 from https://registry.npmjs.org/pdfjs-dist/-/pdfjs-dist-4.4.168.tgz. The legacy builds include compatibility support for older WebViews and Node 20. Apache-2.0; see `pdfjs/LICENSE`, `pdfjs/cmaps/LICENSE`, and `pdfjs/standard_fonts/LICENSE_*`.
- `pdf-lib/`: pdf-lib 1.17.1 from https://registry.npmjs.org/pdf-lib/-/pdf-lib-1.17.1.tgz. MIT; see `pdf-lib/LICENSE.md`. The ESM distribution contains its runtime dependencies.

The desktop resource mapping includes the complete ui directory, including these assets. Keep workers and auxiliary data at the same version as the PDF.js main module.
