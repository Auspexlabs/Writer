WMFJS.bundle.js is the unmodified WMF renderer from rtf.js 3.0.9, MIT licensed.
Source: https://github.com/tbluemel/rtf.js
Package: https://registry.npmjs.org/rtf.js/-/rtf.js-3.0.9.tgz
Bundle SHA-256: cbd48013970b474e82fc7515b6977ca79d620ea65688d4c4f996d87973ba6ed3

symbol.txt is the Adobe Symbol encoding map from Unicode, with its redistribution notice intact:
https://www.unicode.org/Public/MAPPINGS/VENDORS/ADOBE/symbol.txt
symbol.js derives that map and replaces obsolete private bracket pieces with standard Unicode characters.

The application lazily loads this local bundle, renders the preview as SVG, and keeps the original WMF/OLE parts in the document. No document bytes leave the application.
