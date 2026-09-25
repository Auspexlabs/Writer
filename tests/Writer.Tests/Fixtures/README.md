# Test fixtures

The Office files in `docx/`, `xlsx/` and `pptx/` and the Markdown files in `md/` are copied from the
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) `examples/` directory and are used here unchanged as
round-trip fidelity samples, with one exception: `pptx/decor.pptx` is our modification of the OfficeCLI
Mars deck. It adds a decorated slide master (bar, connector, group, footer placeholder, background), layout
pictures and backgrounds, `showMasterSp="0"` on a layout and a slide, a slow fade, and speaker notes, a hidden
slide and a timed push transition on slide 2. They are licensed under the Apache License 2.0; see `LICENSE-OfficeCLI` and
`NOTICE-OfficeCLI` in this directory.

`pptx/text-runs.pptx` is ours: a blank deck from the engine with one text box whose runs each carry their own fonts,
colours, highlight, double and wavy underlines, double strike, sub- and superscript, caps, spacing, kerning, languages, a
link with a tooltip and an `extLst`, whose paragraphs carry their own alignment, spacing, bullet, numbering, level and
end marks around a line break and a slide-number field, and a vertical box with its own insets and list style.
