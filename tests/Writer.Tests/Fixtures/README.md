# Test fixtures

The Office files in `docx/`, `xlsx/` and `pptx/` and the Markdown files in `md/` are copied from the
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) `examples/` directory and are used here unchanged as
round-trip fidelity samples, with one exception: `pptx/decor.pptx` is our modification of the OfficeCLI
Mars deck. It adds a decorated slide master (bar, connector, group, footer placeholder, background), layout
pictures and backgrounds, `showMasterSp="0"` on a layout and a slide, a slow fade, and speaker notes, a hidden
slide and a timed push transition on slide 2. They are licensed under the Apache License 2.0; see `LICENSE-OfficeCLI` and
`NOTICE-OfficeCLI` in this directory.
