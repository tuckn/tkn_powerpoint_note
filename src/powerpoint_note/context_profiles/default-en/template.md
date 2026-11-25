---
version: 2.0.0
language: English
labels:
  personal: "Add your own notes below this line."
  slide: "Slide"
  column: "Column"
  merged_cells: "This table has merged cells; each value appears in the top-left cell of its merge."
  chart: "Chart"
  smartart: "SmartArt"
  picture: "Picture"
  location_outside: "off slide"
  location_partial: "partly off slide"
  location_unknown: "position unknown"
  stale: "The slide changed after this explanation was generated. Run with `--context` to update it."
  skipped: "No explanation was generated for this slide"
---
#### Slide context (AI generated)

{{#stale}}
> [!WARNING]
> {{stale}}
{{/stale}}

{{summary}}

{{#sections}}
{{sections}}
{{/sections}}

{{#uncertainties}}
##### Uncertainties

{{uncertainties}}
{{/uncertainties}}

{{#image}}
##### Source image

{{image}}
{{/image}}
