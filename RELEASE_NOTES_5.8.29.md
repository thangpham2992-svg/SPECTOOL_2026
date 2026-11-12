# SpecWebTool v5.8.29 — Check Spec Error Navigation Fix

- Fixed tapping a dimension error (AH/BH/OH/PD) not scrolling to its field after the responsive dimension layout redesign.
- Scroll targeting now uses the visible `.dimension-entry` instead of its `.field` wrapper with `display: contents`.
- Focus and highlight both the erroneous row and input; avoid extra browser automatic scroll.
- Same behavior supported for both NEXIEZ-MR and NEXIEZ-MRL Version02.
- No technical data, spec validation or export mapping changes.
