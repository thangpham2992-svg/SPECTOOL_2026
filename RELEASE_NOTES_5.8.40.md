# SpecWebTool v5.8.40 – N-series COP image preview for MRL Version02

- Recognize COP model N-prefixed codes (`N1010E`, `N1030E`, etc.) when resolving image previews.
- For MRL Version02 only, use the matching M-prefixed COP model image if there is no matching N image. Also support shortened codes (`N1010` -> `M1010E`).
- Preview only: actual selected code, source database, Check Spec and Excel/PDF exports are unchanged.
- MR unaffected.
- Passed JS syntax and mapping regression tests; browser interaction not verified.
