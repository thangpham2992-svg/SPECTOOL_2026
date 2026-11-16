# SpecWebTool v5.8.67 – SETS-L suggestion, GitHub <100 files

Base: v5.8.66. No changes to the Maximum Decoration Weight 5% margin, to
AH/BH/OH/PD formulas, any user inputs, or model MR validation.

For NEXIEZ-MRL Version02, when the current OH/PD is below the standard minimum,
show a read-only hint suggesting the permitted SETS-L minimum calculated by
re-evaluating the *same* SpecForm F36/F37 formulas with A157 enabled. The hint
is shown only when the SETS-L option is allowed by SpecForm E157 and
DataTable D520 and when it actually reduces the corresponding minimum.
The original Check Spec failure is retained and clicking it still navigates to
that dimension. If the proposed minimum is above actual input, warning states
that the input must still be increased. At 60 m/min no SETS-L advice is shown.

Removed only obsolete testing JS and replaced prior release note with this one.
No referenced assets or runtime JS files removed.
