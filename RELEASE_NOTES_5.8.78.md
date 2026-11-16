# v5.8.78 – Flooring weight exceeds Technical Study maximum

- Preserve the Technical Study maximum even if flooring alone exceeds it.
- Display the total technical allowance (e.g. 118 kg) rather than --- or 0 in that case.
- Check Spec FAIL at Flooring Weight (D56) with the instruction to reduce stone-flooring mass.
- Decoration Weight = 0 is accepted as an input, but the flooring-overweight failure remains.
- Preserve existing ordinary decoration calculations and 5% reserve when flooring is within the allowance.
- Scope limited to NEXIEZ-MRL Version02; NEXIEZ-MR unchanged.
