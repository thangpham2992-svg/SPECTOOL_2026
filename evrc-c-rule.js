'use strict';
// Required EVRC-C for special COPs when any of four features is selected.
// Keep model-specific COP formats and do not auto-apply the feature.
window.SPEC_EVRC_C_RULE = Object.freeze({
 check(model, copResult, selected){
  const cop = String(copResult?.ok ? copResult.value ?? '' : '').trim().toUpperCase();
  const special = model === 'MR'
   ? /^(?:C)?(?:760|770|780)(?:\b|$)/.test(cop)
   : model === 'MRL' && /^[MP]\d{3,4}(?:ETL|E)?$/.test(cop);
  const triggers = ['NSCR-C','AS','SCS-B','NS-CB'];
  const triggered = special ? triggers.filter(code => !!selected(code)) : [];
  return {cop, special, triggered, required: triggered.length > 0,
   evrcC: !!selected('EVRC-C')};
 }
});
