// relabel_step7_to_5.js — rename visible "7) ..." heading to "5) ..." (v1)
(function(){
  function relabel(h, newNum){
    const rest = (h.textContent || '').replace(/^\s*\d+\)\s*/, '');
    h.textContent = `${newNum}) ${rest}`;
  }
  function run(){
    const hs = Array.from(document.querySelectorAll('.card h2, h2'));
    const cand = hs.find(h => /^\s*7\)\s*/.test((h.textContent||'')));
    if (cand) relabel(cand, '5');
  }
  document.addEventListener('DOMContentLoaded', run);
  run();
})();
