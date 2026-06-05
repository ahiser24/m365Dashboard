"use strict";
/* Product switcher for the combined M365 Adoption board.
   The Viva module (dashboard.js) auto-boots on load. The Copilot module
   (copilot.js) does NOT auto-boot — it exposes window.CopilotBoard and is
   booted lazily the first time the Copilot tab is selected, so its fetch
   probing only runs when the user actually wants to see it. */
(function(){
  function $(id){ return document.getElementById(id); }

  function show(product){
    var viva = product === "viva";
    $("viva-app").classList.toggle("hidden", !viva);
    $("cp-app").classList.toggle("hidden", viva);
    if(!viva && window.CopilotBoard){ window.CopilotBoard.boot(); }
  }

  function init(){
    var sw = $("productSwitch");
    if(!sw) return;
    var btns = sw.querySelectorAll("button");
    btns.forEach(function(b){
      b.addEventListener("click", function(){
        btns.forEach(function(x){ x.setAttribute("aria-selected","false"); });
        b.setAttribute("aria-selected","true");
        show(b.dataset.product);
      });
    });
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
