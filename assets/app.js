"use strict";
/* Product switcher for the combined M365 Adoption board.
   The page opens on a neutral Home chooser — no product loads until the user
   picks one. Each product module (Viva, Copilot, Teams) exposes a lazy
   window.*Board with a boot() that runs its loader exactly once, so selecting
   a tab the first time triggers its fetch/upload flow and re-selecting it just
   shows the already-loaded board. */
(function(){
  function $(id){ return document.getElementById(id); }

  var PANES = {
    home:    "home-app",
    viva:    "viva-app",
    copilot: "cp-app",
    teams:   "tm-app"
  };
  function board(product){
    if(product==="viva")    return window.VivaBoard;
    if(product==="copilot") return window.CopilotBoard;
    if(product==="teams")   return window.TeamsBoard;
    return null;
  }

  function show(product){
    if(!PANES[product]) product="home";
    // toggle panes
    Object.keys(PANES).forEach(function(p){
      var el=$(PANES[p]); if(el) el.classList.toggle("hidden", p!==product);
    });
    // keep both the top switcher and the home chooser tiles in sync
    syncAria("productSwitch", product);
    syncAria("appChooser", product);
    // lazily boot the chosen product (no-op after the first time)
    var b=board(product);
    if(b) b.boot();
    window.scrollTo(0,0);
  }

  function syncAria(containerId, product){
    var c=$(containerId); if(!c) return;
    Array.prototype.forEach.call(c.querySelectorAll("button"), function(btn){
      btn.setAttribute("aria-selected", btn.dataset.product===product ? "true":"false");
    });
  }

  function wire(containerId){
    var c=$(containerId); if(!c) return;
    Array.prototype.forEach.call(c.querySelectorAll("button[data-product]"), function(btn){
      btn.addEventListener("click", function(){ show(btn.dataset.product); });
    });
  }

  function init(){
    wire("productSwitch");
    wire("appChooser");
    var home=$("homeLink");
    if(home) home.addEventListener("click", function(){ show("home"); });
    
    // Global click listener to close dropdowns
    document.addEventListener("click", function(e){
      if(!e.target.closest(".refresh-dropdown-container")){
        Array.prototype.forEach.call(document.querySelectorAll(".refresh-dropdown"), function(d){
          d.classList.remove("show");
        });
      }
    });

    show("home");
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
