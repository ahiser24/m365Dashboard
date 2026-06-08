"use strict";
// Handles tab switching between the dashboard panels (Home, Viva, Copilot, Teams).
// Each product is lazy-loaded so we only fetch/boot it when clicked for the first time.
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
    
    // Show the active pane and hide the others
    Object.keys(PANES).forEach(function(p){
      var el=$(PANES[p]); if(el) el.classList.toggle("hidden", p!==product);
    });
    
    // Keep switcher tabs and home grid tiles in sync
    syncAria("productSwitch", product);
    syncAria("appChooser", product);
    
    // Boot the product dashboard if it hasn't loaded yet
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

  function initTheme() {
    var themeToggle = $("themeToggle");
    if (!themeToggle) return;

    function updateThemeUI(theme) {
      var icon = theme === "dark" ? "☀️" : "🌙";
      var title = theme === "dark" ? "Switch to Light Mode" : "Switch to Dark Mode";
      var iconEl = themeToggle.querySelector(".theme-icon");
      if (iconEl) iconEl.textContent = icon;
      themeToggle.title = title;
    }

    var savedTheme = localStorage.getItem("theme");
    if (savedTheme) {
      document.documentElement.setAttribute("data-theme", savedTheme);
      updateThemeUI(savedTheme);
    } else {
      var isDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
      updateThemeUI(isDark ? "dark" : "light");
    }

    themeToggle.addEventListener("click", function() {
      var current = document.documentElement.getAttribute("data-theme") || 
                    (window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
      var next = current === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      localStorage.setItem("theme", next);
      updateThemeUI(next);
    });
  }

  function init(){
    wire("productSwitch");
    wire("appChooser");
    initTheme();
    var home=$("homeLink");
    if(home) home.addEventListener("click", function(){ show("home"); });
    
    // Close dropdowns if the user clicks anywhere else
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
