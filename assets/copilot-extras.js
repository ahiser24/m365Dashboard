"use strict";
// Extension reports for Copilot Chat and Agents tabs
(function(){
  var RAW = {};   // declAgents, agents, chatAdopt, chatPrompts, chatUsers
  var TBL = {};   // SortTable instances, kept so search re-renders them

  // Helper functions
  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,function(c){
    return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]; }); }
  function num(v){ var n=parseFloat(String(v==null?"":v).replace(/,/g,"")); return isFinite(n)?n:0; }
  function fmt(n){ n=Math.round(n); return n.toLocaleString("en-US"); }
  function getUserField(row, possibleFields) {
    if (!row) return "";
    for (var i = 0; i < possibleFields.length; i++) {
      if (row[possibleFields[i]] !== undefined) return row[possibleFields[i]];
    }
    var rowKeys = Object.keys(row);
    for (var i = 0; i < possibleFields.length; i++) {
      var target = possibleFields[i].toLowerCase().replace(/[^a-z0-9]/g, "");
      for (var j = 0; j < rowKeys.length; j++) {
        var key = rowKeys[j].toLowerCase().replace(/[^a-z0-9]/g, "");
        if (key === target) return row[rowKeys[j]];
      }
    }
    return "";
  }
  function fmtShort(n){
    n=Math.round(n);
    if(Math.abs(n)>=1e6) return (n/1e6).toFixed(n%1e6?1:0)+"M";
    if(Math.abs(n)>=1e3) return (n/1e3).toFixed(n%1e3?1:0)+"k";
    return String(n);
  }
  function pct(a,b){ return b? Math.round(a/b*1000)/10 : 0; }
  function clip(s,n){ s=String(s); return s.length>n? s.slice(0,n-1)+"…" : s; }
  var COLORS = ["var(--c1)","var(--c2)","var(--c3)","var(--c4)","var(--c5)","var(--c6)","var(--c7)","var(--c8)"];

  // Copilot Chat surfaces
  var EDP_APPS = [
    {key:"Microsoft 365 Copilot (app)", short:"Copilot app",       color:"var(--accent)"},
    {key:"Teams",                       short:"Teams",             color:"var(--c2)"},
    {key:"Outlook",                     short:"Outlook",           color:"var(--c6)"},
    {key:"Word",                        short:"Word",              color:"var(--c4)"},
    {key:"Excel",                       short:"Excel",             color:"var(--c3)"},
    {key:"PowerPoint",                  short:"PowerPoint",        color:"var(--c5)"},
    {key:"OneNote",                     short:"OneNote",           color:"var(--c7)"},
    {key:"Edge",                        short:"Edge",              color:"var(--muted)"},
    {key:"Copilot.cloud.microsoft",     short:"Copilot Chat (web)",color:"var(--c1)"}
  ];

  // CSV parser
  function parseCSV(text){
    text = text.replace(/^﻿/,"");
    var rows=[], row=[], field="", i=0, inQ=false, c, n=text.length;
    for(; i<n; i++){
      c=text[i];
      if(inQ){
        if(c==='"'){ if(text[i+1]==='"'){ field+='"'; i++; } else inQ=false; }
        else field+=c;
      } else {
        if(c==='"') inQ=true;
        else if(c===','){ row.push(field); field=""; }
        else if(c==='\n'){ row.push(field); rows.push(row); row=[]; field=""; }
        else if(c==='\r'){ /* skip */ }
        else field+=c;
      }
    }
    if(field.length||row.length){ row.push(field); rows.push(row); }
    if(!rows.length) return {headers:[],rows:[]};
    var headers=rows[0].map(function(h){return h.trim();});
    var out=[];
    for(var r=1;r<rows.length;r++){
      if(rows[r].length===1 && rows[r][0]==="") continue;
      var o={}; for(var k=0;k<headers.length;k++){ o[headers[k]]= rows[r][k]!==undefined? rows[r][k] : ""; }
      out.push(o);
    }
    return {headers:headers, rows:out};
  }

  // SVG horizontal bar chart generator
  function hBar(items, opt){
    opt=opt||{};
    var n=items.length, rowH=opt.rowH||20, gap=opt.gap||9, pl=opt.labelW||130, pr=58;
    var W=opt.w||540, H=Math.max(1,n)*(rowH+gap)+8, iw=W-pl-pr;
    var max=Math.max.apply(null, items.map(function(d){return d.v;}).concat([1]));
    var out='<svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="bar chart" preserveAspectRatio="xMinYMin meet">';
    items.forEach(function(d,i){
      var y=i*(rowH+gap)+6, w=Math.max(2,(d.v/max)*iw);
      var col=d.color||opt.color||COLORS[i%COLORS.length];
      var vlabel=opt.fmt? opt.fmt(d.v) : fmtShort(d.v);
      var tipv=opt.tipv? opt.tipv(d) : fmt(d.v);
      out+='<rect x="'+pl+'" y="'+y+'" width="'+w.toFixed(1)+'" height="'+rowH+'" rx="5" fill="'+col+'"'+
           ' data-tip="'+esc(d.label)+'" data-tipv="'+esc(tipv)+'" data-tipc="'+col+'"></rect>';
      out+='<text x="'+(pl-8)+'" y="'+(y+rowH/2+4)+'" text-anchor="end" font-size="12" fill="var(--ink)">'+esc(clip(d.label,28))+'</text>';
      out+='<text x="'+(pl+w+7).toFixed(1)+'" y="'+(y+rowH/2+4)+'" font-size="11.5" fill="var(--muted)">'+esc(vlabel)+'</text>';
    });
    return out+'</svg>';
  }

  // SortTable class definition
  function addResizer(th, handle, table){
    var startX=0, startW=0, active=false;
    handle.addEventListener("pointerdown", function(e){
      e.preventDefault(); e.stopPropagation();
      active=true; startX=e.clientX; startW=th.getBoundingClientRect().width;
      Array.prototype.forEach.call(table.querySelectorAll("thead th"), function(h){
        if(h===th) return;
        var cw=h.getBoundingClientRect().width;
        h.style.width=cw+"px"; h.style.minWidth=cw+"px"; h.style.maxWidth=cw+"px";
      });
      table.style.width="auto"; table.style.minWidth="100%";
      handle.classList.add("active"); table.classList.add("resizing");
      try{ handle.setPointerCapture(e.pointerId); }catch(_){}
    });
    handle.addEventListener("pointermove", function(e){
      if(!active) return;
      var w=Math.max(48, startW+(e.clientX-startX));
      th.style.width=w+"px"; th.style.minWidth=w+"px"; th.style.maxWidth=w+"px";
    });
    function end(e){
      if(!active) return; active=false;
      handle.classList.remove("active"); table.classList.remove("resizing");
      try{ handle.releasePointerCapture(e.pointerId); }catch(_){}
    }
    handle.addEventListener("pointerup", end);
    handle.addEventListener("pointercancel", end);
  }
  function SortTable(tableEl, columns, opts){
    this.table=tableEl; this.columns=columns; this.opts=opts||{};
    this.sort=this.opts.initSort? {key:this.opts.initSort.key, dir:this.opts.initSort.dir} : {key:null,dir:-1};
    this.built=false; this.thRow=null;
    this.selectable=this.opts.selectable!==false;
    this.selected={}; this.selOrder=[]; this.rowMap={}; this.cmpUI=null;
  }
  SortTable.prototype.buildHead=function(){
    var self=this, thead=this.table.querySelector("thead");
    this.table.classList.add("resizable");
    var tr=document.createElement("tr");
    if(this.selectable){
      var selTh=document.createElement("th");
      selTh.className="selcol"; selTh.setAttribute("aria-label","Select");
      tr.appendChild(selTh);
    }
    this.columns.forEach(function(c){
      var th=document.createElement("th");
      if(c.num) th.classList.add("num");
      if(c.sort!==false) th.classList.add("sortable");
      th.dataset.k=c.k;
      var lbl=document.createElement("span"); lbl.className="thlabel"; lbl.textContent=c.t;
      th.appendChild(lbl);
      if(c.sort!==false){
        th.addEventListener("click", function(e){
          if(e.target.classList.contains("col-resizer")) return;
          self.clickSort(c);
        });
      }
      var rz=document.createElement("div"); rz.className="col-resizer"; rz.title="Drag to resize";
      th.appendChild(rz); addResizer(th, rz, self.table);
      tr.appendChild(th);
    });
    if(this.selectable){
      var tb=this.table.querySelector("tbody");
      if(tb && !tb._selWired){
        tb._selWired=true;
        tb.addEventListener("change", function(e){
          var cb=e.target;
          if(cb && cb.classList && cb.classList.contains("rowsel"))
            self.toggle(cb.getAttribute("data-key"), cb.checked);
        });
      }
    }
    thead.innerHTML=""; thead.appendChild(tr);
    this.thRow=tr; this.built=true;
  };
  SortTable.prototype.clickSort=function(c){
    if(this.sort.key===c.k) this.sort.dir*=-1;
    else this.sort={key:c.k, dir:(c.sort==="string"?1:-1)};
    this.render();
  };
  SortTable.prototype.syncAria=function(){
    var self=this;
    if(!this.thRow) return;
    Array.prototype.forEach.call(this.thRow.querySelectorAll("th"), function(th){
      if(th.dataset.k===self.sort.key) th.setAttribute("aria-sort", self.sort.dir<0?"descending":"ascending");
      else th.removeAttribute("aria-sort");
    });
  };
  SortTable.prototype.render=function(){
    if(!this.built) this.buildHead();
    var self=this, rows=this.opts.getRows()||[];
    var k=this.sort.key;
    if(k){
      var col=null; this.columns.forEach(function(c){ if(c.k===k) col=c; });
      var typ=col? col.sort : "number";
      rows=rows.slice().sort(function(a,b){
        var av=a[k], bv=b[k];
        if(typ==="string"){
          av=(av==null?"":String(av)).toLowerCase(); bv=(bv==null?"":String(bv)).toLowerCase();
          return self.sort.dir*(av<bv?-1:av>bv?1:0);
        }
        return self.sort.dir*((+av||0)-(+bv||0));
      });
    }
    var total=rows.length;
    var max=this.opts.max||total;
    var shown=rows.slice(0,max);
    this.rowMap={}; rows.forEach(function(r){ self.rowMap[self.keyOf(r)]=r; });
    var tbody=this.table.querySelector("tbody");
    tbody.innerHTML=shown.map(function(r){
      var key=self.keyOf(r), sel=self.selectable && !!self.selected[key];
      var lead=self.selectable? '<td class="selcol"><input type="checkbox" class="rowsel"'+(sel?" checked":"")+' data-key="'+esc(key)+'"></td>' : "";
      return "<tr"+(sel?' class="selrow"':"")+">"+lead+self.columns.map(function(c){
        var cell=c.render? c.render(r) : esc(r[c.k]);
        return "<td"+(c.num?' class="num"':"")+">"+cell+"</td>";
      }).join("")+"</tr>";
    }).join("");
    this.syncAria();
    if(this.opts.onCount) this.opts.onCount(total, shown.length);
    if(this.selectable) this.renderCompareBar();
  };
  SortTable.prototype.keyOf=function(r){
    var kf=this.opts.rowKey;
    if(kf) return String(typeof kf==="function"? kf(r) : r[kf]);
    var k0=this.columns[0].k;
    return String(r[k0]==null?"":r[k0]);
  };
  SortTable.prototype.toggle=function(key, on){
    if(on){ if(!this.selected[key]){ this.selected[key]=this.rowMap[key]||{}; this.selOrder.push(key); } }
    else { delete this.selected[key]; var i=this.selOrder.indexOf(key); if(i>=0) this.selOrder.splice(i,1); }
    this.render();
  };
  SortTable.prototype.ensureCompareUI=function(){
    if(this.cmpUI) return this.cmpUI;
    var self=this, wrap=this.table.closest? this.table.closest(".tablewrap") : null;
    var anchor=wrap||this.table;
    var bar=document.createElement("div"); bar.className="cmp-bar hidden";
    var info=document.createElement("span"); info.className="cmp-info";
    var go=document.createElement("button"); go.type="button"; go.className="btn ghost cmp-go"; go.textContent="Compare";
    var clr=document.createElement("button"); clr.type="button"; clr.className="btn ghost cmp-clear"; clr.textContent="Clear";
    bar.appendChild(info); bar.appendChild(go); bar.appendChild(clr);
    var panel=document.createElement("div"); panel.className="cmp-panel hidden";
    anchor.parentNode.insertBefore(bar, anchor.nextSibling);
    bar.parentNode.insertBefore(panel, bar.nextSibling);
    go.addEventListener("click", function(){ self.renderComparePanel(); });
    clr.addEventListener("click", function(){ self.clearSelection(); });
    this.cmpUI={bar:bar, info:info, panel:panel, go:go};
    return this.cmpUI;
  };
  SortTable.prototype.renderCompareBar=function(){
    if(!this.selectable) return;
    var ui=this.ensureCompareUI(), n=this.selOrder.length;
    if(n>0){
      ui.bar.classList.remove("hidden");
      ui.info.textContent=n+(n===1?" row selected":" rows selected");
      ui.go.disabled=n<2; ui.go.textContent=n<2?"Select 2+ to compare":"Compare "+n;
    } else {
      ui.bar.classList.add("hidden"); ui.panel.classList.add("hidden");
    }
  };
  SortTable.prototype.clearSelection=function(){
    this.selected={}; this.selOrder=[];
    if(this.cmpUI) this.cmpUI.panel.classList.add("hidden");
    this.render();
  };
  SortTable.prototype.renderComparePanel=function(){
    var self=this, ui=this.ensureCompareUI(), keys=this.selOrder.slice();
    if(keys.length<2){ ui.panel.classList.add("hidden"); return; }
    var rows=keys.map(function(k){ return self.selected[k]||self.rowMap[k]||{}; });
    var head='<th class="cmp-field">Field</th>'+rows.map(function(r){
      return "<th>"+esc(self.keyOf(r))+"</th>";
    }).join("");
    var body=this.columns.map(function(c){
      var cells=rows.map(function(r){
        var v=c.render? c.render(r) : esc(r[c.k]);
        return "<td"+(c.num?' class="num"':"")+">"+v+"</td>";
      }).join("");
      return '<tr><td class="cmp-field">'+esc(c.t)+"</td>"+cells+"</tr>";
    }).join("");
    ui.panel.innerHTML='<div class="cmp-head"><strong>Comparison</strong>'+
      '<button type="button" class="btn ghost cmp-close">Close</button></div>'+
      '<div class="cmp-scroll"><table class="cmp-table"><thead><tr>'+head+
      "</tr></thead><tbody>"+body+"</tbody></table></div>";
    ui.panel.classList.remove("hidden");
    var cl=ui.panel.querySelector(".cmp-close");
    if(cl) cl.addEventListener("click", function(){ ui.panel.classList.add("hidden"); });
  };

  function kpiCards(host, cards){
    $(host).innerHTML = cards.map(function(c){
      return '<div class="card kpi"><h3>'+esc(c.h)+'</h3><div class="num">'+esc(c.n)+'</div>'+
        '<div class="sub">'+esc(c.s||"")+'</div>'+
        (c.bar!=null?'<div class="bar"><span style="width:'+Math.min(100,c.bar)+'%"></span></div>':'')+'</div>';
    }).join("");
  }

  // Declarative Agent usage analysis
  function computeDecl(){
    return RAW.declAgents.rows.map(function(r){
      var lic = num(getUserField(r, ["Active users (licensed)", "Active users licensed", "licensedActiveUsers"]));
      var unlic = num(getUserField(r, ["Active users (unlicensed)", "Active users unlicensed", "unlicensedActiveUsers"]));
      var id = getUserField(r, ["Agent ID", "AgentID", "id"]);
      var name = (getUserField(r, ["Agent name", "AgentName", "name"]) || "").trim() || "(unnamed)";
      var creator = getUserField(r, ["Creator type", "CreatorType", "creator"]);
      var responses = num(getUserField(r, ["Responses sent to users", "responses"]));
      var last = getUserField(r, ["Last activity date (UTC)", "Last Activity Date", "lastActivityDate"]);
      return {
        id: id,
        name: name,
        creator: creator,
        lic: lic, unlic: unlic, active: lic+unlic,
        responses: responses,
        last: last
      };
    });
  }
  function renderDecl(){
    var rows=computeDecl();
    var totResp=0, totActive=0, withUse=0, top=null;
    rows.forEach(function(a){
      totResp+=a.responses; totActive+=a.active;
      if(a.responses>0 || a.active>0) withUse++;
      if(!top || a.responses>top.responses) top=a;
    });
    kpiCards("cpxDeclKpis",[
      {h:"Declarative agents", n:fmt(rows.length), s:"in the 30-day usage report"},
      {h:"Agents with usage", n:fmt(withUse), s:pct(withUse,rows.length)+"% had any activity", bar:pct(withUse,rows.length)},
      {h:"Responses sent", n:fmt(totResp), s:"to users across all agents"},
      {h:"Top agent", n:top?clip(top.name,18):"—", s:top?(fmt(top.responses)+" responses"):""}
    ]);
    var bars=rows.slice().sort(function(a,b){return b.responses-a.responses;}).slice(0,12)
              .filter(function(a){return a.responses>0;})
              .map(function(a,i){return {label:a.name, v:a.responses, color:COLORS[i%COLORS.length]};});
    $("cpxDeclBars").innerHTML = bars.length? hBar(bars,{labelW:150,fmt:function(v){return fmt(v);},
      tipv:function(d){return fmt(d.v)+" responses";}}) : '<p class="muted">No agent responses recorded in this window.</p>';

    var cols=[
      {k:"name", t:"Agent", sort:"string", render:function(r){return esc(r.name);}},
      {k:"creator", t:"Creator type", sort:"string", render:function(r){return esc(r.creator||"—");}},
      {k:"lic", t:"Licensed", num:true, sort:"number", render:function(r){return fmt(r.lic);}},
      {k:"unlic", t:"Unlicensed", num:true, sort:"number", render:function(r){return fmt(r.unlic);}},
      {k:"active", t:"Active users", num:true, sort:"number", render:function(r){return fmt(r.active);}},
      {k:"responses", t:"Responses", num:true, sort:"number", render:function(r){return fmt(r.responses);}},
      {k:"last", t:"Last activity", sort:"string", render:function(r){return esc(r.last||"—");}}
    ];
    TBL.decl=new SortTable($("cpxDeclTable"), cols, {
      initSort:{key:"responses",dir:-1}, max:300, rowKey:"name",
      getRows:function(){
        var q=($("cpxDeclSearch").value||"").toLowerCase().trim();
        return rows.filter(function(a){ return !q || a.name.toLowerCase().indexOf(q)>=0 || (a.creator||"").toLowerCase().indexOf(q)>=0; });
      },
      onCount:function(total,shown){ $("cpxDeclCount").textContent = fmt(total)+" agents"+(total>shown?" (top "+shown+")":""); }
    });
    TBL.decl.render();
    $("cpxDeclWrap").classList.remove("hidden");
  }

  // Exhaustive Agent inventory analytics
  function computeAll(){
    return RAW.agents.rows.map(function(r){
      var name = (getUserField(r, ["Name", "name"]) || "").trim() || "(unnamed)";
      var status = getUserField(r, ["Status", "status"]);
      var channel = getUserField(r, ["Channel", "channel"]);
      var created = (getUserField(r, ["Date created", "DateCreated", "created"]) || "").slice(0, 10);
      var modified = (getUserField(r, ["Last Modified", "LastModified", "modified"]) || "").slice(0, 10);
      var publisher = getUserField(r, ["Publisher", "publisher"]);
      var pubType = getUserField(r, ["Publisher Type", "PublisherType", "publisherType"]);
      var version = getUserField(r, ["Version", "version"]);
      var owner = getUserField(r, ["Owner", "owner"]);
      var platform = getUserField(r, ["Platform", "platform"]);
      var sensitivity = getUserField(r, ["Sensitivity", "sensitivity"]);
      return {
        name: name,
        status: status,
        channel: channel,
        created: created,
        modified: modified,
        publisher: publisher,
        pubType: pubType,
        version: version,
        owner: owner,
        platform: platform,
        sensitivity: sensitivity
      };
    });
  }
  function tally(rows, field){
    var m={}; rows.forEach(function(r){ var v=(r[field]||"").trim()||"(none)"; m[v]=(m[v]||0)+1; });
    return Object.keys(m).map(function(k){return {label:k, v:m[k]};}).sort(function(a,b){return b.v-a.v;});
  }
  function renderAll(){
    var rows=computeAll();
    var plats=tally(rows,"platform"), stats=tally(rows,"status");
    var avail=0; rows.forEach(function(r){ if((r.status||"").toLowerCase()==="available") avail++; });
    kpiCards("cpxAllKpis",[
      {h:"Total agents", n:fmt(rows.length), s:"declarative + non-declarative"},
      {h:"Available", n:fmt(avail), s:pct(avail,rows.length)+"% published & available", bar:pct(avail,rows.length)},
      {h:"Platforms", n:fmt(plats.length), s:"distinct build platforms"},
      {h:"Top platform", n:plats[0]?clip(plats[0].label,16):"—", s:plats[0]?(fmt(plats[0].v)+" agents"):""}
    ]);
    $("cpxPlatBars").innerHTML = hBar(plats.slice(0,12).map(function(d,i){return {label:d.label,v:d.v,color:COLORS[i%COLORS.length]};}),
      {labelW:170,fmt:function(v){return fmt(v);},tipv:function(d){return fmt(d.v)+" agents";}});
    $("cpxStatusBars").innerHTML = hBar(stats.map(function(d,i){return {label:d.label,v:d.v,color:COLORS[i%COLORS.length]};}),
      {labelW:120,fmt:function(v){return fmt(v);},tipv:function(d){return fmt(d.v)+" agents";}});

    var cols=[
      {k:"name", t:"Agent", sort:"string", render:function(r){return esc(r.name);}},
      {k:"status", t:"Status", sort:"string", render:function(r){return esc(r.status||"—");}},
      {k:"channel", t:"Channel", sort:"string", render:function(r){return esc(r.channel||"—");}},
      {k:"platform", t:"Platform", sort:"string", render:function(r){return esc(r.platform||"—");}},
      {k:"pubType", t:"Publisher type", sort:"string", render:function(r){return esc(r.pubType||"—");}},
      {k:"owner", t:"Owner", sort:"string", render:function(r){return esc(r.owner||"—");}},
      {k:"version", t:"Version", sort:"string", render:function(r){return esc(r.version||"—");}},
      {k:"created", t:"Created", sort:"string", render:function(r){return esc(r.created||"—");}},
      {k:"modified", t:"Last modified", sort:"string", render:function(r){return esc(r.modified||"—");}}
    ];
    TBL.all=new SortTable($("cpxAllTable"), cols, {
      initSort:{key:"modified",dir:-1}, max:400, rowKey:"name",
      getRows:function(){
        var q=($("cpxAllSearch").value||"").toLowerCase().trim();
        return rows.filter(function(r){
          if(!q) return true;
          return r.name.toLowerCase().indexOf(q)>=0 || (r.owner||"").toLowerCase().indexOf(q)>=0 ||
                 (r.platform||"").toLowerCase().indexOf(q)>=0 || (r.status||"").toLowerCase().indexOf(q)>=0;
        });
      },
      onCount:function(total,shown){ $("cpxAllCount").textContent = fmt(total)+" agents"+(total>shown?" (top "+shown+")":""); }
    });
    TBL.all.render();
    $("cpxAllWrap").classList.remove("hidden");
  }

  // Copilot Chat adoption & prompts analysis
  function pickPeriodRow(parsed, anyCol, appSuffix){
    // Select the longest reporting period from row list
    var best=null;
    parsed.rows.forEach(function(r){
      var p = num(getUserField(r, ["Report period", "reportPeriod"]));
      if(!best || p>best.period){
        var refresh = getUserField(r, ["Report refresh date", "reportRefreshDate", "refreshDate"]);
        var anyVal = num(getUserField(r, [anyCol]));
        var o={ period:p, refresh:refresh, any:anyVal, app:{} };
        EDP_APPS.forEach(function(a){
          o.app[a.key] = num(getUserField(r, [a.key + appSuffix, a.short + appSuffix]));
        });
        best=o;
      }
    });
    return best;
  }
  function renderChatTop(){
    var adopt = RAW.chatAdopt? pickPeriodRow(RAW.chatAdopt,"Any app active users"," active users") : null;
    var prompts = RAW.chatPrompts? pickPeriodRow(RAW.chatPrompts,"Any app prompts submitted"," prompts submitted") : null;
    if(!adopt && !prompts) return;
    var period=(adopt&&adopt.period)||(prompts&&prompts.period)||0;
    var refresh=(adopt&&adopt.refresh)||(prompts&&prompts.refresh)||"";
    var cards=[];
    if(adopt) cards.push({h:"Any-app active users", n:fmt(adopt.any), s:"Copilot Chat over "+period+" days"});
    if(prompts) cards.push({h:"Prompts submitted", n:fmt(prompts.any), s:"across all apps, "+period+" days"});
    if(adopt && prompts) cards.push({h:"Prompts per active user", n:fmt(adopt.any? prompts.any/adopt.any : 0), s:"average over the window"});
    if(adopt){
      var topApp=null; EDP_APPS.forEach(function(a){ if(!topApp || adopt.app[a.key]>adopt.app[topApp.key]) topApp=a; });
      cards.push({h:"Top surface", n:topApp?topApp.short:"—", s:topApp?(fmt(adopt.app[topApp.key])+" active users"):""});
    }
    kpiCards("cpxChatKpis", cards);

    if(adopt){
      var ab=EDP_APPS.map(function(a){return {label:a.short,v:adopt.app[a.key],color:a.color};})
              .filter(function(d){return d.v>0;}).sort(function(a,b){return b.v-a.v;});
      $("cpxAdoptBars").innerHTML = hBar(ab,{labelW:140,fmt:function(v){return fmt(v);},tipv:function(d){return fmt(d.v)+" active users";}});
      $("cpxAdoptNote").textContent = "Active users per surface · "+period+"-day window"+(refresh?(" · refreshed "+refresh):"");
    } else { $("cpxAdoptBars").innerHTML='<p class="muted">Drop the adoption-by-period export to see this.</p>'; }

    if(prompts){
      var pb=EDP_APPS.map(function(a){return {label:a.short,v:prompts.app[a.key],color:a.color};})
              .filter(function(d){return d.v>0;}).sort(function(a,b){return b.v-a.v;});
      $("cpxPromptBars").innerHTML = hBar(pb,{labelW:140,fmt:function(v){return fmtShort(v);},tipv:function(d){return fmt(d.v)+" prompts";}});
      $("cpxPromptNote").textContent = "Prompts submitted per surface · "+period+"-day window";
    } else { $("cpxPromptBars").innerHTML='<p class="muted">Drop the prompts-by-period export to see this.</p>'; }

    $("cpxChatTopWrap").classList.remove("hidden");
  }

  // Copilot Chat per-person usage details
  function computeChatUsers(){
    return RAW.chatUsers.rows.map(function(r){
      var upn = getUserField(r, ["User principal name", "userPrincipalName", "upn"]);
      var name = getUserField(r, ["Display name", "displayName"]) || upn;
      var period = num(getUserField(r, ["Report period", "reportPeriod"]));
      var prompts = num(getUserField(r, ["Prompts submitted", "promptsSubmitted", "prompts"]));
      var days = num(getUserField(r, ["Active usage days", "activeUsageDays", "days"]));
      var last = getUserField(r, ["Last activity date", "lastActivityDate", "last"]);
      var o={
        name: name,
        upn: upn,
        period: period,
        prompts: prompts,
        days: days,
        last: last,
        app:{}
      };
      EDP_APPS.forEach(function(a){
        o.app[a.key] = getUserField(r, [
          "Last activity date of " + a.key + " (UTC)",
          "Last activity date of " + a.short + " (UTC)",
          a.key + " Last Activity Date",
          a.short + " Last Activity Date",
          a.key,
          a.short
        ]);
      });
      return o;
    });
  }
  function renderChatUsers(){
    var rows=computeChatUsers();
    var totPrompts=0, maxDays=0, period=0;
    rows.forEach(function(u){ totPrompts+=u.prompts; if(u.days>maxDays) maxDays=u.days; if(u.period>period) period=u.period; });
    var topUser=rows.slice().sort(function(a,b){return b.prompts-a.prompts;})[0];
    kpiCards("cpxChatKpis2",[
      {h:"Chat end users", n:fmt(rows.length), s:"in the "+period+"-day detail export"},
      {h:"Total prompts", n:fmt(totPrompts), s:"submitted across these users"},
      {h:"Avg prompts / user", n:fmt(rows.length? totPrompts/rows.length : 0), s:"mean over the window"},
      {h:"Most active user", n:topUser?clip(topUser.name,16):"—", s:topUser?(fmt(topUser.prompts)+" prompts"):""}
    ]);
    var reach=EDP_APPS.map(function(a){
      var c=0; rows.forEach(function(u){ if(u.app[a.key]) c++; });
      return {label:a.short, v:c, color:a.color};
    }).filter(function(d){return d.v>0;}).sort(function(a,b){return b.v-a.v;});
    $("cpxChatReach").innerHTML = hBar(reach,{labelW:140,fmt:function(v){return fmt(v);},tipv:function(d){return fmt(d.v)+" people";}});

    var cols=[
      {k:"name", t:"Person", sort:"string", render:function(u){return esc(u.name);}},
      {k:"prompts", t:"Prompts", num:true, sort:"number", render:function(u){return fmt(u.prompts);}},
      {k:"days", t:"Active days", num:true, sort:"number", render:function(u){return fmt(u.days);}},
      {k:"last", t:"Last activity", sort:"string", render:function(u){return esc(u.last||"—");}}
    ];
    EDP_APPS.forEach(function(a){
      cols.push({k:"d_"+a.key, t:a.short, sort:"string", render:function(u){
        var v=u.app[a.key]; return v? esc(v) : '<span class="muted">—</span>';
      }});
    });
    TBL.chatUsers=new SortTable($("cpxChatUsrTable"), cols, {
      initSort:{key:"prompts",dir:-1}, max:300, rowKey:"upn",
      getRows:function(){
        var q=($("cpxChatUsrSearch").value||"").toLowerCase().trim();
        return rows.filter(function(u){
          return !q || u.name.toLowerCase().indexOf(q)>=0 || u.upn.toLowerCase().indexOf(q)>=0;
        }).map(function(u){
          var o={name:u.name, upn:u.upn, prompts:u.prompts, days:u.days, last:u.last, app:u.app};
          EDP_APPS.forEach(function(a){ o["d_"+a.key]=u.app[a.key]||""; });
          return o;
        });
      },
      onCount:function(total,shown){ $("cpxChatUsrCount").textContent = fmt(total)+" people"+(total>shown?" (top "+shown+")":""); }
    });
    TBL.chatUsers.render();
    $("cpxChatUsersWrap").classList.remove("hidden");
  }

  // Drag-and-drop file router
  function nameToKey(fn){
    fn=fn.toLowerCase();
    if(fn.indexOf("declarativeagents")>=0) return "declAgents";
    if(fn.indexOf("adoptionbyperiod")>=0) return "chatAdopt";
    if(fn.indexOf("promptssubmittedbyperiod")>=0) return "chatPrompts";
    // Route to Copilot Chat user list if EDP suffix present
    if(fn.indexOf("activityuserdetail")>=0 && fn.indexOf("edp")>=0) return "chatUsers";
    if(fn.indexOf("agents")>=0) return "agents";   // exhaustive list, checked last
    return null;
  }
  function rerender(){
    if(RAW.declAgents) renderDecl();
    else { $("cpxDeclWrap").classList.add("hidden"); }
    if(RAW.agents) renderAll();
    else { $("cpxAllWrap").classList.add("hidden"); }
    if(RAW.chatAdopt || RAW.chatPrompts) renderChatTop();
    else { $("cpxChatTopWrap").classList.add("hidden"); }
    if(RAW.chatUsers) renderChatUsers();
    else { $("cpxChatUsersWrap").classList.add("hidden"); }
    updateChecklists();

    var hasAgents = !!(RAW.declAgents || RAW.agents);
    var hasChat = !!(RAW.chatAdopt || RAW.chatPrompts || RAW.chatUsers);

    var agentPlaceholder = $("cpxAgentsPlaceholder");
    if(agentPlaceholder) agentPlaceholder.classList.toggle("hidden", hasAgents);

    var chatPlaceholder = $("cpxChatPlaceholder");
    if(chatPlaceholder) chatPlaceholder.classList.toggle("hidden", hasChat);
  }
  function ingest(fileList){
    var arr=Array.prototype.slice.call(fileList);
    var jobs=arr.map(function(f){
      var key=nameToKey(f.name);
      if(!key) return Promise.resolve();
      return f.text().then(function(txt){ RAW[key]=parseCSV(txt); });
    });
    return Promise.all(jobs).then(rerender);
  }

  var AGENT_FILES = { declAgents:"Declarative agents — 30-day usage", agents:"All agents — inventory" };
  var CHAT_FILES  = { chatAdopt:"Adoption by app (by period)", chatPrompts:"Prompts submitted by app", chatUsers:"End-user usage details" };
  function checklistHtml(map){
    return Object.keys(map).map(function(k){
      var ok=!!RAW[k];
      return '<li class="'+(ok?"ok":"pending")+'">'+(ok?"✔":"○")+" "+esc(map[k])+"</li>";
    }).join("");
  }
  function updateChecklists(){
    if($("cpxAgentsCheck")) $("cpxAgentsCheck").innerHTML=checklistHtml(AGENT_FILES);
    if($("cpxChatCheck"))   $("cpxChatCheck").innerHTML=checklistHtml(CHAT_FILES);
  }

  function wireDrop(zoneId, inputId){
    var dz=$(zoneId), fi=$(inputId);
    if(!dz||!fi) return;
    dz.addEventListener("click",function(){ fi.click(); });
    fi.addEventListener("change",function(){ ingest(fi.files); });
    ["dragover","dragenter"].forEach(function(ev){
      dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.add("drag"); });
    });
    ["dragleave","drop"].forEach(function(ev){
      dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.remove("drag"); });
    });
    dz.addEventListener("drop",function(e){
      if(e.dataTransfer&&e.dataTransfer.files) ingest(e.dataTransfer.files);
    });
  }

  // Setup button to reveal uploader drawer
  function makeReveal(which, label){
    var load=$("cpx"+(which==="agents"?"Agents":"Chat")+"Load");
    if(!load || !load.parentNode) return;
    var id="cpx"+(which==="agents"?"Agents":"Chat")+"Reveal";
    if($(id)) return;
    var btn=document.createElement("button");
    btn.type="button"; btn.id=id; btn.className="btn ghost hidden";
    btn.textContent=label;
    btn.addEventListener("click",function(){ setLoadVisible(which, true); });
    load.parentNode.insertBefore(btn, load);
  }
  function init(){
    if(!$("cptab-agents")) return;   // panels not on the page
    ["cpxDeclSearch","cpxAllSearch"].forEach(function(id){
      var el=$(id); if(el) el.addEventListener("input",function(){
        if(id==="cpxDeclSearch" && TBL.decl) TBL.decl.render();
        if(id==="cpxAllSearch" && TBL.all) TBL.all.render();
      });
    });
    var cu=$("cpxChatUsrSearch"); if(cu) cu.addEventListener("input",function(){ if(TBL.chatUsers) TBL.chatUsers.render(); });
    updateChecklists();
  }

  if(document.readyState==="loading") document.addEventListener("DOMContentLoaded", init);
  else init();

  function status(){
    return { declAgents:!!RAW.declAgents, agents:!!RAW.agents,
             chatAdopt:!!RAW.chatAdopt, chatPrompts:!!RAW.chatPrompts, chatUsers:!!RAW.chatUsers };
  }

  function clear(){
    RAW = {};
    TBL = {};
    rerender();
  }

  // expose for diagnostics / optional external boot
  window.CopilotExtras = { ingest:ingest, rerender:rerender, status:status, clear:clear };
})();
