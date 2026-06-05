"use strict";
/* Microsoft Teams adoption board.
   Self-contained module: loads the single Teams user-activity export
   (Teams_UserActivityUserDetail), computes leadership-facing datasets, and
   renders into the #tm-app container. Element IDs are prefixed "tm" so this
   can coexist with the Viva and Copilot boards on the same page. */
(function(){
  // The export shares a base name plus a period suffix. Microsoft's portal
  // writes the suffix a few different ways ( _D180 , _'D180' , _180 ), so we
  // probe several styles and auto-detect whichever window is present.
  var FOLDER = "Usage%20Reports/";
  var BASE = "Teams_UserActivityUserDetail";
  var PERIODS = [30,60,90,180];
  function suffixes(p){ return ["_D"+p, "_'D"+p+"'", "_"+p]; }
  var LABEL = "Teams user activity";

  var detectedPeriod = 0;
  var COLORS = ["var(--c1)","var(--c2)","var(--c3)","var(--c4)","var(--c5)","var(--c6)","var(--c7)","var(--c8)"];
  var RAW = {}, DATA = {}, usrTbl = null;
  var hasData = false;   // true once a dashboard has rendered (enables "Back to current reports")

  /* ---------- helpers ---------- */
  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,function(c){
    return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]; }); }
  function num(v){ var n=parseFloat(v); return isFinite(n)?n:0; }
  function fmt(n){ n=Math.round(n); return n.toLocaleString("en-US"); }
  function fmtShort(n){
    n=Math.round(n);
    if(Math.abs(n)>=1e6) return (n/1e6).toFixed(n%1e6?1:0)+"M";
    if(Math.abs(n)>=1e3) return (n/1e3).toFixed(n%1e3?1:0)+"k";
    return String(n);
  }
  function pct(a,b){ return b? Math.round(a/b*1000)/10 : 0; }
  function clip(s,n){ s=String(s); return s.length>n? s.slice(0,n-1)+"…" : s; }
  // Friendly short name from a UPN (no Display Name column in this export).
  function shortName(upn){
    var local=String(upn||"").split("@")[0];
    if(!local) return upn||"";
    return local.split(/[._-]/).map(function(p){
      return p? p.charAt(0).toUpperCase()+p.slice(1) : p;
    }).join(" ");
  }
  function hrs(sec){ return sec/3600; }
  function fmtHrs(sec){
    var h=sec/3600;
    if(h>=1000) return fmtShort(h)+" h";
    if(h>=10) return Math.round(h)+" h";
    return (Math.round(h*10)/10)+" h";
  }

  /* ---------- floating tooltip (own element id so it never clashes) ---------- */
  var Tip = {
    el:null,
    ensure:function(){
      if(!this.el){ this.el=document.createElement("div"); this.el.id="tmtip"; document.body.appendChild(this.el); }
      return this.el;
    },
    show:function(html,x,y){
      var e=this.ensure(); e.innerHTML=html; e.style.display="block";
      var w=e.offsetWidth, h=e.offsetHeight, vw=window.innerWidth, vh=window.innerHeight;
      var left=x+14, top=y+14;
      if(left+w>vw-8) left=x-w-14; if(left<8) left=8;
      if(top+h>vh-8) top=y-h-14; if(top<8) top=8;
      e.style.left=left+"px"; e.style.top=top+"px";
    },
    hide:function(){ if(this.el) this.el.style.display="none"; }
  };

  /* ---------- robust CSV parser ---------- */
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

  /* ---------- compute derived datasets ---------- */
  function compute(){
    var rows = RAW.detail.rows;
    var users = rows.map(function(r){
      var upn=r["User Principal Name"]||r["User Id"]||"";
      var teamChat=num(r["Team Chat Message Count"]);
      var privChat=num(r["Private Chat Message Count"]);
      var calls=num(r["Call Count"]);
      var meetings=num(r["Meeting Count"]);
      var orgCount=num(r["Meetings Organized Count"]);
      var attCount=num(r["Meetings Attended Count"]);
      var post=num(r["Post Messages"]);
      var reply=num(r["Reply Messages"]);
      var urgent=num(r["Urgent Messages"]);
      var audioS=num(r["Audio Duration In Seconds"]);
      var videoS=num(r["Video Duration In Seconds"]);
      var shareS=num(r["Screen Share Duration In Seconds"]);
      return {
        upn:upn, name:shortName(upn),
        last:r["Last Activity Date"]||"",
        licensed:(r["Is Licensed"]||"").toLowerCase()==="yes",
        teamChat:teamChat, privChat:privChat, chat:teamChat+privChat,
        calls:calls, meetings:meetings, org:orgCount, att:attCount,
        adHocOrg:num(r["Ad Hoc Meetings Organized Count"]),
        adHocAtt:num(r["Ad Hoc Meetings Attended Count"]),
        schedOrg:num(r["Scheduled One-time Meetings Organized Count"]),
        schedAtt:num(r["Scheduled One-time Meetings Attended Count"]),
        recurOrg:num(r["Scheduled Recurring Meetings Organized Count"]),
        recurAtt:num(r["Scheduled Recurring Meetings Attended Count"]),
        post:post, reply:reply, urgent:urgent,
        audioS:audioS, videoS:videoS, shareS:shareS,
        commS:audioS+videoS+shareS,
        total:teamChat+privChat+calls+meetings+post+reply
      };
    });

    var T = {
      teamChat:0, privChat:0, calls:0, meetings:0, org:0, att:0,
      adHocAtt:0, schedAtt:0, recurAtt:0,
      post:0, reply:0, urgent:0, audioS:0, videoS:0, shareS:0
    };
    var active=0, licensed=0;
    users.forEach(function(u){
      if(u.licensed) licensed++;
      if(u.total>0 || u.commS>0) active++;
      T.teamChat+=u.teamChat; T.privChat+=u.privChat; T.calls+=u.calls;
      T.meetings+=u.meetings; T.org+=u.org; T.att+=u.att;
      T.adHocAtt+=u.adHocAtt; T.schedAtt+=u.schedAtt; T.recurAtt+=u.recurAtt;
      T.post+=u.post; T.reply+=u.reply; T.urgent+=u.urgent;
      T.audioS+=u.audioS; T.videoS+=u.videoS; T.shareS+=u.shareS;
    });

    var first = rows[0]||{};
    var refresh = first["Report Refresh Date"]||"";
    var windowDays = num(first["Report Period"]) || detectedPeriod || 30;

    DATA = {
      users:users, totals:T,
      totalUsers:users.length, licensed:licensed, active:active,
      refresh:refresh, windowDays:windowDays
    };
  }

  /* ---------- bar + donut (data-tip driven hover) ---------- */
  function hBarChart(items, opt){
    opt=opt||{};
    var rowH=opt.rowH||26, gap=8, pl=opt.labelW||180, pr=54;
    var W=opt.w||520, H=items.length*(rowH+gap)+8;
    var max=Math.max.apply(null, items.map(function(d){return d.v;}).concat([1]));
    var iw=W-pl-pr;
    var out='<svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="bar chart">';
    items.forEach(function(d,i){
      var y=i*(rowH+gap)+6, w=Math.max(2,(d.v/max)*iw);
      var col=d.color||opt.color||COLORS[i%COLORS.length];
      var vlabel=opt.fmt? opt.fmt(d.v) : fmtShort(d.v);
      var tipv=opt.tipv? opt.tipv(d) : fmt(d.v);
      out+='<rect x="'+pl+'" y="'+y+'" width="'+w.toFixed(1)+'" height="'+rowH+'" rx="5" fill="'+col+'"'+
           ' data-tip="'+esc(d.label)+'" data-tipv="'+esc(tipv)+'" data-tipc="'+col+'"></rect>';
      out+='<text x="'+(pl-8)+'" y="'+(y+rowH/2+4)+'" text-anchor="end" font-size="12" fill="var(--ink)">'+esc(opt.clip!==false? clip(d.label,26):d.label)+'</text>';
      out+='<text x="'+(pl+w+7).toFixed(1)+'" y="'+(y+rowH/2+4)+'" font-size="11.5" fill="var(--muted)">'+esc(vlabel)+'</text>';
    });
    return out+'</svg>';
  }
  function donut(segs, opt){
    opt=opt||{}; var size=opt.size||190, r=70, cx=size/2, cy=size/2, sw=26;
    var total=segs.reduce(function(a,s){return a+s.v;},0)||1;
    var circ=2*Math.PI*r, out='<svg viewBox="0 0 '+size+' '+size+'" role="img" aria-label="donut chart">';
    out+='<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="var(--chip)" stroke-width="'+sw+'"/>';
    var off=0;
    segs.forEach(function(s,i){
      var frac=s.v/total, len=frac*circ, col=s.color||COLORS[i%COLORS.length];
      out+='<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="'+col+'" stroke-width="'+sw+
           '" stroke-dasharray="'+len.toFixed(2)+' '+(circ-len).toFixed(2)+'" stroke-dashoffset="'+(-off).toFixed(2)+
           '" transform="rotate(-90 '+cx+' '+cy+')"'+
           ' data-tip="'+esc(s.label)+'" data-tipv="'+fmt(s.v)+' ('+pct(s.v,total)+'%)" data-tipc="'+col+'"></circle>';
      off+=len;
    });
    out+='<text x="'+cx+'" y="'+(cy-2)+'" text-anchor="middle" font-size="22" font-weight="700" fill="var(--ink)">'+fmtShort(opt.center!=null?opt.center:total)+'</text>';
    out+='<text x="'+cx+'" y="'+(cy+16)+'" text-anchor="middle" font-size="10.5" fill="var(--muted)">'+esc(opt.centerLabel||"")+'</text>';
    return out+'</svg>';
  }
  function legend(items){
    return items.map(function(s,i){
      return '<span><i class="swatch" style="background:'+(s.color||COLORS[i%COLORS.length])+'"></i>'+esc(s.label)+'</span>';
    }).join("");
  }

  /* ---------- resizable, sortable table (same engine as the other boards) ---------- */
  function addResizer(th, handle, table){
    var startX=0, startW=0, active=false;
    handle.addEventListener("pointerdown", function(e){
      e.preventDefault(); e.stopPropagation();
      active=true; startX=e.clientX; startW=th.getBoundingClientRect().width;
      // Freeze every other column to its current pixel width and let the table
      // grow past its container, so widening one column scrolls (via the
      // .tablewrap overflow) instead of squishing the rest.
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
    thead.innerHTML=""; thead.appendChild(tr);
    this.thRow=tr; this.built=true;
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

  /* ---------- renderers ---------- */
  function renderHeader(){
    var wd=DATA.windowDays;
    $("tmSubtitle").innerHTML = "Microsoft Teams usage across The Mosaic Company &middot; rolling "+wd+"-day window";
    var pills=[];
    pills.push('<span class="pill">Window: '+wd+' days</span>');
    if(DATA.refresh) pills.push('<span class="pill">Data refreshed '+esc(DATA.refresh)+'</span>');
    pills.push('<span class="pill">'+fmt(DATA.licensed)+' licensed users</span>');
    pills.push('<span class="pill">'+fmt(DATA.active)+' active users</span>');
    $("tmMetaPills").innerHTML=pills.join("");
    $("tmFootNote").innerHTML = "Generated from the Microsoft Teams user-activity export &middot; "+wd+
      "-day window &middot; an active user is anyone with at least one recorded chat, call, meeting, or channel message.";
  }

  function renderKpis(){
    var d=DATA, T=d.totals;
    var adoption=pct(d.active,d.licensed);
    var chatTot=T.teamChat+T.privChat;
    var commHrs=hrs(T.audioS+T.videoS+T.shareS);
    var cards=[
      {h:"Active users", n:fmt(d.active), s:adoption+"% of "+fmt(d.licensed)+" licensed", bar:adoption},
      {h:"Chat messages", n:fmtShort(chatTot), s:fmt(T.teamChat)+" team · "+fmt(T.privChat)+" private", bar:pct(T.teamChat,chatTot)},
      {h:"Meetings attended", n:fmtShort(T.att), s:fmt(T.org)+" organized", bar:pct(T.org,T.att)},
      {h:"Calls", n:fmtShort(T.calls), s:"across all users", bar:100},
      {h:"Channel messages", n:fmtShort(T.post+T.reply), s:fmt(T.post)+" posts · "+fmt(T.reply)+" replies", bar:pct(T.post,T.post+T.reply)},
      {h:"Communication time", n:fmtShort(commHrs)+" h", s:"audio, video & screen share", bar:100}
    ];
    $("tmKpiRow").innerHTML = cards.map(function(c){
      return '<div class="card kpi"><h3>'+esc(c.h)+'</h3><div class="num">'+esc(c.n)+'</div>'+
        '<div class="sub">'+esc(c.s)+'</div><div class="bar"><span style="width:'+Math.min(100,c.bar)+'%"></span></div></div>';
    }).join("");
  }

  function renderOverview(){
    var T=DATA.totals;
    var mix=[
      {label:"Team chat",    v:T.teamChat, color:"var(--c1)"},
      {label:"Private chat", v:T.privChat, color:"var(--c2)"},
      {label:"Channel posts",v:T.post,     color:"var(--c3)"},
      {label:"Channel replies",v:T.reply,  color:"var(--c4)"},
      {label:"Calls",        v:T.calls,    color:"var(--c5)"},
      {label:"Meetings",     v:T.meetings, color:"var(--c6)"}
    ].sort(function(a,b){return b.v-a.v;});
    $("tmActivityMix").innerHTML = hBarChart(mix,{labelW:130});

    var top=DATA.users.slice().sort(function(a,b){return b.meetings-a.meetings;})
            .filter(function(u){return u.meetings>0;}).slice(0,10)
            .map(function(u){return {label:u.name,v:u.meetings,color:"var(--c2)"};});
    $("tmTopMeet").innerHTML = top.length? hBarChart(top,{labelW:150}) :
      '<p class="cnote">No meeting activity recorded.</p>';
  }

  function renderMeetings(){
    var T=DATA.totals;
    var oa=[
      {label:"Organized",v:T.org,color:"var(--c3)"},
      {label:"Attended", v:T.att,color:"var(--c2)"}
    ];
    $("tmOrgAtt").innerHTML = hBarChart(oa,{labelW:120});

    var byType=[
      {label:"Ad-hoc",            v:T.adHocAtt,color:"var(--c1)"},
      {label:"Scheduled one-time",v:T.schedAtt,color:"var(--c4)"},
      {label:"Scheduled recurring",v:T.recurAtt,color:"var(--c6)"}
    ];
    $("tmMeetType").innerHTML = hBarChart(byType,{labelW:160});

    var durs=[
      {label:"Audio",        v:hrs(T.audioS),color:"var(--c1)"},
      {label:"Video",        v:hrs(T.videoS),color:"var(--c5)"},
      {label:"Screen share", v:hrs(T.shareS),color:"var(--c3)"}
    ];
    $("tmDurations").innerHTML = hBarChart(durs,{labelW:130,
      fmt:function(v){return fmtShort(v)+" h";}, tipv:function(d){return Math.round(d.v).toLocaleString("en-US")+" hours";}});

    var topOrg=DATA.users.slice().sort(function(a,b){return b.org-a.org;})
            .filter(function(u){return u.org>0;}).slice(0,10)
            .map(function(u){return {label:u.name,v:u.org,color:"var(--c3)"};});
    $("tmTopOrg").innerHTML = topOrg.length? hBarChart(topOrg,{labelW:150}) :
      '<p class="cnote">No meetings were organized in this window.</p>';
  }

  function renderMessaging(){
    var T=DATA.totals;
    var chatSegs=[
      {label:"Team chat",   v:T.teamChat,color:"var(--c1)"},
      {label:"Private chat",v:T.privChat,color:"var(--c2)"}
    ];
    $("tmChatDonut").innerHTML = donut(chatSegs,{center:T.teamChat+T.privChat,centerLabel:"chat messages"});
    $("tmChatLegend").innerHTML = legend(chatSegs);

    var chan=[
      {label:"Posts",  v:T.post,  color:"var(--c3)"},
      {label:"Replies",v:T.reply, color:"var(--c4)"},
      {label:"Urgent", v:T.urgent,color:"var(--c5)"}
    ];
    $("tmChannelBars").innerHTML = hBarChart(chan,{labelW:120});

    var topMsg=DATA.users.slice().sort(function(a,b){return b.chat-a.chat;})
            .filter(function(u){return u.chat>0;}).slice(0,10)
            .map(function(u){return {label:u.name,v:u.chat,color:"var(--c1)"};});
    $("tmTopMsg").innerHTML = topMsg.length? hBarChart(topMsg,{labelW:150}) :
      '<p class="cnote">No chat messages recorded.</p>';

    var topCall=DATA.users.slice().sort(function(a,b){return b.calls-a.calls;})
            .filter(function(u){return u.calls>0;}).slice(0,10)
            .map(function(u){return {label:u.name,v:u.calls,color:"var(--c5)"};});
    $("tmTopCall").innerHTML = topCall.length? hBarChart(topCall,{labelW:150}) :
      '<p class="cnote">No calls recorded.</p>';
  }

  function buildUsersTable(){
    var cols=[
      {k:"name", t:"Person", sort:"string", render:function(u){return esc(u.name);}},
      {k:"upn", t:"User principal name", sort:"string", render:function(u){return '<span class="muted">'+esc(u.upn)+'</span>';}},
      {k:"last", t:"Last activity", sort:"string", render:function(u){return esc(u.last||"—");}},
      {k:"teamChat", t:"Team chat", num:true, sort:"number", render:function(u){return fmt(u.teamChat);}},
      {k:"privChat", t:"Private chat", num:true, sort:"number", render:function(u){return fmt(u.privChat);}},
      {k:"calls", t:"Calls", num:true, sort:"number", render:function(u){return fmt(u.calls);}},
      {k:"meetings", t:"Meetings", num:true, sort:"number", render:function(u){return fmt(u.meetings);}},
      {k:"org", t:"Organized", num:true, sort:"number", render:function(u){return fmt(u.org);}},
      {k:"post", t:"Posts", num:true, sort:"number", render:function(u){return fmt(u.post);}},
      {k:"reply", t:"Replies", num:true, sort:"number", render:function(u){return fmt(u.reply);}},
      {k:"commS", t:"Comm time", num:true, sort:"number", render:function(u){return u.commS? fmtHrs(u.commS):"—";}},
      {k:"total", t:"Total actions", num:true, sort:"number", render:function(u){return "<strong>"+fmt(u.total)+"</strong>";}}
    ];
    usrTbl=new SortTable($("tmUsrTable"), cols, {
      initSort:{key:"total",dir:-1}, max:300,
      getRows:function(){
        var q=($("tmUsrSearch").value||"").toLowerCase().trim();
        var activeOnly=$("tmActiveOnly").checked;
        return DATA.users.filter(function(u){
          if(activeOnly && !(u.total>0 || u.commS>0)) return false;
          if(q) return u.name.toLowerCase().indexOf(q)>=0 || u.upn.toLowerCase().indexOf(q)>=0;
          return true;
        });
      },
      onCount:function(total,shown){
        $("tmUsrCount").textContent = fmt(total)+" people"+(total>shown?" (showing top "+shown+")":"");
      }
    });
  }
  function renderUsers(){ if(!usrTbl) buildUsersTable(); usrTbl.render(); }

  function renderAll(){
    compute();
    renderHeader(); renderKpis(); renderOverview();
    renderMeetings(); renderMessaging(); renderUsers();
    $("tmStatus").classList.add("hidden");
    $("tmDash").classList.remove("hidden");
    hasData = true;
  }

  /* ---------- add/replace files via the header Refresh button ---------- */
  // Reveal the upload panel without discarding the rendered dashboard or the
  // already-loaded files. When data is present we also expose a "Back" button so
  // a mis-click can return to the current reports without re-uploading.
  function openUploader(){
    $("tmStatus").classList.remove("hidden");
    $("tmLoadingBox").classList.add("hidden");
    $("tmFallbackBox").classList.remove("hidden");
    $("tmDash").classList.add("hidden");
    updateChecklist();
    toggleBack();
  }
  function backToReports(){
    if(!hasData) return;
    $("tmStatus").classList.add("hidden");
    $("tmFallbackBox").classList.add("hidden");
    $("tmDash").classList.remove("hidden");
  }
  function toggleBack(){
    var b=$("tmBackBtn"); if(!b) return;
    if(hasData) b.classList.remove("hidden"); else b.classList.add("hidden");
  }

  /* ---------- tabs ---------- */
  function initTabs(){
    var nav=$("tmTabs"); if(!nav) return;
    var btns=nav.querySelectorAll("button");
    btns.forEach(function(b){
      b.addEventListener("click",function(){
        btns.forEach(function(x){x.setAttribute("aria-selected","false");});
        b.setAttribute("aria-selected","true");
        $("tm-app").querySelectorAll(".tm-tabpanel").forEach(function(p){p.classList.add("hidden");});
        $("tmtab-"+b.dataset.tab).classList.remove("hidden");
      });
    });
  }
  function clearData(){
    RAW = {};
    DATA = {};
    fileMap = {};
    hasData = false;
    usrTbl = null;
    updateChecklist();
    toggleBack();
    $("tmStatus").classList.remove("hidden");
    $("tmLoadingBox").classList.add("hidden");
    $("tmFallbackBox").classList.remove("hidden");
    $("tmDash").classList.add("hidden");
  }

  function initControls(){
    $("tmUsrSearch").addEventListener("input",renderUsers);
    $("tmActiveOnly").addEventListener("change",renderUsers);
    
    // Wire up dropdown controls
    var tmBtn = $("tmReloadBtn");
    var tmDropdown = $("tmReloadDropdown");
    if(tmBtn && tmDropdown){
      tmBtn.addEventListener("click", function(e){
        e.stopPropagation();
        Array.prototype.forEach.call(document.querySelectorAll(".refresh-dropdown"), function(d){
          if(d !== tmDropdown) d.classList.remove("show");
        });
        tmDropdown.classList.toggle("show");
      });
    }
    var tmRefFolder = $("tmRefreshFolder");
    if(tmRefFolder){
      tmRefFolder.addEventListener("click", function(){
        tmDropdown.classList.remove("show");
        boot(true);
      });
    }
    var tmUpReplace = $("tmUploadReplace");
    if(tmUpReplace){
      tmUpReplace.addEventListener("click", function(){
        tmDropdown.classList.remove("show");
        openUploader();
      });
    }
    var tmClearData = $("tmClearData");
    if(tmClearData){
      tmClearData.addEventListener("click", function(){
        tmDropdown.classList.remove("show");
        if(confirm("Are you sure you want to clear all loaded data? This will reset the dashboard.")){
          clearData();
        }
      });
    }

    var bb=$("tmBackBtn"); if(bb) bb.addEventListener("click",backToReports);
    // hover tooltips for bar/donut, scoped to the Teams container
    var root=$("tm-app")||document;
    root.addEventListener("mousemove", function(e){
      var t = e.target.closest && e.target.closest("[data-tip]");
      if(t){
        var title=esc(t.getAttribute("data-tip")||"");
        var v=esc(t.getAttribute("data-tipv")||"");
        var c=t.getAttribute("data-tipc")||"var(--c1)";
        var html='<div class="tt-row"><span class="sw" style="background:'+c+'"></span>'+
          '<span>'+title+'</span><span class="tt-val">'+v+'</span></div>';
        Tip.show(html, e.clientX, e.clientY); return;
      }
      Tip.hide();
    });
  }

  /* ---------- loading ---------- */
  function fetchText(url){
    return fetch(url,{cache:"no-store"}).then(function(res){
      if(!res.ok) throw new Error(url+" ("+res.status+")");
      return res.text();
    });
  }
  // Probe period (30/60/90/180) x suffix style ( _D180 / _'D180' / _180 ).
  function fetchAll(){
    detectedPeriod=0;
    var combos=[];
    PERIODS.forEach(function(p){ suffixes(p).forEach(function(sx){ combos.push({p:p,sx:sx}); }); });
    var i=0;
    function tryNext(){
      if(i>=combos.length) return Promise.reject(new Error("No Teams report file was found in this folder."));
      var combo=combos[i++], sx=combo.sx;
      return fetchText(FOLDER+BASE+sx+".csv").then(function(txt){
        detectedPeriod=combo.p; RAW.detail=parseCSV(txt);
      }).catch(function(e){
        return tryNext();
      });
    }
    return tryNext();
  }
  function boot(isReload){
    $("tmStatus").classList.remove("hidden");
    $("tmLoadingBox").classList.remove("hidden");
    $("tmFallbackBox").classList.add("hidden");
    if(isReload) $("tmDash").classList.add("hidden");
    fetchAll().then(function(){ renderAll(); }).catch(function(err){ showFallback(err); });
  }

  /* ---------- manual fallback (file:// or missing file) ---------- */
  var fileMap = {};
  function nameToKey(fn){
    fn=fn.toLowerCase();
    if(fn.indexOf("useractivityuserdetail")>=0 || fn.indexOf("teams")>=0) return "detail";
    return null;
  }
  function showFallback(){
    $("tmLoadingBox").classList.add("hidden");
    $("tmFallbackBox").classList.remove("hidden");
    updateChecklist();
    toggleBack();
  }
  function updateChecklist(){
    var ok=!!fileMap.detail;
    $("tmFileChecklist").innerHTML = '<li class="'+(ok?"ok":"pending")+'">'+(ok?"✔":"○")+" "+esc(LABEL)+"</li>";
  }
  function ingestFiles(fileList){
    var arr=Array.prototype.slice.call(fileList);
    var jobs=arr.map(function(f){
      var key=nameToKey(f.name);
      if(!key) return Promise.resolve();
      return f.text().then(function(txt){ fileMap[key]=parseCSV(txt); });
    });
    Promise.all(jobs).then(function(){
      updateChecklist();
      if(fileMap.detail){ RAW=fileMap; renderAll(); }
      else { $("tmFallbackErr").textContent = "That doesn't look like the Teams user-activity export. Please add the Teams_UserActivityUserDetail CSV."; }
    });
  }
  function initFallback(){
    var dz=$("tmDropZone"), fi=$("tmFileInput");
    dz.addEventListener("click",function(){ fi.click(); });
    fi.addEventListener("change",function(){ ingestFiles(fi.files); });
    ["dragover","dragenter"].forEach(function(ev){
      dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.add("drag"); });
    });
    ["dragleave","drop"].forEach(function(ev){
      dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.remove("drag"); });
    });
    dz.addEventListener("drop",function(e){
      if(e.dataTransfer&&e.dataTransfer.files) ingestFiles(e.dataTransfer.files);
    });
  }

  /* ---------- expose boot for the product switcher ---------- */
  function init(){ initTabs(); initControls(); initFallback(); }
  window.TeamsBoard = {
    booted:false,
    init:init,
    boot:function(){ if(this.booted) return; this.booted=true; init(); boot(false); }
  };
})();
