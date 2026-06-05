"use strict";
(function(){
  // All five exports share one base name plus a _D## suffix that reflects the
  // report period (30/60/90/180 days). The board auto-detects whichever period
  // is present in the folder, so no edits are needed when the window changes.
  // The CSVs live in a "Usage Reports" subfolder next to this page. The space is
  // pre-encoded so the relative URL works over both http(s) and file://.
  var FOLDER = "Usage%20Reports/";
  var BASES = {
    byDate: "VivaEngage_ActivityByDate",
    counts: "VivaEngage_UserCounts",
    byUser: "VivaEngage_ActivityByUser",
    device: "VivaEngage_DeviceUsageByUser",
    groups: "VivaEngage_GroupActivityDetail"
  };
  var PERIODS = [30,60,90,180];
  var LABELS = {
    byDate:"Activity by date", counts:"User counts", byUser:"Activity by user",
    device:"Device usage by user", groups:"Group activity detail"
  };
  var detectedPeriod = 0;
  var COLORS = ["var(--c1)","var(--c2)","var(--c3)","var(--c4)","var(--c5)","var(--c6)"];
  var RAW = {}, DATA = {}, loadedFiles = {};
  var grpTbl = null, usrTbl = null;

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

  /* ---------- floating tooltip ---------- */
  var Tip = {
    el:null,
    ensure:function(){
      if(!this.el){ this.el=document.createElement("div"); this.el.id="tip"; document.body.appendChild(this.el); }
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

  /* ---------- robust CSV parser (quoted fields, embedded commas/quotes, BOM) ---------- */
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
    /* ByDate: daily totals */
    var bd = RAW.byDate.rows.map(function(r){
      return { date:r["Report Date"], liked:num(r["Liked"]), posted:num(r["Posted"]), read:num(r["Read"]) };
    }).filter(function(r){return r.date;}).sort(function(a,b){return a.date<b.date?-1:1;});

    /* UserCounts: daily unique users per action */
    var uc = RAW.counts.rows.map(function(r){
      return { date:r["Report Date"], liked:num(r["Liked"]), posted:num(r["Posted"]), read:num(r["Read"]) };
    }).filter(function(r){return r.date;}).sort(function(a,b){return a.date<b.date?-1:1;});

    var totRead=0, totPost=0, totLike=0;
    bd.forEach(function(d){ totRead+=d.read; totPost+=d.posted; totLike+=d.liked; });

    /* ByUser: engagement per person */
    var users = RAW.byUser.rows.map(function(r){
      var posted=num(r["Posted Count"]), read=num(r["Read Count"]), liked=num(r["Liked Count"]);
      return {
        name:r["Display Name"]||r["User Principal Name"]||"",
        upn:r["User Principal Name"]||"",
        state:(r["User State"]||"").toLowerCase(),
        last:r["Last Activity Date"]||"",
        posted:posted, read:read, liked:liked, total:posted+read+liked
      };
    });
    var totalUsers=users.length;
    var readers=0, likers=0, posters=0, engaged=0;
    users.forEach(function(u){
      if(u.total>0) engaged++;
      if(u.read>0) readers++;
      if(u.liked>0) likers++;
      if(u.posted>0) posters++;
    });

    /* Devices: per active user platform flags */
    var devCols=[
      {key:"Used Web",label:"Web"},
      {key:"Used Android Phone",label:"Android"},
      {key:"Used iPhone",label:"iPhone"},
      {key:"Used iPad",label:"iPad"},
      {key:"Used Windows Phone",label:"Windows Phone"},
      {key:"Used Others",label:"Other"}
    ];
    var devCount={}, devActive=0, mobileUsers=0, webUsers=0;
    devCols.forEach(function(c){ devCount[c.label]=0; });
    RAW.device.rows.forEach(function(r){
      var anyMobile=false, anyWeb=(r["Used Web"]==="Yes"), anyUse=false;
      devCols.forEach(function(c){
        if(r[c.key]==="Yes"){ devCount[c.label]++; anyUse=true;
          if(c.label!=="Web") anyMobile=true; }
      });
      if(anyUse) devActive++;
      if(anyMobile) mobileUsers++;
      if(anyWeb) webUsers++;
    });
    var webOnly = Math.max(0, devActive - mobileUsers);

    /* Groups */
    var groups = RAW.groups.rows.map(function(r){
      var posted=num(r["Posted Count"]), read=num(r["Read Count"]), liked=num(r["Liked Count"]);
      return {
        name:r["Group Display Name"]||"",
        owner:r["Owner Principal Name"]||"",
        type:(r["Group Type"]||"").toLowerCase(),
        deleted:(r["Is Deleted"]||"").toLowerCase()==="true",
        members:num(r["Member Count"]),
        last:r["Last Activity Date"]||"",
        posted:posted, read:read, liked:liked, total:posted+read+liked
      };
    });
    var liveGroups = groups.filter(function(g){return !g.deleted;});
    var activeGroups = liveGroups.filter(function(g){return g.total>0;}).length;

    var refresh = RAW.byDate.rows.length? RAW.byDate.rows[0]["Report Refresh Date"] : "";
    var dmin = bd.length? bd[0].date : "", dmax = bd.length? bd[bd.length-1].date : "";
    // Window length, in priority order: the "Report Period" column in the data,
    // then the filename suffix we detected, then the actual span of dates present.
    var reportPeriod = RAW.byDate.rows.length? num(RAW.byDate.rows[0]["Report Period"]) : 0;
    var daySpan = (dmin&&dmax)? Math.round((Date.parse(dmax)-Date.parse(dmin))/86400000)+1 : bd.length;
    var windowDays = reportPeriod || detectedPeriod || daySpan || 30;

    DATA = {
      byDate:bd, counts:uc,
      totals:{read:totRead,post:totPost,like:totLike},
      users:users, totalUsers:totalUsers, engaged:engaged,
      readers:readers, likers:likers, posters:posters,
      devCols:devCols, devCount:devCount, devActive:devActive,
      mobileUsers:mobileUsers, webUsers:webUsers, webOnly:webOnly,
      groups:liveGroups, allGroups:groups, totalGroups:liveGroups.length, activeGroups:activeGroups,
      refresh:refresh, period:{from:dmin,to:dmax}, windowDays:windowDays, daySpan:daySpan
    };
  }

  /* ---------- line chart (geometry + SVG + interactive hover) ---------- */
  function lineGeom(data, opt){
    opt=opt||{};
    var W=opt.w||520, H=opt.h||190, pl=42, pr=12, pt=12, pb=26;
    var iw=W-pl-pr, ih=H-pt-pb;
    var vals=data.map(function(d){return d.v;});
    var max=Math.max.apply(null,vals.concat([1])), min=0;
    var n=data.length;
    function X(i){ return pl + (n<=1?iw/2 : i/(n-1)*iw); }
    function Y(v){ return pt + ih - (v-min)/(max-min||1)*ih; }
    return {W:W,H:H,pl:pl,pr:pr,pt:pt,pb:pb,iw:iw,ih:ih,max:max,min:min,n:n,
            X:X,Y:Y,color:opt.color||"var(--c1)",area:!!opt.area};
  }
  function lineSVG(data, g){
    var W=g.W,H=g.H,pl=g.pl,pr=g.pr,n=g.n,max=g.max,color=g.color;
    var area="", line="";
    data.forEach(function(d,i){
      var x=g.X(i),y=g.Y(d.v);
      line += (i?"L":"M")+x.toFixed(1)+" "+y.toFixed(1)+" ";
      area += (i?"L":"M")+x.toFixed(1)+" "+y.toFixed(1)+" ";
    });
    if(g.area && n){ area += "L"+g.X(n-1).toFixed(1)+" "+(g.pt+g.ih)+" L"+g.X(0).toFixed(1)+" "+(g.pt+g.ih)+" Z"; }
    var dots="";
    data.forEach(function(d,i){
      var x=g.X(i),y=g.Y(d.v);
      dots += '<circle cx="'+x.toFixed(1)+'" cy="'+y.toFixed(1)+'" r="2.4" fill="'+color+'"></circle>';
    });
    var grid="", ylab="";
    for(var q=0; q<=2; q++){
      var v=max*q/2, y=g.Y(v);
      grid += '<line x1="'+pl+'" y1="'+y.toFixed(1)+'" x2="'+(W-pr)+'" y2="'+y.toFixed(1)+'" stroke="var(--line)" stroke-width="1"/>';
      ylab += '<text x="'+(pl-6)+'" y="'+(y+3).toFixed(1)+'" text-anchor="end" font-size="10" fill="var(--muted)">'+fmtShort(v)+'</text>';
    }
    var xl="";
    [0, Math.floor((n-1)/2), n-1].forEach(function(i){
      if(i<0||i>=n) return;
      var lab=(data[i].label||"").slice(5); // MM-DD
      xl += '<text x="'+g.X(i).toFixed(1)+'" y="'+(H-7)+'" text-anchor="middle" font-size="10" fill="var(--muted)">'+esc(lab)+'</text>';
    });
    var uid="g"+Math.random().toString(36).slice(2,8);
    var fill = g.area? '<defs><linearGradient id="'+uid+'" x1="0" x2="0" y1="0" y2="1">'+
      '<stop offset="0" stop-color="'+color+'" stop-opacity="0.28"/>'+
      '<stop offset="1" stop-color="'+color+'" stop-opacity="0.02"/></linearGradient></defs>'+
      '<path d="'+area+'" fill="url(#'+uid+')" stroke="none"/>' : "";
    return '<svg viewBox="0 0 '+W+' '+H+'" role="img" aria-label="line chart" preserveAspectRatio="xMidYMid meet">'+
      grid+ fill +
      '<path d="'+line+'" fill="none" stroke="'+color+'" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>'+
      dots + ylab + xl + '</svg>';
  }
  function renderLine(el, data, opt){
    opt=opt||{};
    var g=lineGeom(data,opt);
    el.innerHTML='<div class="chartbox">'+lineSVG(data,g)+
      '<div class="crosshair"></div><div class="hoverdot"></div></div>';
    var box=el.querySelector(".chartbox");
    box._chart={ data:data, g:g, color:opt.color||"var(--c1)", title:opt.seriesLabel||"Value" };
    attachLineHover(box);
  }
  function attachLineHover(box){
    var ch=box.querySelector(".crosshair"), dot=box.querySelector(".hoverdot");
    function move(e){
      var c=box._chart, g=c.g, rect=box.getBoundingClientRect();
      if(!rect.width || !g.n) return;
      var s=rect.width/g.W;
      var vx=(e.clientX-rect.left)/s;
      var idx=Math.round((vx-g.pl)/(g.iw||1)*(g.n-1));
      if(idx<0) idx=0; if(idx>g.n-1) idx=g.n-1;
      var d=c.data[idx]; if(!d) return;
      var px=g.X(idx)*s, py=g.Y(d.v)*s;
      ch.style.display="block"; ch.style.left=px+"px";
      dot.style.display="block"; dot.style.left=px+"px"; dot.style.top=py+"px";
      var html='<div class="tt-title">'+esc(d.label||"")+'</div>'+
        '<div class="tt-row"><span class="sw" style="background:'+c.color+'"></span>'+
        '<span>'+esc(c.title)+'</span><span class="tt-val">'+fmt(d.v)+'</span></div>';
      Tip.show(html, e.clientX, e.clientY);
    }
    function leave(){ ch.style.display="none"; dot.style.display="none"; Tip.hide(); }
    box.addEventListener("mousemove", move);
    box.addEventListener("mouseleave", leave);
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
      var col=opt.color||COLORS[i%COLORS.length];
      out+='<rect x="'+pl+'" y="'+y+'" width="'+w.toFixed(1)+'" height="'+rowH+'" rx="5" fill="'+col+'"'+
           ' data-tip="'+esc(d.label)+'" data-tipv="'+fmt(d.v)+'" data-tipc="'+col+'"></rect>';
      out+='<text x="'+(pl-8)+'" y="'+(y+rowH/2+4)+'" text-anchor="end" font-size="12" fill="var(--ink)">'+esc(opt.clip!==false? clip(d.label,26):d.label)+'</text>';
      out+='<text x="'+(pl+w+7).toFixed(1)+'" y="'+(y+rowH/2+4)+'" font-size="11.5" fill="var(--muted)">'+fmtShort(d.v)+'</text>';
    });
    return out+'</svg>';
  }

  function donut(segs, opt){
    opt=opt||{}; var size=opt.size||190, r=70, cx=size/2, cy=size/2, sw=26;
    var total=segs.reduce(function(a,s){return a+s.v;},0)||1;
    var ang=-Math.PI/2, circ=2*Math.PI*r, out='<svg viewBox="0 0 '+size+' '+size+'" role="img" aria-label="donut chart">';
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

  /* ---------- resizable, click-sortable table ---------- */
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

  /* ---------- renderers ---------- */
  function renderHeader(){
    var p=DATA.period, wd=DATA.windowDays;
    $("subtitle").innerHTML = "Community engagement across The Mosaic Company &middot; rolling "+wd+"-day window";
    var pills=[];
    pills.push('<span class="pill">Window: '+wd+' days</span>');
    if(p.from&&p.to) pills.push('<span class="pill">'+esc(p.from)+' to '+esc(p.to)+'</span>');
    if(DATA.refresh) pills.push('<span class="pill">Data refreshed '+esc(DATA.refresh)+'</span>');
    pills.push('<span class="pill">'+fmt(DATA.totalUsers)+' licensed users</span>');
    $("metaPills").innerHTML=pills.join("");
    $("footNote").innerHTML = "Generated from Viva Engage usage exports &middot; data refreshes automatically when the source files in this folder are updated &middot; "+wd+"-day window, "+esc(p.from)+" &ndash; "+esc(p.to);
  }

  function renderKpis(){
    var d=DATA, adoption=pct(d.engaged,d.totalUsers);
    var cards=[
      {h:"Active participants", n:fmt(d.engaged), s:adoption+"% of "+fmt(d.totalUsers)+" licensed users", bar:adoption},
      {h:"Posts created", n:fmt(d.totals.post), s:fmt(d.posters)+" unique contributors", bar:pct(d.posters,d.engaged)},
      {h:"Reads", n:fmt(d.totals.read), s:fmt(d.readers)+" unique readers", bar:pct(d.readers,d.engaged)},
      {h:"Reactions", n:fmt(d.totals.like), s:fmt(d.likers)+" unique likers", bar:pct(d.likers,d.engaged)},
      {h:"Active communities", n:fmt(d.activeGroups), s:"of "+fmt(d.totalGroups)+" total communities", bar:pct(d.activeGroups,d.totalGroups)},
      {h:"Mobile adoption", n:pct(d.mobileUsers,d.devActive)+"%", s:fmt(d.mobileUsers)+" of "+fmt(d.devActive)+" active users", bar:pct(d.mobileUsers,d.devActive)}
    ];
    $("kpiRow").innerHTML = cards.map(function(c){
      return '<div class="card kpi"><h3>'+esc(c.h)+'</h3><div class="num">'+c.n+'</div>'+
        '<div class="sub">'+esc(c.s)+'</div><div class="bar"><span style="width:'+Math.min(100,c.bar)+'%"></span></div></div>';
    }).join("");
  }

  function renderOverview(){
    renderLine($("ovReads"), DATA.byDate.map(function(d){return {label:d.date,v:d.read};}),
               {area:true,color:"var(--c1)",h:200,seriesLabel:"Reads"});
    var fl=$("funnelLede"); if(fl) fl.textContent="How many unique users read, react, and create over the "+DATA.windowDays+"-day window.";
    // funnel
    var d=DATA, steps=[
      {l:"Read",v:d.readers,c:"var(--c1)"},
      {l:"Reacted",v:d.likers,c:"var(--c2)"},
      {l:"Posted",v:d.posters,c:"var(--c3)"}
    ];
    var top=Math.max.apply(null,steps.map(function(s){return s.v;}).concat([1]));
    $("funnel").innerHTML = steps.map(function(s){
      return '<div class="step"><div class="lbl">'+s.l+'</div>'+
        '<div class="track"><div class="fill" style="width:'+Math.max(6,s.v/top*100)+'%;background:'+s.c+'">'+fmt(s.v)+'</div></div>'+
        '<div class="pct">'+pct(s.v,d.engaged)+'%</div></div>';
    }).join("") + '<p class="cnote" style="margin-top:10px">Percentages are of the '+fmt(d.engaged)+' active participants.</p>';
    // top groups
    var tg=DATA.groups.slice().sort(function(a,b){return b.read-a.read;}).slice(0,10)
            .map(function(g){return {label:g.name,v:g.read};});
    $("ovTopGroups").innerHTML = hBarChart(tg,{color:"var(--c2)",labelW:190});
    // devices
    var segs=DATA.devCols.map(function(c){return {label:c.label,v:DATA.devCount[c.label]};})
              .filter(function(s){return s.v>0;}).sort(function(a,b){return b.v-a.v;});
    $("ovDevices").innerHTML = hBarChart(segs,{labelW:110});
    $("ovDevLegend").innerHTML = "";
  }

  function renderTrends(){
    var bd=DATA.byDate, uc=DATA.counts;
    function L(sel,arr,key,color,label){
      renderLine($(sel), arr.map(function(d){return {label:d.date,v:d[key]};}),
                 {color:color,area:true,h:150,seriesLabel:label});
    }
    L("trReadTot",bd,"read","var(--c1)","Reads");
    L("trPostTot",bd,"posted","var(--c3)","Posts");
    L("trLikeTot",bd,"liked","var(--c2)","Likes");
    L("trReadU",uc,"read","var(--c1)","Active readers");
    L("trPostU",uc,"posted","var(--c3)","Active posters");
    L("trLikeU",uc,"liked","var(--c2)","Active likers");
  }

  /* groups table */
  function buildGroupsTable(){
    var cols=[
      {k:"name", t:"Community", sort:"string", render:function(g){return esc(g.name);}},
      {k:"type", t:"Type", sort:"string", render:function(g){return '<span class="tag '+(g.type==="private"?"private":"public")+'">'+esc(g.type||"-")+"</span>";}},
      {k:"members", t:"Members", num:true, sort:"number", render:function(g){return fmt(g.members);}},
      {k:"posted", t:"Posts", num:true, sort:"number", render:function(g){return fmt(g.posted);}},
      {k:"read", t:"Reads", num:true, sort:"number", render:function(g){return fmt(g.read);}},
      {k:"liked", t:"Likes", num:true, sort:"number", render:function(g){return fmt(g.liked);}},
      {k:"last", t:"Last activity", sort:"string", render:function(g){return esc(g.last||"—");}}
    ];
    grpTbl=new SortTable($("grpTable"), cols, {
      initSort:{key:"read",dir:-1}, max:300,
      getRows:function(){
        var q=($("grpSearch").value||"").toLowerCase().trim();
        var type=$("grpType").value;
        var activeOnly=$("grpActiveOnly").checked;
        return DATA.groups.filter(function(g){
          if(type && g.type!==type) return false;
          if(activeOnly && g.total<=0) return false;
          if(q && g.name.toLowerCase().indexOf(q)<0 && g.owner.toLowerCase().indexOf(q)<0) return false;
          return true;
        });
      },
      onCount:function(total){
        $("grpCount").textContent = total+" communities"+(total>300?" (showing top 300)":"");
      }
    });
  }
  function renderGroups(){ if(!grpTbl) buildGroupsTable(); grpTbl.render(); }

  /* users table */
  function buildUsersTable(){
    var cols=[
      {k:"name", t:"Person", sort:"string", render:function(u){return esc(u.name);}},
      {k:"state", t:"State", sort:"string", render:function(u){return esc(u.state||"—");}},
      {k:"last", t:"Last activity", sort:"string", render:function(u){return esc(u.last||"—");}},
      {k:"posted", t:"Posts", num:true, sort:"number", render:function(u){return fmt(u.posted);}},
      {k:"read", t:"Reads", num:true, sort:"number", render:function(u){return fmt(u.read);}},
      {k:"liked", t:"Likes", num:true, sort:"number", render:function(u){return fmt(u.liked);}},
      {k:"total", t:"Total", num:true, sort:"number", render:function(u){return "<strong>"+fmt(u.total)+"</strong>";}}
    ];
    usrTbl=new SortTable($("usrTable"), cols, {
      initSort:{key:"total",dir:-1}, max:200,
      getRows:function(){
        var q=($("usrSearch").value||"").toLowerCase().trim();
        return DATA.users.filter(function(u){
          if(u.total<=0 && !q) return false;
          if(q) return u.name.toLowerCase().indexOf(q)>=0 || u.upn.toLowerCase().indexOf(q)>=0;
          return true;
        });
      },
      onCount:function(total,shown){
        $("usrCount").textContent = fmt(total)+" people"+(total>shown?" (showing top "+shown+")":"");
      }
    });
  }
  function renderUsers(){ if(!usrTbl) buildUsersTable(); usrTbl.render(); }

  function renderDevices(){
    var segs=DATA.devCols.map(function(c){return {label:c.label,v:DATA.devCount[c.label]};})
              .filter(function(s){return s.v>0;}).sort(function(a,b){return b.v-a.v;});
    $("dvBars").innerHTML=hBarChart(segs,{labelW:120});
    var donutSegs=[
      {label:"Used mobile",v:DATA.mobileUsers,color:"var(--c2)"},
      {label:"Web only",v:DATA.webOnly,color:"var(--c1)"}
    ];
    $("dvDonut").innerHTML=donut(donutSegs,{center:DATA.devActive,centerLabel:"active users"});
    $("dvDonutLegend").innerHTML=legend(donutSegs);
  }

  var hasData = false;

  function renderAll(){
    compute();
    renderHeader(); renderKpis(); renderOverview(); renderTrends();
    renderGroups(); renderUsers(); renderDevices();
    $("status").classList.add("hidden");
    $("dash").classList.remove("hidden");
    hasData = true;
  }

  /* ---------- tabs ---------- */
  function initTabs(){
    var btns=$("tabs").querySelectorAll("button");
    btns.forEach(function(b){
      b.addEventListener("click",function(){
        btns.forEach(function(x){x.setAttribute("aria-selected","false");});
        b.setAttribute("aria-selected","true");
        document.querySelectorAll(".tabpanel").forEach(function(p){p.classList.add("hidden");});
        $("tab-"+b.dataset.tab).classList.remove("hidden");
      });
    });
  }
  
  function openUploader(){
    $("status").classList.remove("hidden");
    $("loadingBox").classList.add("hidden");
    $("fallbackBox").classList.remove("hidden");
    $("dash").classList.add("hidden");
    updateChecklist();
    toggleBack();
  }
  function backToReports(){
    if(!hasData) return;
    $("status").classList.add("hidden");
    $("fallbackBox").classList.add("hidden");
    $("dash").classList.remove("hidden");
  }
  function toggleBack(){
    var b=$("vivaBackBtn"); if(!b) return;
    if(hasData) b.classList.remove("hidden"); else b.classList.add("hidden");
  }
  function clearData(){
    RAW = {};
    DATA = {};
    fileMap = {};
    hasData = false;
    grpTbl = null;
    usrTbl = null;
    updateChecklist();
    toggleBack();
    $("status").classList.remove("hidden");
    $("loadingBox").classList.add("hidden");
    $("fallbackBox").classList.remove("hidden");
    $("dash").classList.add("hidden");
  }

  function initControls(){
    ["grpSearch","grpType","grpActiveOnly"].forEach(function(id){
      $(id).addEventListener("input",renderGroups); $(id).addEventListener("change",renderGroups);
    });
    $("usrSearch").addEventListener("input",renderUsers);
    
    // Wire up dropdown controls
    var rBtn = $("reloadBtn");
    var rDropdown = $("reloadDropdown");
    if(rBtn && rDropdown){
      rBtn.addEventListener("click", function(e){
        e.stopPropagation();
        Array.prototype.forEach.call(document.querySelectorAll(".refresh-dropdown"), function(d){
          if(d !== rDropdown) d.classList.remove("show");
        });
        rDropdown.classList.toggle("show");
      });
    }
    var vRefFolder = $("vivaRefreshFolder");
    if(vRefFolder){
      vRefFolder.addEventListener("click", function(){
        rDropdown.classList.remove("show");
        boot(true);
      });
    }
    var vUpReplace = $("vivaUploadReplace");
    if(vUpReplace){
      vUpReplace.addEventListener("click", function(){
        rDropdown.classList.remove("show");
        openUploader();
      });
    }
    var vClearData = $("vivaClearData");
    if(vClearData){
      vClearData.addEventListener("click", function(){
        rDropdown.classList.remove("show");
        if(confirm("Are you sure you want to clear all loaded data? This will reset the dashboard.")){
          clearData();
        }
      });
    }
    var vBack = $("vivaBackBtn");
    if(vBack){
      vBack.addEventListener("click", backToReports);
    }

    // chart hover for bar/donut via data-tip delegation.
    // Scoped to the Viva container so it never fires over the Copilot section
    // (both dashboards share one page in the combined app).
    var hoverRoot = document.getElementById("viva-app") || document;
    hoverRoot.addEventListener("mousemove", function(e){
      var t = e.target.closest && e.target.closest("[data-tip]");
      if(t){
        var title=esc(t.getAttribute("data-tip")||"");
        var v=esc(t.getAttribute("data-tipv")||"");
        var c=t.getAttribute("data-tipc")||"var(--c1)";
        var html='<div class="tt-row"><span class="sw" style="background:'+c+'"></span>'+
          '<span>'+title+'</span><span class="tt-val">'+v+'</span></div>';
        Tip.show(html, e.clientX, e.clientY); return;
      }
      // line charts manage their own tooltip on mouseleave
      if(e.target.closest && e.target.closest(".chartbox")) return;
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
  // Probe the report-period suffix (_D30 -> _D60 -> _D90 -> _D180) until the
  // activity-by-date file is found, then load the rest of the set at that period.
  function fetchAll(){
    detectedPeriod=0;
    var i=0;
    function tryNext(){
      if(i>=PERIODS.length) return Promise.reject(new Error("No Viva Engage report files (_D30/_D60/_D90/_D180) were found in this folder."));
      var p=PERIODS[i++];
      return fetchText(BASES.byDate+"_D"+p+".csv").then(function(txt){
        detectedPeriod=p; RAW.byDate=parseCSV(txt);
        return Promise.all(["counts","byUser","device","groups"].map(function(k){
          return fetchText(BASES[k]+"_D"+p+".csv").then(function(t){ RAW[k]=parseCSV(t); });
        }));
      }).catch(function(e){
        if(detectedPeriod) throw e;   // period found but a sibling file is missing -> surface it
        return tryNext();             // this period's file absent -> try the next window
      });
    }
    return tryNext();
  }
  function boot(isReload){
    $("status").classList.remove("hidden");
    $("loadingBox").classList.remove("hidden");
    $("fallbackBox").classList.add("hidden");
    if(isReload) $("dash").classList.add("hidden");
    fetchAll().then(function(){
      renderAll();
    }).catch(function(err){
      // fetch blocked (file://) or files missing -> manual fallback
      showFallback(err);
    });
  }

  /* ---------- manual fallback ---------- */
  var fileMap = {};
  function nameToKey(fn){
    fn=fn.toLowerCase();
    if(fn.indexOf("activitybydate")>=0) return "byDate";
    if(fn.indexOf("usercounts")>=0) return "counts";
    if(fn.indexOf("activitybyuser")>=0) return "byUser";
    if(fn.indexOf("deviceusage")>=0) return "device";
    if(fn.indexOf("groupactivity")>=0) return "groups";
    return null;
  }
  function showFallback(err){
    $("loadingBox").classList.add("hidden");
    $("fallbackBox").classList.remove("hidden");
    updateChecklist();
    toggleBack();
  }
  function updateChecklist(){
    $("fileChecklist").innerHTML = Object.keys(LABELS).map(function(k){
      var ok=!!fileMap[k];
      return '<li class="'+(ok?"ok":"pending")+'">'+(ok?"✔":"○")+" "+esc(LABELS[k])+"</li>";
    }).join("");
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
      var need=["byDate","counts","byUser","device","groups"];
      var have=need.filter(function(k){return fileMap[k];});
      if(have.length===need.length){
        RAW=fileMap; renderAll();
      } else {
        $("fallbackErr").textContent = "Loaded "+have.length+" of 5 files. Please add the remaining "+(5-have.length)+".";
      }
    });
  }
  function initFallback(){
    var dz=$("dropZone"), fi=$("fileInput");
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

  /* ---------- init / expose boot for the product switcher ---------- */
  function init(){ initTabs(); initControls(); initFallback(); }
  window.VivaBoard = {
    booted:false,
    init:init,
    boot:function(){ if(this.booted) return; this.booted=true; init(); boot(false); }
  };
})();
