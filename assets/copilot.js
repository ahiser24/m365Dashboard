"use strict";
// Copilot dashboard loader and UI renderer
(function(){
  // Probe several filename suffix styles to auto-detect the period
  var FOLDER = "Usage%20Reports/";
  var BASES = {
    byProduct:  "CopilotAdoptionByProduct",
    trend:      "CopilotAdoptionTrend",
    userDetail: "CopilotUsageUserDetail"
  };
  var PERIODS = [30,60,90,180];
  function suffixes(p){ return ["_D"+p, "_'D"+p+"'", "_"+p]; }
  var LABELS = { byProduct:"Adoption by product", trend:"Adoption trend", userDetail:"Usage by user" };

  // Map columns for Copilot apps tracked in standard reports
  var APPS = [
    {key:"Copilot Chat",                 short:"Copilot Chat", color:"var(--c1)",  det:"copilotChatLastActivityDate",      inProd:false},
    {key:"Teams",                        short:"Teams",        color:"var(--c2)",  det:"microsoftTeamsCopilotLastActivityDate", inProd:true},
    {key:"Outlook",                      short:"Outlook",      color:"var(--c6)",  det:"outlookCopilotLastActivityDate",   inProd:true},
    {key:"Word",                         short:"Word",         color:"var(--c4)",  det:"wordCopilotLastActivityDate",      inProd:true},
    {key:"Excel",                        short:"Excel",        color:"var(--c3)",  det:"excelCopilotLastActivityDate",     inProd:true},
    {key:"PowerPoint",                   short:"PowerPoint",   color:"var(--c5)",  det:"powerPointCopilotLastActivityDate",inProd:true},
    {key:"OneNote",                      short:"OneNote",      color:"var(--c7)",  det:"oneNoteCopilotLastActivityDate",   inProd:true},
    {key:"Loop",                         short:"Loop",         color:"var(--c8)",  det:"loopCopilotLastActivityDate",      inProd:true},
    {key:"Edge",                         short:"Edge",         color:"var(--muted)", det:null,                            inProd:true},
    {key:"Microsoft 365 Copilot (app)",  short:"Copilot app",  color:"var(--accent)", det:null,                          inProd:true}
  ];

  // Map columns for richer activity detail reports
  var RICH_APPS = [
    {key:"Copilot Chat (work)", short:"Chat (work)",   color:"var(--c1)",     col:"Last activity date of Copilot Chat (work) (UTC)"},
    {key:"Copilot Chat (web)",  short:"Chat (web)",    color:"var(--c1)",     col:"Last activity date of Copilot Chat (web) (UTC)"},
    {key:"Teams",               short:"Teams",         color:"var(--c2)",     col:"Last activity date of Teams Copilot (UTC)"},
    {key:"Outlook",             short:"Outlook",       color:"var(--c6)",     col:"Last activity date of Outlook Copilot (UTC)"},
    {key:"Word",                short:"Word",          color:"var(--c4)",     col:"Last activity date of Word Copilot (UTC)"},
    {key:"Excel",               short:"Excel",         color:"var(--c3)",     col:"Last activity date of Excel Copilot (UTC)"},
    {key:"PowerPoint",          short:"PowerPoint",    color:"var(--c5)",     col:"Last activity date of PowerPoint Copilot (UTC)"},
    {key:"OneNote",             short:"OneNote",       color:"var(--c7)",     col:"Last activity date of OneNote Copilot (UTC)"},
    {key:"Loop",                short:"Loop",          color:"var(--c8)",     col:"Last activity date of Loop Copilot (UTC)"},
    {key:"Copilot app",         short:"Copilot app",   color:"var(--accent)", col:"Last activity date of Microsoft 365 Copilot (app) (UTC)"},
    {key:"Edge",                short:"Edge",          color:"var(--muted)",  col:"Last activity date of Edge (UTC)"},
    {key:"Copilot Agent",       short:"Copilot Agent", color:"var(--c2)",     col:"Last activity date of Copilot Agent (UTC)"}
  ];

  var detectedPeriod = 0, detectedSuffix = "_D30";
  var COLORS = ["var(--c1)","var(--c2)","var(--c3)","var(--c4)","var(--c5)","var(--c6)","var(--c7)","var(--c8)"];
  var RAW = {}, DATA = {}, usrTbl = null, usrTblRich = null, prodTbl = null;

  // Helpers
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

  // Custom tooltip popup positioner for Copilot charts
  var Tip = {
    el:null,
    ensure:function(){
      if(!this.el){ this.el=document.createElement("div"); this.el.id="cptip"; document.body.appendChild(this.el); }
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

  // CSV Parser
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

  // Process CSV data into metrics
  function colEnabled(app){ return app+" Enabled Users"; }
  function colActive(app){ return app+" Active Users"; }

  function compute(){
    /* Trend: one row per day. Each app contributes an enabled/active pair. */
    var trend = RAW.trend.rows.map(function(r){
      var date = getUserField(r, ["reportDate", "Report Date"]);
      var period = num(getUserField(r, ["reportPeriod", "Report Period"]));
      var allEnabled = num(getUserField(r, ["All Enabled Users", "Any App Enabled Users"]));
      var allActive = num(getUserField(r, ["All Active Users", "Any App Active Users"]));
      var o={ date:date, period:period, allEnabled:allEnabled, allActive:allActive, app:{} };
      APPS.forEach(function(a){
        o.app[a.key]={
          enabled: num(getUserField(r, [a.key + " Enabled Users", "Microsoft " + a.key + " Enabled Users", a.short + " Enabled Users"])),
          active: num(getUserField(r, [a.key + " Active Users", "Microsoft " + a.key + " Active Users", a.short + " Active Users"]))
        };
      });
      return o;
    }).filter(function(r){return r.date;}).sort(function(a,b){return a.date<b.date?-1:1;});

    /* By-product snapshot: the most recent (or only) row of the adoption export. */
    var prodRow = RAW.byProduct.rows.length? RAW.byProduct.rows[RAW.byProduct.rows.length-1] : {};
    var snapshot = APPS.map(function(a){
      var enabled = num(getUserField(prodRow, [a.key + " Enabled Users", "Microsoft " + a.key + " Enabled Users", a.short + " Enabled Users"]));
      var active = num(getUserField(prodRow, [a.key + " Active Users", "Microsoft " + a.key + " Active Users", a.short + " Active Users"]));
      return { key:a.key, short:a.short, color:a.color, enabled:enabled, active:active };
    }).filter(function(s){ return s.enabled>0 || s.active>0; });
    var allEnabledSnap = num(getUserField(prodRow, ["All Enabled Users", "Any App Enabled Users"]));
    var allActiveSnap  = num(getUserField(prodRow, ["All Active Users", "Any App Active Users"]));
    var refresh = getUserField(prodRow, ["Report Refresh Date", "refreshDate"]) || getUserField(RAW.trend.rows[0]||{}, ["reportDate", "Report Date"]) || "";

    /* Per-app average / peak daily active, from the trend file (clearly Copilot usage). */
    var appStats = APPS.map(function(a){
      var sum=0, peak=0, peakDate="";
      trend.forEach(function(d){
        var v=d.app[a.key].active; sum+=v;
        if(v>peak){ peak=v; peakDate=d.date; }
      });
      return { key:a.key, short:a.short, color:a.color, det:a.det,
               avg: trend.length? sum/trend.length : 0, peak:peak, peakDate:peakDate,
               latest: trend.length? trend[trend.length-1].app[a.key].active : 0 };
    });

    /* All-active daily stats */
    var allSum=0, allPeak=0, allPeakDate="";
    trend.forEach(function(d){ allSum+=d.allActive; if(d.allActive>allPeak){ allPeak=d.allActive; allPeakDate=d.date; } });
    var allAvg = trend.length? allSum/trend.length : 0;
    var enabledLatest = trend.length? trend[trend.length-1].allEnabled : allEnabledSnap;

    /* User detail: per-person last-activity dates. Two possible sources —
       the richer CopilotActivityUserDetail export (prompt counts, active days,
       more app columns) takes precedence over the basic usage export. */
    var usersRich = !!RAW.userDetailV4, userApps, users, richStats=null;
    if(usersRich){
      userApps = RICH_APPS.map(function(a){ return {key:a.key, short:a.short, color:a.color}; });
      users = RAW.userDetailV4.rows.map(function(r){
        var upn = getUserField(r, ["User Principal Name", "userPrincipalName"]);
        var name = getUserField(r, ["Display Name", "displayName"]) || upn;
        var last = getUserField(r, ["Last Activity Date", "lastActivityDate"]);
        var prompts = num(getUserField(r, ["Prompts submitted for All Apps", "promptsAllApps", "Prompts submitted"]));
        var promptsWork = num(getUserField(r, ["Prompts submitted for Copilot Chat (work)", "promptsWork"]));
        var promptsWeb = num(getUserField(r, ["Prompts submitted for Copilot Chat (web)", "promptsWeb"]));
        var days = num(getUserField(r, ["Active Usage Days for All Apps", "activeUsageDaysAllApps", "Active usage days"]));
        var u={ name:name, upn:upn, last:last, prompts:prompts, promptsWork:promptsWork, promptsWeb:promptsWeb, days:days, app:{} };
        RICH_APPS.forEach(function(a){
          u.app[a.key] = getUserField(r, [a.col, a.key + " Last Activity Date", a.short + " Last Activity Date", a.key, a.short]);
        });
        return u;
      }).filter(function(u){ return u.name || u.upn; });
      var tp=0, td=0, topU=null;
      users.forEach(function(u){ tp+=u.prompts; td+=u.days; if(!topU||u.prompts>topU.prompts) topU=u; });
      richStats = { totalPrompts:tp, avgPrompts: users.length? tp/users.length : 0,
                    avgDays: users.length? td/users.length : 0, top:topU };
    } else {
      userApps = APPS.filter(function(a){return a.det;}).map(function(a){ return {key:a.key, short:a.short, color:a.color}; });
      users = RAW.userDetail.rows.map(function(r){
        var upn = getUserField(r, ["userPrincipalName", "User Principal Name"]);
        var name = getUserField(r, ["displayName", "Display Name"]) || upn;
        var last = getUserField(r, ["lastActivityDate", "Last Activity Date"]);
        var u={ name:name, upn:upn, last:last, prompts:0, days:0, app:{} };
        APPS.forEach(function(a){
          if(a.det) {
            u.app[a.key] = getUserField(r, [a.det, a.key + " Last Activity Date", a.short + " Last Activity Date", a.key, a.short]);
          }
        });
        return u;
      });
    }
    /* Reach by app = how many detailed users have any recorded activity in that app. */
    var reach = userApps.map(function(a){
      var c=0; users.forEach(function(u){ if(u.app[a.key]) c++; });
      return { key:a.key, short:a.short, color:a.color, v:c };
    });

    var windowDays = (trend.length && trend[0].period) || num(prodRow["Report Period"]) || detectedPeriod || 30;
    var dmin = trend.length? trend[0].date : "", dmax = trend.length? trend[trend.length-1].date : "";

    DATA = {
      trend:trend, snapshot:snapshot, appStats:appStats, reach:reach,
      allEnabledSnap:allEnabledSnap, allActiveSnap:allActiveSnap,
      enabledLatest:enabledLatest, allAvg:allAvg, allPeak:allPeak, allPeakDate:allPeakDate,
      users:users, totalUsers:users.length,
      usersRich:usersRich, userApps:userApps, richStats:richStats,
      refresh:refresh, windowDays:windowDays, period:{from:dmin,to:dmax}
    };
  }

  // Line chart SVG generator and interactive crosshair tracker
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
      var lab=(data[i].label||"").slice(5);
      xl += '<text x="'+g.X(i).toFixed(1)+'" y="'+(H-7)+'" text-anchor="middle" font-size="10" fill="var(--muted)">'+esc(lab)+'</text>';
    });
    var uid="cg"+Math.random().toString(36).slice(2,8);
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

  // Bar and donut SVG chart templates
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

  // SortTable class definition
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

  // Renderers
  function renderHeader(){
    var p=DATA.period, wd=DATA.windowDays;
    $("cpSubtitle").innerHTML = "Microsoft 365 Copilot adoption across The Mosaic Company &middot; rolling "+wd+"-day window";
    var pills=[];
    pills.push('<span class="pill">Window: '+wd+' days</span>');
    if(p.from&&p.to) pills.push('<span class="pill">'+esc(p.from)+' to '+esc(p.to)+'</span>');
    if(DATA.refresh) pills.push('<span class="pill">Data refreshed '+esc(DATA.refresh)+'</span>');
    pills.push('<span class="pill">'+fmt(DATA.enabledLatest||DATA.allEnabledSnap)+' Copilot-enabled users</span>');
    $("cpMetaPills").innerHTML=pills.join("");
    $("cpFootNote").innerHTML = "Generated from Microsoft 365 Copilot usage exports &middot; "+wd+"-day window, "+esc(p.from)+" &ndash; "+esc(p.to)+
      " &middot; the three exports use different definitions of “active” — each panel notes its source.";
  }

  function renderKpis(){
    var d=DATA;
    var enabled = d.enabledLatest || d.allEnabledSnap;
    var winAdopt = pct(d.allActiveSnap, d.allEnabledSnap);
    var top = d.appStats.slice().filter(function(a){return a.avg>0;}).sort(function(a,b){return b.avg-a.avg;})[0];
    var chat = d.appStats.filter(function(a){return a.key==="Copilot Chat";})[0] || {peak:0,peakDate:""};
    var cards=[
      {h:"Copilot-enabled users", n:fmt(enabled), s:"licensed in this period", bar:100},
      {h:"Active in window", n:fmt(d.allActiveSnap), s:winAdopt+"% of "+fmt(d.allEnabledSnap)+" enabled (any app)", bar:winAdopt},
      {h:"Peak daily active", n:fmt(d.allPeak), s:d.allPeakDate?("on "+d.allPeakDate):"", bar:pct(d.allPeak,enabled)},
      {h:"Avg daily active", n:fmt(d.allAvg), s:"across "+d.trend.length+" days", bar:pct(d.allAvg,enabled)},
      {h:"Top Copilot app", n:top?top.short:"—", s:top?(fmt(top.avg)+" avg daily active"):"", bar:top?pct(top.avg,d.allAvg):0},
      {h:"Copilot Chat", n:fmt(chat.peak), s:"peak daily active"+(chat.peakDate?(" · "+chat.peakDate):""), bar:pct(chat.peak,d.allPeak)}
    ];
    $("cpKpiRow").innerHTML = cards.map(function(c){
      return '<div class="card kpi"><h3>'+esc(c.h)+'</h3><div class="num">'+esc(c.n)+'</div>'+
        '<div class="sub">'+esc(c.s)+'</div><div class="bar"><span style="width:'+Math.min(100,c.bar)+'%"></span></div></div>';
    }).join("");
  }

  function renderOverview(){
    renderLine($("cpAllTrend"), DATA.trend.map(function(d){return {label:d.date,v:d.allActive};}),
               {area:true,color:"var(--c1)",h:200,seriesLabel:"Active users"});
    // Average daily active users per app (from the trend export = Copilot usage).
    var bars=DATA.appStats.filter(function(a){return a.avg>0;}).sort(function(a,b){return b.avg-a.avg;})
              .map(function(a){return {label:a.short,v:a.avg,color:a.color};});
    $("cpByApp").innerHTML = hBarChart(bars,{labelW:120,fmt:function(v){return fmt(v);},
      tipv:function(d){return fmt(d.v)+" avg / day";}});
  }

  function renderTrends(){
    var series=[{el:"cpTrAll",key:"__all",color:"var(--c1)",label:"All active"}];
    DATA.appStats.filter(function(a){return a.peak>0;}).forEach(function(a){
      series.push({el:"cp_"+a.key, key:a.key, color:a.color, label:a.short});
    });
    // Build the per-app chart cards dynamically so we only show apps with data.
    var host=$("cpTrendGrid");
    host.innerHTML = series.map(function(s){
      var note = s.key==="__all"? '<p class="cnote">unique people using any Copilot capability each day</p>' : '';
      return '<div class="card chart-card"><p class="ctitle">'+esc(s.label)+'</p>'+note+'<div id="'+esc(s.el)+'"></div></div>';
    }).join("");
    series.forEach(function(s){
      var data = DATA.trend.map(function(d){
        return {label:d.date, v: s.key==="__all"? d.allActive : d.app[s.key].active};
      });
      renderLine($(s.el), data, {color:s.color, area:true, h:150, seriesLabel:s.label});
    });
  }

  function buildProductTable(){
    var cols=[
      {k:"short", t:"Product", sort:"string", render:function(r){return '<span class="dot" style="background:'+r.color+'"></span>'+esc(r.short);}},
      {k:"enabled", t:"Enabled", num:true, sort:"number", render:function(r){return fmt(r.enabled);}},
      {k:"active", t:"Active", num:true, sort:"number", render:function(r){return fmt(r.active);}},
      {k:"adoption", t:"Adoption", num:true, sort:"number", render:function(r){
        var a=r.adoption; return '<div class="minibar"><span style="width:'+Math.min(100,a)+'%;background:'+r.color+'"></span></div><span class="mbv">'+a+'%</span>';}},
      {k:"avgDaily", t:"Avg daily active", num:true, sort:"number", render:function(r){return fmt(r.avgDaily);}}
    ];
    prodTbl=new SortTable($("cpProdTable"), cols, {
      initSort:{key:"active",dir:-1},
      getRows:function(){
        var statByKey={}; DATA.appStats.forEach(function(a){ statByKey[a.key]=a; });
        return DATA.snapshot.map(function(s){
          var st=statByKey[s.key]||{avg:0};
          return { key:s.key, short:s.short, color:s.color, enabled:s.enabled, active:s.active,
                   adoption:pct(s.active,s.enabled), avgDaily:st.avg };
        });
      },
      onCount:function(total){ $("cpProdCount").textContent = total+" products"; }
    });
  }
  function renderProducts(){
    if(!prodTbl) buildProductTable();
    prodTbl.render();
    // adoption% bars per product (snapshot)
    var bars=DATA.snapshot.map(function(s){return {label:s.short,v:pct(s.active,s.enabled),color:s.color};})
              .sort(function(a,b){return b.v-a.v;});
    $("cpAdoptBars").innerHTML = hBarChart(bars,{labelW:120,fmt:function(v){return v+"%";},
      tipv:function(d){return d.v+"% adoption";}});
  }

  function buildUsersTable(){
    var rich=DATA.usersRich, dcols=DATA.userApps;
    var cols=[ {k:"name", t:"Person", sort:"string", render:function(u){return esc(u.name);}} ];
    if(rich){
      cols.push({k:"prompts", t:"Prompts", num:true, sort:"number", render:function(u){return fmt(u.prompts);}});
      cols.push({k:"days", t:"Active days", num:true, sort:"number", render:function(u){return fmt(u.days);}});
    }
    cols.push({k:"last", t:"Last activity", sort:"string", render:function(u){return esc(u.last||"—");}});
    dcols.forEach(function(a){
      cols.push({k:"d_"+a.key, t:a.short, sort:"string", render:function(u){
        var v=u.app[a.key]; return v? esc(v) : '<span class="muted">—</span>';
      }});
    });
    usrTbl=new SortTable($("cpUsrTable"), cols, {
      initSort:{key: rich?"prompts":"last", dir:-1}, max:300,
      getRows:function(){
        var q=($("cpUsrSearch").value||"").toLowerCase().trim();
        return DATA.users.filter(function(u){
          if(!q) return true;
          return u.name.toLowerCase().indexOf(q)>=0 || u.upn.toLowerCase().indexOf(q)>=0;
        }).map(function(u){
          // flatten per-app dates so the sort engine can read them by key
          var o={name:u.name, upn:u.upn, last:u.last, prompts:u.prompts||0, days:u.days||0, app:u.app};
          dcols.forEach(function(a){ o["d_"+a.key]=u.app[a.key]||""; });
          return o;
        });
      },
      onCount:function(total,shown){
        $("cpUsrCount").textContent = fmt(total)+" people"+(total>shown?" (showing top "+shown+")":"");
      }
    });
  }
  function renderPeopleKpis(){
    var host=$("cpPeopleKpis"); if(!host) return;
    if(!DATA.usersRich || !DATA.richStats){ host.classList.add("hidden"); host.innerHTML=""; return; }
    var s=DATA.richStats;
    var cards=[
      {h:"People in report", n:fmt(DATA.totalUsers), s:"from the detailed per-user export"},
      {h:"Total prompts", n:fmt(s.totalPrompts), s:"submitted across all apps"},
      {h:"Avg prompts / person", n:fmt(s.avgPrompts), s:"mean over the window"},
      {h:"Avg active days", n:fmt(s.avgDays), s:"per person, all apps"},
      {h:"Most active", n:s.top?clip(s.top.name,18):"—", s:s.top?(fmt(s.top.prompts)+" prompts"):""}
    ];
    host.innerHTML = cards.map(function(c){
      return '<div class="card kpi"><h3>'+esc(c.h)+'</h3><div class="num">'+esc(c.n)+'</div>'+
        '<div class="sub">'+esc(c.s)+'</div></div>';
    }).join("");
    host.classList.remove("hidden");
  }
  function renderUsers(){
    // Column layout differs between the basic and richer exports; rebuild the
    // table if the source kind changed (e.g. after a Refresh with a new file).
    if(usrTbl && usrTblRich!==DATA.usersRich){ usrTbl=null; }
    if(!usrTbl){ buildUsersTable(); usrTblRich=DATA.usersRich; }
    usrTbl.render();
    renderPeopleKpis();
    var bars=DATA.reach.slice().sort(function(a,b){return b.v-a.v;})
              .map(function(r){return {label:r.short,v:r.v,color:r.color};});
    $("cpReachBars").innerHTML = hBarChart(bars,{labelW:120});
    $("cpReachNote").textContent = "Based on the "+fmt(DATA.totalUsers)+" people in the per-user export. "+
      "A person is counted for an app if the export records any Copilot activity there.";
  }

  var hasData = false;

  function renderAll(){
    compute();
    renderHeader(); renderKpis(); renderOverview(); renderTrends();
    renderProducts(); renderUsers();
    $("cpStatus").classList.add("hidden");
    $("cpDash").classList.remove("hidden");
    hasData = true;
  }

  // Tab navigation switching
  function initTabs(){
    var nav=$("cpTabs"); if(!nav) return;
    var btns=nav.querySelectorAll("button");
    btns.forEach(function(b){
      b.addEventListener("click",function(){
        btns.forEach(function(x){x.setAttribute("aria-selected","false");});
        b.setAttribute("aria-selected","true");
        $("cp-app").querySelectorAll(".cp-tabpanel").forEach(function(p){p.classList.add("hidden");});
        $("cptab-"+b.dataset.tab).classList.remove("hidden");
      });
    });
  }

  function openUploader(){
    $("cpStatus").classList.remove("hidden");
    $("cpLoadingBox").classList.add("hidden");
    $("cpFallbackBox").classList.remove("hidden");
    $("cpDash").classList.add("hidden");
    updateChecklist();
    toggleBack();
  }

  function backToReports(){
    if(!hasData) return;
    $("cpStatus").classList.add("hidden");
    $("cpFallbackBox").classList.add("hidden");
    $("cpDash").classList.remove("hidden");
  }

  function toggleBack(){
    var b=$("cpBackBtn"); if(!b) return;
    if(hasData) b.classList.remove("hidden"); else b.classList.add("hidden");
  }

  function clearData(){
    RAW = {};
    DATA = {};
    fileMap = {};
    hasData = false;
    usrTbl = null;
    usrTblRich = null;
    prodTbl = null;
    if(window.CopilotExtras && window.CopilotExtras.clear){
      window.CopilotExtras.clear();
    }
    updateChecklist();
    toggleBack();
    $("cpStatus").classList.remove("hidden");
    $("cpLoadingBox").classList.add("hidden");
    $("cpFallbackBox").classList.remove("hidden");
    $("cpDash").classList.add("hidden");
  }

  function initControls(){
    $("cpUsrSearch").addEventListener("input",renderUsers);
    
    // Wire up dropdown controls
    var cpBtn = $("cpReloadBtn");
    var cpDropdown = $("cpReloadDropdown");
    if(cpBtn && cpDropdown){
      cpBtn.addEventListener("click", function(e){
        e.stopPropagation();
        Array.prototype.forEach.call(document.querySelectorAll(".refresh-dropdown"), function(d){
          if(d !== cpDropdown) d.classList.remove("show");
        });
        cpDropdown.classList.toggle("show");
      });
    }
    var cpUpReplace = $("cpUploadReplace");
    if(cpUpReplace){
      cpUpReplace.addEventListener("click", function(){
        cpDropdown.classList.remove("show");
        openUploader();
      });
    }
    var cpClearData = $("cpClearData");
    if(cpClearData){
      cpClearData.addEventListener("click", function(){
        cpDropdown.classList.remove("show");
        if(confirm("Are you sure you want to clear all loaded data? This will reset the dashboard.")){
          clearData();
        }
      });
    }
    var cpBack = $("cpBackBtn");
    if(cpBack){
      cpBack.addEventListener("click", backToReports);
    }

    // hover tooltips for bar/donut, scoped to the Copilot container
    var root=$("cp-app")||document;
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
      if(e.target.closest && e.target.closest(".chartbox")) return;
      Tip.hide();
    });
  }

  // Loading logic
  function fetchText(url){
    return fetch(url,{cache:"no-store"}).then(function(res){
      if(!res.ok) throw new Error(url+" ("+res.status+")");
      return res.text();
    });
  }
  // Probe period (30/60/90/180) x suffix style ( _D30 / _'D30' / _30 ) using the
  // trend export, then load the matching by-product + user-detail files.
  function fetchAll(){
    detectedPeriod=0;
    var combos=[];
    PERIODS.forEach(function(p){ suffixes(p).forEach(function(sx){ combos.push({p:p,sx:sx}); }); });
    var i=0;
    function tryNext(){
      if(i>=combos.length) return Promise.reject(new Error("No Copilot report files were found in this folder."));
      var combo=combos[i++], sx=combo.sx;
      return fetchText(FOLDER+BASES.trend+sx+".csv").then(function(txt){
        detectedPeriod=combo.p; detectedSuffix=sx; RAW.trend=parseCSV(txt);
        return Promise.all(["byProduct","userDetail"].map(function(k){
          return fetchText(FOLDER+BASES[k]+sx+".csv").then(function(t){ RAW[k]=parseCSV(t); });
        }));
      }).catch(function(e){
        if(detectedPeriod) throw e;  // window found but a sibling missing -> surface
        return tryNext();
      });
    }
    return tryNext();
  }
  function boot(isReload){
    $("cpStatus").classList.remove("hidden");
    $("cpLoadingBox").classList.remove("hidden");
    $("cpFallbackBox").classList.add("hidden");
    if(isReload) $("cpDash").classList.add("hidden");
    fetchAll().then(function(){ renderAll(); }).catch(function(err){ showFallback(err); });
  }

  // Fallback uploader
  var fileMap = {};
  function nameToKey(fn){
    fn=fn.toLowerCase();
    if(fn.indexOf("adoptionbyproduct")>=0) return "byProduct";
    if(fn.indexOf("adoptiontrend")>=0)     return "trend";
    // The richer per-user export — "CopilotActivityUserDetail" (no "EDP") — takes
    // precedence over the basic usage export. The EDP "ActivityUserDetail" file
    // also matches "activityuserdetail" but carries the "edp" marker and belongs
    // to the optional Copilot Chat set, so route it to CopilotExtras instead.
    if(fn.indexOf("activityuserdetail")>=0 && fn.indexOf("edp")<0) return "userDetailV4";
    if(fn.indexOf("usageuserdetail")>=0 || (fn.indexOf("userdetail")>=0 && fn.indexOf("activityuserdetail")<0)) return "userDetail";
    return null;
  }
  function showFallback(){
    $("cpLoadingBox").classList.add("hidden");
    $("cpFallbackBox").classList.remove("hidden");
    updateChecklist();
    toggleBack();
  }
  // Optional Agent and Chat reports
  var OPTIONAL = [
    {k:"declAgents",  label:"Declarative agents — 30-day usage"},
    {k:"agents",      label:"All agents — inventory"},
    {k:"chatAdopt",   label:"Copilot Chat — adoption by app"},
    {k:"chatPrompts", label:"Copilot Chat — prompts submitted by app"},
    {k:"chatUsers",   label:"Copilot Chat — end-user usage details"}
  ];
  function updateChecklist(){
    var st = (window.CopilotExtras && window.CopilotExtras.status) ? window.CopilotExtras.status() : {};
    var req = Object.keys(LABELS).map(function(k){
      var ok = k==="userDetail" ? !!(fileMap.userDetail||fileMap.userDetailV4) : !!fileMap[k];
      return '<li class="'+(ok?"ok":"pending")+'">'+(ok?"✔":"○")+" "+esc(LABELS[k])+'<span class="muted"> &middot; required</span></li>';
    }).join("");
    var opt = OPTIONAL.map(function(o){
      var ok=!!st[o.k];
      return '<li class="'+(ok?"ok":"pending")+'">'+(ok?"✔":"○")+" "+esc(o.label)+'<span class="muted"> &middot; optional</span></li>';
    }).join("");
    $("cpFileChecklist").innerHTML = req + opt;
  }
  function ingestFiles(fileList){
    // Forward everything to the extras module too; it picks out the optional
    // Agents / Copilot Chat exports and ignores the rest.
    var extra = (window.CopilotExtras && window.CopilotExtras.ingest)
      ? window.CopilotExtras.ingest(fileList) : Promise.resolve();
    extra.then(updateChecklist);
    var arr=Array.prototype.slice.call(fileList);
    var jobs=arr.map(function(f){
      var key=nameToKey(f.name);
      if(!key) return Promise.resolve();
      return f.text().then(function(txt){ fileMap[key]=parseCSV(txt); });
    });
    Promise.all(jobs).then(function(){
      updateChecklist();
      // The People tab accepts either per-user export; userDetailV4 (the richer
      // CopilotActivityUserDetail) satisfies the same slot as the basic file.
      var need=[ !!fileMap.byProduct, !!fileMap.trend, !!(fileMap.userDetail||fileMap.userDetailV4) ];
      var have=need.filter(function(ok){return ok;}).length;
      if(have===3){ RAW=fileMap; renderAll(); }
      else { $("cpFallbackErr").textContent = "Loaded "+have+" of 3 required files. Please add the remaining "+(3-have)+"."; }
    });
  }
  function initFallback(){
    var dz=$("cpDropZone"), fi=$("cpFileInput");
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

  // Module export
  function init(){ initTabs(); initControls(); initFallback(); }
  window.CopilotBoard = {
    booted:false,
    init:init,
    // Boot exactly once, the first time the Copilot tab is opened. Re-selecting
    // the tab must NOT re-run the loader, otherwise the "add the data files"
    // panel reappears above an already-loaded dashboard. (Use the Refresh
    // button to deliberately reload.)
    boot:function(){ if(this.booted) return; this.booted=true; init(); boot(false); }
  };
})();
