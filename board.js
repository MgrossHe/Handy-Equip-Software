/* Office Board — main board (runs after sign-in). */
(function(){
  var OL = "https://outlook.office365.com/owa/?ItemID=";
  var TAIL = "&exvsurl=1&viewmodel=ReadMessageItem";
  function mail(id){ return OL + id + TAIL; }

  // Promise status: your notes, the Squid export of Nov 23, and your email (read Nov 24).
  var P0 = (window.__OB && window.__OB.promises) || {};
  var PROMISES = P0.PROMISES || [], KEPT = P0.KEPT || [], SEEN = P0.SEEN || [];

  var D = null, view = "overview";
  var invState = {q:"", bucket:"", status:"Open", sort:"bal", dir:-1, cust:null};
  var custState = {q:"", sort:"open", dir:-1};
  var BUCKETS = ["1-30","31-60","61-90","91-180","181-365","Over 1 year"];
  function tot(a){ return a.reduce(function(s,r){ return s + (r.bal || 0); }, 0); }
  // Split each customer's invoices into approved (Open) and In Review (not approved, not due).
  function prepare(){
    var by = {};
    D.invoices.forEach(function(r){
      var k = by[r.customer] || (by[r.customer] = {open:0, review:0, nOpen:0, oldestOpen:null});
      if (r.status === "In Review") k.review += r.bal || 0;
      else { k.open += r.bal || 0; k.nOpen++; if (r.dpd != null && (k.oldestOpen == null || r.dpd > k.oldestOpen)) k.oldestOpen = r.dpd; }
    });
    D.customers.forEach(function(c){
      var k = by[c.customer] || {open:0, review:0, nOpen:0, oldestOpen:null};
      c.open = Math.round(k.open * 100) / 100; c.review = Math.round(k.review * 100) / 100; c.nOpen = k.nOpen; c.oldestOpen = k.oldestOpen;
    });
  }
  var OVER90 = {"91-180":1,"181-365":1,"Over 1 year":1};

  function esc(s){ return String(s == null ? "" : s).replace(/[&<>"']/g, function(ch){ return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[ch]; }); }
  function money(n){ if (n == null || n === "") return ""; var neg = n < 0; var s = Math.abs(n).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}); return (neg ? "-$" : "$") + s; }
  function money0(n){ return "$" + Math.round(n).toLocaleString("en-US"); }
  function $(id){ return document.getElementById(id); }
  function custByName(n){ if (!D) return null; for (var i = 0; i < D.customers.length; i++){ if (D.customers[i].customer === n) return D.customers[i]; } return null; }
  function findCust(name){
    if (!D) return null;
    var key = name.toLowerCase().replace(/[^a-z0-9]/g,"");
    for (var i = 0; i < D.customers.length; i++){
      var k = D.customers[i].customer.toLowerCase().replace(/[^a-z0-9]/g,"");
      if (k === key || k.indexOf(key) === 0 || key.indexOf(k) === 0) return D.customers[i];
    }
    return null;
  }

  /* ---------- routing ---------- */
  var VIEWS = {home:1, overview:1, invoices:1, customers:1, promises:1, hershy:1, email:1, sales:1, service:1, purchasing:1, report:1};
  function show(v){
    view = v;
    var vs = document.querySelectorAll(".view");
    for (var i = 0; i < vs.length; i++) vs[i].hidden = vs[i].getAttribute("data-view") !== v;
    var ns = document.querySelectorAll("[data-nav]");
    for (var j = 0; j < ns.length; j++){
      if (ns[j].getAttribute("data-nav") === v) ns[j].setAttribute("aria-current","page"); else ns[j].removeAttribute("aria-current");
    }
    if (v === "invoices") renderInvoices();
    if (v === "customers") renderCustomers();
    if (v === "email") renderEmail();
    var sec = v === "home" ? "home" : v === "email" ? "email" : (v === "sales" || v === "service" || v === "report") ? "cust" : v === "purchasing" ? "purch" : "ar";
    if (KINDS[v]) renderTracker(v);
    if (v === "report") renderReport();
    document.body.setAttribute("data-sec", sec);
    var ms = document.querySelectorAll("[data-sec]");
    for (var m = 0; m < ms.length; m++){ if (ms[m] === document.body) continue;
      if (ms[m].getAttribute("data-sec") === sec) ms[m].setAttribute("aria-current","page"); else ms[m].removeAttribute("aria-current"); }
    if (v === "home") renderHome();
    window.scrollTo(0, 0);
  }
  function route(){
    var h = (location.hash || "").replace(/^#/, "");
    if (h === "over90"){ invState.cust = null; invState.bucket = "90+"; invState.status = "Open"; $("invBucket").value = "90+"; $("invStatus").value = "Open"; show("invoices"); return; }
    if (h === "review"){ invState.cust = null; invState.bucket = ""; invState.status = "In Review"; $("invBucket").value = ""; $("invStatus").value = "In Review"; show("invoices"); return; }
    if (h === "invoices" && !pendingCust){ /* keep current filters */ }
    show(VIEWS[h] ? h : "home");
  }
  var pendingCust = false;
  function openCustomer(name){
    invState.cust = name; invState.q = ""; invState.bucket = ""; invState.status = "";
    $("invQ").value = ""; $("invBucket").value = ""; $("invStatus").value = "";
    pendingCust = true;
    if (location.hash === "#invoices") show("invoices"); else location.hash = "invoices";
    pendingCust = false;
  }
  window.addEventListener("hashchange", route);

  /* ---------- overview ---------- */
  function renderOverview(){
    var open = D.invoices.filter(function(r){ return r.status !== "In Review"; });
    var rev = D.invoices.filter(function(r){ return r.status === "In Review"; });
    var o90 = open.filter(function(r){ return OVER90[r.bucket]; });
    $("tTotal").textContent = money0(tot(open)); $("tInv").textContent = open.length;
    $("tCust").textContent = D.customers.filter(function(c){ return c.open > 0; }).length;
    $("t90").textContent = money0(tot(o90)); $("t90n").textContent = o90.length + " invoices";
    $("tRev").textContent = money0(tot(rev)); $("tRevN").textContent = rev.length + " invoices · not approved, not due";
    var aging = BUCKETS.map(function(b){ var rs = open.filter(function(r){ return r.bucket === b; }); return {bucket: b, amount: tot(rs), n: rs.length}; });
    var max = 0; aging.forEach(function(a){ if (a.amount > max) max = a.amount; });
    var html = "";
    aging.forEach(function(a){
      if (!a.amount) return;
      var n = a.n;
      var old = OVER90[a.bucket] ? " old" : "";
      html += '<button type="button" class="agrow' + old + '" data-bucket="' + esc(a.bucket) + '"><span class="lb">' + esc(a.bucket === "Over 1 year" ? "1 year +" : a.bucket + " days") +
        '</span><span class="track"><span class="fill" style="display:block;width:' + (a.amount / max * 100).toFixed(1) + '%"></span></span>' +
        '<span class="amt">' + money0(a.amount) + '</span><span class="cnt">' + esc(n) + ' inv</span></button>';
    });
    $("aging").innerHTML = html;
    var top = D.customers.slice().sort(function(a,b){ return (b.open||0) - (a.open||0); }).slice(0, 10);
    $("topCust").innerHTML = top.map(function(c){
      return '<tr class="click" data-cust="' + esc(c.customer) + '"><td>' + esc(c.customer) + '</td><td class="num">' + money(c.open) +
        '</td><td class="num">' + esc(c.nOpen) + '</td><td class="num">' + esc(c.oldestOpen) + '</td><td class="wrapc muted">' + esc(c.notes) + '</td></tr>';
    }).join("");
  }

  /* ---------- invoices ---------- */
  function filteredInvoices(){
    var q = invState.q.toLowerCase();
    return D.invoices.filter(function(r){
      if (invState.cust && r.customer !== invState.cust) return false;
      if (invState.bucket){ if (invState.bucket === "90+"){ if (!OVER90[r.bucket]) return false; } else if (r.bucket !== invState.bucket) return false; }
      if (invState.status && r.status !== invState.status) return false;
      if (q){ var hay = (r.customer + " " + r.inv + " " + r.so + " " + r.po + " " + r.email + " " + r.rep + " " + r.note).toLowerCase(); if (hay.indexOf(q) < 0) return false; }
      return true;
    });
  }
  function sortRows(rows, key, dir){
    return rows.sort(function(a,b){
      var x = a[key], y = b[key];
      if (typeof x === "number" || typeof y === "number"){ return ((x||0) - (y||0)) * dir; }
      return String(x||"").localeCompare(String(y||"")) * dir;
    });
  }
  function renderInvoices(){
    if (!D) return;
    var rows = sortRows(filteredInvoices(), invState.sort, invState.dir);
    var total = rows.reduce(function(s,r){ return s + (r.bal||0); }, 0);
    var c = invState.cust ? custByName(invState.cust) : null;
    $("backCust").hidden = !invState.cust;
    $("invTitle").textContent = invState.cust ? invState.cust : invState.status === "In Review" ? "In Review invoices (not approved, not due)" :
      (invState.bucket === "90+" ? "Invoices over 90 days" : "Open invoices");
    if (c){
      $("custProfile").innerHTML = '<div class="card profile">' +
        prof("Owed, approved (Nov 23)", money(c.open)) + prof("In Review (not due)", money(c.review)) + prof("Squid balance (net of credits)", money(c.squid)) +
        prof("Oldest approved invoice", (c.oldestOpen != null ? c.oldestOpen + " days" : "—")) + prof("Your sheet note", c.notes || "—") +
        prof("Follow up", c.follow || "—") + prof("Email", c.email || "—") + prof("Phone", c.phone || "—") + prof("Sales rep", c.rep || "—") + '</div>';
    } else { $("custProfile").innerHTML = ""; }
    renderNotesPanel(); renderFollow();
    var apr = rows.filter(function(r){ return r.status !== "In Review"; }), nRev = rows.length - apr.length, aprT = tot(apr);
    $("invSum").innerHTML = (apr.length || !nRev ? "<b>" + apr.length + "</b> approved · <b>" + money(aprT) + "</b> owed" : "") +
      (nRev ? (apr.length ? " &nbsp;·&nbsp; " : "") + "<b>" + nRev + "</b> In Review · " + money(total - aprT) + " (not due yet)" : "");
    var tb = $("invTable").querySelector("tbody");
    if (!rows.length){ tb.innerHTML = '<tr><td colspan="9" class="empty">No invoices match.</td></tr>'; }
    else tb.innerHTML = rows.map(function(r){
      var st = r.status === "In Review" ? '<span class="chip amber">In Review</span>' : '<span class="chip slate">Open</span>';
      var days = (r.dpd != null && r.dpd > 90) ? '<span style="color:var(--red);font-weight:600">' + r.dpd + '</span>' : esc(r.dpd);
      return '<tr' + (invState.cust ? '' : ' class="click" data-cust="' + esc(r.customer) + '"') + '><td>' + esc(r.customer) + '</td><td>' + esc(r.inv) +
        (r.so ? ' <span class="muted">' + esc(r.so) + '</span>' : '') + '</td><td>' + st + '</td><td>' + esc(r.date) + '</td><td class="num">' + days +
        '</td><td class="num">' + money(r.orig) + '</td><td class="num">' + money(r.paid) + '</td><td class="num"><b>' + money(r.bal) + '</b></td><td>' + esc(r.rep) + '</td></tr>';
    }).join("");
    markSort($("invTable"), invState);
  }
  function prof(k, v){ return '<div><div class="k">' + esc(k) + '</div><div class="v">' + esc(v) + '</div></div>'; }
  function markSort(table, st){
    var ths = table.querySelectorAll("th[data-sort]");
    for (var i = 0; i < ths.length; i++){
      var a = ths[i].querySelector(".arr"); if (a) a.remove();
      if (ths[i].getAttribute("data-sort") === st.sort){ var s = document.createElement("span"); s.className = "arr"; s.textContent = st.dir > 0 ? "▲" : "▼"; ths[i].appendChild(s); }
    }
  }

  /* ---------- customers ---------- */
  function renderCustomers(){
    if (!D) return;
    var q = custState.q.toLowerCase();
    var rows = D.customers.filter(function(c){
      if (!q) return true;
      return (c.customer + " " + c.notes + " " + c.follow + " " + c.rep + " " + c.email + " " +
        notesFor(c.customer).map(function(n){ return n.text; }).join(" ")).toLowerCase().indexOf(q) >= 0;
    });
    rows = sortRows(rows, custState.sort, custState.dir);
    var total = rows.reduce(function(s,c){ return s + (c.open||0); }, 0);
    var owing = rows.filter(function(c){ return c.open > 0; }).length, revOnly = rows.length - owing;
    $("custSum").innerHTML = "<b>" + owing + "</b> customers owe · <b>" + money(total) + "</b> approved" +
      (revOnly ? " &nbsp;·&nbsp; " + revOnly + " more have only In Review invoices" : "");
    var tb = $("custTable").querySelector("tbody");
    tb.innerHTML = rows.length ? rows.map(function(c){
      return '<tr class="click" data-cust="' + esc(c.customer) + '"><td>' + esc(c.customer) + '</td><td class="num"><b>' + money(c.open) + '</b></td><td class="num muted">' +
        (c.review ? money(c.review) : '') + '</td><td class="num">' + esc(c.nOpen) +
        '</td><td class="num">' + esc(c.oldestOpen) + '</td><td class="num">' + (c.credit ? money(c.credit) : '') + '</td><td class="wrapc">' + esc(c.notes) +
        '</td><td class="wrapc muted">' + esc(c.follow) + '</td><td class="wrapc">' + latestNoteCell(c.customer) + '</td><td>' + esc(c.rep) + '</td></tr>';
    }).join("") : '<tr><td colspan="10" class="empty">No customers match.</td></tr>';
    markSort($("custTable"), custState);
  }

  /* ---------- promises ---------- */
  function linkHtml(pair){ return '<a href="' + mail(pair[1]) + '" target="_blank" rel="noopener">✉ ' + esc(pair[0]) + '</a>'; }
  function renderPromises(){
    $("pOpen").innerHTML = PROMISES.map(function(p){
      var c = findCust(p.c);
      var amt = c ? money(c.open) : "";
      return '<div class="pc ' + p.cls + '"><div class="row1"><button type="button" class="nm" data-cust="' + esc(c ? c.customer : "") + '">' + esc(p.c) + '</button>' +
        '<span class="amt">' + amt + '</span></div><div><span class="chip ' + p.cls + '">' + esc(p.chip) + '</span></div>' +
        '<div class="q">' + esc(p.q) + '</div><div class="d">' + esc(p.d) + (amt ? ' Balance shown is Nov 23.' : '') + '</div>' +
        (p.links ? '<div class="links">' + p.links.map(linkHtml).join("") + '</div>' : '') + '</div>';
    }).join("");
    $("pKept").innerHTML = KEPT.map(function(p){
      return '<div class="pc green"><div class="row1"><span style="font-weight:600">' + esc(p.c) + '</span><span class="chip green">Kept</span></div><div class="d">' + esc(p.d) + '</div></div>';
    }).join("");
    $("pSeen").innerHTML = SEEN.map(function(p){
      var c = findCust(p.c);
      return '<div class="pc slate"><div class="row1">' + (c ? '<button type="button" class="nm" data-cust="' + esc(c.customer) + '">' + esc(p.c) + '</button>' : '<span style="font-weight:600">' + esc(p.c) + '</span>') +
        '<span class="chip slate">Verify</span></div><div class="d">' + esc(p.d) + '</div>' + (p.l ? '<div class="links">' + linkHtml(p.l) + '</div>' : '') + '</div>';
    }).join("");
    $("tPromise").textContent = PROMISES.length;
  }

  /* ---------- hershy ---------- */
  function notApplied(){
    return D.customers.filter(function(c){ return (c.owed || 0) > 0 && c.squid != null && (c.owed - c.squid) > 0.5; })
      .map(function(c){ return {c: c, gap: Math.round((c.owed - c.squid) * 100) / 100, full: c.squid <= 10}; })
      .sort(function(a,b){ return b.gap - a.gap; });
  }
  function renderHershy(){
    var ap = notApplied(), apTot = ap.reduce(function(s,x){ return s + x.gap; }, 0), apFull = ap.filter(function(x){ return x.full; }).length;
    $("hApSum").textContent = ap.length + " customers · " + money(apTot) + " · " + apFull + " fully paid";
    $("hAp").innerHTML = ap.map(function(x){
      return '<tr class="click" data-cust="' + esc(x.c.customer) + '"><td>' + esc(x.c.customer) + '</td><td>' +
        (x.full ? '<span class="chip green">Paid in full, not applied</span>' : '<span class="chip amber">Partly applied</span>') +
        '</td><td class="num">' + money(x.c.owed) + '</td><td class="num">' + money(x.c.squid) + '</td><td class="num"><b>' + money(x.gap) + '</b></td><td class="num">' + esc(x.c.n) + '</td></tr>';
    }).join("");
    $("tHershy").textContent = money0(apTot);
    $("tHershyN").textContent = ap.length + " customers paid but not applied";
    $("hChecks").innerHTML = D.checks.map(function(t){ return "<li>" + esc(t) + "</li>"; }).join("");
    $("hcCount").textContent = D.checks.length + " items from the Nov 23 export";
    function sum(a){ return a.reduce(function(s,x){ return s + (x.amount||0); }, 0); }
    var un = D.unapplied.slice().sort(function(a,b){ return b.amount - a.amount; });
    $("hUnSum").textContent = un.length + " payments · " + money(sum(un)) + " · money received but not matched to an invoice";
    $("hUn").innerHTML = un.map(function(r){ return '<tr><td>' + esc(r.customer) + '</td><td class="num">' + money(r.amount) + '</td><td>' + esc(r.ref) + '</td><td>' + esc(r.date) + '</td></tr>'; }).join("");
    $("hNiSum").textContent = D.nonInvoice.length + " items · " + money(sum(D.nonInvoice)) + " · inflate Squid’s Open balance";
    $("hNi").innerHTML = D.nonInvoice.map(function(r){ return '<tr><td>' + esc(r.customer) + '</td><td class="num">' + money(r.amount) + '</td><td>' + esc(r.type) + '</td><td>' + esc(r.date) + '</td></tr>'; }).join("");
    $("hCrSum").textContent = D.credits.length + " customers · " + money(sum(D.credits)) + " · money Handy may owe back or need to apply";
    $("hCr").innerHTML = D.credits.map(function(r){ return '<tr><td>' + esc(r.customer) + '</td><td class="num">' + money(r.amount) + '</td><td>' + esc(r.email) + '</td><td>' + esc(r.phone) + '</td></tr>'; }).join("");
    $("hPoSum").textContent = D.paidStillOpen.length ? D.paidStillOpen.length + " invoices · close them out in Squid" : "Not in the Nov 23 export";
    if (!D.paidStillOpen.length) { $("hPo").innerHTML = '<tr><td colspan="5" class="muted">The Nov 23 Customer Balance export doesn’t show paid-in-full invoices, so this list is empty until the next full pull.</td></tr>'; return; }
    $("hPo").innerHTML = D.paidStillOpen.map(function(r){ return '<tr><td>' + esc(r.customer) + '</td><td>' + esc(r.inv) + '</td><td class="num">' + money(r.amount) + '</td><td>' + esc(r.date) + '</td><td>' + esc(r.status) + '</td></tr>'; }).join("");
  }

  /* ---------- Excel ---------- */
  function sheet(rows, cols){ var ws = XLSX.utils.aoa_to_sheet(rows); ws["!cols"] = cols.map(function(w){ return {wch:w}; }); return ws; }
  function hershyWorkbook(){
    var wb = XLSX.utils.book_new();
    var ap = notApplied(), apNames = {};
    ap.forEach(function(x){ apNames[x.c.customer] = x; });
    XLSX.utils.book_append_sheet(wb, sheet([["Customer","Status","Open invoices","Squid balance","Not applied","# invoices"]].concat(ap.map(function(x){
      return [x.c.customer, x.full ? "Paid in full, not applied" : "Partly applied", x.c.owed, x.c.squid, x.gap, x.c.n]; })), [34,24,14,14,14,10]), "Apply payments");
    XLSX.utils.book_append_sheet(wb, sheet([["Customer","Invoice #","Invoice date","Days","Balance on invoice","Customer's not-applied total"]].concat(
      D.invoices.filter(function(r){ return apNames[r.customer]; }).sort(function(a,b){ return a.customer.localeCompare(b.customer) || String(a.date).localeCompare(String(b.date)); })
        .map(function(r){ return [r.customer, r.inv, r.date, r.dpd, r.bal, apNames[r.customer].gap]; })), [34,14,12,8,16,22]), "Invoices to apply to");
    XLSX.utils.book_append_sheet(wb, sheet([["Things to check (from the Squid export of Nov 23)"]].concat(D.checks.map(function(t,i){ return [(i+1) + ". " + t]; })), [140]), "Things to check");
    XLSX.utils.book_append_sheet(wb, sheet([["Customer","Amount","Payment #","Date"]].concat(D.unapplied.map(function(r){ return [r.customer, r.amount, r.ref, r.date]; })), [34,14,14,12]), "Unapplied payments");
    XLSX.utils.book_append_sheet(wb, sheet([["Customer","Amount","Type","Date"]].concat(D.nonInvoice.map(function(r){ return [r.customer, r.amount, r.type, r.date]; })), [34,14,20,12]), "Non-invoice on Open tab");
    XLSX.utils.book_append_sheet(wb, sheet([["Customer","Credit balance","Email","Phone"]].concat(D.credits.map(function(r){ return [r.customer, r.amount, r.email, r.phone]; })), [34,14,36,22]), "Credit balances");
    XLSX.utils.book_append_sheet(wb, sheet([["Customer","Invoice #","Invoice total","Date","Status"]].concat(D.paidStillOpen.map(function(r){ return [r.customer, r.inv, r.amount, r.date, r.status]; })), [34,14,14,12,12]), "Paid but still Open");
    return XLSX.write(wb, {type:"array", bookType:"xlsx"});
  }
  function invoicesWorkbook(){
    var rows = sortRows(filteredInvoices(), invState.sort, invState.dir);
    var aoa = [["Customer","Invoice #","SO #","Status","Invoice date","Days past due","Age","Original","Paid","Balance","Email","Phone","Rep","Note"]].concat(rows.map(function(r){
      return [r.customer, r.inv, r.so, r.status, r.date, r.dpd, r.bucket, r.orig, r.paid, r.bal, r.email, r.phone, r.rep, r.note]; }));
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet(aoa, [30,12,12,10,12,8,12,12,12,12,30,18,16,24]), "Open invoices");
    return XLSX.write(wb, {type:"array", bookType:"xlsx"});
  }
  /* ---------- requests tracker: Sales, Service, Purchasing ---------- */
  var KINDS = {
    sales: {title:"Customers — Sales", one:"sales request", who:"Customer", statuses:["New","Quoting","Quote sent","Approved","Ordered","Delivered"], stop:["Lost"],
            ph:"Paste the WhatsApp message, or write what they need. e.g. 10 short u-boats, deliver to Goshen next week"},
    service: {title:"Customers — Service", one:"service request", who:"Customer", statuses:["New","Scheduled","In progress","Waiting on parts","Done"], stop:["Cancelled"],
            ph:"Paste the message, or write what's wrong. e.g. Crown pallet jack not lifting, needs a tech this week"},
    purchasing: {title:"Purchasing", one:"purchase", who:"Vendor", statuses:["Need to order","Quote requested","Ordered","Shipped","Received"], stop:["Cancelled"],
            ph:"What we need to buy. e.g. 40 beams 96\" orange for Bingo job, ask Ridg-U-Rak for price"}
  };
  var CHANNELS = ["WhatsApp","Phone","Email","Text","In person","Other","From Mark"];
  var RQ = [], rqDb = null, rqState = "loading", RS = {}, drafts = {}, rqPending = false;
  Object.keys(KINDS).forEach(function(k){ RS[k] = {q:"", pill:"active", open:{}, adding:false, msg:""}; });
  function allSt(k){ return KINDS[k].statuses.concat(KINDS[k].stop); }
  function isClosed(k, s){ var K = KINDS[k]; return s === K.statuses[K.statuses.length - 1] || K.stop.indexOf(s) >= 0; }
  function stCls(k, s){ var K = KINDS[k]; if (s === K.statuses[K.statuses.length - 1]) return "green"; if (K.stop.indexOf(s) >= 0) return "slate"; if (s === K.statuses[0]) return "amber"; return "acc"; }
  function today(){ var d = new Date(); return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2); }
  function isLate(r){ return r.due && !isClosed(r.kind, r.status) && r.due < today(); }
  function fmtDue(d){ if (!d) return ""; var p = d.split("-"); return p.length === 3 ? (+p[1]) + "/" + (+p[2]) : d; }
  function ago(iso){ var t = Date.parse(iso); if (isNaN(t)) return ""; var m = Math.round((Date.now() - t) / 60000);
    if (m < 1) return "just now"; if (m < 60) return m + " min ago"; var h = Math.round(m / 60); if (h < 24) return h + "h ago"; var dd = Math.round(h / 24); return dd === 1 ? "yesterday" : dd + " days ago"; }
  function rqOf(k){ return RQ.filter(function(r){ return r.kind === k; }); }
  function rqSort(a, b){
    var la = isLate(a) ? 0 : 1, lb = isLate(b) ? 0 : 1; if (la !== lb) return la - lb;
    var ca = isClosed(a.kind, a.status) ? 1 : 0, cb = isClosed(b.kind, b.status) ? 1 : 0; if (ca !== cb) return ca - cb;
    return a.updatedAt < b.updatedAt ? 1 : a.updatedAt > b.updatedAt ? -1 : 0;
  }
  function dval(id, f, fallback){ var d = drafts[id]; return d && d[f] != null ? d[f] : (fallback == null ? "" : fallback); }
  function inp(id, f, val, attrs){ return '<input data-draft="' + esc(id) + '" data-f="' + f + '" value="' + esc(dval(id, f, val)) + '" ' + (attrs || "") + '>'; }
  function formFields(k, id, r){
    var K = KINDS[k], st = dval(id, "status", r ? r.status : K.statuses[0]), ch = dval(id, "channel", r ? r.channel : "");
    return '<div class="fgrid">' +
      '<label class="fl"><span>' + K.who + ' *</span>' + inp(id, "who", r && r.who, 'list="rqNames" maxlength="120" autocomplete="off" placeholder="Name"') + '</label>' +
      '<label class="fl"><span>Contact / phone</span>' + inp(id, "contact", r && r.contact, 'maxlength="120" placeholder="e.g. Yossi 845-…"') + '</label>' +
      '<div class="fl wide"><span>Came in by</span><div class="chans">' + CHANNELS.map(function(c){
        return '<button type="button" class="chan" data-chan="' + esc(c) + '" data-id="' + esc(id) + '" aria-pressed="' + (ch === c) + '">' + esc(c) + '</button>'; }).join("") + '</div></div>' +
      '<label class="fl wide"><span>' + (k === "purchasing" ? "What we need *" : "Their message / what they need *") + '</span><textarea data-draft="' + esc(id) + '" data-f="msg" maxlength="4000" placeholder="' + esc(K.ph) + '">' + esc(dval(id, "msg", r && r.msg)) + '</textarea></label>' +
      '<label class="fl"><span>Status</span><select data-draft="' + esc(id) + '" data-f="status">' + allSt(k).map(function(s){ return '<option' + (s === st ? ' selected' : '') + '>' + esc(s) + '</option>'; }).join("") + '</select></label>' +
      '<label class="fl"><span>Next step</span>' + inp(id, "next", r && r.next, 'maxlength="300" placeholder="e.g. Send quote, call back Thursday"') + '</label>' +
      '<label class="fl"><span>Who’s on it</span>' + inp(id, "owner", r && r.owner, 'list="rqOwners" maxlength="60" placeholder="Mendel"') + '</label>' +
      '<label class="fl"><span>Due</span>' + inp(id, "due", r && r.due, 'type="date"') + '</label>' +
      '<label class="fl"><span>' + (k === "purchasing" ? "PO #" : "SO / Invoice #") + '</span>' + inp(id, "ref", r && r.ref, 'maxlength="60" placeholder="' + (k === "purchasing" ? "PO 6700" : "SO-101704-D") + '"') + '</label>' +
      (k === "purchasing" ? '<label class="fl"><span>For customer / job</span>' + inp(id, "forCust", r && r.forCust, 'list="rqNames" maxlength="120"') + '</label>' : '') +
      '</div>';
  }
  function cardHtml(r){
    var k = r.kind, open = RS[k].open[r.id], late = isLate(r);
    var h = '<article class="rq' + (late ? ' late' : '') + (isClosed(k, r.status) ? ' closed' : '') + '">' +
      '<button type="button" class="rqhead" data-toggle="' + esc(r.id) + '" aria-expanded="' + !!open + '">' +
        '<span class="rqwho">' + esc(r.who) + (r.forCust ? ' <span class="muted">· for ' + esc(r.forCust) + '</span>' : '') + '</span>' +
        '<span class="chip ' + stCls(k, r.status) + '">' + esc(r.status) + '</span>' +
        (r.due ? '<span class="rqdue' + (late ? ' bad' : '') + '">' + (late ? 'Late · ' : 'Due ') + esc(fmtDue(r.due)) + '</span>' : '') +
        '<span class="rqmsg">' + esc(r.msg) + '</span>' +
        '<span class="rqmeta">' + (r.source === "email" ? '✉ from email · ' : '') + (r.next ? '<b>Next:</b> ' + esc(r.next) + ' · ' : '') + (r.owner ? esc(r.owner) + ' · ' : '') + (r.channel ? esc(r.channel) + ' · ' : '') + (r.ref ? esc(r.ref) + ' · ' : '') + 'updated ' + esc(ago(r.updatedAt)) + '</span>' +
      '</button>';
    if (r.emailLink || r.draftLink) h += '<div class="rqlinks">' + (r.draftLink ? '<a class="dlbtn sm" href="' + esc(r.draftLink) + '" target="_blank" rel="noopener">✎ Open draft reply (edit &amp; send in Outlook)</a>' : '') +
      (r.emailLink ? '<a class="btn2" href="' + esc(r.emailLink) + '" target="_blank" rel="noopener">Open email</a>' : '') + '</div>';
    if (open){
      h += '<div class="rqbody">';
      if (canWrite && rqDb){
        h += '<div class="upd"><label class="fl wide"><span>What happened / where it’s up to</span><textarea data-draft="' + esc(r.id) + '" data-f="note" maxlength="4000" placeholder="e.g. Sent quote, he’ll confirm Monday">' + esc(dval(r.id, "note", "")) + '</textarea></label>' +
          formFields(k, r.id, r) +
          '<div class="crow"><button type="button" class="dlbtn" data-saveupd="' + esc(r.id) + '">Save update</button><span class="dlmsg" id="um-' + esc(r.id) + '" role="status"></span></div></div>';
      } else {
        h += '<p class="rqfull">' + esc(r.msg) + '</p>' + (r.contact ? '<p class="muted">Contact: ' + esc(r.contact) + '</p>' : '');
      }
      h += '<h4 class="tlh">History</h4><ol class="tl">' + (r.log || []).slice().reverse().map(function(l){
        return '<li><span class="ndate">' + esc(fmtWhen(l.at)) + ' · ' + esc(who(l.by)) + (l.status ? ' · <span class="chip ' + stCls(k, l.status) + '">' + esc(l.status) + '</span>' : '') + '</span>' +
          (l.text ? '<span class="tltext">' + esc(l.text) + '</span>' : '') + (l.changes ? '<span class="tlch">' + esc(l.changes) + '</span>' : '') + '</li>'; }).join("") + '</ol></div>';
    }
    return h + '</article>';
  }
  function renderTracker(k){
    var box = $("trk-" + k); if (!box) return;
    var ae = document.activeElement;
    if (ae && box.contains(ae) && /INPUT|TEXTAREA|SELECT/.test(ae.tagName)){ rqPending = true; return; }
    var K = KINDS[k], S = RS[k], items = rqOf(k), q = S.q.trim().toLowerCase();
    var counts = {active:0, done:0, late:0, all:items.length}; allSt(k).forEach(function(s){ counts["s:" + s] = 0; });
    items.forEach(function(r){ counts["s:" + r.status] = (counts["s:" + r.status] || 0) + 1; if (isClosed(k, r.status)) counts.done++; else counts.active++; if (isLate(r)) counts.late++; });
    var shown = items.filter(function(r){
      if (S.pill === "active" && isClosed(k, r.status)) return false;
      if (S.pill === "done" && !isClosed(k, r.status)) return false;
      if (S.pill === "late" && !isLate(r)) return false;
      if (S.pill === "mark" && r.channel !== "From Mark") return false;
      if (S.pill.indexOf("s:") === 0 && r.status !== S.pill.slice(2)) return false;
      if (q && [r.who, r.msg, r.next, r.owner, r.ref, r.contact, r.forCust].join(" ").toLowerCase().indexOf(q) < 0) return false;
      return true; }).sort(rqSort);
    counts.mark = items.filter(function(r){ return r.channel === "From Mark" && !isClosed(k, r.status); }).length;
    var pills = [["active","Open"],["late","Late"],["mark","From Mark"]].concat(K.statuses.slice(0, -1).map(function(s){ return ["s:" + s, s]; })).concat([["done","Done"],["all","All"]]);
    var h = '<div class="vhead"><h2>' + esc(K.title) + '</h2><span class="count">' + counts.active + ' open' + (counts.late ? ' · <b class="bad">' + counts.late + ' late</b>' : '') + '</span></div>';
    if (rqState === "none") h += '<div class="card empty">Requests can’t be reached right now. Reload the page.</div>';
    h += '<div class="controls">' + (canWrite && rqDb ? '<button type="button" class="dlbtn" data-add="' + k + '">+ New ' + esc(K.one) + '</button>' : '') +
      '<input type="search" data-q="' + k + '" value="' + esc(S.q) + '" placeholder="Search name, item, SO #, who’s on it…" aria-label="Search">' +
      '<a class="btn2" href="#report">Status report</a></div>';
    if (S.adding){
      h += '<div class="card rqform"><h3>New ' + esc(K.one) + '</h3>' + formFields(k, "add-" + k, null) +
        '<div class="crow"><button type="button" class="dlbtn" data-saveadd="' + k + '">Save</button><button type="button" class="btn2" data-canceladd="' + k + '">Cancel</button><span class="dlmsg" id="am-' + k + '" role="status">' + esc(S.msg) + '</span></div></div>';
    }
    h += '<div class="pills" role="group" aria-label="Filter">' + pills.map(function(p){
      var n = p[0] === "active" ? counts.active : p[0] === "done" ? counts.done : p[0] === "late" ? counts.late : p[0] === "mark" ? counts.mark : p[0] === "all" ? counts.all : counts[p[0]] || 0;
      if ((p[0] === "late" || p[0] === "mark") && !n && S.pill !== p[0]) return "";
      return '<button type="button" class="pill" data-pill="' + esc(p[0]) + '" data-k="' + k + '" aria-pressed="' + (S.pill === p[0]) + '">' + esc(p[1]) + ' <span>' + n + '</span></button>'; }).join("") + '</div>';
    if (rqState === "loading") h += '<div class="card empty">Loading…</div>';
    else if (!shown.length) h += '<div class="card empty">' + (items.length ? 'Nothing here with this filter.' : 'No ' + esc(K.one) + 's yet. Tap “+ New ' + esc(K.one) + '” when a message comes in.') + '</div>';
    else h += '<div class="rqlist">' + shown.map(cardHtml).join("") + '</div>';
    box.innerHTML = h;
  }
  function renderReport(){
    var box = $("rpBody"); if (!box) return;
    var t = today(), wk = new Date(Date.now() - 7 * 864e5).toISOString();
    var h = '<div class="tiles">' + Object.keys(KINDS).map(function(k){
      var it = rqOf(k), open = it.filter(function(r){ return !isClosed(k, r.status); }), late = open.filter(isLate);
      return '<a class="tile' + (late.length ? ' red' : '') + '" href="#' + k + '"><div class="k">' + esc(k === "purchasing" ? "Purchasing" : KINDS[k].title.replace("Customers — ", "")) + ' · open</div><div class="v">' + open.length + '</div><div class="n">' + (late.length ? late.length + ' late · ' : '') +
        it.filter(function(r){ return isClosed(k, r.status) && r.updatedAt >= wk; }).length + ' finished this week</div></a>'; }).join("") + '</div>';
    Object.keys(KINDS).forEach(function(k){
      var K = KINDS[k], it = rqOf(k), open = it.filter(function(r){ return !isClosed(k, r.status); }).sort(rqSort);
      h += '<section class="blk"><div class="sec-head"><h3>' + esc(K.title) + '</h3><span class="count">' + allSt(k).map(function(s){ var n = it.filter(function(r){ return r.status === s; }).length; return n ? esc(s) + ' ' + n : ''; }).filter(Boolean).join(' · ') + '</span><span class="rule"></span></div>';
      h += open.length ? '<div class="tblwrap"><table><thead><tr><th>' + K.who + '</th><th>What</th><th>Status</th><th>Next step</th><th>Who</th><th>Due</th><th>Last update</th></tr></thead><tbody>' +
        open.map(function(r){ var last = (r.log || [])[ (r.log || []).length - 1 ] || {};
          return '<tr><td>' + esc(r.who) + '</td><td class="wrapc">' + esc(r.msg.length > 140 ? r.msg.slice(0, 137) + "…" : r.msg) + '</td><td><span class="chip ' + stCls(k, r.status) + '">' + esc(r.status) + '</span></td><td class="wrapc">' + esc(r.next) + '</td><td>' + esc(r.owner) + '</td><td' + (isLate(r) ? ' class="bad"' : '') + '>' + esc(fmtDue(r.due)) + '</td><td class="wrapc">' + esc(ago(r.updatedAt)) + (last.text ? ' — ' + esc(last.text) : '') + '</td></tr>'; }).join("") + '</tbody></table></div>'
        : '<div class="card empty">Nothing open.</div>';
      h += '</section>';
    });
    box.innerHTML = h;
  }
  function reportWorkbook(){
    var wb = XLSX.utils.book_new(), sum = [["Section","Status","Count"]];
    Object.keys(KINDS).forEach(function(k){
      var K = KINDS[k], it = rqOf(k).slice().sort(rqSort);
      allSt(k).forEach(function(s){ sum.push([K.title, s, it.filter(function(r){ return r.status === s; }).length]); });
      var aoa = [["Added", K.who, "Contact", "Came in by", "Request", "Status", "Next step", "Who’s on it", "Due", k === "purchasing" ? "PO #" : "SO / Invoice #"].concat(k === "purchasing" ? ["For customer / job"] : []).concat(["Last update", "Latest note", "Late?"])];
      it.forEach(function(r){ var notes = (r.log || []).filter(function(l){ return l.text; }); var last = notes[notes.length - 1];
        aoa.push([String(r.createdAt || "").slice(0, 10), r.who, r.contact, r.channel, r.msg, r.status, r.next, r.owner, r.due, r.ref].concat(k === "purchasing" ? [r.forCust] : []).concat([String(r.updatedAt || "").slice(0, 10), last ? last.text : "", isLate(r) ? "LATE" : ""])); });
      XLSX.utils.book_append_sheet(wb, sheet(aoa, [11,28,22,11,50,14,34,12,11,14].concat(k === "purchasing" ? [24] : []).concat([11,40,7])), k === "purchasing" ? "Purchasing" : k === "sales" ? "Sales" : "Service");
    });
    XLSX.utils.book_append_sheet(wb, sheet(sum, [24,18,8]), "Summary");
    wb.SheetNames.unshift(wb.SheetNames.pop());
    return XLSX.write(wb, {type:"array", bookType:"xlsx"});
  }
  function refreshTrackers(){
    if (KINDS[view]) renderTracker(view);
    if (view === "report") renderReport();
    if (view === "home") renderHome();
    var dl = $("rqNames"); if (dl){ var seen = {}, opts = [];
      (D ? D.customers.map(function(c){ return c.customer; }) : []).concat(RQ.map(function(r){ return r.who; }), RQ.map(function(r){ return r.forCust; }))
        .forEach(function(n){ if (n && !seen[n]){ seen[n] = 1; opts.push('<option value="' + esc(n) + '">'); } });
      dl.innerHTML = opts.join(""); }
  }
  function rqErr(e){ var c = e && e.code; if (c === "invalid_argument"){ canWrite = false; return "You don’t have permission to change this board."; }
    if (c === "unavailable" || c === "resource_exhausted") return "Couldn’t save just now. Try again in a moment."; return "Couldn’t save."; }
  function clean(s){ return String(s == null ? "" : s).trim(); }
  var FIELDS = ["who","contact","channel","msg","status","next","owner","due","ref","forCust"];
  var LABEL = {who:"Name", contact:"Contact", channel:"Came in by", msg:"Message", status:"Status", next:"Next step", owner:"Who’s on it", due:"Due", ref:"Ref #", forCust:"For"};
  function saveAdd(k){
    var id = "add-" + k, d = drafts[id] || {}, K = KINDS[k], msg = $("am-" + k);
    var doc = {kind: k}; FIELDS.forEach(function(f){ doc[f] = clean(d[f]); });
    if (!doc.status) doc.status = K.statuses[0];
    if (k !== "purchasing") delete doc.forCust;
    if (!doc.who || !doc.msg){ msg.className = "dlmsg warn"; msg.textContent = "Fill in the " + K.who.toLowerCase() + " and what they need."; return; }
    var now = new Date().toISOString();
    doc.createdAt = now; doc.createdBy = myId || null; doc.updatedAt = now; doc.updatedBy = myId || null;
    doc.log = [{at: now, by: myId || null, status: doc.status, text: "Added" + (doc.channel ? " · came in by " + doc.channel : ""), changes: ""}];
    var btn = document.querySelector('[data-saveadd="' + k + '"]'); if (btn) btn.disabled = true;
    msg.className = "dlmsg"; msg.textContent = "Saving…";
    rqDb.collection("requests").add(doc).then(function(){ delete drafts[id]; RS[k].adding = false; RS[k].msg = ""; if (document.activeElement) document.activeElement.blur(); renderTracker(k); })
      .catch(function(e){ msg.className = "dlmsg warn"; msg.textContent = rqErr(e); if (btn) btn.disabled = false; });
  }
  function saveUpd(id){
    var r = RQ.filter(function(x){ return x.id === id; })[0]; if (!r || !rqDb) return;
    var d = drafts[id] || {}, upd = {}, ch = [], msg = $("um-" + id);
    FIELDS.forEach(function(f){ if (d[f] == null) return; var v = clean(d[f]); if (v !== clean(r[f])){ if ((f === "who" || f === "msg") && !v) return; upd[f] = v;
      ch.push(f === "msg" ? "Message edited" : LABEL[f] + ": " + (clean(r[f]) || "—") + " → " + (v || "—")); } });
    var note = clean(d.note);
    if (!ch.length && !note){ msg.className = "dlmsg warn"; msg.textContent = "Nothing changed yet. Write what happened or change the status."; return; }
    var now = new Date().toISOString();
    upd.updatedAt = now; upd.updatedBy = myId || null;
    upd.log = (r.log || []).concat([{at: now, by: myId || null, status: upd.status || r.status, text: note, changes: ch.join("; ")}]);
    var btn = document.querySelector('[data-saveupd="' + id + '"]'); if (btn) btn.disabled = true;
    msg.className = "dlmsg"; msg.textContent = "Saving…";
    rqDb.doc("requests/" + id).update(upd).then(function(){ delete drafts[id]; if (document.activeElement) document.activeElement.blur(); renderTracker(r.kind); var m = $("um-" + id); if (m){ m.className = "dlmsg ok"; m.textContent = "Saved."; } })
      .catch(function(e){ msg.className = "dlmsg warn"; msg.textContent = rqErr(e); if (btn) btn.disabled = false; });
  }
  document.addEventListener("input", function(e){
    var t = e.target;
    if (t.getAttribute("data-draft")){ var id = t.getAttribute("data-draft"); (drafts[id] = drafts[id] || {})[t.getAttribute("data-f")] = t.value; return; }
    if (t.getAttribute("data-q")){ var k = t.getAttribute("data-q"); RS[k].q = t.value; rqPending = false;
      var pos = t.selectionStart; t.blur(); renderTracker(k); var n = document.querySelector('[data-q="' + k + '"]'); if (n){ n.focus(); try { n.setSelectionRange(pos, pos); } catch (x){} } }
  });
  document.addEventListener("change", function(e){ var t = e.target; if (t.getAttribute("data-draft")){ var id = t.getAttribute("data-draft"); (drafts[id] = drafts[id] || {})[t.getAttribute("data-f")] = t.value; } });
  document.addEventListener("focusout", function(){ setTimeout(function(){ if (rqPending){ var ae = document.activeElement; if (!ae || !/INPUT|TEXTAREA|SELECT/.test(ae.tagName)){ rqPending = false; refreshTrackers(); } } }, 150); });
  document.addEventListener("click", function(e){
    var t = e.target.closest ? e.target.closest("[data-add],[data-canceladd],[data-saveadd],[data-saveupd],[data-toggle],[data-pill],[data-chan]") : null; if (!t) return;
    var a;
    if ((a = t.getAttribute("data-add"))){ RS[a].adding = true; RS[a].msg = ""; renderTracker(a); var w = document.querySelector('[data-draft="add-' + a + '"][data-f="who"]'); if (w) w.focus(); return; }
    if ((a = t.getAttribute("data-canceladd"))){ RS[a].adding = false; delete drafts["add-" + a]; t.blur(); renderTracker(a); return; }
    if ((a = t.getAttribute("data-saveadd"))){ saveAdd(a); return; }
    if ((a = t.getAttribute("data-saveupd"))){ saveUpd(a); return; }
    if ((a = t.getAttribute("data-toggle"))){ var r = RQ.filter(function(x){ return x.id === a; })[0]; if (!r) return; RS[r.kind].open[a] = !RS[r.kind].open[a]; renderTracker(r.kind); return; }
    if ((a = t.getAttribute("data-pill"))){ RS[t.getAttribute("data-k")].pill = a; renderTracker(t.getAttribute("data-k")); return; }
    if ((a = t.getAttribute("data-chan"))){ var id = t.getAttribute("data-id"), d = (drafts[id] = drafts[id] || {}), on = t.getAttribute("aria-pressed") !== "true";
      d.channel = on ? a : ""; var sib = t.parentNode.querySelectorAll(".chan"); for (var i = 0; i < sib.length; i++) sib[i].setAttribute("aria-pressed", String(on && sib[i] === t)); }
  });
  function startRequests(db){
    rqDb = db;
    db.collection("requests").orderBy("updatedAt", "desc").limit(1000).onSnapshot(function(snap){
      RQ = snap.docs.map(function(d){ var b = d.data() || {}; var o = {id: d.id, kind: b.kind, log: Array.isArray(b.log) ? b.log : [], createdAt: b.createdAt, updatedAt: String(b.updatedAt || ""),
          emailLink: /^https:\/\/outlook\.office365\.com\//.test(b.emailLink || "") ? String(b.emailLink) : "", draftLink: /^https:\/\/outlook\.office365\.com\//.test(b.draftLink || "") ? String(b.draftLink) : "", source: b.source ? String(b.source) : ""};
        FIELDS.forEach(function(f){ o[f] = b[f] == null ? "" : String(b[f]); }); return o; }).filter(function(r){ return KINDS[r.kind] && r.who; });
      if (!snap.metadata.fromCache) rqState = "ready"; else if (rqState === "loading" && RQ.length) rqState = "ready";
      var ids = {}; RQ.forEach(function(r){ (r.log || []).forEach(function(l){ if (l.by && l.by !== myId && !names[l.by]) ids[l.by] = 1; }); });
      var list = Object.keys(ids);
      if (list.length && userApi && typeof userApi.profiles === "function") Promise.resolve(userApi.profiles(list)).then(function(ps){ list.forEach(function(id){ var p = ps && ps[id]; names[id] = (p && p.name) || "Someone"; }); refreshTrackers(); }).catch(function(){});
      refreshTrackers();
    }, function(){ rqState = "none"; refreshTrackers(); });
    setTimeout(function(){ if (rqState === "loading"){ rqState = "ready"; refreshTrackers(); } }, 8000);
  }

  /* ---------- follow-up email (Outlook draft via the viewer's Microsoft 365 connector) ---------- */
  var MS = "Microsoft 365", mcpApi = null, FU = null;
  var SIG = (window.__OB && window.__OB.settings && window.__OB.settings.signature) || "";
  function fmtD(d){ var p = String(d || "").split("-"); return p.length === 3 ? (+p[1]) + "/" + (+p[2]) + "/" + p[0] : String(d || ""); }
  function fuTemplate(c){
    var inv = D.invoices.filter(function(r){ return r.customer === c.customer && r.status !== "In Review"; }).sort(function(a,b){ return String(a.date).localeCompare(String(b.date)); });
    var t = inv.reduce(function(s,r){ return s + (r.bal || 0); }, 0), old = inv.reduce(function(m,r){ return Math.max(m, r.dpd || 0); }, 0);
    var lines = inv.map(function(r){ return "• " + r.inv + " — dated " + fmtD(r.date) + " — balance " + money(r.bal); }).join("\n");
    var body = "Hi,\n\nI hope you're doing well. I'm following up on the open balance for " + c.customer + ":\n\n" + lines +
      "\n\nTotal open: " + money(t) + "\n\n" +
      (old > 60 ? "Some of these are past due. Could you please let me know when we can expect payment?" : "Could you please let me know when we can expect payment?") +
      " You can pay through the secure link on the invoice, by ACH, or by check. If you need a copy of any invoice, just let me know.\n\nThank you\nMendel\n\n" + SIG;
    var to = c.email || "";
    if (!fuEmails(to).length){ var en = notesFor(c.customer).filter(function(n){ return n.source === "email" && /@/.test(n.from); })[0];
      if (en){ var m = en.from.match(/[^\s()<>]+@[^\s()<>]+\.[^\s()<>]+/); if (m) to = m[0]; } }
    return {cust: c.customer, to: to, subject: "Open invoices — " + c.customer + " — Handy Equip", body: body, n: inv.length, total: t,
            draftId: "", draftLink: "", sent: false, busy: false, msg: "", msgCls: ""};
  }
  function fuHtml(s){ return s.split(/\n{2,}/).map(function(p){ return "<p>" + esc(p).replace(/\n/g, "<br>") + "</p>"; }).join(""); }
  function fuEmails(s){ return String(s || "").split(/[;,\s]+/).filter(function(x){ return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x); }); }
  function mcpErrText(e, action){
    var c = e && e.code;
    if (c === "server_not_connected") return "Outlook isn’t connected to Office Board yet.";
    if (c === "needs_reauth") return "Outlook isn’t connected to Office Board yet.";
    if (c === "not_in_manifest" || c === "approval_required" || c === "consent_required") return "The board needs your OK to use Outlook. Press the button again and choose Allow.";
    if (c === "blocked_by_policy") return "Your organization blocks this Outlook action.";
    if (c === "selection_required") return "Choose which Microsoft 365 account to use in the prompt, then try again.";
    if (c === "tool_error") return "Outlook refused: " + String(e.message || "").slice(0, 200);
    if (c === "server_unavailable" || c === "upstream_error") return "Outlook didn’t answer. Check your Drafts folder before trying again" + (action === "send" ? " — it may have been sent." : " — it may already be there.");
    return "Outlook isn’t connected to Office Board yet. Copy the message into Outlook for now.";
  }
  function renderFollow(){
    var box = $("custFollow"); if (!box) return;
    var c = invState.cust ? custByName(invState.cust) : null;
    if (!c || !(c.open > 0)){ box.innerHTML = ""; return; }
    if (FU && FU.cust !== c.customer) FU = null;
    var ae = document.activeElement; if (ae && box.contains(ae) && /INPUT|TEXTAREA/.test(ae.tagName)) return;
    var h = '<div class="card fu">';
    if (!FU){
      h += '<div class="crow"><button type="button" class="dlbtn" data-fu="open">✉ Draft follow-up email</button>' +
        '<span class="muted" style="font-size:.8rem">Writes a reminder for the approved invoices. You review it, then save it to Outlook and send.</span></div>';
    } else {
      h += '<h3>Follow-up email — ' + esc(FU.cust) + '</h3>' +
        '<div class="fgrid"><label class="fl wide"><span>To</span><input data-fuf="to" value="' + esc(FU.to) + '" placeholder="customer@email.com"' + (FU.draftId ? ' disabled' : '') + '></label>' +
        '<label class="fl wide"><span>Subject</span><input data-fuf="subject" maxlength="255" value="' + esc(FU.subject) + '"' + (FU.draftId ? ' disabled' : '') + '></label>' +
        '<label class="fl wide"><span>Message</span><textarea data-fuf="body" style="min-height:260px"' + (FU.draftId ? ' disabled' : '') + '>' + esc(FU.body) + '</textarea></label></div>';
      if (FU.sent){
        h += '<div class="crow"><span class="chip green">Sent</span><span class="dlmsg ok">Sent from your Outlook. A note was added to this customer.</span><button type="button" class="btn2" data-fu="close">Close</button></div>';
      } else if (FU.draftId){
        h += '<div class="crow"><button type="button" class="dlbtn" data-fu="send"' + (FU.busy ? ' disabled' : '') + '>' + (FU.confirm ? 'Confirm: send to ' + esc(FU.to) : 'Send now') + '</button>' +
          (FU.draftLink ? '<a class="btn2" href="' + esc(FU.draftLink) + '" target="_blank" rel="noopener" style="text-decoration:none;padding:6px 11px">Open draft in Outlook to edit</a>' : '') +
          '<button type="button" class="btn2" data-fu="close">Close</button><span class="dlmsg ' + FU.msgCls + '" role="status">' + esc(FU.msg) + '</span></div>' +
          '<p class="muted" style="font-size:.78rem;margin-top:8px">Saved in your Outlook Drafts. “Send now” sends this exact draft. If you edit it in Outlook, send it from Outlook.</p>';
      } else {
        h += '<div class="crow"><button type="button" class="dlbtn" data-fu="save"' + (FU.busy ? ' disabled' : '') + '>Save to Outlook Drafts</button>' +
          '<button type="button" class="btn2" data-fu="reset">Start over</button><button type="button" class="btn2" data-fu="close">Cancel</button>' +
          '<span class="dlmsg ' + FU.msgCls + '" role="status">' + esc(FU.msg) + '</span></div>';
      }
    }
    box.innerHTML = h + '</div>';
  }
  function fuNote(text, link){
    if (!notesDb || !FU) return;
    notesDb.collection("callnotes").add({customer: FU.cust, text: text, at: new Date().toISOString(), by: myId || null, source: "followup", link: link || ""}).catch(function(){});
  }
  function fuSave(){
    if (!FU || FU.busy) return;
    var to = fuEmails(FU.to);
    if (!to.length){ FU.msg = "Add the customer’s email address first."; FU.msgCls = "warn"; renderFollow(); return; }
    if (!mcpApi){ FU.msg = "Outlook isn’t connected to Office Board yet. Copy the message into Outlook for now."; FU.msgCls = "warn"; renderFollow(); return; }
    FU.busy = true; FU.msg = "Saving to Outlook Drafts…"; FU.msgCls = ""; renderFollow();
    mcpApi.callTool(MS, "outlook_create_draft", {to: to, subject: FU.subject.slice(0, 255), body: fuHtml(FU.body), bodyType: "html"}, {cache: false}).then(function(res){
      var p = res && res.payload, txt = typeof p === "string" ? p : (res && res.content && res.content[0] && res.content[0].text) || "";
      var id = (p && typeof p === "object" && (p.id || p.messageId)) || ((txt.match(/\bid:\s*([A-Za-z0-9+\/=_-]+)/) || [])[1]) || "";
      var link = (p && typeof p === "object" && p.webLink) || ((txt.match(/webLink:\s*(https:\/\/\S+)/) || [])[1]) || "";
      FU.busy = false; FU.draftId = id; FU.draftLink = /^https:\/\/outlook\.office365\.com\//.test(link) ? link : "";
      FU.msg = id ? "Saved to Drafts. Review it, then press Send now." : "Saved to Drafts. Send it from Outlook."; FU.msgCls = "ok";
      if (!id) FU.draftId = "unknown";
      fuNote("Follow-up email drafted (" + FU.n + " invoice" + (FU.n === 1 ? "" : "s") + ", " + money(FU.total) + ") — waiting in Outlook Drafts.", FU.draftLink);
      renderFollow();
    }).catch(function(e){ FU.busy = false; FU.msg = mcpErrText(e, "save"); FU.msgCls = "warn"; renderFollow(); });
  }
  function fuSend(){
    if (!FU || FU.busy || !FU.draftId || FU.draftId === "unknown") return;
    if (!FU.confirm){ FU.confirm = true; FU.msg = ""; renderFollow(); return; }
    FU.confirm = false;
    FU.busy = true; FU.msg = "Sending…"; FU.msgCls = ""; renderFollow();
    mcpApi.callTool(MS, "outlook_send_draft", {messageId: FU.draftId}, {cache: false}).then(function(){
      FU.busy = false; FU.sent = true;
      fuNote("Follow-up email sent to " + FU.to + " (" + FU.n + " invoice" + (FU.n === 1 ? "" : "s") + ", " + money(FU.total) + ").", "");
      renderFollow();
    }).catch(function(e){ FU.busy = false; FU.msg = mcpErrText(e, "send"); FU.msgCls = "warn"; renderFollow(); });
  }
  document.addEventListener("input", function(e){ var f = e.target.getAttribute && e.target.getAttribute("data-fuf"); if (f && FU) FU[f] = e.target.value; });
  document.addEventListener("click", function(e){
    var t = e.target.closest ? e.target.closest("[data-fu]") : null; if (!t) return;
    var a = t.getAttribute("data-fu"), c = invState.cust ? custByName(invState.cust) : null;
    if (a === "open" && c && D){ FU = fuTemplate(c); renderFollow(); var ta = document.querySelector('[data-fuf="to"]'); if (ta && !FU.to) ta.focus(); }
    else if (a === "reset" && c){ FU = fuTemplate(c); t.blur(); renderFollow(); }
    else if (a === "close"){ FU = null; t.blur(); renderFollow(); }
    else if (a === "save"){ t.blur(); fuSave(); }
    else if (a === "send"){ t.blur(); fuSend(); }
  });

  /* ---------- email tab ---------- */
  var E = null;
  function olink(u, t){ return /^https:\/\//.test(u || "") ? '<a href="' + esc(u) + '" target="_blank" rel="noopener">' + (t || "Open email") + '</a>' : ""; }
  function shortDate(d){ return d ? String(d).slice(5, 10).replace("-", "/") : ""; }
  function renderHome(){
    if (rqDb){ var op = function(k){ return rqOf(k).filter(function(r){ return !isClosed(k, r.status); }); };
      var sa = op("sales"), se = op("service"), pu = op("purchasing");
      $("hmCu").textContent = sa.length + se.length; $("hmCuS").textContent = sa.length; $("hmCuV").textContent = se.length;
      $("hmCuL").textContent = sa.concat(se).filter(isLate).length;
      $("hmPu").textContent = pu.length; $("hmPuN").textContent = pu.filter(function(r){ return r.status === "Need to order" || r.status === "Quote requested"; }).length;
      $("hmPuW").textContent = pu.filter(function(r){ return r.status === "Ordered" || r.status === "Shipped"; }).length; $("hmPuL").textContent = pu.filter(isLate).length; }
    if (D){ $("hmAr").textContent = $("tTotal").textContent; $("hmAr90").textContent = $("t90").textContent; $("hmArH").textContent = $("tHershy").textContent; }
    if (E){ $("hmEm").textContent = E.waiting.length; $("hmEmO").textContent = E.offers.length; $("hmEmD").textContent = E.draftsTotal; $("hmEmP").textContent = E.payments.length; }
  }
  function renderEmail(){
    if (!E) return;
    $("eAsOf").textContent = "from your Outlook · checked to " + shortDate(E.asOf);
    var nItems = E.offers.reduce(function(s, o){ return s + o.items.length; }, 0);
    $("eOffSum").textContent = E.offers.length + " offers · " + nItems + " line items";
    $("eOffers").innerHTML = E.offers.map(function(o){
      return '<div class="card" style="margin-bottom:10px"><div style="display:flex;gap:8px;flex-wrap:wrap;align-items:baseline;justify-content:space-between">' +
        '<div><b>' + esc(o.company) + '</b> · ' + esc(o.from) + ' <span class="muted">· ' + esc(shortDate(o.date)) + '</span><div style="font-size:.86rem">' + esc(o.title) + '</div></div>' +
        '<div>' + olink(o.link) + '</div></div>' +
        (o.warn ? '<div style="margin:6px 0"><span class="chip amber">Check</span> <span style="font-size:.82rem">' + esc(o.warn) + '</span></div>' : '') +
        '<div class="tblwrap offer" style="margin-top:6px"><table><thead><tr><th>Location</th><th>Item</th><th class="num">Qty</th><th class="num">Price</th><th>Notes</th></tr></thead><tbody>' +
        o.items.map(function(i){ return '<tr><td>' + esc(i.location) + '</td><td>' + esc(i.item) + '</td><td class="num">' + esc(i.qty) + '</td><td class="num">' + (i.price === "" ? "" : money(i.price)) + '</td><td>' + esc(i.notes) + '</td></tr>'; }).join("") +
        '</tbody></table></div></div>';
    }).join("");
    var reps = NOTES.filter(function(n){ return n.source === "email"; });
    $("eRepSum").textContent = notesState === "ready" || reps.length ? reps.length + " replies · also saved in each customer’s notes" : "loading…";
    $("eRep").innerHTML = reps.length ? reps.map(function(n){
      return '<tr><td>' + esc(shortDate(n.at)) + '</td><td><a href="#" data-cust="' + esc(n.customer) + '">' + esc(n.customer) + '</a></td><td>' + esc(n.text) + '</td><td>' + olink(n.link) + '</td></tr>'; }).join("")
      : '<tr><td colspan="4" class="muted">No customer replies saved yet.</td></tr>';
    $("eWaitSum").textContent = E.waiting.length + " people";
    $("eWait").innerHTML = E.waiting.map(function(w){ return '<tr><td>' + esc(shortDate(w.date)) + '</td><td>' + esc(w.who) + '</td><td>' + esc(w.what) + '</td><td>' + olink(w.link) + '</td></tr>'; }).join("");
    $("eDrSum").textContent = E.drafts.length + " recent of " + E.draftsTotal + " in Drafts";
    $("eDr").innerHTML = E.drafts.map(function(d){ return '<tr><td>' + esc(shortDate(d.date)) + '</td><td>' + esc(d.to) + '</td><td>' + esc(d.subject) + '</td><td>' + esc(d.note) + '</td></tr>'; }).join("");
    var appr = E.payments.filter(function(p){ return p.status === "Approved" && p.amount > 0; }).reduce(function(s, p){ return s + p.amount; }, 0);
    $("ePaySum").textContent = E.payments.length + " items · " + money(appr) + " Sola-approved";
    $("ePay").innerHTML = E.payments.map(function(p){
      var cls = /Declined|Refund|Wrong/.test(p.status) ? "red" : p.status === "Approved" ? "green" : "amber";
      return '<tr><td>' + esc(shortDate(p.date)) + '</td><td>' + esc(p.customer) + '</td><td class="num">' + money(p.amount) + '</td><td>' + esc(p.how) + '</td><td><span class="chip ' + cls + '">' + esc(p.status) + '</span></td><td>' + esc(p.flag) + '</td><td>' + olink(p.link) + '</td></tr>'; }).join("");
  }
  document.addEventListener("click", function(e){
    var a = e.target.closest && e.target.closest("#eRep a[data-cust]");
    if (a){ e.preventDefault(); openCustomer(a.getAttribute("data-cust")); }
  });
  function offersWorkbook(){
    var aoa = [["Offer date","From","Company","Email","Location","Item","Qty","Price (each)","Notes","Check","Source email"]];
    E.offers.forEach(function(o){ o.items.forEach(function(i){ aoa.push([o.date, o.from, o.company, o.email, i.location, i.item, i.qty, i.price, i.notes, o.warn, o.link]); }); });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, sheet(aoa, [11,22,24,30,26,40,7,11,22,40,30]), "Rack offers");
    return XLSX.write(wb, {type:"array", bookType:"xlsx"});
  }

  function wireDownload(downloads, btnId, msgId, build, name){
    var btn = $(btnId), msg = $(msgId);
    btn.hidden = false;
    btn.addEventListener("click", function(){
      if (typeof XLSX === "undefined" || !D){ msg.className = "dlmsg warn"; msg.textContent = "Still loading — try again in a moment."; return; }
      var buf; try { buf = build(); } catch (e){ msg.className = "dlmsg warn"; msg.textContent = "Couldn’t build the file."; return; }
      btn.disabled = true; msg.className = "dlmsg"; msg.textContent = "Preparing…";
      downloads.save({filename: name(), data: buf}).then(function(){ msg.className = "dlmsg ok"; msg.textContent = "Saved."; })
        .catch(function(err){ var c = err && err.code; msg.className = "dlmsg" + (c === "declined" ? "" : " warn");
          msg.textContent = c === "declined" ? "Cancelled." : c === "rate_limited" ? "One moment, then try again." : "Download didn’t start. Try again."; })
        .then(function(){ btn.disabled = false; });
    });
  }

  /* ---------- call notes (saved in the board's database) ---------- */
  var NOTES = [], notesDb = null, notesState = "loading", myId = null, canWrite = true, userApi = null, names = {};
  var editing = null, notesPanelFor = null, focusComposer = false;
  function fmtWhen(iso){
    var d = new Date(iso); if (isNaN(d)) return String(iso || "");
    return d.toLocaleString("en-US", {weekday:"short", month:"short", day:"numeric", year:"numeric", hour:"numeric", minute:"2-digit"});
  }
  function notesFor(c){
    return NOTES.filter(function(n){ return n.customer === c; }).sort(function(a,b){ return a.at < b.at ? 1 : a.at > b.at ? -1 : 0; });
  }
  function who(id){ if (!id) return "Someone"; if (id === myId) return "You"; return names[id] || "Someone"; }
  function latestNoteCell(c){
    var ns = notesFor(c);
    if (!ns.length) return (notesDb && canWrite) ? '<button type="button" class="btn2" data-cust="' + esc(c) + '" data-focus="1">+ Add note</button>' : '';
    var n = ns[0], txt = (n.source === "email" ? "✉ " : "") + (n.text.length > 120 ? n.text.slice(0, 117) + "…" : n.text);
    return '<span class="ndate">' + esc(fmtWhen(n.at)) + (ns.length > 1 ? ' · ' + ns.length + ' notes' : '') + '</span><span class="nsnip">' + esc(txt) + '</span>';
  }
  function renderNotesPanel(){
    var box = $("custNotes");
    if (!invState.cust){ box.innerHTML = ""; notesPanelFor = null; editing = null; return; }
    if (notesPanelFor !== invState.cust){
      notesPanelFor = invState.cust; editing = null;
      box.innerHTML = '<div class="card notes"><h3>Call notes</h3>' +
        '<div class="composer" id="nComposer"><label for="nText" class="muted" style="font-size:.78rem">What did the customer say? Each note is saved with the date and time.</label>' +
        '<textarea id="nText" maxlength="4000" placeholder="e.g. Spoke to Kevin. Rest of the balance by check, mailed Friday."></textarea>' +
        '<div class="crow"><button type="button" class="dlbtn" id="nSave">Save note</button><span class="dlmsg" id="nMsg" role="status"></span></div></div>' +
        '<p class="dlmsg" id="nStatus" style="margin:8px 0 0"></p><ul class="nlist" id="nList"></ul></div>';
      $("nSave").addEventListener("click", saveNote);
    }
    var comp = $("nComposer");
    comp.hidden = !(notesDb && canWrite);
    var st = $("nStatus");
    st.className = "dlmsg";
    st.textContent = !notesDb ? (notesState === "none" ? "Notes can’t be reached right now. Reload the page." : "Loading notes…") :
      notesState === "slow" ? "Notes can’t be reached right now. They’ll appear when the connection returns." :
      notesState === "error" ? "Notes stopped updating. Reload the page to reconnect." :
      (canWrite ? "" : "You can read these notes but not add to them.");
    var ns = notesFor(invState.cust), ul = $("nList");
    var act = document.activeElement, focusedEdit = act && act.id === "nEditText";
    var caret = focusedEdit ? act.selectionStart : null;
    ul.innerHTML = ns.length ? ns.map(function(n){
      var hist = n.history || [], isEd = editing && editing.id === n.id;
      var src = n.source === "email"
        ? '<span class="chip slate">From email</span><span>' + esc(n.from) + '</span>' +
          (/^https:\/\//.test(n.link) ? '<a href="' + esc(n.link) + '" target="_blank" rel="noopener">Open email</a>' : '')
        : '<span>' + esc(who(n.by)) + '</span>';
      return '<li class="nitem"><div class="nmeta"><b>' + esc(fmtWhen(n.at)) + '</b>' + src +
        (n.editedAt ? '<span>· edited ' + esc(fmtWhen(n.editedAt)) + '</span>' : '') +
        (notesDb && canWrite && !isEd ? '<button type="button" class="btn2" data-edit="' + esc(n.id) + '">Edit</button>' : '') + '</div>' +
        (isEd ? '<div class="nedit"><textarea id="nEditText" maxlength="4000" aria-label="Edit note">' + esc(editing.draft) + '</textarea>' +
                '<div class="crow" style="margin-top:6px"><button type="button" class="dlbtn" data-editsave="' + esc(n.id) + '">Save change</button>' +
                '<button type="button" class="btn2" data-editcancel="1">Cancel</button></div></div>'
              : '<div class="ntext">' + esc(n.text) + '</div>') +
        (hist.length ? '<details class="nhist"><summary>Earlier wording' + (hist.length > 1 ? ' (' + hist.length + ')' : '') + '</summary>' +
          hist.slice().reverse().map(function(h){ return '<div><b>' + esc(fmtWhen(h.at)) + '</b>\n' + esc(h.text) + '</div>'; }).join("") + '</details>' : '') +
        '</li>';
    }).join("") : '<li class="nitem muted" style="font-size:.84rem">No call notes yet for this customer.</li>';
    var ta = $("nEditText");
    if (ta){
      ta.addEventListener("input", function(){ if (editing) editing.draft = this.value; });
      if (focusedEdit){ ta.focus(); try { ta.setSelectionRange(caret, caret); } catch (e) {} }
    }
    if (focusComposer && !comp.hidden){ focusComposer = false; $("nText").focus(); }
  }
  function noteErr(e){
    var c = e && e.code;
    if (c === "invalid_argument"){ canWrite = false; return "You don’t have permission to add notes on this board."; }
    if (c === "quota_exceeded") return "Couldn’t save the note. Try again.";
    if (c === "unavailable" || c === "resource_exhausted") return "Couldn’t save just now. Try again in a moment.";
    return "Couldn’t save the note.";
  }
  function saveNote(){
    var ta = $("nText"), msg = $("nMsg"), btn = $("nSave");
    var text = ta.value.trim();
    if (!text){ msg.className = "dlmsg warn"; msg.textContent = "Type what the customer said first."; return; }
    if (!notesDb) return;
    btn.disabled = true; msg.className = "dlmsg"; msg.textContent = "Saving…";
    notesDb.collection("callnotes").add({customer: invState.cust, text: text, at: new Date().toISOString(), by: myId || null})
      .then(function(){ ta.value = ""; msg.className = "dlmsg ok"; msg.textContent = "Saved."; })
      .catch(function(e){ msg.className = "dlmsg warn"; msg.textContent = noteErr(e); if (!canWrite) renderNotesPanel(); })
      .then(function(){ btn.disabled = false; });
  }
  function saveEdit(id){
    var n = NOTES.filter(function(x){ return x.id === id; })[0];
    if (!n || !editing || !notesDb) return;
    var text = editing.draft.trim();
    if (!text || text === n.text){ editing = null; renderNotesPanel(); return; }
    var hist = (n.history || []).concat([{text: n.text, at: n.editedAt || n.at}]);
    var btn = document.querySelector("[data-editsave]"); if (btn) btn.disabled = true;
    notesDb.doc("callnotes/" + id).update({text: text, editedAt: new Date().toISOString(), editedBy: myId || null, history: hist})
      .then(function(){ editing = null; renderNotesPanel(); })
      .catch(function(e){ var m = noteErr(e); if (btn) btn.disabled = false; var st = $("nStatus"); if (st){ st.className = "dlmsg warn"; st.textContent = m; } });
  }
  function refreshNotesUI(){ if (view === "customers") renderCustomers(); if (view === "invoices") renderNotesPanel(); if (view === "email") renderEmail(); }
  function resolveNames(){
    if (!userApi || typeof userApi.profiles !== "function") return;
    var ids = {}; NOTES.forEach(function(n){ if (n.by && n.by !== myId && !names[n.by]) ids[n.by] = 1; });
    var list = Object.keys(ids); if (!list.length) return;
    Promise.resolve(userApi.profiles(list)).then(function(ps){
      list.forEach(function(id){ var p = ps && ps[id]; names[id] = (p && p.name) || "Someone"; });
      refreshNotesUI();
    }).catch(function(){});
  }
  function startNotes(db){
    notesDb = db;
    var timer = setTimeout(function(){ if (notesState === "loading"){ notesState = "slow"; refreshNotesUI(); } }, 10000);
    db.collection("callnotes").orderBy("at", "desc").limit(1000).onSnapshot(function(snap){
      NOTES = snap.docs.map(function(d){
        var b = d.data() || {};
        return {id: d.id, customer: b.customer, text: String(b.text || ""), at: b.at, by: b.by, editedAt: b.editedAt,
                source: b.source, from: b.from ? String(b.from) : "", link: b.link ? String(b.link) : "",
                history: Array.isArray(b.history) ? b.history : []};
      }).filter(function(n){ return n.customer && n.at; });
      if (!snap.metadata.fromCache){ clearTimeout(timer); notesState = "ready"; }
      resolveNames();
      refreshNotesUI();
    }, function(){ clearTimeout(timer); notesState = "error"; refreshNotesUI(); });
    refreshNotesUI();
  }

  /* ---------- events ---------- */
  document.addEventListener("click", function(e){
    var cl0 = e.target.closest ? function(s){ return e.target.closest(s); } : function(){ return null; };
    var ed = cl0("[data-edit]");
    if (ed){ var nid = ed.getAttribute("data-edit"); var nn = NOTES.filter(function(x){ return x.id === nid; })[0];
      if (nn){ editing = {id: nid, draft: nn.text}; renderNotesPanel(); var ta = $("nEditText"); if (ta) ta.focus(); } return; }
    var es = cl0("[data-editsave]");
    if (es){ saveEdit(es.getAttribute("data-editsave")); return; }
    if (cl0("[data-editcancel]")){ editing = null; renderNotesPanel(); return; }
    var t = cl0("[data-cust]");
    if (t && t.getAttribute("data-cust")){ e.preventDefault(); if (t.getAttribute("data-focus")) focusComposer = true; openCustomer(t.getAttribute("data-cust")); return; }
    var ag = e.target.closest ? e.target.closest("[data-bucket]") : null;
    if (ag){ invState.cust = null; invState.bucket = ag.getAttribute("data-bucket"); invState.status = "Open"; $("invBucket").value = invState.bucket; $("invStatus").value = "Open"; location.hash = "invoices"; if (view === "invoices") renderInvoices(); return; }
    var th = e.target.closest ? e.target.closest("th[data-sort]") : null;
    if (th){
      var key = th.getAttribute("data-sort"), st = th.closest("#custTable") ? custState : invState;
      if (st.sort === key) st.dir = -st.dir; else { st.sort = key; st.dir = (key === "customer" || key === "inv" || key === "rep" || key === "status" || key === "notes" || key === "follow") ? 1 : -1; }
      if (st === custState) renderCustomers(); else renderInvoices();
    }
  });
  $("backCust").addEventListener("click", function(){ invState.cust = null; invState.status = "Open"; $("invStatus").value = "Open"; renderInvoices(); });
  $("invQ").addEventListener("input", function(){ invState.q = this.value; renderInvoices(); });
  $("invBucket").addEventListener("change", function(){ invState.bucket = this.value; renderInvoices(); });
  $("invStatus").addEventListener("change", function(){ invState.status = this.value; renderInvoices(); });
  $("custQ").addEventListener("input", function(){ custState.q = this.value; renderCustomers(); });

  /* ---------- boot ---------- */
  renderPromises();
  route();
  fetch("email.json").then(function(r){ return r.json(); }).then(function(e){ E = e; if (view === "email") renderEmail(); renderHome(); }).catch(function(){
    $("eOffers").innerHTML = '<div class="empty">Couldn’t load the email data. Reload the page.</div>'; });
  fetch("ar_data.json").then(function(r){ return r.json(); }).then(function(data){
    D = data;
    prepare();
    renderOverview(); renderPromises(); renderHershy(); renderHome(); refreshTrackers();
    if (view === "invoices") renderInvoices();
    if (view === "customers") renderCustomers();
    var dn = D.droppedNotes || [];
    if (dn.length){ $("clearedBlk").hidden = false; $("clearedSum").textContent = dn.length + " customers paid in full · kept so your notes aren’t lost";
      $("clearedTb").innerHTML = dn.map(function(r){ return '<tr><td>' + esc(r.customer) + '</td><td>' + esc(r.notes) + '</td><td>' + esc(r.follow) + '</td></tr>'; }).join(""); }
  }).catch(function(){
    $("aging").innerHTML = '<div class="empty">Couldn’t load the invoice data. Reload the page.</div>';
  });

  var cl = window.claude;
  if (cl && typeof cl.use === "function"){
    cl.use("downloads").then(function(downloads){
      if (!downloads) return;
      wireDownload(downloads, "hDl", "hMsg", hershyWorkbook, function(){ return "Handy-Equip-Hershy-fixlist-2026-11-23.xlsx"; });
      wireDownload(downloads, "rpDl", "rpMsg", reportWorkbook, function(){ return "Handy-Equip-status-report-" + today() + ".xlsx"; });
      wireDownload(downloads, "eDl", "eMsg", offersWorkbook, function(){ return "Handy-Equip-rack-offers-" + (E ? E.asOf : "") + ".xlsx"; });
      wireDownload(downloads, "invDl", "invMsg", invoicesWorkbook, function(){ return "Handy-Equip-invoices" + (invState.cust ? "-" + invState.cust.replace(/[^A-Za-z0-9]+/g,"-") : "") + "-2026-11-23.xlsx"; });
    }).catch(function(){});
    cl.use("mcp").then(function(m){ mcpApi = m; }).catch(function(){});
    Promise.all([cl.use("db"), cl.use("user")]).then(function(r){
      var db = r[0]; userApi = r[1];
      var idP = (userApi && typeof userApi.id === "function") ? Promise.resolve(userApi.id()) : Promise.resolve(null);
      var canP = (userApi && typeof userApi.can === "function") ? Promise.resolve(userApi.can("data.write")) : Promise.resolve(null);
      return Promise.all([idP, canP]).then(function(v){
        myId = v[0] || null;
        if (v[1] === false) canWrite = false;
        if (db){ startNotes(db); startRequests(db); } else { notesState = "none"; rqState = "none"; refreshNotesUI(); refreshTrackers(); }
      });
    }).catch(function(){ notesState = "none"; rqState = "none"; refreshNotesUI(); refreshTrackers(); });
  } else {
    notesState = "none"; rqState = "none";
  }
})();
