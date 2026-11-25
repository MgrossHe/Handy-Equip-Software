/* Office Board — sign-in gate and the storage adapter the board runs on.
 *
 * The board (board.js) was first built as a claude.ai artifact and talks to a small
 * `window.claude.use(...)` API for its database, the signed-in user and file downloads.
 * This file provides that same API on top of Supabase, so board.js runs unchanged.
 *
 * The publishable key below is meant to be public. Every table is protected by
 * row-level security: only signed-in users whose email is on the allow list can read
 * or write anything.
 */
(function () {
  "use strict";

  var SUPABASE_URL = "https://aylvakfpaehrzgwgyoub.supabase.co";
  var SUPABASE_KEY = "sb_publishable_zc41vUfNsAvUi9-xWzxXAw_CMQ4ZQ1c";
  var FILES = ["ar_data", "email", "promises", "settings"];

  var sb = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
    auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true, flowType: "implicit" }
  });

  function $(id) { return document.getElementById(id); }
  function show(id) {
    ["gate", "app"].forEach(function (x) { $(x).hidden = x !== id; });
  }
  function gateMsg(text, cls) { var m = $("gateMsg"); m.textContent = text || ""; m.className = "dlmsg" + (cls ? " " + cls : ""); }
  function gateView(name) {
    ["gForm", "gLink", "gDenied", "gLoading"].forEach(function (x) { $(x).hidden = x !== name; });
    show("gate");
    var focus = { gForm: "gEmail", gLink: "gLinkBtn" }[name];
    if (focus) setTimeout(function () { $(focus).focus(); }, 0);
  }

  /* ---------- sign-in ----------
   * Work email + the shared office password. The password is checked on the server
   * (Supabase function "shared-login"), which only accepts @handyequip.com emails,
   * locks out repeated wrong guesses, and returns a one-time token that is exchanged
   * here for a normal signed-in session.
   */
  $("gForm").addEventListener("submit", function (e) {
    e.preventDefault();
    var email = $("gEmail").value.trim().toLowerCase(), password = $("gPass").value;
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { gateMsg("Enter your work email.", "warn"); return; }
    if (!password) { gateMsg("Enter the password.", "warn"); return; }
    var btn = $("gBtn"); btn.disabled = true; gateMsg("Signing in…");
    fetch(SUPABASE_URL + "/functions/v1/shared-login", {
      method: "POST",
      headers: { "Content-Type": "application/json", apikey: SUPABASE_KEY },
      body: JSON.stringify({ email: email, password: password })
    }).then(function (res) {
      return res.json().catch(function () { return {}; }).then(function (body) { return { ok: res.ok, body: body }; });
    }).then(function (r) {
      if (!r.ok || !r.body.token_hash) throw new Error(r.body.error || "Couldn’t sign in. Try again.");
      return sb.auth.verifyOtp({ token_hash: r.body.token_hash, type: "magiclink" });
    }).then(function (r) {
      if (r.error) throw new Error("Couldn’t sign in. Try again.");
      $("gPass").value = ""; gateMsg("");
      if (r.data && r.data.session) start(r.data.session);
    }).catch(function (err) {
      gateMsg(err && /fetch|network/i.test(err.message) ? "Can’t reach the server. Check your connection." : (err.message || "Couldn’t sign in."), "warn");
    }).then(function () { btn.disabled = false; });
  });
  // Emailed link: ?token_hash=…&type=… — wait for a person to press the button.
  var qs = new URLSearchParams(location.search), linkHash = qs.get("token_hash"), linkType = qs.get("type") || "email";
  $("gLinkBtn").addEventListener("click", function () {
    var btn = this, msg = $("gLinkMsg");
    btn.disabled = true; msg.className = "dlmsg"; msg.textContent = "Signing in…";
    sb.auth.verifyOtp({ token_hash: linkHash, type: linkType }).then(function (r) {
      history.replaceState(null, "", location.pathname);
      if (r.error) {
        gateView("gForm");
        gateMsg("That sign-in link didn’t work or has expired. Send a new one.", "warn");
        return;
      }
      if (r.data && r.data.session) start(r.data.session);
    });
  });
  $("gOut").addEventListener("click", signOut);
  $("signOut").addEventListener("click", signOut);
  function signOut() { sb.auth.signOut().then(function () { location.hash = ""; location.reload(); }); }

  /* ---------- the window.claude adapter ---------- */
  function rid() {
    var a = new Uint8Array(15), s = "", c = "abcdefghijklmnopqrstuvwxyz0123456789";
    crypto.getRandomValues(a);
    for (var i = 0; i < a.length; i++) s += c[a[i] % c.length];
    return s;
  }
  function fail(error) { var e = new Error((error && error.message) || "failed"); e.code = "unavailable"; return e; }

  function makeDb(user) {
    var listeners = {};   // collection -> [refresh fns]
    function refreshAll(coll) { (listeners[coll] || []).forEach(function (f) { f(); }); }

    sb.channel("docs-live")
      .on("postgres_changes", { event: "*", schema: "public", table: "docs" }, function (p) {
        var coll = (p.new && p.new.collection) || (p.old && p.old.collection);
        if (coll) refreshAll(coll);
      })
      .subscribe();

    function query(coll, field, dir, lim) {
      return {
        limit: function (n) { return query(coll, field, dir, n); },
        onSnapshot: function (cb, errCb) {
          function load() {
            sb.from("docs").select("id,data").eq("collection", coll).limit(lim || 1000).then(function (r) {
              if (r.error) { if (errCb) errCb(fail(r.error)); return; }
              var rows = r.data.slice().sort(function (a, b) {
                var x = String((a.data || {})[field] || ""), y = String((b.data || {})[field] || "");
                return (x < y ? -1 : x > y ? 1 : 0) * (dir === "desc" ? -1 : 1);
              });
              cb({
                docs: rows.map(function (row) { return { id: row.id, data: function () { return row.data; } }; }),
                metadata: { fromCache: false }
              });
            });
          }
          (listeners[coll] = listeners[coll] || []).push(load);
          load();
          // Safety net in case the live connection drops.
          setInterval(load, 60000);
          return function () {};
        }
      };
    }

    return {
      collection: function (coll) {
        return {
          add: function (doc) {
            var id = rid();
            return sb.from("docs").insert({ collection: coll, id: id, data: doc }).then(function (r) {
              if (r.error) throw fail(r.error);
              refreshAll(coll);
              return { id: id };
            });
          },
          orderBy: function (field, dir) { return query(coll, field, dir, 1000); }
        };
      },
      doc: function (path) {
        var parts = path.split("/"), coll = parts[0], id = parts[1];
        return {
          update: function (fields) {
            return sb.from("docs").select("data").eq("collection", coll).eq("id", id).single().then(function (r) {
              if (r.error) throw fail(r.error);
              var merged = Object.assign({}, r.data.data, fields);
              return sb.from("docs").update({ data: merged, updated_at: new Date().toISOString() })
                .eq("collection", coll).eq("id", id).select("id");
            }).then(function (r) {
              if (r.error) throw fail(r.error);
              if (!r.data || !r.data.length) { var e = new Error("not saved"); e.code = "invalid_argument"; throw e; }
              refreshAll(coll);
            });
          }
        };
      }
    };
  }

  function makeUser(user) {
    return {
      id: function () { return user.id; },
      can: function () { return true; },
      profiles: function (ids) {
        return sb.from("profiles").select("id,name").in("id", ids).then(function (r) {
          var out = {};
          (r.data || []).forEach(function (p) { out[p.id] = { name: p.name }; });
          return out;
        });
      }
    };
  }

  var downloads = {
    save: function (o) {
      var blob = new Blob([o.data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = o.filename || "download.xlsx";
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
      return Promise.resolve();
    }
  };

  function installAdapter(user) {
    var db = makeDb(user), userApi = makeUser(user);
    window.claude = {
      use: function (name) {
        if (name === "db") return Promise.resolve(db);
        if (name === "user") return Promise.resolve(userApi);
        if (name === "downloads") return Promise.resolve(downloads);
        if (name === "mcp") return Promise.resolve(outlookApi);
        return Promise.resolve(null);
      }
    };
  }

  /* ---------- Outlook (Microsoft 365) through the ms-api server function ----------
   * Each person connects their own Outlook once ("Connect Outlook"). The board's
   * follow-up email buttons then save drafts to, and send from, that person's mailbox.
   */
  function msCall(body) {
    return sb.functions.invoke("ms-api", { body: body }).then(function (r) {
      if (!r.error) return r.data;
      var ctx = r.error.context;
      return (ctx && typeof ctx.json === "function" ? ctx.json().catch(function () { return {}; }) : Promise.resolve({}))
        .then(function (b) {
          var e = new Error((b.error && b.error.message) || "Outlook didn’t answer.");
          e.code = (b.error && b.error.code) || "upstream_error";
          // The board sign-in ended on the server (e.g. signed out elsewhere): sign in again.
          if (e.code === "unauthenticated") {
            sb.auth.signOut().then(function () { location.reload(); });
          }
          throw e;
        });
    });
  }
  // Same shape as the claude.ai connector the board was written for.
  var outlookApi = {
    callTool: function (server, tool, args) {
      if (tool === "outlook_create_draft") {
        return msCall({ action: "create_draft", to: args.to, subject: args.subject, body: args.body, bodyType: args.bodyType })
          .then(function (d) { return { payload: { id: d.id, webLink: d.webLink } }; });
      }
      if (tool === "outlook_send_draft") return msCall({ action: "send_draft", messageId: args.messageId }).then(function () { return { payload: {} }; });
      var e = new Error("Not available"); e.code = "not_in_manifest"; return Promise.reject(e);
    }
  };
  function olNotice(text) { $("olMsg").textContent = text; $("olBar").hidden = false; }
  function setupOutlook() {
    $("olClose").addEventListener("click", function () { $("olBar").hidden = true; });
    $("olBtn").addEventListener("click", function () {
      var b = this; b.disabled = true; b.textContent = "Opening Microsoft…";
      msCall({ action: "start" }).then(function (d) { location.href = d.url; })
        .catch(function (e) { b.disabled = false; b.textContent = "Connect Outlook"; olNotice(e.message); });
    });
    var qs2 = new URLSearchParams(location.search);
    if (qs2.get("outlook") === "connected") olNotice("Outlook is connected. Follow-up emails now save to your Outlook Drafts.");
    if (qs2.get("outlook_error")) olNotice(qs2.get("outlook_error"));
    if (qs2.has("outlook") || qs2.has("outlook_error")) history.replaceState(null, "", location.pathname + location.hash);
    msCall({ action: "status" }).then(function (d) {
      $("olBtn").hidden = !!d.connected;
      $("olState").hidden = !d.connected;
      if (d.ms_email) $("olState").title = d.ms_email;
    }).catch(function () { $("olBtn").hidden = false; });
  }

  /* ---------- "as of" times and new-data notice ---------- */
  function fmtStamp(iso) {
    var d = new Date(iso); if (isNaN(d)) return String(iso || "");
    var day = d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
    return /T\d/.test(String(iso)) ? day + ", " + d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" }) : day;
  }
  function showAsOf(ob) {
    if (ob.ar_data && ob.ar_data.asOf) $("asofSquid").textContent = fmtStamp(ob.ar_data.asOf);
    if (ob.email) $("asofEmail").textContent = fmtStamp(ob.email.checkedAt || ob.email.asOf);
  }
  // The scheduled email refresh rewrites board_files; tell people with the board open.
  function watchForNewData(rows) {
    var seen = {};
    rows.forEach(function (r) { seen[r.name] = r.updated_at; });
    $("freshBtn").addEventListener("click", function () { location.reload(); });
    setInterval(function () {
      sb.from("board_files").select("name,updated_at").in("name", FILES).then(function (r) {
        if (r.error || !r.data) return;
        var changed = r.data.some(function (x) { return seen[x.name] && x.updated_at !== seen[x.name]; });
        if (changed) $("freshBar").hidden = false;
      });
    }, 3 * 60 * 1000);
  }

  /* ---------- start the board once signed in and approved ---------- */
  var started = false;
  function start(session) {
    if (started) return;
    var user = session.user;
    gateView("gLoading");
    sb.rpc("is_board_member").then(function (r) {
      if (r.error || r.data !== true) {
        $("gDeniedEmail").textContent = user.email || "";
        gateView("gDenied");
        return;
      }
      started = true;
      var name = (user.email || "").split("@")[0];
      sb.from("profiles").upsert({ id: user.id, name: name.charAt(0).toUpperCase() + name.slice(1) }, { onConflict: "id", ignoreDuplicates: true }).then(function () {});
      return sb.from("board_files").select("name,data,updated_at").in("name", FILES).then(function (f) {
        if (f.error) { gateView("gForm"); gateMsg("Couldn’t load the board. Reload the page.", "warn"); started = false; return; }
        var ob = {};
        f.data.forEach(function (row) { ob[row.name] = row.data; });
        window.__OB = ob;
        // board.js loads its snapshot files by name; serve them from Supabase.
        var realFetch = window.fetch.bind(window);
        window.fetch = function (url, opts) {
          var m = typeof url === "string" && url.match(/^(ar_data|email)\.json$/);
          if (m) {
            var body = ob[m[1] === "ar_data" ? "ar_data" : "email"];
            return body ? Promise.resolve(new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } }))
                        : Promise.reject(new Error("missing " + m[1]));
          }
          return realFetch(url, opts);
        };
        installAdapter(user);
        $("whoami").textContent = user.email || "";
        showAsOf(ob, f.data);
        setupOutlook();
        watchForNewData(f.data);
        if (location.hash === "" && /access_token=/.test(location.href)) history.replaceState(null, "", location.pathname);
        show("app");
        var s = document.createElement("script");
        s.src = "board.js?v=2";
        document.body.appendChild(s);
      });
    });
  }

  sb.auth.onAuthStateChange(function (event, session) {
    if (session && session.user) start(session);
  });
  sb.auth.getSession().then(function (r) {
    var session = r.data && r.data.session;
    if (session) { if (linkHash) history.replaceState(null, "", location.pathname); start(session); }
    else if (linkHash) gateView("gLink");
    else if (!/access_token=|error_description=/.test(location.hash)) gateView("gForm");
    if (/error_description=/.test(location.hash)) {
      gateView("gForm");
      gateMsg("That sign-in link didn’t work or has expired. Send a new one.", "warn");
      history.replaceState(null, "", location.pathname);
    }
  });
})();
