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
    ["gForm", "gSent", "gLink", "gDenied", "gLoading"].forEach(function (x) { $(x).hidden = x !== name; });
    show("gate");
    var focus = { gForm: "gEmail", gSent: "gCode", gLink: "gLinkBtn" }[name];
    if (focus) setTimeout(function () { $(focus).focus(); }, 0);
  }
  function authErr(error, fallback) {
    var m = (error && error.message) || "";
    if (/rate limit/i.test(m)) return "Too many sign-in emails were sent in the last hour. Wait a while, then try again.";
    if (/expired|invalid/i.test(m)) return "That code didn’t work or has expired. Check it, or send a new one.";
    return m || fallback;
  }

  /* ---------- sign-in ----------
   * Sign-in uses a 6-digit code typed on this page, or an emailed link that opens this
   * page with a "Sign in" button. Either way the one-time token is only used when a
   * person acts, so company email link scanners (e.g. Microsoft Safe Links) that open
   * links ahead of the reader can't use it up first.
   */
  var sentTo = "";
  $("gForm").addEventListener("submit", function (e) {
    e.preventDefault();
    var email = $("gEmail").value.trim().toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { gateMsg("Enter your work email.", "warn"); return; }
    var btn = $("gBtn"); btn.disabled = true; gateMsg("Sending…");
    sb.auth.signInWithOtp({ email: email, options: { emailRedirectTo: location.origin + location.pathname } })
      .then(function (r) {
        btn.disabled = false;
        if (r.error) { gateMsg(authErr(r.error, "Couldn’t send the email. Try again."), "warn"); return; }
        gateMsg("");
        sentTo = email;
        $("gSentTo").textContent = email;
        $("gCode").value = ""; $("gCodeMsg").textContent = "";
        gateView("gSent");
      });
  });
  $("gSent").addEventListener("submit", function (e) {
    e.preventDefault();
    var code = $("gCode").value.replace(/\D/g, ""), msg = $("gCodeMsg"), btn = $("gCodeBtn");
    if (code.length < 6) { msg.className = "dlmsg warn"; msg.textContent = "Type the code from the email."; return; }
    btn.disabled = true; msg.className = "dlmsg"; msg.textContent = "Checking…";
    sb.auth.verifyOtp({ email: sentTo, token: code, type: "email" }).then(function (r) {
      btn.disabled = false;
      if (r.error) { msg.className = "dlmsg warn"; msg.textContent = authErr(r.error, "Couldn’t sign in. Try again."); return; }
      msg.textContent = "";
      if (r.data && r.data.session) start(r.data.session);
    });
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
  $("gBack").addEventListener("click", function () { gateView("gForm"); });
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
        // Outlook (Microsoft 365) is not connected to the website yet.
        return Promise.resolve(null);
      }
    };
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
      return sb.from("board_files").select("name,data").in("name", FILES).then(function (f) {
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
        if (location.hash === "" && /access_token=/.test(location.href)) history.replaceState(null, "", location.pathname);
        show("app");
        var s = document.createElement("script");
        s.src = "board.js?v=1";
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
