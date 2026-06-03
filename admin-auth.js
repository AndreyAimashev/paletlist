(function () {
  "use strict";

  var AUTH_STORAGE_KEY = "paletlist_auth_token";

  function redirectToLogin() {
    window.location.replace("index.html");
  }

  function stripLegacyCredentialsFromUrl() {
    try {
      var params = new URLSearchParams(window.location.search);
      if (!params.has("login") && !params.has("password")) {
        return;
      }
      params.delete("login");
      params.delete("password");
      var qs = params.toString();
      var next = window.location.pathname + (qs ? "?" + qs : "") + window.location.hash;
      window.history.replaceState({}, "", next);
    } catch (e) {
      /* ignore */
    }
  }

  function readBootstrapToken() {
    var token = null;
    try {
      token = sessionStorage.getItem(AUTH_STORAGE_KEY);
      sessionStorage.removeItem(AUTH_STORAGE_KEY);
    } catch (e) {
      token = null;
    }
    return token;
  }

  function validateTokenSync(token) {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", "/api/auth/me", false);
    xhr.setRequestHeader("Authorization", "Bearer " + token);
    try {
      xhr.send();
    } catch (e) {
      return null;
    }
    if (xhr.status !== 200) {
      return null;
    }
    try {
      return JSON.parse(xhr.responseText);
    } catch (e2) {
      return null;
    }
  }

  stripLegacyCredentialsFromUrl();

  var token = readBootstrapToken();
  if (!token) {
    redirectToLogin();
    return;
  }

  var user = validateTokenSync(token);
  if (!user) {
    redirectToLogin();
    return;
  }

  window.__PALETLIST_AUTH = {
    token: token,
    user_id: user.user_id != null ? Number(user.user_id) : 0,
    login: user.login || "",
    display_name: user.display_name || "",
    is_admin: !!user.is_admin,
  };

  var nativeFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {
    init = init || {};
    var url =
      typeof input === "string"
        ? input
        : input && input.url
          ? input.url
          : "";
    if (url.indexOf("/api/") === 0 && url.indexOf("/api/auth/login") !== 0) {
      var headers = new Headers(init.headers || {});
      if (!headers.has("Authorization") && window.__PALETLIST_AUTH && window.__PALETLIST_AUTH.token) {
        headers.set("Authorization", "Bearer " + window.__PALETLIST_AUTH.token);
      }
      init = Object.assign({}, init, { headers: headers });
    }
    return nativeFetch(input, init).then(function (res) {
      if (res.status === 401 && url.indexOf("/api/auth/me") !== 0) {
        redirectToLogin();
      }
      return res;
    });
  };
})();
