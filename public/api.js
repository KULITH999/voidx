const API_URL = "https://script.google.com/macros/s/AKfycbyg-DT4YrB1mqI5QfNkeYDukkFP706ExyYosJuIcGzIQXumLQfXHMD4Xqf2kkbxhL2Q/exec";

function makeRunner(onSuccess, onFailure) {
  return new Proxy({}, {
    get(_, prop) {
      if (prop === "withSuccessHandler") return fn => makeRunner(fn, onFailure);
      if (prop === "withFailureHandler") return fn => makeRunner(onSuccess, fn);
      return (...args) => {
        fetch(API_URL, {
          method: "POST",
          // text/plain avoids a CORS preflight, which Apps Script can't answer
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ method: prop, args })
        })
          .then(r => r.json())
          .then(res => res.ok
            ? onSuccess && onSuccess(res.result)
            : onFailure && onFailure({ message: res.error }))
          .catch(err => onFailure && onFailure(err));
      };
    }
  });
}

window.google = { script: { run: makeRunner() } };
