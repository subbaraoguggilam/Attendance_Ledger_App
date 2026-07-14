// The original component was built for Claude.ai's artifact sandbox, which
// injects a `window.storage` key-value API (backed by Anthropic's servers).
// Outside claude.ai that API doesn't exist, so this file provides a
// drop-in browser localStorage replacement with the same shape:
//   get(key, shared?)    -> { key, value, shared } | null
//   set(key, value, shared?) -> { key, value, shared } | null
//   delete(key, shared?) -> { key, deleted, shared } | null
//   list(prefix?, shared?) -> { keys, prefix?, shared } | null
//
// NOTE: "shared" here just means a separate namespace in the SAME browser's
// localStorage — it is NOT actually shared across users/devices. For real
// cross-device/cross-user persistence (e.g. saving registers under an email
// so they show up on another computer), replace this with calls to your own
// backend/database. See README.md for pointers.

function nsKey(key, shared) {
  return `${shared ? "shared" : "personal"}:${key}`;
}

async function get(key, shared = false) {
  try {
    const raw = window.localStorage.getItem(nsKey(key, shared));
    if (raw === null) return null;
    return { key, value: raw, shared };
  } catch (e) {
    return null;
  }
}

async function set(key, value, shared = false) {
  try {
    window.localStorage.setItem(nsKey(key, shared), value);
    return { key, value, shared };
  } catch (e) {
    return null;
  }
}

async function del(key, shared = false) {
  try {
    window.localStorage.removeItem(nsKey(key, shared));
    return { key, deleted: true, shared };
  } catch (e) {
    return null;
  }
}

async function list(prefix = "", shared = false) {
  try {
    const nsPrefix = `${shared ? "shared" : "personal"}:${prefix}`;
    const keys = [];
    for (let i = 0; i < window.localStorage.length; i++) {
      const k = window.localStorage.key(i);
      if (k && k.startsWith(nsPrefix)) {
        keys.push(k.slice(nsPrefix.length - prefix.length));
      }
    }
    return { keys, prefix, shared };
  } catch (e) {
    return null;
  }
}

if (typeof window !== "undefined" && !window.storage) {
  window.storage = { get, set, delete: del, list };
}
