// Mirrors the main app's LinkifiedText regex (sans @mentions, which the
// add-in doesn't need) so the preview matches what users will see in
// Waypoint after submitting.

const TLDS = [
  "co.uk", "co.nz", "co.jp", "co.kr",
  "com.au", "com.br", "com.mx", "com.sg",
  "com", "org", "net", "io", "co", "ai", "app", "dev", "design",
  "gov", "edu", "info", "biz", "me", "tv", "xyz", "tech",
  "uk", "au", "us", "ca", "de", "fr", "jp", "in", "br",
  "mx", "eu", "nz", "ie", "nl", "se", "no", "fi", "dk",
  "es", "it", "ru", "cn", "kr", "sg", "hk", "za",
  "online", "store", "site", "shop", "news", "blog",
  "cloud", "global", "world", "agency", "studio",
];
const TLD = `(?:${TLDS.map((t) => t.replace(/\./g, "\\.")).join("|")})`;
const DOMAIN_BODY = `(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\\.)+${TLD}\\b(?:/[^\\s<>"'\`]*)?`;
const HTTP_URL = `https?:\\/\\/[^\\s<>"'\`]+`;

const LINK_RE = new RegExp(
  `\\[([^\\]]+)\\]\\((${HTTP_URL}|${DOMAIN_BODY})\\)` +
    `|(${HTTP_URL})` +
    `|(?<![\\w@/])(${DOMAIN_BODY})`,
  "g",
);

const TRAILING_PUNCT = /[.,;:!?)\]}'"]$/;

function normalizeHref(url: string): string {
  return /^https?:\/\//i.test(url) ? url : `https://${url}`;
}

function makeAnchor(href: string, text: string): HTMLAnchorElement {
  const a = document.createElement("a");
  a.href = normalizeHref(href);
  a.target = "_blank";
  a.rel = "noreferrer";
  a.textContent = text;
  return a;
}

// Extract just the links from `text` — used for a compact preview that
// confirms detected links without echoing the whole overview body.
export function extractLinks(text: string): { href: string; label: string }[] {
  if (!text) return [];
  const links: { href: string; label: string }[] = [];
  LINK_RE.lastIndex = 0;
  for (let m = LINK_RE.exec(text); m !== null; m = LINK_RE.exec(text)) {
    const [, mdLabel, mdUrl, httpUrl, domain] = m;
    if (mdUrl) {
      links.push({ href: normalizeHref(mdUrl), label: mdLabel });
    } else {
      let url = httpUrl || domain;
      while (url.length > 0 && TRAILING_PUNCT.test(url)) {
        url = url.slice(0, -1);
      }
      links.push({ href: normalizeHref(url), label: url });
    }
  }
  return links;
}

// Replace `target`'s contents with `text` rendered so URLs and `[label](url)`
// links appear as real <a> elements. Newlines are preserved (the target
// should have `white-space: pre-wrap`).
export function renderLinkified(text: string, target: HTMLElement): void {
  target.replaceChildren();
  if (!text) return;

  LINK_RE.lastIndex = 0;
  let lastIndex = 0;

  for (let m = LINK_RE.exec(text); m !== null; m = LINK_RE.exec(text)) {
    const [match, label, mdUrl, httpUrl, domain] = m;
    if (m.index > lastIndex) {
      target.appendChild(
        document.createTextNode(text.slice(lastIndex, m.index)),
      );
    }

    if (mdUrl) {
      target.appendChild(makeAnchor(mdUrl, label));
    } else {
      let url = httpUrl || domain;
      let trailing = "";
      while (url.length > 0 && TRAILING_PUNCT.test(url)) {
        trailing = url.slice(-1) + trailing;
        url = url.slice(0, -1);
      }
      target.appendChild(makeAnchor(url, url));
      if (trailing) target.appendChild(document.createTextNode(trailing));
    }

    lastIndex = m.index + match.length;
  }

  if (lastIndex < text.length) {
    target.appendChild(document.createTextNode(text.slice(lastIndex)));
  }
}
