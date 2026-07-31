/**
 * XML well-formedness checks for generated parts.
 *
 * Excel does not report "your escaping is wrong" — it reports "we found
 * unreadable content" and offers to repair the file, discarding data. These
 * assertions catch the three ways this package can produce that outcome:
 * illegal characters, broken surrogate pairs, and unbalanced tags.
 */

// XML 1.0 forbids every C0 control character except tab, LF and CR. There is no
// escape sequence that makes them legal — even a &#0; character reference is
// invalid, so these have to be removed upstream rather than encoded.
const ILLEGAL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;

function findIllegalChar(xml) {
  const match = ILLEGAL_CHARS.exec(xml);
  if (!match) return null;
  return { index: match.index, code: match[0].codePointAt(0) };
}

function findLoneSurrogate(xml) {
  for (let i = 0; i < xml.length; i++) {
    const code = xml.charCodeAt(i);
    const isHigh = code >= 0xd800 && code <= 0xdbff;
    const isLow = code >= 0xdc00 && code <= 0xdfff;
    if (!isHigh && !isLow) continue;
    if (isLow) return { index: i, code };
    const next = xml.charCodeAt(i + 1);
    if (!(next >= 0xdc00 && next <= 0xdfff)) return { index: i, code };
    i++;
  }
  return null;
}

/**
 * Tag-balance check. Deliberately small: it understands the subset of XML this
 * package emits (declaration, elements, attributes, text) and nothing else, so
 * it stays readable and needs no dependency.
 */
function checkTagBalance(xml) {
  const stack = [];
  const tag = /<\?[^>]*\?>|<!--[\s\S]*?-->|<(\/)?([A-Za-z_][\w.:-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/)?>/g;
  let match;
  let consumed = 0;

  while ((match = tag.exec(xml)) !== null) {
    // Any '<' that did not start a construct we recognise is malformed.
    const gap = xml.slice(consumed, match.index);
    if (gap.includes("<")) return `unparseable markup near index ${consumed + gap.indexOf("<")}`;
    consumed = match.index + match[0].length;

    const [, closing, name, , selfClosing] = match;
    if (!name) continue; // declaration or comment
    if (closing) {
      const open = stack.pop();
      if (open !== name) return `</${name}> closes <${open ?? "nothing"}>`;
    } else if (!selfClosing) {
      stack.push(name);
    }
  }

  if (xml.slice(consumed).includes("<")) return "trailing unparseable markup";
  if (stack.length) return `unclosed element(s): ${stack.join(", ")}`;
  return null;
}

/**
 * @param {typeof import("node:assert")} assert passed in so failures point at the calling test
 * @param {string} xml
 * @param {string} label part name, used in the failure message
 */
function assertWellFormedXml(assert, xml, label) {
  const illegal = findIllegalChar(xml);
  assert.equal(
    illegal,
    null,
    illegal &&
      `${label}: illegal XML character U+${illegal.code.toString(16).padStart(4, "0").toUpperCase()} at index ${illegal.index} — Excel will refuse to open this file`,
  );

  const surrogate = findLoneSurrogate(xml);
  assert.equal(
    surrogate,
    null,
    surrogate &&
      `${label}: lone surrogate U+${surrogate.code.toString(16).toUpperCase()} at index ${surrogate.index}`,
  );

  const imbalance = checkTagBalance(xml);
  assert.equal(imbalance, null, imbalance && `${label}: ${imbalance}`);
}

module.exports = { assertWellFormedXml, findIllegalChar, findLoneSurrogate, checkTagBalance };
