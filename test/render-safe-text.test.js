/**
 * renderSafeText helper — sub-project A Zone 1
 *
 * Tests the pure string→DocumentFragment helper that normalizes literal
 * '\n' (backslash+n) and real '\n' newlines into <br> nodes while HTML-
 * escaping everything else. Helper lives in src/index.html; we extract it
 * by bracket-depth the same way fuzzy-match.test.js does.
 */

const fs = require('fs');
const path = require('path');

function loadRenderSafeText() {
  var src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'),
    'utf8'
  );
  var startIdx = src.indexOf('function renderSafeText(');
  if (startIdx === -1) throw new Error('renderSafeText not found in src/index.html');
  var depth = 0;
  var inFn = false;
  var endIdx = startIdx;
  for (var i = startIdx; i < src.length; i++) {
    var ch = src.charAt(i);
    if (ch === '{') { depth++; inFn = true; continue; }
    if (ch === '}') {
      depth--;
      if (inFn && depth === 0) { endIdx = i + 1; break; }
    }
  }
  var body = src.substring(startIdx, endIdx);
  // The helper uses document.createDocumentFragment and document.createElement.
  // Tests run under node without jsdom, so stub a minimal document before eval.
  var stub = `
    var document = {
      createDocumentFragment: function() {
        var nodes = [];
        return {
          __nodes: nodes,
          appendChild: function(n) { nodes.push(n); return n; },
          get childNodes() { return nodes; }
        };
      },
      createElement: function(tag) {
        return { __tag: tag, nodeName: tag.toUpperCase(), childNodes: [] };
      },
      createTextNode: function(text) {
        return { __text: text, nodeName: '#text', nodeValue: text };
      }
    };
  `;
  var wrapper = stub + body + '\n; return renderSafeText;';
  return new Function(wrapper)();
}

const renderSafeText = loadRenderSafeText();

describe('renderSafeText (sub-project A Zone 1)', () => {
  test('empty string returns empty fragment', () => {
    const frag = renderSafeText('');
    expect(frag.childNodes.length).toBe(0);
  });

  test('null returns empty fragment', () => {
    const frag = renderSafeText(null);
    expect(frag.childNodes.length).toBe(0);
  });

  test('undefined returns empty fragment', () => {
    const frag = renderSafeText(undefined);
    expect(frag.childNodes.length).toBe(0);
  });

  test('plain text returns a single text node', () => {
    const frag = renderSafeText('hello world');
    expect(frag.childNodes.length).toBe(1);
    expect(frag.childNodes[0].nodeName).toBe('#text');
    expect(frag.childNodes[0].nodeValue).toBe('hello world');
  });

  test('literal backslash-n is split into two text nodes with a BR between', () => {
    const frag = renderSafeText('line one\\nline two');
    expect(frag.childNodes.length).toBe(3);
    expect(frag.childNodes[0].nodeValue).toBe('line one');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('line two');
  });

  test('real newline is split into two text nodes with a BR between', () => {
    const frag = renderSafeText('line one\nline two');
    expect(frag.childNodes.length).toBe(3);
    expect(frag.childNodes[0].nodeValue).toBe('line one');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('line two');
  });

  test('Windows CRLF is normalized to a single break', () => {
    const frag = renderSafeText('a\r\nb');
    expect(frag.childNodes.length).toBe(3);
    expect(frag.childNodes[0].nodeValue).toBe('a');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('b');
  });

  test('multi-line content produces alternating text/BR nodes', () => {
    const frag = renderSafeText('1. Was the employee warned?\\n2. Was the rule reasonable?\\n3. Was an investigation done?');
    // 3 text nodes + 2 BRs = 5 children
    expect(frag.childNodes.length).toBe(5);
    expect(frag.childNodes[0].nodeValue).toBe('1. Was the employee warned?');
    expect(frag.childNodes[1].nodeName).toBe('BR');
    expect(frag.childNodes[2].nodeValue).toBe('2. Was the rule reasonable?');
    expect(frag.childNodes[3].nodeName).toBe('BR');
    expect(frag.childNodes[4].nodeValue).toBe('3. Was an investigation done?');
  });

  test('script-tag content stays as escaped text (XSS defense)', () => {
    const frag = renderSafeText('<script>alert(1)</script>');
    expect(frag.childNodes.length).toBe(1);
    expect(frag.childNodes[0].nodeName).toBe('#text');
    expect(frag.childNodes[0].nodeValue).toBe('<script>alert(1)</script>');
  });

  test('numeric input is stringified, not rejected', () => {
    const frag = renderSafeText(42);
    expect(frag.childNodes.length).toBe(1);
    expect(frag.childNodes[0].nodeValue).toBe('42');
  });
});
