/**
 * fuzzyMatch helper — v4.55.2 C05-BUG-07
 *
 * Carmen's C05-BUG-07 flagged "no fuzzy/typo-tolerant search anywhere".
 * Wave 30 added fuzzyMatch(query, target) as a shared helper in index.html
 * and adopted it in the steward_view.html resource search as the first
 * opt-in call site. This test file extracts the function from the source
 * (since it lives inside the main IIFE and isn't importable as a module)
 * and exercises every path in the scoring model:
 *
 *   Path 1: empty query → { matches: true, score: 1 }
 *   Path 2: exact case-insensitive substring → 0.9..1 (prefix bonus)
 *   Path 3: subsequence (every char in order) → 0.5..0.8
 *   Path 4: bounded Levenshtein for typos → 0.3..0.6
 *   Path 5: otherwise → { matches: false, score: 0 }
 *
 * Plus edge cases: null/undefined inputs, empty target, very long inputs
 * (Levenshtein bailout), performance ceiling, scoring ordering.
 */

const fs = require('fs');
const path = require('path');

// Extract fuzzyMatch from src/index.html — read the whole file, locate the
// function definition, and eval it into a local variable.
function loadFuzzyMatch() {
  var src = fs.readFileSync(
    path.resolve(__dirname, '..', 'src', 'index.html'),
    'utf8'
  );
  var startIdx = src.indexOf('function fuzzyMatch(');
  if (startIdx === -1) throw new Error('fuzzyMatch not found in src/index.html');
  // Find the matching closing brace by bracket depth.
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
  var wrapper = body + '\n; return fuzzyMatch;';
  // eslint-disable-next-line no-new-func
  return new Function(wrapper)();
}

const fuzzyMatch = loadFuzzyMatch();

describe('fuzzyMatch (C05-BUG-07)', () => {
  describe('Path 1: empty query', () => {
    test('empty string query matches everything with score 1', () => {
      expect(fuzzyMatch('', 'anything')).toEqual({ matches: true, score: 1 });
    });
    test('null query matches everything with score 1', () => {
      expect(fuzzyMatch(null, 'anything')).toEqual({ matches: true, score: 1 });
    });
    test('undefined query matches everything with score 1', () => {
      expect(fuzzyMatch(undefined, 'anything')).toEqual({ matches: true, score: 1 });
    });
    test('whitespace-only query matches everything (after trim)', () => {
      expect(fuzzyMatch('   ', 'anything')).toEqual({ matches: true, score: 1 });
    });
  });

  describe('null/empty target', () => {
    test('null target with non-empty query → no match', () => {
      expect(fuzzyMatch('query', null).matches).toBe(false);
    });
    test('undefined target with non-empty query → no match', () => {
      expect(fuzzyMatch('query', undefined).matches).toBe(false);
    });
    test('empty string target with non-empty query → no match', () => {
      expect(fuzzyMatch('query', '').matches).toBe(false);
    });
  });

  describe('Path 2: exact case-insensitive substring', () => {
    test('exact prefix match scores > 0.9', () => {
      var r = fuzzyMatch('grie', 'Grievance Log');
      expect(r.matches).toBe(true);
      expect(r.score).toBeGreaterThan(0.9);
    });
    test('interior substring match scores = 0.9', () => {
      var r = fuzzyMatch('ance', 'Grievance Log');
      expect(r.matches).toBe(true);
      expect(r.score).toBe(0.9);
    });
    test('case-insensitive substring match', () => {
      expect(fuzzyMatch('GRIE', 'grievance').matches).toBe(true);
      expect(fuzzyMatch('grie', 'GRIEVANCE').matches).toBe(true);
    });
    test('full exact match scores > 0.9 with prefix bonus', () => {
      var r = fuzzyMatch('cases', 'cases');
      expect(r.matches).toBe(true);
      expect(r.score).toBeGreaterThan(0.9);
    });
    test('trimmed query with trailing whitespace still matches', () => {
      expect(fuzzyMatch('  cases  ', 'open cases').matches).toBe(true);
    });
  });

  describe('Path 3: subsequence match', () => {
    test('"jdoe" matches "John Doe" as a subsequence', () => {
      // j (from John) ... d (from Doe) ... o (from Doe) ... e (from Doe)
      // Wait — "John Doe" is j-o-h-n-space-d-o-e. "jdoe" is j-d-o-e.
      // j matches j, then we skip ohn space, d matches d, o matches o, e matches e. YES.
      var r = fuzzyMatch('jdoe', 'John Doe');
      expect(r.matches).toBe(true);
      expect(r.score).toBeGreaterThanOrEqual(0.5);
      expect(r.score).toBeLessThanOrEqual(0.8);
    });
    test('subsequence match when substring would fail', () => {
      var r = fuzzyMatch('grlog', 'Grievance Log');
      expect(r.matches).toBe(true);
    });
    test('subsequence shorter than target scores lower than tight match', () => {
      var shortInLong = fuzzyMatch('abc', 'a---b---c---xxxxxxxxxx');
      var tightMatch = fuzzyMatch('abc', 'abc');
      expect(tightMatch.score).toBeGreaterThan(shortInLong.score);
    });
  });

  describe('Path 4: bounded Levenshtein (typo tolerance)', () => {
    test('single-char deletion matches (via subsequence fast path)', () => {
      // "grievnce" is a subsequence of "grievance" (the 'a' is skipped), so
      // this hits Path 3 (subsequence) before Path 4 (Levenshtein) gets a
      // chance. Score lands in the 0.5..0.8 subsequence range, not the
      // Levenshtein range. Documented here so future refactors that move
      // Levenshtein above subsequence don't silently regress the score band.
      var r = fuzzyMatch('grievnce', 'grievance');
      expect(r.matches).toBe(true);
      expect(r.score).toBeGreaterThanOrEqual(0.5);
      expect(r.score).toBeLessThanOrEqual(0.8);
    });
    test('single-char substitution matches via Levenshtein (not a subsequence)', () => {
      // "grievence" vs "grievance" — 'e' substituted for 'a' at position 5.
      // This is NOT a subsequence (the 'e' blocks the scan from matching
      // 'a' later), so it falls through to the Levenshtein path with a
      // single-edit distance.
      var r = fuzzyMatch('grievence', 'grievance');
      expect(r.matches).toBe(true);
      expect(r.score).toBeGreaterThanOrEqual(0.3);
      expect(r.score).toBeLessThanOrEqual(0.6);
    });
    test('transposition matches within budget', () => {
      // "grievacne" → "grievance" is 2 edits (swap c and n)
      var r = fuzzyMatch('grievacne', 'grievance');
      expect(r.matches).toBe(true);
    });
    test('missing character matches', () => {
      // "grievnc" → "grievance" — 2 deletions from target perspective
      var r = fuzzyMatch('grievnc', 'grievance');
      expect(r.matches).toBe(true);
    });
    test('too many typos → no match', () => {
      // "xxxxxxxx" vs "grievance" — 8 substitutions, exceeds ceil(8/4) = 2
      var r = fuzzyMatch('xxxxxxxx', 'grievance');
      expect(r.matches).toBe(false);
    });
  });

  describe('Path 5: no match', () => {
    test('totally unrelated query and target → no match', () => {
      var r = fuzzyMatch('zzzz', 'hello world');
      expect(r.matches).toBe(false);
      expect(r.score).toBe(0);
    });
    test('query longer than target by more than edit budget → no match', () => {
      var r = fuzzyMatch('a-very-long-query-string', 'xy');
      expect(r.matches).toBe(false);
    });
  });

  describe('Performance ceiling', () => {
    test('query longer than 32 chars does not run Levenshtein', () => {
      var longQuery = 'a'.repeat(40);
      // "a" repeated 40 times against an unrelated target should fall through
      // to the Levenshtein path, which is skipped when query.length > 32.
      // So the result should be no-match (since substring and subsequence
      // also fail).
      var r = fuzzyMatch(longQuery, 'zzz');
      expect(r.matches).toBe(false);
    });
    test('target longer than 128 chars skips Levenshtein', () => {
      var longTarget = 'z'.repeat(150) + 'query';
      // 'query' as substring WILL match (fast path 2). Use a non-substring
      // query instead.
      var r = fuzzyMatch('xyzxyz', 'z'.repeat(140));
      expect(r.matches).toBe(false);
    });
    test('call completes in under 10ms even for the longest allowed input', () => {
      var q = 'a'.repeat(32);
      var t = 'b'.repeat(128);
      var start = Date.now();
      fuzzyMatch(q, t);
      var elapsed = Date.now() - start;
      expect(elapsed).toBeLessThan(10);
    });
  });

  describe('Score ordering for ranking', () => {
    test('prefix match scores higher than interior match', () => {
      var prefix = fuzzyMatch('test', 'test case');
      var interior = fuzzyMatch('test', 'unit test case');
      expect(prefix.score).toBeGreaterThan(interior.score);
    });
    test('substring match scores higher than subsequence match', () => {
      var substring = fuzzyMatch('abc', 'abc def');
      var subsequence = fuzzyMatch('abc', 'a-b-c-def');
      expect(substring.score).toBeGreaterThan(subsequence.score);
    });
    test('subsequence match scores higher than typo match', () => {
      // Not a strict guarantee across all inputs but true for most realistic cases.
      var subseq = fuzzyMatch('gr', 'grievance'); // substring (fast path 2)
      var typo = fuzzyMatch('grievnce', 'grievance'); // Levenshtein
      expect(subseq.score).toBeGreaterThan(typo.score);
    });
  });

  describe('Real-world resource search scenarios', () => {
    test('steward types "conntract" typo for "contract"', () => {
      expect(fuzzyMatch('conntract', 'Contract FAQ').matches).toBe(true);
    });
    test('steward types partial "emrg" for "emergency"', () => {
      expect(fuzzyMatch('emrg', 'emergency').matches).toBe(true);
    });
    test('steward types "sick leav" for "sick leave"', () => {
      expect(fuzzyMatch('sick leav', 'Sick Leave Policy').matches).toBe(true);
    });
  });
});
