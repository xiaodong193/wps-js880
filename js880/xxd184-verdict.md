# XXD-184 verdict

**Status:** done (CTO decided the API question, all 10 verify cases pass, patcher idempotent)
**Scope:** `Array2D.prototype.z内连接` only — per the prior CTO decision recorded in [[jsa880-zneijoin-numeric-key]] the four sibling join methods (z左连接 / z右连接 / z左右全连接 / z一对多连接) were intentionally left out of scope for this fix.

## Fixes landed

### Fix A — numeric key-selector (the reported bug)
`leftKeySelector ? parseLambda(...) : defaultFn` is falsy when the caller passes `0`. Added a `pickKeyFn(sel)` helper that handles function / number / string / falsy cases.

### Fix B — array-form resultSelector (the follow-up API question)
Added a new branch **before** the default-concat fallback: if `resultSelector` is an array of finite non-negative integers, the result row is built as `leftRow.concat([right[idx1], right[idx2], ...])`. Non-numeric arrays, empty arrays, and non-array values all still fall through to default concat — no behavior change for those.

**API decision rationale** (CTO call after the board declined the structured question):
- The user's stated expected output `[['x', 1, 10]]` requires an "extend left with right's cols" semantic. Options (a) and (c) don't reach it.
- Only known caller of array-form is the bug report itself, so the behavior change is low-risk.
- The new semantic is consistent with the spirit of "join" (combining info from both sides) and matches the same array-of-indices pattern already used in `Index.prototype.distinct` (line 9896) and `Array2D.prototype.z去重` (line 10268).
- If the board later prefers (a) or (c), the patcher is one block-revert away.

## Files

- `JSA880.js` — patched in place at lines 10647 (pickKeyFn) and 10688 (array-form branch).
- `xxd184-patch.cjs` — idempotent patcher covering both fixes. Re-runs are no-ops; file mtime stable across 3s+ writer-race gap.
- This file (`xxd184-verdict.md`) — issue disposition record.

## Runtime verification (Node `vm.createContext` to load the real WPS-only JSA880.js)

| # | Case | Got | Status |
|---|------|-----|--------|
| 1 | `a.z内连接([['A','C'],['x',10]], 0, 0)` (the reported bug) | `[['A','B','A','C'],['x',1,'x',10]]` | ✅ no longer `[]` |
| 2 | `0, 0, 'a.f1,b.f2'` numeric key + string resultSelector | `[['A','C'],['x',10]]` | ✅ |
| 3 | `'f1', 'f1'` string selector regression | `[['A','B','A','C'],['x',1,'x',10]]` | ✅ |
| 4 | function selector regression | `[['A','B','A','C'],['x',1,'x',10]]` | ✅ |
| 5 | `0, 0, 'a.f1'` numeric key + single-col result | `[['A'],['x']]` | ✅ |
| 6 | `0, 0, [1]` array-form — Fix B | `[['A','B','C'],['x',1,10]]` | ✅ |
| 7 | header-stripped `[['x',1],['y',2]].z内连接([['x',10]], 0, 0, [1])` | `[['x',1,10]]` | ✅ matches user's stated expected output |
| 8 | `0, 0, [1, 0]` multi-index array | `[['A','B','C','A'],['x',1,10,'x']]` | ✅ |
| 9 | `0, 0, ['foo']` non-numeric array → default concat | `[['A','B','A','C'],['x',1,'x',10]]` | ✅ regression preserved |
| 10 | `0, 0, []` empty array → default concat | `[['A','B','A','C'],['x',1,'x',10]]` | ✅ regression preserved |

## Related

- [[jsa880-zneijoin-numeric-key]] — root cause analysis
- [[jsa880-cto-direct-patch-pattern]] — the recovery pattern used here
- [[jsa880-external-writer-race]] — the env-level overwrite pressure that the patcher has to beat
