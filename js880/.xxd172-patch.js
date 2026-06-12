// XXD-172 direct patch: attach map2d / forEach2d onto JSA namespace.
// Issue: JSA.map2d / JSA.forEach2d are missing despite v4.1.0 changelog claim.
// The two functions exist as top-level definitions at L20080/L20098 but were
// never aliased onto JSA, so user code calls throw "JSA.map2d is not a function".
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const FILE = path.join(__dirname, 'JSA880.js');
const TMP = FILE + '.xxd172.tmp';
const MARKER = '/* XXD-172 final fix */';
const EXPECT_MAP = 'JSA.map2d = map2d;';
const EXPECT_FE = 'JSA.forEach2d = forEach2d;';

function patch() {
  let src = fs.readFileSync(FILE, 'utf8');
  if (src.indexOf(MARKER) !== -1 && src.indexOf(EXPECT_MAP) !== -1) {
    return { alreadyPatched: true, size: src.length };
  }
  const block = `
// ${MARKER}
// [XXD-172] attach top-level map2d / forEach2d onto JSA namespace
// Issue: R2.1 测试 JSA.map2d([[1,2],[3,4]], fn) THROW; v4.1.0 changelog 谎报已整合
// Fix: 函数体本身已存在(L20080/L20098), 这里补导出
${EXPECT_MAP}
${EXPECT_FE}
`;
  // Anchor: insert right after the forEach2d function body closes
  const anchor = '    for (var j = 0; j < arr[i].length; j++) {\n            fn(arr[i][j], i, j, arr);\n        }\n    }\n}\n';
  if (src.indexOf(anchor) === -1) {
    throw new Error('anchor for forEach2d close brace not found');
  }
  src = src.replace(anchor, anchor + block);
  fs.writeFileSync(TMP, src);
  fs.renameSync(TMP, FILE);
  return { patched: true, size: src.length };
}

function verify() {
  const src = fs.readFileSync(FILE, 'utf8');
  const ctx = { console, JSA: undefined };
  vm.createContext(ctx);
  vm.runInContext(src, ctx, { filename: 'JSA880.js' });
  const JSA = ctx.JSA;
  if (!JSA) throw new Error('JSA namespace not found after eval');
  if (typeof JSA.map2d !== 'function') throw new Error('JSA.map2d still missing');
  if (typeof JSA.forEach2d !== 'function') throw new Error('JSA.forEach2d still missing');
  const r1 = JSA.map2d([[1,2],[3,4]], function(x){ return x*2; });
  if (JSON.stringify(r1) !== '[[2,4],[6,8]]') throw new Error('map2d result wrong: ' + JSON.stringify(r1));
  const seen = [];
  JSA.forEach2d([[1,2],[3,4]], function(v, r, c){ seen.push([v, r, c]); });
  if (seen.length !== 4 || seen[0][0] !== 1 || seen[3][0] !== 4) throw new Error('forEach2d result wrong: ' + JSON.stringify(seen));
  return { map2d: r1, forEach2d: seen };
}

// Patch + verify, then re-check after 15s watch window (writer race)
const out = patch();
const v1 = verify();
console.log('[XXD-172] patch:', out);
console.log('[XXD-172] verify:', JSON.stringify(v1));

setTimeout(() => {
  try {
    const reread = fs.readFileSync(FILE, 'utf8');
    if (reread.indexOf(MARKER) === -1) {
      console.log('[XXD-172] marker gone after 15s, re-patching...');
      patch();
    }
    const v2 = verify();
    console.log('[XXD-172] post-watch verify OK:', JSON.stringify(v2));
  } catch (e) {
    console.log('[XXD-172] post-watch verify FAILED:', e.message);
    process.exit(2);
  }
}, 15000);
