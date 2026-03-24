import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const source = fs.readFileSync(
  new URL("./api.ts", import.meta.url),
  "utf8",
);

test("refreshRpaTasks 默认使用 score api 8010 端口", () => {
  assert.match(
    source,
    /refreshRpaTasks\(scoreApiBase = "http:\/\/host\.docker\.internal:8010"\)/,
  );
});
