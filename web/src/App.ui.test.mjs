import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const source = fs.readFileSync(
  new URL("./App.tsx", import.meta.url),
  "utf8",
);

test("RPA 参数配置区域提供刷新发送任务按钮", () => {
  assert.match(source, /刷新发送任务/);
});

test("刷新发送任务按钮绑定 onRefreshRpaTasks", () => {
  assert.match(
    source,
    /<button\s+type="button"\s+onClick=\{onRefreshRpaTasks\}\s+disabled=\{rpaLoading\}>[\s\S]*?刷新发送任务[\s\S]*?<\/button>/,
  );
});

test("RPA 页面提供日常操作区标题", () => {
  assert.match(source, /<h3>日常操作<\/h3>/);
});

test("RPA 参数区折叠进高级设置", () => {
  assert.match(source, /<details[\s\S]*?<summary>高级设置<\/summary>/);
});
