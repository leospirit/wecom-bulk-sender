# 企业微信批量图片发送

功能：
- 递归扫描本地目录图片
- 文件名开头连续中文作为学生姓名
- 自动匹配“学生名+妈妈”到通讯录 userId
- 批量发送 / 勾选发送
- 发送状态监控
- 自动监控目录（可开关）

## 目录结构

```
wecom-bulk-sender/
  api/
  web/
  data/
```

`data/` 会挂载到容器内的 `/data`。

## 快速开始（Windows/Mac）

1. 安装 Docker Desktop
2. 准备通讯录 Excel（企业微信导出）
3. 把图片放到 `data/inbox/`（可包含子文件夹）
4. 配置企业微信参数 `data/config.yaml`

```yaml
corp_id: "你的corpId"
agent_id: "你的agentId"
secret: "你的secret"
root_path: "/data/inbox"
rate_limit_per_sec: 1.0
max_concurrency: 2
```

5. 启动：

```bash
docker compose up --build
```

6. 打开前端：`http://localhost:5173`

## 使用流程

1. 上传通讯录 Excel（可选，如果已放到 `data/contacts.xlsx`）
2. 设置根目录（默认 `/data/inbox`）
3. 点击“扫描”
4. 选择：
   - “批量发送” → 发送全部待发送
   - 勾选后点“发送勾选” → 只发选中
5. 可打开“自动监控”自动监听新增文件

## 注意

- 文件名开头必须是学生姓名（中文）
- 只发送“妈妈”
- 未匹配到妈妈的会标记为跳过
