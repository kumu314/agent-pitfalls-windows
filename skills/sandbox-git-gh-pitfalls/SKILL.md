---
name: sandbox-git-gh-pitfalls
description: 沙箱内 git/gh 协作操作的已知坑与可靠写法。在沙箱里做 git 分支/commit/push、gh 开 PR、多代理协作推送时先读。覆盖：git ref 回滚（checkout -b 产生 unborn 分支、commit 后 ref 消失）、gh 未登录用 GCM 取 token、worktree 丢文件、代理 push 可靠写法、push OOM（http.postBuffer 500MB）、gh pr merge 静默失败、gh --body 反引号陷阱、**git push 彻底失效（exit 128 输出全空）时改走 Contents API**。触发词：Out of memory、malloc failed、postBuffer、unborn branch、ref 消失、gh auth、403、push 失败、push exit 128、TLS handshake、worktree 文件丢失、开 PR、merge 没生效、推送无输出。
---

# 沙箱 git / gh 协作坑（可靠写法）

> 跨项目通用。来源：2026-09 一次三人数模协作仓库的多轮 PR 实战。

## 1. git ref 回滚 bug（沙箱特有，最高频）
- **现象**：`git checkout -b` 后分支 unborn；`git commit` 成功但 `git branch` 看不到 ref；`git update-ref` 写入后调用一结束 ref 消失。**objects 和普通文件写入都还在**，只有 ref 层回滚。
- **对策**（实测可靠）：手动建 ref 文件：
  ```bash
  mkdir -p .git/refs/heads/agent/modeler
  printf '%s\n' "<SHA>" > .git/refs/heads/agent/modeler/<task>
  ```
  fetch 后恢复 origin 引用同理：`printf '%s\n' "$SHA" > .git/refs/remotes/origin/main`。

## 2. worktree 随机丢文件
- **现象**：`git status` 突然显示一堆 `D`（文件被删），实际没动过。
- **对策**：`git checkout -- .` 恢复；每次动手前先 `git status` 快照，异常立即恢复再继续。

## 3. gh CLI 未登录（hosts.yml 缺失）
- **现象**：`gh` 报未认证，但 git push 正常（token 在 Git Credential Manager）。
- **对策**：从 GCM 捞 token 注入环境变量：
  ```bash
  export GH_TOKEN=$(printf "protocol=https\nhost=github.com\n\n" | git credential fill | grep '^password=' | cut -d= -f2)
  ```
  （username 应为 `x-access-token`；scopes 含 repo 即可开 PR/合并。）

## 4. push 网络失败（本机有系统代理时）
- **可靠写法**：`git -c http.proxy=http://127.0.0.1:<端口> push`（实测一次成功；端口以注册表 `ProxyServer` 为准，别照抄别人的）。
- `HTTPS_PROXY=... git push` **有时不生效**（仍报 schannel handshake failed），别只依赖环境变量。
- gh 用 `HTTPS_PROXY=... HTTP_PROXY=... gh <cmd>` 可用。
- 偶发 `TLS handshake timeout` → 多试几次（本机网络间歇性，批量重试 5-10 次常过）。
- 注意：有些 agent 客户端内置的 GitHub 连接器是**只读**的，写操作（`create_or_update_file` 到 main 会 403）要改走 gh CLI。
- ⭐ **`fatal: Out of memory, malloc failed (tried to allocate 524288000 bytes)` 的根因是全局 `http.postbuffer` 被设成 500MB**，不是 pack 参数问题。
  调 `pack.windowMemory` / `pack.packSizeLimit` **完全无效**（2026-09-10 实测连失败 6 次，白耗轮次）。
  判定：`git config --list --show-origin | grep -i postbuffer`。
  正解（push 时覆盖，不必改全局配置）：
  ```bash
  git -c http.proxy=http://127.0.0.1:<端口> -c http.postBuffer=1048576 push -u origin <branch>
  ```
  **gh 命令同样中招**——`gh pr merge --delete-branch` 会触发本地 fetch/prune 而 OOM。用 git 的环境变量注入，gh 内层 git 也读得到：
  ```bash
  export GIT_CONFIG_COUNT=1 GIT_CONFIG_KEY_0=http.postBuffer GIT_CONFIG_VALUE_0=1048576
  ```
- **`gh pr merge` 会静默失败**：报 `Post "https://api.github.com/graphql": EOF`，或**退出码 0 但 PR 仍是 OPEN**（2026-09-10 撞过一次，以为合了其实没合）。
  → **合并动作后必须复核**：`gh pr view <n> --json state,mergedAt`，见 `"MERGED"` 才算数。别信退出码，也别信命令自己的输出。
  → `--delete-branch` 放在**确认 MERGED 之后**再补一次调用（合并前加它容易连带失败）。

## 5. gh pr 的 --body 反引号陷阱
- **现象**：`gh pr create --body "含 \`code\` 的文本"`，bash 把反引号当命令替换执行，正文内容被吃。
- **对策**：正文写临时文件 `--body-file`，或正文里不用反引号。

## 6. PR 流程纪律（多代理协作）
- 分支命名 `agent/<role>/<task>`；PR 描述里列出改动清单与验证结果。
- merge 前在本地核对 diff 范围 = 声称范围（防止夹带/漏带）。
- 合并后同步 main 到本地：fetch + 手动恢复 origin ref（见坑 1）再 rebase。

## 7. `exit=$?` 取到的是管道最后一个命令的退出码（踩 3 次！）
- **现象**：`python x.py | tail -3; echo $?` → `$?` 是 tail 的（永远 0），脚本失败也报成功。2026-09-10 轮询 CI 时第三次踩。
- **对策（动作前自检）**：要拿真实退出码就**不用管道**；必须用时取 `${PIPESTATUS[0]}`（bash）。
- 判定脚本成败一律以 exit code + 关键输出行双证据，别只看"有没有报错文案"。

## 8. Git Bash 的 /tmp，Windows Python 看不见
- **现象**：bash 里 `curl ... -o /tmp/a.json` 成功，Python 报 `FileNotFoundError: /tmp/a.json`；测试脚本两次撞上。
- **对策**：跨 bash/Python 传递的中间文件一律写 **Windows 绝对路径**（如 `D:/.../tmp_xxx.json`）。

## 9. origin/main 卡在 packed-refs 不更新（假同步，最危险）
- **现象**：`git fetch` 打印「`a..b main -> origin/main`」看似成功，但 `git rev-parse origin/main` 纹丝不动；`git update-ref` 也静默无效；偶尔伴随 `ambiguous argument 'origin/xxx'`。
- **根因**：ref 只存在于 `.git/packed-refs`，而对应的 loose 目录（`.git/refs/remotes/origin/`）不存在，写入静默失败。
- **危害**：此时 `git reset --hard origin/main` 会**把工作树静默拉回旧提交**——陈旧基线推上去 = 删队友产出。
- **判定**：`git show-ref | grep origin/main` 看真实存储值；`grep origin/main .git/packed-refs`；对比 fetch 输出。
- **修法**（坑 1 手法的完整版）：
  ```bash
  mkdir -p .git/refs/remotes/origin
  printf '%s\n' "<远程真实SHA>" > .git/refs/remotes/origin/main
  git fetch origin main && git rev-parse origin/main   # 验证跟手
  ```
- **规矩**：每次 fetch 后 `git rev-parse origin/main` 与远程对一次账，再动手 reset。

## 10. fetch 单个含斜杠分支名后，`origin/<name>` 不可用
- **现象**：`git fetch origin agent/writer/claim-role` 成功，但 `git log origin/agent/writer/claim-role` 报 ambiguous argument。
- **对策**：直接用 `FETCH_HEAD`：`git log FETCH_HEAD -3`、`git diff origin/main FETCH_HEAD`。

## 11. 手拼 SHA 错一位 → 422 "No commit found"
- **现象**：把 `18b321f1` 记成 `18b321f0` 去 API 查 CI，白查 8 轮。
- **对策（动作前自检）**：SHA **永远动态取**（`git rev-parse HEAD` / API 返回值），禁止手敲短 SHA 拼长。

## 12. 重试循环别把"空结果"当网络失败
- **现象**：`if [ -n "$OUT" ]` 判成功——仓库确实无 open PR 时输出为空，循环误报网络失败 6 次。
- **对策**：区分「网络失败」和「结果为空」：失败看 curl/gh 的 exit code 或 HTTP 码；空结果用 `--jq 'length'` 显式拿到 `0`。

## 13. `git push` 彻底失效（exit 128、stdout/stderr 全空）→ 改走 Contents API ★ 2026-09-13 实测

- **现象（与坑 4 不同层次的恶化）**：坑 4 的 push 至少会**打印报错**（handshake failed / OOM）。
  这里 push **exit 128 但一个字都不输出**：`git push origin main`、`git push -u origin <b>`、
  加 `-c http.proxy` — 全一样；`Start-Process` 重定向捕获也是空；代理重试 5 次同败。
  **注意：fetch 仍然正常。** 即"只读链路通、写链路死"。
- **诊断顺序**：先 `curl -s -m 12 -o /dev/null -w "%{http_code}" https://api.github.com/user`
  → 有 HTTP 码说明网络与凭证都没问题，**是 git 的 push 通道问题，别再改代理/重试 push**，
  改走 GitHub Contents API 逐文件提交（`PUT /repos/{owner}/{repo}/contents/{path}`，带 `sha` 做乐观并发），
  绕开 git 的写链路；只要 API 通就能把改动落进仓库。
- **止损线**：push 静默失败**重试超过 2 次就停**。继续重试纯烧轮次，且这次连续失败曾把
  一整轮时间耗光（同一会话同类调用 >3 次即停，见 `agent-exec-discipline`）。
- **不要做的**：不要为了"让 push 能跑"去 reset/rebase/换 remote 地址——**工作树本身没问题，
  动的越多风险越大**；也不要 `push --force`。

## 复用信号
出现"分支不见了""commit 凭空消失""文件莫名被删""gh 报未认证""push 一直握手失败"→ 先对照本 skill，按可靠写法重来，不盲目重试原始命令。
