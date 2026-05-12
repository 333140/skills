🧹 neat-freak（洁癖）
"每次任务做完要退出窗口的时候，如果不跑一遍 /neat，我就浑身难受，如坐针毡如芒刺背如鲠在喉。"

每次你在 Agent 里干完一件事，跑一下 /neat，它会把你这次会话改的东西，跟项目里的文档、CLAUDE.md / AGENTS.md、Agent 记忆全部对齐一遍，最后给你一份变更摘要。

为什么需要这个

你大概也遇到过：代码都迭代了七八轮，文档还是最初那一版；记忆里写着用 SQLite，其实你早换 PostgreSQL 了；CLAUDE.md 里的接口列表跟实际路由对不上。Agent 看着这些过期信息，越用越笨。

不是模型变笨，是文档和记忆脑腐了。neat-freak 就是清这个的。

它会动哪三层东西

项目根的 CLAUDE.md / AGENTS.md（给当前 AI 看的）
项目的 docs/ 和 README（给同事和其他人看的）
Agent 自己的记忆系统（给跨会话的自己看的）
这三层受众不同，职责不重叠，得分别处理。这也是我当时不满意 Claude Code 那个 AutoDream 的原因——它只动记忆，不动文档。

怎么触发

/neat            # 直接命令
整理一下          # 自然语言
同步一下          # 自然语言
sync up          # English
🌐 跨平台：Claude Code · Codex · OpenCode · OpenClaw
