# 后续 Codex 开发约定

先阅读 README.md、docs/REQUIREMENTS.md、docs/HANDOFF.md 和 CHANGELOG.md。
这是独立 Python/Tkinter 项目，不依赖其他前端仓库。
不要修改、覆盖用户源账单或参考样例。真实账单不得提交到 Git。
分类规则必须同时检测老师类和学生类，冲突不可按优先级静默归类。
账单计算使用 Decimal，按行四舍五入到分。金额口径由用户选择。
核心逻辑在 bill_splitter/core.py；界面在 app.py；两者分离。
修改核心逻辑时运行 `.venv/Scripts/python.exe -m unittest discover -s tests -v`。
完成变更后更新 CHANGELOG.md 和 docs/HANDOFF.md，记录验证结果与限制。
除非明确要求，不添加联网、云端存储、登录或自动上传。
