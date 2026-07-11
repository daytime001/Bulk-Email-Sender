## 任务目标

判断 Windows 用户在解析收件人 Excel 时提示 `Recipient file not found` 的原因。

## 当前状态

- 已完成：定位错误来自 Python 收件人加载器的文件存在性检查。
- 已完成：定位并修复 XLSX 解析缺少 `openpyxl` 的运行时依赖问题。
- 已完成：复查非技术用户首装路径，改为 Windows 直接下载国内镜像 Python 安装器 + 国内 PyPI 镜像安装依赖的一键配置路径。
- 已完成：定位并修复部分 Windows 电脑上中文路径经 worker 管道传输后被 Python 误解码，导致真实文件被判定不存在的问题。
- 已完成：定位并修复 JSON 导入时 worker 响应含中文导致 Rust 读取 stdout 报非 UTF-8 的问题。
- 已完成：新增收件人“研究方向”字段，支持 JSON/XLSX 解析与 `{research_direction}` 模板渲染。
- 进行中：无。
- 待开始：如需进一步发布，重新打包 Windows 客户端给用户验证。

## 已做决策

| 决策 | 理由 | 排除方案 |
| --- | --- | --- |
| 先按路径不可见问题分析，不按 Excel 内容解析问题分析 | 错误在 `Path.exists()` 处抛出，尚未进入 XLSX 解析 | 排查表头、单元格内容、openpyxl 解析格式 |
| 暂不修改代码 | 用户当前询问“是什么情况”，先给诊断结论 | 直接实现修复 |
| 运行时配置改为应用专用 venv | 只安装 Python 本体会导致 `openpyxl` 缺失；专用 venv 可避免污染用户系统 Python | 直接向用户系统 Python 全局安装依赖 |
| runtime 包构建阶段校验必需包 | 发布包缺依赖应在构建时失败，而不是到用户解析 XLSX 时失败 | 只在用户侧报错 |
| 设置页自动配置不再要求 manifest | 用户目标是点击按钮后程序自动用国内镜像装 Python 和依赖，不应要求发布者配置额外地址 | 让用户或发布者手动配置 `VITE_RUNTIME_MANIFEST_URL` |
| pip 安装增加官方源 + 国内镜像兜底 | 用户网络可能无法稳定访问默认 PyPI | 只使用默认 PyPI |
| Windows 用户路径不使用 uv | 用户不需要感知 uv；Windows 可直接下载官方 Python 安装器静默安装并创建应用专用 venv | 把 72MB 的 uv.exe 打进安装包；Windows 自动安装失败后再悄悄安装 uv |
| Windows 直接使用国内 Python 镜像 | 南京大学与阿里云镜像均可下载官方 Windows Python 安装器，更符合“一键自动安装 Python”的诉求 | 默认走 GitHub / uv |
| worker 进程强制 UTF-8 I/O | Rust/Tauri 通过 stdin 发送 UTF-8 JSON；Windows Python 默认管道编码可能跟随系统代码页，中文路径会被误解码 | 继续依赖 Python 默认编码 |
| loader 入口统一清洗路径文本 | 手动输入或平台返回值可能含外层引号、`file://` URI 或不可见字符 | 只在前端做 trim |
| worker 协议输出改为 ASCII-safe JSON | Rust 读取 worker stdout 要求 UTF-8；Windows Python 若按本地代码页输出中文，会产生非 UTF-8 字节 | 继续用 `ensure_ascii=False` 直接输出中文 |
| JSON 文件读取增加 GB18030 兜底 | 国内 Windows 用户可能用旧工具保存 ANSI/GBK 中文 JSON | 只支持 UTF-8 JSON |
| 研究方向占位符设为可选 | 老模板不应因为新增字段而突然无法发送；用户需要时可在正文或主题使用 `{research_direction}` | 把研究方向加入必填占位符 |

## 待解决问题

- [x] 问题：用户通过选择器选择的中文路径在部分 Windows 上仍被 Python 判定不存在。处理：worker 进程设置 `PYTHONIOENCODING=utf-8` 与 `PYTHONUTF8=1`，并在 loader 中清洗外层引号、不可见字符和 `file://` URI。
- [x] 问题：读取 JSON 时 worker 响应可能包含中文，部分 Windows Python 按本地编码写 stdout，Rust 按 UTF-8 读取失败。处理：worker 协议输出统一使用 ASCII-safe JSON 转义。
- [x] 问题：用户 JSON 文件可能是 GBK/GB18030 编码。处理：读取 JSON 时优先 `utf-8-sig`，再尝试 `gb18030`。
- [x] 问题：XLSX 需要新增“研究方向”列并可用于邮件模板。处理：收件人模型新增 `research_direction`，XLSX/JSON 解析保留该字段，发送模板新增 `{research_direction}` 变量，前端提示和预览表同步展示。
- [x] 问题：打包后的 Windows app 只安装 Python 解释器但未安装 XLSX 解析依赖。处理：配置 Python 时创建应用专用 venv 并安装 `openpyxl`。
- [x] 问题：弱网下载可能失败或卡住。处理：runtime 下载增加 HTTP 超时与重试，pip 安装增加重试、超时和多镜像兜底。
- [x] 问题：弱网中断可能留下半个安装包。处理：下载先写入 `.download` 临时文件，完整写完后再替换正式文件；Python 未安装成功时会重新下载安装器。
- [x] 问题：全新电脑没有 Python。处理：Windows 下直接从南京大学 / 阿里云 / 官方源下载 Python 3.11.9 安装器并静默安装到应用私有目录。
- [x] 问题：是否只需要 `openpyxl`。处理：检查 `pyproject.toml`、`uv.lock` 与业务代码 import，确认业务运行时唯一第三方依赖为 `openpyxl`，其传递依赖 `et-xmlfile` 由 pip 自动安装。

## 关键约束

- 只能基于截图与本地代码判断，暂未拿到用户机器日志。
- Windows 路径包含空格与中文，需考虑路径复制、同步盘、权限与输入法字符差异。
- XLSX 解析依赖 `openpyxl`，运行时就绪必须包含解释器和依赖两个层面。

## 验证与结果

- 已验证：`bulk_email_sender/recipients_loader.py` 在 `path.exists()` 为 false 时抛出截图中的英文错误。
- 已验证：前端选择文件后将路径原样传给 Tauri，Tauri 再原样传给 Python worker。
- 已验证：`No module named 'openpyxl'` 是运行时缺少 XLSX 解析依赖，不是文件内容问题。
- 已验证：`uv run pytest tests/test_runtime_packager.py tests/test_runtime_smoke.py -q` 通过，7 passed。
- 已验证：`cargo test` 通过，14 passed。
- 已验证：`npm run build` 通过。
- 已验证：`uv run pytest -q` 通过，46 passed。
- 已验证：`uv run ruff check bulk_email_sender tests` 通过。
- 已验证：`npm run lint` 通过。
- 已验证：南京大学 Python 镜像 `https://mirror.nju.edu.cn/python/3.11.9/python-3.11.9-amd64.exe` 可访问。
- 已验证：阿里云 Python 镜像 `https://mirrors.aliyun.com/python-release/windows/python-3.11.9-amd64.exe` 可访问。
- 已验证：使用清华 PyPI 镜像安装 `openpyxl>=3.1.5,<4` 成功，并可 `import openpyxl`。
- 已验证：设置页自动配置现在只调用后端自动配置命令；Windows 由后端直接下载 Python 安装器并安装依赖，不再自动安装 uv。
- 已验证：下载文件采用临时文件写入，避免中断后的半包被当成可用安装包。
- 已验证：新增测试覆盖 worker Python UTF-8 环境变量，避免 Windows 中文路径在 stdin/stdout 管道中被系统代码页误解码。
- 已验证：新增测试覆盖外层引号、BOM/零宽字符、`file://` URI 形式的收件人文件路径。
- 已验证：新增测试覆盖 worker 响应协议输出为纯 ASCII，中文姓名仍可被 JSON 正确还原。
- 已验证：新增测试覆盖 GB18030 编码 JSON 文件读取。
- 已验证：新增测试覆盖 JSON/XLSX 研究方向解析、worker 透传和邮件模板渲染。
- 结果：路径不可见和缺依赖是两个不同问题；缺依赖已修复。

## 风险与异常

- 风险：用户手动输入路径时可能含有错误斜杠、旧文件名或实际文件扩展名不同。
- 风险：文件在桌面但处于 OneDrive 未本地下载、快捷方式、微信临时目录、权限受限或文件被移动时，路径检查会失败。
- 异常：截图路径显示为 `C:\Users\Duan risheng\Desktop\发邮件测试.xlsx`，需要用户确认资源管理器中实际文件名和扩展名完全一致。
- 风险：首次自动配置运行环境需要联网安装 `openpyxl`，网络受限时会报“安装 Python 依赖失败”。
- 风险：南京大学 / 阿里云 Python 镜像若临时不可用，会回退官方 Python 源；仍需提示用户切换网络后重试。

## 关键变更记录

- 2026-07-11：创建元文件，记录 Windows 收件人路径不可见问题的初步诊断。
- 2026-07-11：修复 Python 运行时配置只安装解释器、不安装业务依赖的问题。
- 2026-07-11：按用户要求简化首装路径，Windows 不内置也不自动安装 uv，直接下载国内镜像 Python 安装器并安装 `openpyxl`。
- 2026-07-11：补充弱网下载保护，避免残留半包影响下一次自动配置。
- 2026-07-11：修复 worker 管道编码与路径文本归一化问题，避免中文路径真实存在但被误报找不到。
- 2026-07-11：修复 JSON 导入时 worker stdout 非 UTF-8 问题，并支持 GB18030 编码 JSON。
- 2026-07-11：新增研究方向字段与 `{research_direction}` 模板变量，并更新内置收件人示例。

## 参考资料

- `bulk_email_sender/recipients_loader.py`
- `bulk_email_sender/worker.py`
- `apps/desktop/src/App.tsx`
- `apps/desktop/src-tauri/src/lib.rs`
- `bulk_email_sender/runtime_packager.py`
- `bulk_email_sender/runtime_smoke.py`
- `tests/test_runtime_packager.py`
- `apps/desktop/src/App.tsx`
- `apps/desktop/src/features/settings/SettingsWorkspace.tsx`
- `apps/desktop/src-tauri/tauri.conf.json`

## 交付与复盘

- 交付物：诊断结论、运行时依赖修复、构建阶段依赖校验、Windows 中文路径误判修复、JSON/worker 编码修复、研究方向占位符功能、验证结果。
- 遗留问题：需要在 Windows 真机上安装打包产物，点击“自动配置运行环境”做一次端到端验收。
- 复盘结论：运行时“就绪”不能只检查 Python 版本，还必须检查业务依赖是否可导入；跨进程传中文路径时不能依赖 Windows Python 默认编码，worker 协议输出也应避免直接暴露本地编码字节。
