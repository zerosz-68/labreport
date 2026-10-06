# 常见问题

## 安装与依赖

**Q：必须装 pandoc / LibreOffice 吗？**
不用。`pandoc` 完全不需要（公式转换是内置的）；LibreOffice 只影响 `pdf` 的排版目检，缺了会优雅降级（退出码 3），填写与自检照常。

**Q：`doctor` 返回 3 是什么意思？**
核心依赖（`python-docx` / `lxml` / `Pillow`）齐全，只是缺可选程序。功能降级但可用。返回 `1` 才是缺核心依赖，先 `pip install`。

**Q：能装在只有标准库的环境里吗？**
不能，`python-docx` 是硬依赖。但**不需要**安装本工具本身——`python run.py <命令>` 是零安装入口，只要求 `src/` 在本地。

---

## 中文与编码

**Q：在英文版 Windows 上跑，报 `UnicodeEncodeError: 'charmap' codec can't encode...`**
0.1.0 已修：`cli.py` 启动时会处理输出——**输出到管道/文件时切成 UTF-8，输出到终端时保留系统编码但把不可编码字符降级为 `?`**。
如果你是从旧版本升级上来的，用 `labreport --version` 确认，或临时设置 `set PYTHONUTF8=1`。

**Q：日志/重定向文件里的中文是乱码？**
管道输出是 UTF-8。用支持 UTF-8 的编辑器打开，或先 `chcp 65001`。

**Q：PDF 里中文变方块？**
系统缺中文字体。Linux/macOS 建议装 Noto Sans CJK 或 Source Han Sans；`doctor` 会给出提示。

---

## 表格与定位

**Q：`fill` 报"超出网格"怎么办？**
说明 `map.json` 里的行列超出了该表实际尺寸。先 `labreport inspect 模板.docx --max-text 20`，看清该表的 `行 × 列` 与合并情况再改。**注意 `row`/`col` 是屏幕上看到的物理行列（1 起），不是 `python-docx` 的 `cells[r][c]`。**

**Q：填进合并单元格会不会错位？**
不会。坐标走的是"物理网格"，程序自己把物理格换算到真实单元格（合并区域的锚点）。

**Q：文字锚点 `find` 找不到？**
确认文字**逐字一致**（含全角/半角、空格）；命中多个时用 `occurrence` 指定第几个；模板里嵌套表的文字也在搜索范围内。

**Q：能不能填表格以外的段落？**
目前只能定位到表格单元格。正文段落请在模板里预留一个单元格，或在模板阶段留好段落（后续版本考虑支持）。

---

## 公式

**Q：怎么确认公式是"Word 原生公式"而不是图片/文本？**
`inspect` 会统计 OMML 公式数（`公式 N 个(OMML)`）；`audit` 还会检查有没有**空公式**和**LaTeX 残留**——如果正文里还出现 `\frac`、`^{`，说明没转成原生公式。

**Q：支持哪些 LaTeX？**
`\frac` `\sqrt`（含 `\sqrt[n]{}`）`^` `_`（可同时，含上下标）`\left(...\right)` 普通括号（会生成 `m:d` 定界符）`\sum` `\int`（带上限）`\bar` `\vec` `\hat` `\dot` 希腊字母 常用运算符 `\mathrm` `\text`（正体） `\,` 等空白。
矩阵、多行对齐等冷门宏未实现，遇到会**退化并提示**，不会静默出错。

**Q：能和 pandoc 的输出混用吗？**
可以。`labreport omml formulas.md -o out --compare <pandoc片段目录>` 会逐条做结构比对（我们用它做过 28/28 的回归）。

---

## PDF 与目检

**Q：`pdf` 返回 3，没生成文件。**
没有可用的转换器。Windows 装 Word，或任意系统装 LibreOffice；也可以用 `--pages` 前先自己转 PDF。
另外：**首次调用 Word COM 需要放开工作区外的写入权限**（Word 要创建自己的用户配置目录），之后就绪。

**Q：`--pages` 报没有 `pdftoppm`。**
它是 MiKTeX/poppler 的一部分。没有时可以先看 PDF，或跳过目检——`audit` 已经覆盖了结构与数据层面的检查。

**Q：生成的 PNG 看不清。**
加 `--dpi 150`（默认 110）；只看某几页用 `--first` / `--last`。

---

## 数据校验

**Q：报"平均值不自洽"。**
你写的 `mean` 与按 `values` 算出来的不一致（容差由 `precision` 决定）。**回看 `source` 指的照片区域逐格重认**——通常是一位数字抄错。

**Q：报"离群（>3σ）"。**
可能是记错，也可能真的是异常数据。工具只负责指出，是否剔除由你决定（并在报告里说明）。

**Q：目标值与 `expected` 差很多，但公式看着没错。**
先查三件事：① 单位（`data.json` 里是不是按记录原样填的）；② 表达式里用的是 SI 值（`g→kg`、`mm→m` 已自动换算）；③ `expected` 的**量级**（转动惯量常按 $10^{-3}$ 或 $10^{-2}$ 记录，别写成绝对值）。

**Q：我测的是 50 个周期的总时间。**
在对应 `measurements` 里写 `"scale": 0.02`（=1/50），或直接把总时间除以 50 后填 `values`。前者可追溯性更好。

**Q：`--json` 输出能干什么？**
流水线判断用（含每项算出的数值）。例如 CI 里断言 `data check --json` 的 `error` 计数为 0。

---

## 填写与自检

**Q：为什么 `fill` 默认不写文件？**
防止误操作改坏排版。默认只打印"旧值 → 新值"的计划，人确认后加 `--apply`；全自动场景用 `--auto`。

**Q：`audit` 报 warn 要不要处理？**
error 必须处理（不可交付）；warn 需人工判断，常见是"某数据没在成品里找到"（漏填）或"过质心那行的 `———` 被当成占位"。确认无误可以带 warn 交付，但要在报告里说得清。

**Q：`audit` 的数值命中为什么不认我的写法？**
它认 `1.4091×10⁻³`、`1.4091e-3`、`0.0014091` 等等价写法（在容差内），但**不认单位换算后的值**：成品里写 `1.4091` 而 `data.json` 是 `1.4091e-3`，需要靠表头量级解释——这种情况请把 `--tolerance` 调大，或在 `data.json` 里按报告口径记录。

---

## 跨平台

| 环境 | 覆盖 | 结果 |
|---|---|---|
| GitHub Actions · ubuntu / macos / windows × Python 3.9 / 3.12 | 安装 + 15 个不需要真实 docx 的用例 + 命令冒烟 | 6/6 全绿 |
| Windows 11 + Python 3.12 + Word + 真实实验文件 | 全量 24 个用例（含 fill / audit / pdf→PNG 目检） | 24/24 通过 |

**Q：macOS/Linux 上能完全替代 Windows 的 Word 路径吗？**
填写、校验、自检完全一致（纯 Python 操作 OOXML）。差异只在 `pdf`：非 Windows 走 `soffice --headless`，**真实转换未在 CI 覆盖**（CI 只验证了"没有转换器时优雅降级为退出码 3"）。

**Q：支持 Python 3.8 吗？**
不支持，最低 3.9。3.9 在 CI 里是常态测试的。

---

## 隐私

**Q：我能把真实报告提交到仓库吗？**
不要。`.docx` / `.pdf` 常有姓名、学号、成绩；`.gitignore` 已默认拦掉 `*.docx` / `*.pdf`，但**提交历史里的东西删不掉**——一旦推上去，改写历史很麻烦。

**Q：提交里的邮箱怎么脱敏？**
用 GitHub 的免回复邮箱：

```bash
git config --global user.email "<ID>+<用户名>@users.noreply.github.com"
# 已有提交要一起改：
git rebase --root --exec "git commit --amend --no-edit --reset-author"   # 需要 shell
# 若环境不支持 --exec，可用 git commit-tree 逐个重建
git push --force-with-lease
```

注意：GitHub 网页端的 "Keep my email addresses private" **不影响本地 git**，必须本机改。

**Q：`data.json` 会泄露实验数据吗？**
它只含你填的数字与 `source` 文字描述（如"img3 表3.4 第1行"），不含照片本身。想更保险可以把数值换成等比缩放后的合成值。

---

## 测试与 CI

**Q：本地怎么跑？**

```bash
python tests/run_all.py                    # 不需要真实报告，16 个用例
python tests/run_all.py --docx-dir <目录> --frags-dir <目录>   # 全量 24 个
```

**Q：CI 里 Windows 作业失败，但本地是好的？**
看是不是编码问题（见上文"中文与编码"）或其他平台差异；CI 里 `PYTHONUTF8=1` 已设置。把这个当作真实用户场景来修，而不是只改 CI——历史上三次 CI 全红，最后一次暴露的正是**真实用户会遇到的 bug**。

**Q：怎么加自己的用例？**
在 `tests/fixtures/` 放脱敏数据，在 `tests/run_all.py` 里加一行 `case(名称, [命令...], 期望退出码, contains=[...])`。
