# finance-sop-kickoff

一个 [Cursor Agent Skill](https://cursor.com)，通过多轮结构化访谈，帮财务/成本人员把脑子里的业务背景、处理流程、数据口径蒸馏成标准化的 SOP 知识包。**只做知识蒸馏，不写清洗脚本、不跑数据、不产出清洗结果**——产出的是 6 份 markdown 文档，供新人接手项目、也供未来和 AI 协作时作为知识底座。

本 skill 是通用方法论，不包含任何具体公司/项目的业务内容，可以用在任何"财务人员手工处理 Excel"的场景。

## 这是什么

Cursor 支持把一套可复用的工作方法封装成"Skill"——一个包含 `SKILL.md` 指令文件的文件夹，agent 读到之后会按里面写的步骤执行。这个仓库就是一个 Skill，装好之后你可以直接在 Cursor 里调用它，帮你或帮同事梳理某个财务/成本项目的 SOP。

访谈分三轮，产出 6 份文档：

| 轮次 | 产出 |
|------|------|
| 第1轮 | 体系地图 + 背景与目标 |
| 第2轮 | 环节与交付标准（agent 先给候选拆分草案，你只需确认） |
| 第3轮（每环节一次） | 业务口径说明 + 软知识与可信度台账 + 工具操作手册（基于你项目文件夹里的真实数据表生成分级提问清单） |

## 安装方法

Skill 要放在 Cursor 能找到的固定位置。有两种用法，选一种：

### 方式A：个人使用（推荐，装一次，所有项目都能用）

打开终端，克隆到你的个人 Skill 目录：

**Windows（PowerShell）：**

```powershell
git clone https://github.com/<your-org-or-username>/finance-sop-kickoff.git "$env:USERPROFILE\.cursor\skills\finance-sop-kickoff"
```

**macOS / Linux：**

```bash
git clone https://github.com/<your-org-or-username>/finance-sop-kickoff.git ~/.cursor/skills/finance-sop-kickoff
```

装好之后，不管你在 Cursor 里打开哪个项目文件夹，都能调用这个 skill。

### 方式B：不用 git，直接下载

1. 打开本仓库 GitHub 页面，点绿色的 `Code` 按钮 → `Download ZIP`
2. 解压后，把整个文件夹重命名为 `finance-sop-kickoff`
3. 放到 `C:\Users\<你的用户名>\.cursor\skills\`（Windows）或 `~/.cursor/skills/`（Mac/Linux）目录下

### 方式C：某个团队项目里所有人共用

如果想让某个具体项目里协作的所有人都能用这个 skill（不限于你自己），把它放进该项目仓库的 `.cursor/skills/finance-sop-kickoff/` 目录并提交到项目仓库里，其他人 pull 一下项目就自动有了，不需要单独装。

## 怎么调用

这个 skill 设置了 `disable-model-invocation: true`，意思是**必须显式点名调用，agent 不会自动触发**（因为它是一个会占用较长交互时间的多轮访谈，不适合看到文件夹里有Excel就自动开始问）。

1. 在 Cursor 里打开你想梳理 SOP 的项目文件夹，把相关的源数据表（Excel等）放进去
2. 对 agent 说类似：

   > "用 finance-sop-kickoff 这个 skill，帮我梳理一下这个项目的SOP"

   或者

   > "@finance-sop-kickoff 帮我开始访谈"

3. Agent 会先检查这个项目文件夹里有没有已经做过一部分的文档（断点续跑），没有就从第1轮开始问
4. 跟着问答走就行，访谈会分好几次进行，不用一次做完

## 产出物

访谈完成后，会在你的项目文件夹里生成（或更新）：

```
00-体系地图.md
01-背景与目标.md
02-环节与交付标准.md
03-业务口径说明.md（每个环节一份）
04-软知识与可信度台账.md
05-工具操作手册.md（每个环节一份）
```

这些文档不是清洗代码，是"知识包"——给新人接手用，也给未来做自动化时当验收标准用。

## 依赖

`scripts/scan_fields.py` 用来扫描 Excel 生成字段统计，需要 Python 3 环境，并安装：

```bash
pip install pandas openpyxl xlrd
```

（`openpyxl` 读 `.xlsx`，`xlrd` 读旧版 `.xls`）

## 目录结构

```
finance-sop-kickoff/
├── SKILL.md              # 主流程指令（agent 读这个执行）
├── reference.md          # 环节切分信号、字段优先级分级规则细节
├── templates/            # 6份产出文档的空白模板
│   ├── 00-体系地图.md
│   ├── 01-背景与目标.md
│   ├── 02-环节与交付标准.md
│   ├── 03-业务口径说明.md
│   ├── 04-软知识与可信度台账.md
│   └── 05-工具操作手册.md
└── scripts/
    └── scan_fields.py    # 扫描Excel，生成分级字段清单
```

## 更新

如果以后方法论迭代了，已经装过的人只需要：

```bash
cd ~/.cursor/skills/finance-sop-kickoff
git pull
```

## 参与改进

欢迎补充/修正方法论细节（信号规则、优先级判断、模板格式），但请注意：**本仓库不接受任何具体公司/项目的业务案例内容**（比如具体的字段名、口径规则、公司内部数据源），这个 skill 的价值就在于通用性——具体业务知识应该沉淀在各自项目自己的文档里，不要写进这个仓库。
