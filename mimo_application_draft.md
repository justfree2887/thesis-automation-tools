# 小米 MiMo 100T 开发者活动 —— 申请草稿

## 📌 表单信息

### 1️⃣ 你的邮箱
> 请填写你的邮箱，建议使用 GitHub 关联邮箱

**填入：** ____________（你的邮箱）

---

### 2️⃣ 你常使用的 AI 开发/Agent 工具
> 可选：OpenClaw, Claude Code, Codex, Hermes Agent, OpenCode, KiloCode, Cursor, Windsurf, Aider, Cline, 其他

**建议勾选：** Cursor ☑️  +  其他 ☑️  → 填写 "WorkBuddy (DeepSeek-V4-Flash)"

（你实际用的是 WorkBuddy + DeepSeek，核心是 AI Agent 辅助开发模式）

---

### 3️⃣ 目前主要使用的底层模型系列
> 可选：Claude 系列, Gemini 系列, GPT 系列, MiMo 系列, DeepSeek 系列, 豆包系列, MiniMax 系列, 其他

**建议勾选：** DeepSeek 系列 ☑️

---

### 4️⃣ 请描述你使用 Agent 或 AI 驱动构建的具体成果（≥100词）
> 需包含：① 项目解决的核心痛点  ② 核心逻辑流

---

#### 📝 草稿文案（中文版，约 340 字）

> **项目名称：Bio-Composite Thesis Automation Tools**
>
> **核心痛点：** 在生物质复合材料（KH570改性纳米SiO₂/竹粉/PLA）的学术研究中，实验数据处理、图表生成和论文格式修改占据了大量重复性劳动。传统工作流需要手动在 Origin/Excel 和 Word 之间反复切换，中英文混排字体规范（宋体+Times New Roman）极易出错，且生成的图表往往不符合黑白打印的投稿要求。
>
> **核心逻辑流：** 我使用 AI Agent（基于 DeepSeek V4 模型）进行长链推理辅助开发，构建了一套端到端的自动化工具链。系统采用三层流水线架构：(1) **数据层** —— 将力学测试、FTIR 光谱、吸湿率等多源实验数据归一化处理；(2) **可视化层** —— 利用 matplotlib 自动生成含误差棒的柱状图、应力-应变曲线、红外光谱图，支持图案填充/纹理区分，适配黑白打印；(3) **文档集成层** —— 通过 python-docx 操作 OOXML，自动刷新文档中的图表链接、修正中英文混排字体、修复页眉格式等。整个开发过程通过多轮 Agent 交互完成，从需求描述到代码实现、调试优化，形成了"需求→生成→验证→迭代"的闭环协作模式。

---

#### 📝 英文版（约 300 words）

> **Project:** Bio-Composite Thesis Automation Tools
>
> **Core Problem:** In academic research on bio-composite materials (KH570-modified nano-SiO₂/bamboo flour/PLA), researchers spend an enormous amount of time on repetitive tasks: manually processing experimental data across different aging conditions, generating publication-quality charts, and wrestling with Word document formatting. Mixed Chinese/English typesetting requirements (Song Ti for Chinese, Times New Roman for English/numbers) are notoriously error-prone, and chart output often fails to meet black-and-white print submission standards.
>
> **Core Logic Flow:** I used an AI Agent (powered by DeepSeek V4) with long-chain reasoning to build an end-to-end automation toolchain. The system follows a three-tier pipeline architecture:
>
> (1) **Data Layer** — Normalizes multi-source experimental data (mechanical testing, FTIR spectroscopy, moisture absorption) and groups them by aging condition (UV, thermal, hygrothermal).
>
> (2) **Visualization Layer** — Uses matplotlib to auto-generate bar charts with error bars, stress-strain curves, and FTIR spectra. All charts support pattern/texture fills for black-and-white print compatibility, with Chinese axis labels and proper font rendering.
>
> (3) **Document Integration Layer** — Leverages python-docx to manipulate OOXML directly: auto-refreshes chart images in the thesis Word document, fixes mixed Chinese/English font specifications (via styles.xml ascii font modification), repairs header formatting, and manages cross-references.
>
> The entire development process was completed through iterative Agent interactions — from requirement description to code generation, debugging, and optimization — forming a closed-loop "demand → generate → verify → iterate" collaboration pattern.

---

### 5️⃣ 使用证明与影响力证明

#### 可提交材料清单

| 材料类型 | 内容 | 状态 |
|---------|------|------|
| ✅ **GitHub 项目链接** | `https://github.com/shiberlin/thesis-automation-tools` | ⏳ 需创建仓库后上传 |
| ✅ **终端运行日志 / 录屏** | 展示脚本运行过程 + 生成的图表效果 | 可选 |
| ☑️ **AI 平台账单截图** | DeepSeek API / WorkBuddy 使用记录（如有） | 可选 |
| ✅ **成果截图** | 生成的 FTIR 光谱图、力学性能柱状图、论文截图 | 🟢 可直接提供 |

#### 推荐上传组合（最多5个）

1. **GitHub 仓库链接**（必填）
2. **生成的 FTIR 光谱图截图**（`gen_ftir_charts.py` 输出）
3. **力学性能柱状图截图**（`gen_v2_charts.py` 输出，展示误差棒+中文标签）
4. **终端运行日志截图**（展示脚本执行过程，如 `refresh_charts_in_doc.py`）
5. *（可选）* AI 平台使用记录截图

---

## 🚀 下一步操作

### Step 1: 创建 GitHub 仓库

1. 打开 https://github.com/new
2. 仓库名：`thesis-automation-tools`
3. 描述：`AI-assisted automation toolkit for bio-composite thesis research — chart generation, document processing, and reference management`
4. 设为 **Public**
5. 不要勾选任何初始化选项（README、.gitignore、License 都已经准备好了）
6. 点击 **Create repository**
7. 创建后接下面的推送命令

### Step 2: 推送到 GitHub

在终端中执行（替换 `YOUR_USERNAME`）：

```bash
cd /Users/shiberlin/Desktop/thesis-automation-tools
git init
git add .
git commit -m "Initial commit: bio-composite thesis automation tools"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/thesis-automation-tools.git
git push -u origin main
```

### Step 3: 填写 MiMo 申请

复制本文件中的第 4 题文案，粘贴到 MiMo 申请表单，上传 GitHub 链接和截图，提交。
