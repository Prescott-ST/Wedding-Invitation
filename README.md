# 婚礼电子请柬项目 · 申屠俊杰 & 熊婧言

> 2026 年 10 月 6 日（星期二 · 农历丙午年八月廿六）晚 19:00
> 扬州西园大酒店 三楼国际厅（邗江区丰乐上街 1 号）

- **线上请柬**：<https://prescott-st.github.io/Wedding-Invitation/>
- **本仓库**：<https://github.com/Prescott-ST/Wedding-Invitation>
- **风格**：复古法式优雅（奶油底 / 墨绿 / 烫金 / ❦ 花饰 / 双层描金边框）

---

## 一、文件清单

| 文件 | 作用 |
|---|---|
| `index.html` | 请柬正式版，**全部内容、样式、交互都在这一个文件里** |
| `风格样图.html` | 设计期的四款风格样图存档（在线：<https://prescott-st.github.io/Wedding-Invitation/风格样图.html>） |
| `photos/` | 请柬用五张照片（网页压缩版，共约 590KB；原片在 `F:\婚礼影像\照片`） |
| `嘉宾回执统计.xlsx` | 宾客回执汇总表（姓名 / 出席人数 / 提交时间 / 合计） |
| `scripts/export_rsvp.py` | 从 QQ 邮箱拉取回执邮件 → 重新生成上面的 Excel；同一宾客多次提交**仅保留最新一笔**（按姓名判定） |
| `.nojekyll` | 跳过 Jekyll，Pages 构建必需 |
| `.github/workflows/rsvp-export.yml` | 每日定时任务：北京时间 08:05 自动汇总 Excel 并提交；也可在仓库 Actions 页手动 Run |

## 二、创作过程回顾

1. **初版**：中式红金风格，含开场封面、飘落花瓣、竖排姓名、倒计时、一键导航、时间轴、照片位、回执、背景音乐位——单文件自包含 HTML。
2. **风格比选**：做了四张样图（新中式雅致 / 水彩花卉 / 简约现代 / 复古法式），选定**复古法式优雅**后整体重做：配色、字体（思源宋体 + Playfair 花体）、❦ 角饰、双层边框、飘落金花（❦）。
3. **信息定制**：替换真实姓名、日期（含农历丙午年八月初二）、酒店与地址（凯悦官网核实为黄浦路 199 号）、导航关键词；婚礼流程精简为「18:30 签到 + 19:00 开席」。
4. **交互迭代**（均为用户反馈驱动）：
   - 回执表单：姓名 + 出席人数（1–12 可调），移除「出席/婉拒」切换（不提交即视为不出席）
   - 应要求移除全部联系方式
   - 板块顺序：喜帖 → 新人 → **幸福瞬间** → 时间地点 → 婚礼流程 → **倒计时** → 静候佳音
   - 日期胶囊强制单行，杜绝少数文字孤行
5. **部署**：GitHub Pages；首次构建失败（Jekyll 问题）→ 加 `.nojekyll` 解决。
6. **回执数据链路**：页面表单 → FormSubmit AJAX（邮件标题「婚礼回执 · 姓名 · N位」）→ QQ 邮箱；再由脚本/定时任务汇总为 Excel 回传仓库。
7. **回执统计增强**：FormSubmit 激活后全链路联调通过；定时任务上线（每日北京 08:05 自动汇总 Excel 并提交，也可在 Actions 页手动触发）；同一宾客重复提交时按姓名去重、仅保留最新一笔（重名宾客如需精确区分，可在表单增加手机号字段）。

## 三、新设备快速启动

1. **拿到项目**：本页绿色 **Code** 按钮 → **Download ZIP**（或 `git clone`）。只改请柬的话，`index.html` 单文件即可独立工作。
2. **本地预览**：双击 `index.html`；或在该目录跑 `python -m http.server 8641` 后访问 `http://localhost:8641`。
3. **让 AI 助手接手**：把本仓库地址发给它并说明要改什么即可——本 README 已包含全部关键信息；`index.html` 头部还有 ★ 自定义指南注释。
4. **推送更新**（三选一）：
   - 网页直接编辑：仓库页点 `index.html` → 铅笔图标 → 改完 Commit（最简单，无需任何本地工具）
   - GitHub CLI：`winget install GitHub.cli` 后 `gh auth login`
   - 上传文件：仓库页 **Add file → Upload files**
5. 推送后 Pages 约 1–3 分钟自动生效，**链接永远不变**。

## 四、关键配置信息

- **倒计时目标**：`index.html` 脚本内 `target = new Date('2026-10-06T19:00:00+08:00')`
- **回执接口**：`https://formsubmit.co/ajax/491042472@qq.com`（RSVP 配置区）
- **GitHub Secrets**（仓库 Settings → Secrets → Actions）：`RSVP_EMAIL`、`RSVP_MAILCODE`（QQ 邮箱 IMAP 授权码）。授权码值不在本文件中；如失效，到 QQ 邮箱 → 设置 → 账户 → 重新生成并更新 Secret
- **手动更新 Excel**：本机 Python 装好 `openpyxl` 后运行
  `set RSVP_EMAIL=... && set RSVP_MAILCODE=... && python scripts/export_rsvp.py`，再把生成的 `嘉宾回执统计.xlsx` 传回仓库
- **一键导航**：高德 `keywords=扬州西园大酒店`

## 五、待办清单

- [x] ~~FormSubmit 激活~~（2026-08-17 完成并联调验证：提交→邮箱→Excel→仓库全链路通过）
- [x] ~~创建定时任务~~（2026-08-17 完成并首次运行成功，rsvp-bot 已自动提交）
- [x] ~~照片 3 张~~（2026-08-22 完成：大图草坪撒花 + 红色电话亭 + 窗边逆光 + 楼梯暖光对视 + 黄昏江边相依，共 5 张已压缩入 `photos/`）
- [ ] 背景音乐（可选，`<audio id="bgm">` 加 `src`）
- [ ] 确认签到时间 18:30 是否准确
- [ ]（可选）真实回执进来后，让助手过滤掉 Excel 中的「联调测试」行

## 六、踩坑记录（新设备续作必读）

| 现象 | 原因与解法 |
|---|---|
| Pages 构建失败（秒败） | Jekyll 处理分支文件导致 → 根目录加 `.nojekyll`（已加） |
| API 上传 `.github/workflows/` 文件报 404 | 令牌缺 `workflow` 权限 → 网页创建，或用含 repo+workflow 权限的 PAT |
| GitHub 设备码授权一直转圈/失败 | 国内网络连不上 github 授权端点（API 正常）→ 用 PAT 或网页操作 |
| QQ 邮箱 IMAP 按 UTF-8 主题搜索返回无关邮件 | QQ 实现不可靠 → 脚本改用「日期过滤 + 逐封核对主题头」 |
| 邮件主题含 `unknown-8bit` 编码报错 | 解码需按 utf-8 / gb18030 逐级回退（脚本已处理） |
| 改了代码线上没变 | Pages 构建约 1–3 分钟 + CDN 缓存，稍等或加查询参数刷新 |

## 附录：定时任务文件内容（.github/workflows/rsvp-export.yml）

```yaml
name: 汇总嘉宾回执到 Excel

on:
  workflow_dispatch:
  schedule:
    # UTC 00:05 = 北京时间 08:05，每天自动汇总一次
    - cron: '5 0 * * *'

permissions:
  contents: write

jobs:
  export:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'

      - name: 安装依赖并生成 Excel
        env:
          RSVP_EMAIL: ${{ secrets.RSVP_EMAIL }}
          RSVP_MAILCODE: ${{ secrets.RSVP_MAILCODE }}
        run: |
          pip install openpyxl
          python scripts/export_rsvp.py

      - name: 提交到仓库
        run: |
          git config user.name "rsvp-bot"
          git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
          git add -f 嘉宾回执统计.xlsx
          if git diff --cached --quiet; then
            echo "没有新回执，无需更新"
          else
            git commit -m "chore: 更新嘉宾回执统计 $(date -u +%F)"
            git push
          fi
```
