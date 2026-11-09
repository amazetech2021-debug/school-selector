---
name: us-school-selector
description: >
  专为「智能美高择校系统」项目设计的开发与维护技能。
  当用户提到以下任何场景时必须立即调用此技能：
  更新学校数据、重新生成HTML、修改筛选逻辑、修复Bug、新增功能、
  学生择校咨询（输入学生情况自动推荐）、更新规则文档或Skill文件。
  关键词触发：美高择校、selector_responsive、school-selector、
  学校信息表.xlsx、SCHOOL_COORDS_DB、智能择校、择校系统、
  GitHub Pages择校、index.html更新。
---

# 智能美高择校系统 · 完整开发维护技能

## 触发场景
- 更新学校数据（收到新 Excel / 学校信息表）
- 重新生成 HTML / 更新系统 / 上传 GitHub
- 修改筛选逻辑、详情页、排序、导出功能
- 修复 Bug（车程、坐标、地址、导出等）
- 新增功能（筛选项、字段、AI报告格式）
- 学生择校咨询（输入学生情况，自动推荐）
- 更新规则文档 / Skill 文件

**关键词**：美高择校、selector_responsive、school-selector、学校信息表、SCHOOL_COORDS_DB、智能择校、index.html

---

## 一、项目概览

| 项目 | 信息 |
|------|------|
| GitHub 仓库 | amazetech2021-debug/school-selector |
| 访问地址 | https://amazetech2021-debug.github.io/school-selector |
| 模板文件 | `/home/claude/selector_responsive.html` |
| 数据文件 | `/home/claude/schools_all.json`（228所，116列） |
| 坐标文件 | `/home/claude/school_coords_final.json`（独立于Excel） |
| 输出文件 | `/mnt/user-data/outputs/智能美高择校系统_响应式版.html` |
| Skill文件 | `/mnt/user-data/outputs/us-school-selector.skill` |

---

## 二、核心工作流

### 2.1 标准构建 HTML

```python
import json, os
with open('/home/claude/schools_all.json','r',encoding='utf-8') as f: data=json.load(f)
with open('/home/claude/school_coords_final.json','r',encoding='utf-8') as f: coords=json.load(f)
with open('/home/claude/selector_responsive.html','r',encoding='utf-8') as f: html=f.read()

marker='const SCHOOL_COORDS_DB = '
idx_s=html.index(marker)+len(marker)
depth=0; idx_e=idx_s
for i,ch in enumerate(html[idx_s:],idx_s):
    if ch=='{': depth+=1
    elif ch=='}':
        depth-=1
        if depth==0: idx_e=i+1; break
html=html[:idx_s]+json.dumps(coords,ensure_ascii=False)+html[idx_e:]

start=html.index('const SCHOOLS = ')+len('const SCHOOLS = ')
end=html.index(';\n',start)
html=html[:start]+json.dumps(data,ensure_ascii=False)+html[end:]

out='/mnt/user-data/outputs/智能美高择校系统_响应式版.html'
with open(out,'w',encoding='utf-8') as f: f.write(html)
print(f'完成: {len(data)}所, {os.path.getsize(out)//1024}KB')
```

### 2.2 模板丢失时恢复

```python
with open('/mnt/user-data/outputs/智能美高择校系统_响应式版.html','r',encoding='utf-8') as f:
    html = f.read()
import json
marker='const SCHOOL_COORDS_DB = '; idx_s=html.index(marker)+len(marker)
depth=0; idx_e=idx_s
for i,ch in enumerate(html[idx_s:],idx_s):
    if ch=='{': depth+=1
    elif ch=='}':
        depth-=1
        if depth==0: idx_e=i+1; break
coords=json.loads(html[idx_s:idx_e])
with open('/home/claude/school_coords_final.json','w',encoding='utf-8') as f: json.dump(coords,f,ensure_ascii=False)
start=html.index('const SCHOOLS = ')+len('const SCHOOLS = ')
end=html.index(';\n',start)
data=json.loads(html[start:end])
with open('/home/claude/schools_all.json','w',encoding='utf-8') as f: json.dump(data,f,ensure_ascii=False)
with open('/home/claude/selector_responsive.html','w',encoding='utf-8') as f: f.write(html[:start]+'[]'+html[end:])
print(f'恢复: {len(data)}所, {len(coords)}坐标')
```

### 2.3 Excel 数据更新

```python
import pandas as pd, json
df = pd.read_excel('学校信息表_xxxx.xlsx')
df2 = df[df['学校名称(School Name）'].notna()].copy()
exclude_cols = {'data_id'}
cols = [c for c in df.columns if c not in exclude_cols]
records = []
for _, row in df2.iterrows():
    r = {}
    for col in cols:
        val = row[col]
        if pd.isna(val): r[col] = ''
        elif isinstance(val, float) and val == int(val): r[col] = int(val)
        else: r[col] = str(val).replace('\xa0', ' ').strip()
    r['学校名称(School Name）'] = r['学校名称(School Name）'].replace('\xa0',' ').strip()
    records.append(r)
with open('/home/claude/schools_all.json','w',encoding='utf-8') as f:
    json.dump(records, f, ensure_ascii=False)
```

⚠️ **\xa0 隐藏字符** 是最大风险：导致坐标查找失败、车程计算回退到州中心。

### 2.4 更新 Skill 和规则文档

```bash
python3 -c "
import zipfile, os
with zipfile.ZipFile('/mnt/user-data/outputs/us-school-selector.skill','w',zipfile.ZIP_DEFLATED) as zf:
    for root,dirs,files in os.walk('/home/claude/skill-us-school-selector'):
        for file in files:
            abs_path=os.path.join(root,file)
            arcname=os.path.relpath(abs_path,os.path.dirname('/home/claude/skill-us-school-selector'))
            zf.write(abs_path,arcname)
print('Skill打包完成')
"
```

---

## 三、当前嵌入配置（已内置）

| 配置项 | 值 | 说明 |
|--------|-----|------|
| DeepSeek API Key | sk-4c3e0cdb678c435fba4d4917dcc16b91 | 唯一 AI 服务 |
| DeepSeek 模型 | deepseek-v4-pro | 最新版本 |
| Google Maps Key | AIzaSyCN8Del7W133_edbGJQkRuNzJXrMNqgacM | 车程筛选，已内置隐藏 input |

---

## 四、学生择校咨询（直接在对话中推荐）

### 信息收集维度

| 维度 | Excel 字段 |
|------|-----------|
| 托福/ESL | `TOEFL Mini.`、`ESL语言支持` |
| 预算 | `费用预估（含学费+食宿）` |
| 地理 | `州 State`、`是否可以直飞` |
| 特色 | `体育`、`艺术`、推文字段 |
| 规模/中国学生 | `全校学生数量`、`中国学生数量` |
| 课程 | `课程体系`、`#AP/#IB/#CC` |
| 年级 | `开设年级 Grades` |

### Python 快速筛选

```python
def filter_schools(data, toefl=None, budget_wan=None, need_esl=False,
                   states=None, school_type=None, grade=None):
    results = []
    for s in data:
        if toefl is not None:
            req = str(s.get('TOEFL Mini. （以9年级为标准）','')).replace('+','').strip()
            try:
                if req and int(req) > toefl: continue
            except: pass
        if budget_wan is not None:
            cost = s.get('费用预估（含学费+食宿）','')
            try:
                if cost and float(cost) > budget_wan * 10000: continue
            except: pass
        if need_esl and '提供' not in str(s.get('ESL语言支持','')): continue
        if states and s.get('州 State','') not in states: continue
        if school_type == 'male' and '女校' in str(s.get('类型 Type','')): continue
        if grade:
            g = str(s.get('开设年级 Grades',''))
            if str(grade) not in g and '9-12' not in g.lower(): continue
        results.append(s)
    return results
```

---

## 五、智能择校报告 — 两阶段 AI 架构

```
用户自然语言指令
    ↓
【第一次调用】指令解析（800 tokens）→ {intent, target_schools, search_keywords, scope}
    ↓
【本地检索】specific模式（全字段）/ all模式（关键词段落，最多25所）
    ↓
【第二次调用】报告生成（8000 tokens）每校200字以内，无数量限制，严禁编造
```

### 报告卡片结构

```
┌─────────────────────────────────────────┐
│ 序号  学校中文名                 查看详情→ │ 标题栏
│       英文名 · 州  [GE徽章]...            │
├─────────────────────────────────────────┤
│ GE评级  类型  宗教  学生人数  国际生年级   │ 固定信息栏
│ 年费(含食宿)$6.5万/年  托福  课程体系     │ 直接从SCHOOLS读取
│ 临近城市                                  │ findSchoolIndex四级模糊匹配
│ 学校亮点（全量，\n→<br>保留分行）          │
├─────────────────────────────────────────┤
│ [AI生成：项目介绍 / 推荐理由]              │ 动态内容
└─────────────────────────────────────────┘
```

### findSchoolIndex 四级模糊匹配

1. 精确匹配
2. 忽略括号/空格/标点后匹配 → 解决 `Southwestern Academy` vs `Southwestern Academy（CA）`
3. 前缀/包含匹配 → 解决 `La Salle College Prep` vs `La Salle College Preparatory（Homestay）`
4. 前3个单词匹配 → 其他简写情况

### 报告下载格式

| 格式 | 实现方式 | 关键细节 |
|------|---------|---------|
| HTML | 截取 `ai-result-inner` innerHTML + 页面CSS | 所有颜色加 `!important` 强制深色主题 |
| Word | 从SCHOOLS直接读取，3行×6列固定信息表 | 动态分析标签去掉数量和"美国高中/报告"等词 |
| Excel | buildXlsx，学校亮点列保留换行 | 含GE/类型/宗教/人数/年级/年费/托福/课程/城市/亮点 |

**Word 动态分析标签：** 去掉 `\d+所`、`美国高中`、`汇总报告/详细报告/报告/介绍` 等词后显示

**⚠️ 重复函数陷阱：** 每次修改下载函数后必须用 `grep -c "function downloadAIHTML"` 确认只有1个，否则浏览器执行旧版

---

## 六、筛选结果导出 Excel（generateFilterReport）

38列（含FS/NICHE排名）+ 车程2列可选：
GE评级/州/英文名/中文名/亮点/托福/Duolingo/SSAT/面试/申请表/申请费/截止/录取/类型/宗教/年级/人数/建校/PG/最高年龄/国际生年级/住校%/国际生%/中国学生/师生比/学制/12年级/SAT/AP-IB-CC/ESL/年费/费用说明/住宿/走读/直飞/官网/**FS排名/NICHE排名**

换行字段：亮点总结（列4）和费用说明（列31），`wrapCols = new Set([4, 31])`

---

## 七、坐标库维护

```json
"Excelsior School": [34.1397, -118.0353],
"Fairmont Schools": [33.8366, -117.9143],
"MARANATHA HIGH SCHOOL": [34.1478, -118.1073],
"Besant Hill School": [34.448, -119.243],
"Chaminade College Preparatory School": [38.5967, -90.4318],
"Whitestone Academy": [40.788, -73.819],
"Austin Waldorf School TX": [30.3076, -97.7547],
"Green Fields School (homestay)": [32.2226, -110.9747],
"Phoenix Christian Preparatory School (Homestay)": [33.4484, -112.0740]
```

---

## 八、常见问题排查

| 问题 | 原因 | 解决 |
|------|------|------|
| 模板文件丢失 | session 重置 | 从最新 output HTML 恢复（见 2.2） |
| AI 固定信息栏空白 | 名称不匹配 | findSchoolIndex 四级模糊匹配 |
| 下载报告是旧版内容 | 存在重复函数 | grep 检查，保留唯一最新版 |
| AI 报告 JSON 截断 | token 超限 | 减少查询学校数量 |
| Excel 导出空白/PK | ZIP extattr 2字节 | DataView.setUint32（4字节） |
| SyntaxError | 重复括号 | node --check 验证 |
| 车程偏小 | 坐标错误/\xa0污染 | 检查坐标库，清理 key 中的 \xa0 |

详细功能规则见 `references/features.md`
