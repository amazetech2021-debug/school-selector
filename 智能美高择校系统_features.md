# 功能详细规则参考

## 当前 API 配置（已内置）

```javascript
const API_CONFIGS = {
  deepseek: {
    url:   'https://api.deepseek.com/chat/completions',
    key:   'sk-4c3e0cdb678c435fba4d4917dcc16b91',
    model: 'deepseek-v4-pro',
    type:  'openai'
  }
};
// Google Maps Key: AIzaSyCN8Del7W133_edbGJQkRuNzJXrMNqgacM
// SiliconFlow 和 Claude 选项已移除，仅保留 DeepSeek
```

---

## AI 报告完整数据流

```
用户指令 → stripMarkdown() → 第一次callAI()（800t）→ 本地检索
→ 第二次callAI()（8000t）→ stripMarkdown() + JSON三层容错
→ renderReport()（含固定信息栏）→ 下载 HTML/Word/Excel
```

## 固定信息栏字段（getFixedInfo / getSchoolInfo）

从 SCHOOLS 数据库直接读取，两种报告模式均显示：

| 字段 | 数据库键 | 格式 |
|------|---------|------|
| GE评级 | `GE评级` | 金色高亮 |
| 类型 | `类型 Type` | |
| 宗教 | `宗教 Religion` | |
| 学生人数 | `全校学生数量 Student Body` | |
| 国际生年级 | `国际生年级Grades open for International students` | |
| 年费（含食宿） | `费用预估（含学费+食宿）` → getCost格式 `$6.5万/年` | 金色 |
| 托福最低 | `TOEFL Mini. （以9年级为标准）` | |
| 课程体系 | `课程体系` | AP学校/IB学校等 |
| 临近城市 | `临近城市(Nearby City)` | 去末尾分号 |
| 学校亮点 | `亮点总结（Highlights）` | 全量，`\n`→`<br>` |

## findSchoolIndex 四级模糊匹配

```javascript
function findSchoolIndex(name_en){
  if(!name_en) return -1;
  const q = name_en.toLowerCase().replace(/[（）()\[\],.，。\s]+/g,'');
  // 1. 精确
  let idx = SCHOOLS.findIndex(sc=>sc['学校名称(School Name）']===name_en);
  if(idx>=0) return idx;
  // 2. 忽略括号标点
  idx = SCHOOLS.findIndex(sc=>{
    const n=(sc['学校名称(School Name）']||'').toLowerCase().replace(/[（）()\[\],.，。\s]+/g,'');
    return n===q;
  });
  if(idx>=0) return idx;
  // 3. 前缀/包含
  idx = SCHOOLS.findIndex(sc=>{
    const n=(sc['学校名称(School Name）']||'').toLowerCase().replace(/[（）()\[\],.，。\s]+/g,'');
    return n.startsWith(q) || q.startsWith(n);
  });
  if(idx>=0) return idx;
  // 4. 前3个单词
  const words = name_en.toLowerCase().split(/\s+/).slice(0,3).join(' ');
  return SCHOOLS.findIndex(sc=>(sc['学校名称(School Name）']||'').toLowerCase().includes(words));
}
```

## callAI / stripMarkdown

```javascript
async function callAI(cfg, prompt, maxTokens, systemPrompt){
  const resp = await fetch(cfg.url, {
    method:'POST',
    headers:{'Content-Type':'application/json','Authorization':`Bearer ${cfg.key}`},
    body: JSON.stringify({model:cfg.model, max_tokens:maxTokens,
      messages:[{role:'system',content:systemPrompt},{role:'user',content:prompt}]})
  });
  const data = await resp.json();
  if(data.error) throw new Error(data.error.message);
  data._text = data.choices?.[0]?.message?.content || '';
  return data;
}

function stripMarkdown(text){
  text = text.replace(/^```[a-zA-Z]*\s*/,'').replace(/\s*```\s*$/,'');
  text = text.replace(/```[a-zA-Z]*/g,'').replace(/```/g,'');
  text = text.replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'');
  return text.trim();
}
```

## 12个检索字段

```javascript
const SEARCH_FIELDS = [
  '学校推文','拓展推文','特色择校','短期项目','录取捷报',
  '学校开设体育项目罗列 Athletic/Sports',
  '学校开设艺术项目罗列（艺术课程）Arts',
  '学校开设特色项目罗列 Signature programs',
  '体育','艺术','课程','亮点总结（Highlights）'
];
```

## KEYWORD_SYNONYMS / STOPWORDS

```javascript
const KEYWORD_SYNONYMS = {
  '高尔夫球':'高尔夫','golf':'高尔夫','skiing':'滑雪','equestrian':'马术',
  'tennis':'网球','basketball':'篮球','soccer':'足球','swimming':'游泳',
  'music':'音乐','dance':'舞蹈','robotics':'机器人','aviation':'航空',
};
const STOPWORDS = new Set([
  '帮我','查询','整理','学校','详细','信息','项目','介绍',
  '有哪些','推荐','什么','怎么','可以','需要','想要','告诉',
]);
```

---

## 筛选结果 Excel — buildXlsx 换行

```javascript
// wrapCols = new Set([4, 31])  →  亮点总结(4) 费用说明(31)
// 换行列: s="1"，\n → &#10;
// xl/styles.xml: xf[1] = wrapText="1" vertical="top"
// external_attributes = 4字节 u32（DataView.setUint32）!
```

## AI报告 Excel — 列顺序

专项查询（17列）：序号/英文名/中文名/州/GE评级/类型/宗教/学生人数/国际生年级/年费/托福/课程体系/临近城市/**学校亮点**（wrapCols=[13]）/项目详情/核心亮点/适合学生

择校推荐（18列）：序号/推荐类型/英文名/中文名/州/GE评级/类型/宗教/学生人数/国际生年级/年费/托福/课程体系/临近城市/**学校亮点**/推荐理由/核心亮点/注意事项

---

## Word 下载结构（每所学校）

```
h3 序号. 学校中文名  推荐类型（择校模式）
p  英文名 · 州
table（3行×6列）：GE/类型/宗教 | 人数/年级/年费 | 托福/课程体系/临近城市
div  学校亮点（按\n分行，金色左边框）
div  动态分析标签（去掉数量/美国高中/报告等词）
p    项目介绍 / 推荐理由
p    亮点
```

**动态标签 sectionLabel：**
```javascript
const t = (report.title||query||'')
  .replace(/\d+\s*所/g,'').replace(/美国高中/g,'')
  .replace(/学校汇总报告|汇总报告|详细报告|专项查询报告|报告|介绍/g,'').trim();
return t || '项目详情';
```

---

## 17 个详情页模块顺序

1.🏫学校概览 2.🏆学校排名 3.📍地理交通 4.📝申请要求 5.💰费用佣金
6.📚学术课程 7.🌟特色方向 8.🎭社团俱乐部 9.🏠住宿安排 10.📅假期安排
11.🎓大学升学 12.⭐杰出校友 13.✨学校亮点 14.📰推文资料 15.🏗️硬件设施
16.📋FAQ备注 17.🔗相关链接

---

## 新增功能规范

| 类型 | 步骤 |
|------|------|
| 带搜索多选下拉 | MS_CONFIG → HTML → filters → matchFilters → reset() → initMultiSelects() |
| Chip 筛选 | FILTER_DEFS → filters → matchFilters → reset() → updateFilterBadge() |
| 详情页字段 | 对应模块 ds-bd 内追加 dItem() 或 dTxt() |
| AI 同义词 | KEYWORD_SYNONYMS 添加 |
| AI 停用词 | STOPWORDS Set 添加 |
| 修改下载函数 | 修改后必须验证函数唯一性：`grep -c "function downloadAIHTML"` 应为1 |
