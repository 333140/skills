# 法律数据库 URL 与查询模板

## 一、国家法律法规数据库（flk.npc.gov.cn）— 最权威来源

### 查询入口
- 首页：`https://flk.npc.gov.cn/`
- 高级检索：`https://flk.npc.gov.cn/index.html`（页面内搜索框）
- 法律文本详情页 URL 模式：`https://flk.npc.gov.cn/detail2.html{参数}`

### 搜索引擎查询模板
```
web_search("site:flk.npc.gov.cn {法律全称}")
web_search("site:flk.npc.gov.cn {法律简称} 全文")
web_search("site:flk.npc.gov.cn {法律全称} 修订")
```

### WebFetch 抓取策略
- 先用 `web_search` 搜索，从结果中提取 flk.npc.gov.cn 的具体页面 URL
- 再用 `web_fetch` 访问该 URL 获取法律全文
- 若直接页面无法访问，尝试搜索结果摘要中的条文内容

---

## 二、全国人民代表大会官网（npc.gov.cn）

### 查询入口
- 首页：`http://www.npc.gov.cn/`
- 法律数据库入口：`http://www.npc.gov.cn/npc/c2190/list.shtml`
- 立法专题：`http://www.npc.gov.cn/npc/c2190/c2191/list.shtml`

### 搜索引擎查询模板
```
web_search("site:npc.gov.cn {法律全称} 全文")
web_search("site:npc.gov.cn {法律全称} 修订 决定")
web_search("site:npc.gov.cn {法律简称} 第{X}条")
```

---

## 三、司法部官网（moj.gov.cn）— 行政法规/部门规章

### 查询入口
- 首页：`http://www.moj.gov.cn/`
- 法规查询：`http://www.moj.gov.cn/Department/index.html`
- 行政法规：`http://www.moj.gov.cn/Department/node_6.html`

### 搜索引擎查询模板
```
web_search("site:moj.gov.cn {行政法规名称}")
web_search("site:moj.gov.cn {部门规章名称} 规定")
```

---

## 四、国务院官网（gov.cn）— 行政法规

### 查询入口
- 首页：`https://www.gov.cn/`
- 政策法规库：`https://www.gov.cn/zhengce/zhengceku/`
- 行政法规：`https://www.gov.cn/zhengce/content/`（按发布日期组织）

### 搜索引擎查询模板
```
web_search("site:gov.cn {行政法规名称} 全文")
web_search("site:gov.cn 国务院 {条例名称} 修订")
web_search("site:gov.cn {法规简称} 令 第{X}号")
```

---

## 五、最高人民法院官网（court.gov.cn）— 司法解释

### 查询入口
- 首页：`https://www.court.gov.cn/`
- 司法解释：`https://www.court.gov.cn/jianshe/xiangqing.htm`
- 最高法公报：`https://www.court.gov.cn/gongbao/`

### 搜索引擎查询模板
```
web_search("site:court.gov.cn {司法解释名称}")
web_search("site:court.gov.cn 最高人民法院 {规定/解释/批复} {关键词}")
web_search("{司法解释名称} 法释〔{年份}〕{X}号")
```

---

## 六、中国裁判文书网（wenshu.court.gov.cn）— 辅助参考

> ⚠️ 裁判文书网不是法律来源，但可用于验证法条在司法实践中的适用。

### 查询入口
- 首页：`https://wenshu.court.gov.cn/`

### 搜索引擎查询模板
```
web_search("site:wenshu.court.gov.cn {法律名称} 第{X}条")
```

---

## 七、地方性法规来源

| 省份/城市 | 查询来源 | 搜索模板 |
|-----------|---------|---------|
| 北京 | `bjrd.gov.cn` | `site:bjrd.gov.cn {法规名称}` |
| 上海 | `spcsc.sh.cn` | `site:spcsc.sh.cn {法规名称}` |
| 广东 | `gdrd.gd.cn` | `site:gdrd.gd.cn {法规名称}` |
| 浙江 | `zjrd.gov.cn` | `site:zjrd.gov.cn {法规名称}` |
| 江苏 | `jsrd.gov.cn` | `site:jsrd.gov.cn {法规名称}` |
| 其他省份 | `web_search("{省份}人大 {法规名称}")` | 通用搜索兜底 |

---

## 八、港澳台法律数据库

### 香港
- 电子版香港法例：`https://www.elegislation.gov.hk/`
- 搜索模板：`web_search("site:elegislation.gov.hk {law name} Cap.{X}")`
- 常用 Cap. 编号：Cap.1《释义及通则条例》、Cap.221《刑事罪行条例》、Cap.45《雇员条例》

### 澳门
- 澳门印务局：`https://bo.io.gov.mo/`
- 搜索模板：`web_search("site:bo.io.gov.mo {法律名称}")`

### 台湾
- 全国法规资料库：`https://law.moj.gov.tw/`
- 搜索模板：`web_search("site:law.moj.gov.tw {法律名称} 全文")`
- 法规查询：`https://law.moj.gov.tw/Law/LawSearch.aspx`（按法规名称/法规类别搜索）

---

## 九、境外法律数据库

### 美国联邦
- 美国法典：`https://uscode.house.gov/`
- 搜索模板：`web_search("site:uscode.house.gov {act name} {title} U.S.C.")`
- Cornell LII：`https://www.law.cornell.edu/uscode/text`（更友好的界面）
- 搜索模板：`web_search("site:law.cornell.edu {act name}")`

### 欧盟
- EUR-Lex：`https://eur-lex.europa.eu/`
- 搜索模板：`web_search("site:eur-lex.europa.eu {regulation/directive number}")`
- 按编号查询：`https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX:{编号}`

### 英国
- legislation.gov.uk：`https://www.legislation.gov.uk/`
- 搜索模板：`web_search("site:legislation.gov.uk {act name} {year}")`

### 日本
- e-Gov 法令检索：`https://elaws.e-gov.go.jp/`
- 搜索模板：`web_search("site:elaws.e-gov.go.jp {法律名}")`

### 其他国家
- 通用搜索模板：`web_search("{country} {law name} official full text site:gov")`

---

## 十、辅助参考来源（非最终权威）

| 来源 | 网址 | 说明 | 可信度 |
|------|------|------|--------|
| 北大法宝 | `pkulaw.cn` | 法律数据库商业平台，覆盖全但非官方 | 中 |
| 法信 | `faxin.cn` | 最高法旗下法律知识平台 | 中高 |
| 威科先行 | `law.wkinfo.com.cn` | 商业法律数据库 | 中 |
| 百度百科法律词条 | `baike.baidu.com` | 仅作初步参考，不可作为权威来源 | 低 |
| 知乎法律回答 | `zhihu.com` | 仅作参考，需交叉验证 | 低 |

> ⚠️ 辅助来源仅用于交叉验证，不可单独作为核验依据。核验报告中的来源必须包含至少一个官方来源。

---

## 十一、搜索效率优化策略

### 搜索顺序规则
1. **先搜官方数据库**：`site:flk.npc.gov.cn` 优先于通用搜索
2. **精确优先**：法律全称搜索优先于简称（如"中华人民共和国劳动合同法"优于"劳动法"）
3. **版本限定**：对已知修订的法律，搜索时加年份限定（如"2023年修订""2024年修正"）
4. **废止检查**：每次核验必须执行 `web_search("{法律名称} 废止 现行有效")` 确认状态

### 常见搜索失败应对
| 失败原因 | 应对策略 |
|---------|---------|
| 官方网站无法访问 | 切换到第二优先来源，记录访问失败状态 |
| 搜索无结果 | 检查法律名称是否正确，尝试简称/别称搜索 |
| 法条编号不匹配 | 可能是版本差异，搜索修订历史确认 |
| 地方性法规无网站 | 使用通用搜索 `web_search("{省份} {法规名称} 全文")` |

### 并行搜索建议
对同一条核验请求，可同时发起多个搜索（使用并行 tool call）：
```
// 同时搜索多个来源
web_search("site:flk.npc.gov.cn {法律名称}")
web_search("site:npc.gov.cn {法律名称} 全文")
web_search("{法律名称} 现行有效 最新版本")
```
