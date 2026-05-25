# Output Language Guide

Use the **contract’s dominant language** for all outputs: comment labels, summary headings, opinion text, and flowchart node text.

## Language Selection Rules

1. **Detect dominant language** by character ratio in the contract body:
   - If Chinese (CJK) characters are the clear majority (≈60%+), use Chinese.
   - If Latin letters are the clear majority (≈60%+), use English.
2. **If bilingual**, follow the language used in section headings/titles.
3. **If still ambiguous**, ask the user; if no response, default to English.

**Do not mix languages** within the same output.

## Comment Label Mapping

**English**
- `[Issue Type]`
- `[Risk Reason]`
- `[Revision Suggestion]`

**Chinese**
- `【问题类型】`
- `【风险原因】`
- `【修订建议】`

## Summary Heading Mapping (Chinese)

When the contract is Chinese, use these headings/labels exactly:

- `一、合同基本信息`
  - `项目` / `内容`
  - `合同名称` / `合同类型` / `合同当事人` / `签订时间` / `合同期限` / `合同金额`
- `二、业务模式概述`
  - `简要描述：...`
- `三、核心条款要素`
  - `3.1 交易要素`
    - `要素` / `具体内容`
    - `标的物/服务内容` / `数量规格` / `价格构成` / `支付方式` / `交付方式`
  - `3.2 权利义务分配`
    - `甲方主要权利义务：...`
    - `乙方主要权利义务：...`
  - `3.3 履行保障条款`
    - `条款类型` / `具体约定`
    - `违约责任` / `担保措施` / `验收标准` / `质量保证`
  - `3.4 风险分担与特殊约定`
    - `风险分担：...`
    - `特殊约定：...`
  - `3.5 争议解决与合同终止`
    - `项目` / `约定内容`
    - `争议解决方式` / `合同变更` / `合同解除` / `适用法律`
- `四、关键时间节点`

## Opinion Wording (Chinese)

For the four-section opinion format:

**第一段（总体评估）：**
- `本合同整体[可接受/存在重大风险/建议重新起草]。共发现 🔴 高风险 [N] 项、🟡 中风险 [N] 项、🔵 低风险 [N] 项。`
- 若 Layer 0 命中：`其中合法性前置审查发现 [N] 项潜在效力瑕疵问题。`

**第二段（重大风险摘要）：**
- 以总结句开头，仅列 🔴 项
- 格式：`- [条款定位]：[风险描述] → 建议：[一句话对策]`

**第三段（谈判优先级建议）：**
- `在进入合同谈判时，建议优先争取以下修改：`
- 编号列出修改点及影响

**第四段（签署前置条件）：**
- `建议在以下条件满足后方可签署：`
- 使用 □ 复选框格式
- 如无前置条件，此段可省略