# 合同审查示例

## 快速开始

```python
# -*- coding: utf-8 -*-
from scripts.workflow import ContractReviewWorkflow

comments = [
    {
        "search": "Total Price",
        "comment": """[问题类型] 付款条款
[风险原因] 第3.2条 stated 总金额为USD 100,000，但付款条款第5.1条列出USD 1,000,000。
[修改建议] 统一各条款中的总金额，并明确是否包含税费。""",
        "risk_level": "High",
    }
]

workflow = ContractReviewWorkflow("Contract.docx", "Reviewer")
workflow.run_full_workflow(comments, "Contract_Reviewed.docx")
```

## 合同概要（英文）

```python
# -*- coding: utf-8 -*-
summary_text = """I. Basic Contract Information
Item\tContent
Contract Name\tNot specified
Contract Type\tNot specified
Parties\tParty A: Not specified
Party B: Not specified
Signing Date\tNot specified
Term\tNot specified
Contract Amount\tNot specified
II. Business Model Overview
Brief description: Not specified
III. Key Clause Elements
3.1 Transaction Elements
Element\tDetails
Subject Matter/Services\tNot specified
Quantity/Specs\tNot specified
Pricing Structure\tNot specified
Payment Terms\tNot specified
Delivery Terms\tNot specified
3.2 Rights and Obligations
Party A main rights/obligations:
Not specified
Party B main rights/obligations:
Not specified
3.3 Performance Safeguards
Clause Type\tDetails
Liability/Default\tNot specified
Guarantees/Security\tNot specified
Acceptance Standards\tNot specified
Quality Warranty\tNot specified
3.4 Risk Allocation & Special Terms
Risk Allocation:
Not specified
Special Terms:
Not specified
3.5 Dispute Resolution & Termination
Item\tDetails
Dispute Resolution\tNot specified
Amendment\tNot specified
Termination\tNot specified
Governing Law\tNot specified
IV. Key Timeline Milestones
Not specified
"""

workflow = ContractReviewWorkflow("Contract.docx", "Reviewer")
workflow.run_full_workflow(
    comments,
    "Contract_Reviewed.docx",
    summary_text=summary_text,
    summary_filename="Contract_Summary.docx",
    summary_font="Times New Roman",
)
```

## 综合审核意见（中文，四段式）

```python
# -*- coding: utf-8 -*-
opinion_text = """本合同整体存在重大风险。共发现 🔴 高风险 3 项、🟡 中风险 2 项、🔵 低风险 1 项。

- 第5.1条付款条款：总金额与第3.2条不一致（100万 vs 1000万），建议统一并明确含税口径。
- 第12条争议解决：仲裁机构名称（"北京仲裁委员会"）存在歧义，建议改为"中国国际经济贸易仲裁委员会"。
- 第8.3条竞业限制：竞业限制期限为36个月，超过法定上限24个月，建议调整为24个月以内。

在进入合同谈判时，建议优先争取以下修改：
1. 统一合同金额并明确含税口径（影响：避免付款执行争议）
2. 明确仲裁机构全称（影响：避免仲裁条款无效）
3. 缩短竞业限制期限至24个月（影响：避免条款部分无效）

建议在以下条件满足后方可签署：
□ 合同金额在各条款间统一
□ 仲裁机构名称确认无误

---

*本审查报告由 AI 辅助生成，仅供参考，不构成法律意见。*"""

workflow = ContractReviewWorkflow("Contract.docx", "Reviewer")
workflow.run_full_workflow(
    comments,
    "Contract_reviewed.docx",
    opinion_text=opinion_text,
    opinion_filename="Consolidated_Opinion.docx",
    opinion_font="Times New Roman",
)
```

## 业务流程图（Mermaid）

```python
# -*- coding: utf-8 -*-
flowchart_mermaid = """flowchart TD
    A[合同签署] -->|?| B[履行]
"""

workflow = ContractReviewWorkflow("Contract.docx", "Reviewer")
workflow.run_full_workflow(
    comments,
    "Contract_reviewed.docx",
    flowchart_mermaid=flowchart_mermaid,
)
```

## 完整工作流程示例

```python
# -*- coding: utf-8 -*-
#!/usr/bin/env python3
import sys
from pathlib import Path

skill_dir = Path(__file__).parent.parent
sys.path.insert(0, str(skill_dir))

from scripts.workflow import ContractReviewWorkflow

contract_path = "path/to/contract.docx"
output_path = "contract_reviewed.docx"
report_path = "review_report.txt"
reviewer = "Reviewer"

comments = [
    {
        "search": "Total Price",
        "comment": """[问题类型] 付款条款
[风险原因] 价格条款与付款条款金额不匹配。
[修改建议] 统一总金额并明确税费。""",
        "risk_level": "High"
    }
]

summary_text = """I. 基本合同信息
项目\t内容
合同名称\t未指定
合同类型\t未指定
当事人\t甲方：未指定
乙方：未指定
签署日期\t未指定
期限\t未指定
合同金额\t未指定
"""

opinion_text = """本合同整体存在重大风险。共发现 🔴 高风险 2 项、🟡 中风险 1 项、🔵 低风险 0 项。

- 付款条款：付款时间不明确，可能导致付款执行争议，建议明确具体付款日期或里程碑。
- 验收标准：验收标准缺失，可能导致无法主张违约，建议补充具体验收标准和程序。

在进入合同谈判时，建议优先争取以下修改：
1. 明确付款时间和里程碑（影响：避免付款执行争议）
2. 补充验收标准条款（影响：确保可主张违约）

建议在以下条件满足后方可签署：
□ 付款时间条款明确
□ 验收标准条款补充完整

---

*本审查报告由 AI 辅助生成，仅供参考，不构成法律意见。*"""

flowchart_mermaid = """flowchart TD
    A[合同签署] -->|?| B[履行]
"""

workflow = ContractReviewWorkflow(
    contract_path=contract_path,
    reviewer_name=reviewer,
    enable_smart_keyword_expansion=False,
)

workflow.run_full_workflow(
    comments=comments,
    output_docx_filename=output_path,
    report_filename=report_path,
    summary_text=summary_text,
    summary_filename="Contract_Summary.docx",
    summary_font="Times New Roman",
    opinion_text=opinion_text,
    opinion_filename="Consolidated_Opinion.docx",
    opinion_font="Times New Roman",
    flowchart_mermaid=flowchart_mermaid,
    render_flowchart=True,
    parallel_outputs=True,
)

print(f"✓ 已添加批注: {len(workflow.comments_added)}")
print(f"✗ 添加失败的批注: {len(workflow.comments_failed)}")
```

## 语言说明

- 输出必须遵循合同的语言。
- 对于中文合同，使用 **[language.md](language.md)** 中的中文标签，并将 `summary_font` / `opinion_font` 设置为 Fangsong (仿宋)。