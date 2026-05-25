# 仅添加批注的OOXML说明

本参考文档记录此技能所需的最低OOXML组件。有意保持轻量级，专注于**仅添加批注**。

## 范围

- **支持**：向现有.docx添加Word批注而不修改文档文本。
- **不支持**：修订跟踪、丰富样式或完整模式验证。
- **验证**：仅文件存在性 + XML格式良好性。

## 所需文件

批注工作流程仅触及.docx包内的以下部分：

- `word/document.xml`（主正文）
- `word/comments.xml`（批注内容）
- `word/_rels/document.xml.rels`（与`comments.xml`的关系）
- `[Content_Types].xml`（`comments.xml`的内容类型覆盖）

## 命名空间

对批注元素使用WordprocessingML命名空间：

- `w` = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`

## 批注注入模型

要添加批注：

1. 在目标段落内容前插入`w:commentRangeStart`元素。
2. 在段落末尾插入`w:commentRangeEnd`元素。
3. 在范围结束后立即插入`w:commentReference`运行。
4. 在`word/comments.xml`中追加`w:comment`条目。
5. 确保`document.xml.rels`和`[Content_Types].xml`包含批注部分。

### 范围标记示例

```xml
<w:commentRangeStart w:id="12"/>
<!-- 段落内容 -->
<w:commentRangeEnd w:id="12"/>
<w:r><w:commentReference w:id="12"/></w:r>
```

### 批注内容示例

```xml
<w:comment w:id="12" w:author="Reviewer" w:initials="RV" w:date="2025-01-10T12:00:00Z">
  <w:p><w:r><w:t>批注文本第一行</w:t></w:r></w:p>
  <w:p><w:r><w:t>批注文本第二行</w:t></w:r></w:p>
</w:comment>
```

如果`w:t`节点以或以空格开头/结尾，请添加`xml:space="preserve"`，以便Word保留间距。

## 关系

`word/_rels/document.xml.rels`必须引用批注部分：

```xml
<Relationship Id="rIdX"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  Target="comments.xml"/>
```

## 内容类型

`[Content_Types].xml`必须包含：

```xml
<Override PartName="/word/comments.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
```

## 常见故障模式

- **批注不显示**：缺少关系或内容类型覆盖。
- **批注显示但为空**：`comments.xml`格式错误或`w:t`缺少文本节点。
- **空格缺失**：在`w:t`节点上添加`xml:space="preserve"`。

## When to Extend

If you need tracked changes, comment replies, or advanced metadata (e.g., commentsExtended), add separate modules and document them here. Keep this file short and only include what this skill actually uses.