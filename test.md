# GitHub Flavored Markdown Test

This document tests all supported GFM formatting features.

---

## 1. Text Formatting

Normal text with **bold**, *italic*, ~~strikethrough~~, and `inline code`.

**Bold and *nested italic* inside bold.**

*Italic with **bold** inside.*

---

## 2. Headings

### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6

---

## 3. Lists

### Unordered List

- Item one
- Item two
  - Nested item A
  - Nested item B
    - Deep nested item
- Item three

### Ordered List

1. First item
2. Second item
   1. Sub-item 2.1
   2. Sub-item 2.2
3. Third item

### Task List

- [x] Completed task
- [ ] Incomplete task
- [x] Another done task

---

## 4. Blockquotes

> This is a blockquote.
> It can span multiple lines.
>
> > Nested blockquote inside.

---

## 5. Code Blocks

Inline code: `const x = 42;`

Fenced code block:

```python
def greet(name: str) -> str:
    """Return a greeting string."""
    return f"Hello, {name}!"

print(greet("World"))
```

```javascript
const add = (a, b) => a + b;
console.log(add(1, 2)); // 3
```

---

## 6. Tables

| Name       | Age | City          |
|------------|-----|---------------|
| Alice      | 30  | New York      |
| Bob        | 25  | San Francisco |
| Charlie    | 35  | London        |

---

## 7. Links and Images

[Visit GitHub](https://github.com)

[Relative link](./README.md)

---

## 8. Horizontal Rule

---

## 9. Mixed Content

Here is a paragraph with **bold**, *italic*, `code`, and a [link](https://example.com) all together.

> Blockquote with **bold** and `code` inside.

1. List item with *italic* text
2. List item with `code` and **bold**
3. List item with a [link](https://example.com)

