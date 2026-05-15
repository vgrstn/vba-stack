# vba-stack
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA code for a generic Stack Class based on an encapsulated VB Collection

A lightweight, **Collection-backed LIFO stack** for VBA with:
- O(1) push/pop at the **top** (stored at position `1`)
- Safe **enumeration** (`For Each`) via a proper COM enumerator
- Clear error semantics (`Peek` / `Pop` on empty stack raise error 5)
- Zero dependencies (pure VBA)

---

## 📦 Features

- **Fast push/pop** (top is position `1` to avoid VB Collection tail-removal penalty)
- **Enumeration**: `For Each item In Stack` (via hidden `[_NewEnum]`)
- **Utility export**: `Items([base])` returns a 0- or 1-based array copy
- Pure VBA, no external references, Rubberduck-friendly annotations

---

## ⚙️ Public Interface

| Member             | Type       | Description |
|-------------------|------------|-------------|
| `Push(Item)`       | `Sub`      | Adds an item at the **top** of the stack. |
| `Pop()`            | `Function` | Returns **and removes** the top item. Raises error 5 if empty. |
| `Peek` *(Default)* | `Property` | Returns the top item **without** removing it. Raises error 5 if empty. |
| `Count`            | `Property` | Number of items. |
| `IsEmpty`          | `Property` | `True` if empty, else `False`. |
| `Clear`            | `Sub`      | Removes all items. |
| `Items([base])`    | `Function` | Returns all items as a `Variant()` array; `base` default = `0`; `arr(base)` = top (most recently pushed). |
| `For Each`         | Enumerator | Iterates **top → bottom** (don’t mutate during enumeration). |

**Error behavior**  
- Empty stack on `Peek` / `Pop` raises **`vbErrorInvalidProcedureCall (=5)`** with source `"Stack.Peek"` or `"Stack.Pop"`.

---

## 🚀 Quick Start

```vb
Dim s As New Stack

s.Push "alpha"
s.Push "beta"
Debug.Print s.Peek        ' -> beta  (top)
Debug.Print s.Pop         ' -> beta  (removed)
Debug.Print s.Pop         ' -> alpha (removed)
Debug.Print s.IsEmpty     ' -> True
```

---

## ⏱️ Performance

Timings (ms) for one `Push` + one `Pop`, measured on Windows x64:

| # | Count   | vba-stack | System.Collections.Stack |
|---|---------|-----------|--------------------------|
| 1 | 10      | 0.00054   | 0.00410                  |
| 2 | 100     | 0.00053   | 0.00396                  |
| 3 | 1,000   | 0.00054   | 0.00364                  |
| 4 | 10,000  | 0.00054   | 0.00390                  |
| 5 | 100,000 | 0.00053   | 0.00383                  |

Performance is consistent regardless of stack size. `System.Collections.Stack` uses late binding, which explains its relatively poor performance.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
