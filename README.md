# vba-stack
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Generic VBA Class based on an encapsulated VB Collection
A lightweight, **Collection-backed LIFO stack** for VBA with:
- O(1) push/pop at the **top** (stored at position `1`)
- Safe **enumeration** (`For Each`) via a proper COM enumerator
- Clear error semantics (`Peek` / `Pop` on empty stack raise error 5)
- Zero dependencies (pure VBA)

> Version: **2025.09.09**  
> Author: **Vincent van Geerestein**

---

## 📦 Features

- **Fast push/pop** (top is position `1` to avoid VB Collection tail-removal penalty)
- **Enumeration**: `For Each item In Stack` (via hidden `[_NewEnum]`)
- **Strong defaults**: `Peek` is the **default member**
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
| `Items([base])`    | `Function` | Returns all items as a `Variant()` array; `base` can be `0` or `1`. |
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
