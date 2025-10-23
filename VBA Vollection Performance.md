# VBA Collection Performance: Tail-Removal Cost

This note documents and reproduces a long-standing performance quirk in `VBA.Collection`:  
**removing the last item** (the “tail”) gets slower as the collection grows, while removing from the **first position** stays fast and size-independent.

## TL;DR

- `Collection` is implemented as a **doubly linked list** with a **hash table** for key lookup.  
- Adding/removing at **position 1** is effectively O(1).  
- Removing the **last item** (by index = `Count`) degrades with size — effectively **O(n)** — because the internal code traverses from the head to reach the tail node.

---

## Why this matters

If you build stacks or queues on top of `Collection` and you treat the “top” as the **last index**, your `Pop`/`Dequeue` cost will grow with `Count`.  
If, instead, you treat the “top” as **index 1**, both `Push` (insert before 1) and `Pop` (remove index 1) stay **constant-time**.

---
