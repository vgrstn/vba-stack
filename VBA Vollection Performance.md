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

## Example Results (ms)

These are representative numbers (Stopwatch timing, averaged):

count    add item    remove count   add item @1   remove 1
10       0.00008     0.00009        0.00005       0.00007
100      0.00007     0.00016        0.00005       0.00007
1000     0.00006     0.00215        0.00005       0.00007
10000    0.00006     0.03449        0.00005       0.00007
100000   0.00008     0.35003        0.00005       0.00008

Interpretation

remove count (remove last) grows ~linearly with count → size-dependent.
remove 1 and add item @1 stay essentially flat → constant-time.
