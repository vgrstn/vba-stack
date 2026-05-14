# VBA Collection Performance: Tail-Removal Cost

This note documents and reproduces a long-standing performance quirk in `VBA.Collection`:  
**removing the last item** (the “tail”) gets slower as the collection grows, while removing from the **first position** stays fast and size-independent.

- `Collection` is implemented as a **doubly linked list** with a **hash table** for key lookup.  
- Adding/removing at **position 1** is effectively O(1).  
- Removing the **last item** (by index = `Count`) degrades with size — effectively **O(n)** — because the internal code traverses from the head to reach the tail node.

---

## Why this matters

If you build stacks or queues on top of `Collection` and you treat the “top” as the **last index**, your `Pop`/`Dequeue` cost will grow with `Count`.  
If, instead, you treat the “top” as **index 1**, both `Push` (insert before 1) and `Pop` (remove index 1) stay **constant-time**.

---

## 📊 Example Results

Representative benchmark results (milliseconds per operation, averaged over multiple runs):

| Count | Add Item | Remove Count | Add Item @1 | Remove 1 |
|:------:|----------:|--------------:|--------------:|----------:|
| 10     | 0.00009  | 0.00008       | 0.00007       | 0.00007 |
| 100    | 0.00007  | 0.00015       | 0.00007       | 0.00007 |
| 1 000  | 0.00008  | 0.00215       | 0.00007       | 0.00006 |
| 10 000 | 0.00006  | 0.03279       | 0.00007       | 0.00006 |
| 100 000| 0.00007  | 0.32448       | 0.00007       | 0.00007 |

### Interpretation

- **`Remove Count`** (removing the last item) becomes slower as the collection grows — roughly linear with size.  
- **`Remove 1`** and **`Add Item @1`** remain constant in speed — true **O(1)** operations.  
- This confirms that `Collection` internally **traverses from the head** to reach the tail when removing the last item.  
- By defining the stack top at **index 1**, this performance issue is completely avoided.
