# VBA Stack Performance Comparison

This benchmark compares the **custom VBA Stack class** (implemented using a `Collection` with the top at index 1)  
against the **.NET `System.Collections.Stack`** accessed via COM interop.

---

## 🧩 Objective

To demonstrate that:

1. The **custom Stack** achieves constant-time `Push` / `Pop` performance independent of size.  
2. The **System.Collections.Stack** (late-bound) incurs noticeable overhead from COM dispatch and object wrapping.

---

## ⚙️ Test Setup

### Environment
- Windows 11 x64  
- Excel VBA 7.1 (x64)  
- Timing via [`Stopwatch`](../Stopwatch.cls) using `QueryPerformanceCounter`

### Stack Implementations

| Implementation | Internal storage | Access binding | Top position | Expected complexity |
|----------------|------------------|----------------|---------------|--------------------|
| **Custom Stack** | `Collection` | Early (VBA) | Index 1 | O(1) |
| **System Stack** | .NET `System.Collections.Stack` | Late (COM) | Top of internal array | O(1) + COM overhead |

---

## 🧪 Benchmark Code

```vb
Public Sub Run_Stack_Benchmarks()
    Dim n As Long, i As Long
    Dim s As Object, v As Variant
    Dim tCustom As Double, tSystem As Double

    ' Prepare data
    n = 10000
    v = Array(1, 2, 3, 4, 5)

    ' --- Custom Stack ---
    Dim cs As New Stack
    Stopwatch.Start
    For i = 1 To n
        cs.Push v(0)
        cs.Pop
    Next
    tCustom = Stopwatch.Halt

    ' --- System.Collections.Stack ---
    Set s = CreateObject("System.Collections.Stack")
    Stopwatch.Start
    For i = 1 To n
        s.Push v(0)
        s.Pop
    Next
    tSystem = Stopwatch.Halt

    Debug.Print "Iterations:", n
    Debug.Print "Custom Stack (s):", Format$(tCustom, "0.000000")
    Debug.Print "System Stack (s):", Format$(tSystem, "0.000000")
    Debug.Print "Speed-up:", Format$(tSystem / tCustom, "0.0x")
End Sub
