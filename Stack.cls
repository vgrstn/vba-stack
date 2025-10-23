VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Generic Stack Class."
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder("Class")
'@ModuleDescription("Generic Stack Class.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Generic Stack Class
' Enumeration: For Each item in object
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.09.09
'
' Methods
' Clear             Removes all items from the stack
' Push item         Adds an item at the top of the stack
' Pop               Returns and removes the item at the top of the stack
' Items([base])     Exports all stack items to an array
'
' Properties (Get)
' Count             Returns the number of items contained in the stack
' IsEmpty           Returns True if the stack is empty or False if not
' Peek              Returns the item at the top of the stack (default)
'
' This generic Stack Class is implemented using an encapsulated VB Collection.
' The top of the stack is coded at position 1 in order to avoid performance
' degradation when removing an item from the last position in a VB Collection.
'
' Timings (ms) for one push + one pop.
'     count      Class       System
' 1      10    0.00049      0.00342
' 2     100    0.00049      0.00341
' 3    1000    0.00050      0.00344
' 4   10000    0.00049      0.00342
' 5  100000    0.00050      0.00343
'
' The results for the Stack Class show a consistent performance independent of
' its actual size. The System.Collections.Stack uses late binding which is most
' likely the explanation for its relative poor performance.
'
' The VB Collection class is implemented as a double linked list with a hash
' table to access an item by its key. As expected the CPU time for adding a
' new item to the first position or to the last position are similar and these
' timings do not depend on the number of items. Removing an item from the first
' position shows similar timings as for adding an item. However, the timings for
' removing an item from the last position proofs to be dependent on the number
' of items already in the VB Collection. The explanation must be that in the
' latter case the internal code doesn't use the tail node but actuallly follows
' the list from the head node until arriving at the to be deleted node. This
' poor performance has been reported since early VB versions, see Hardcore
' Visual Basic version 5.0 by Bruce McKinney (Microsoft Press, 1997)
'
' Timings (ms) for VB Collection
'     count    add item    remove count     add item @1     remove 1
' 1      10     0.00008         0.00009         0.00005      0.00007
' 2     100     0.00007         0.00016         0.00005      0.00007
' 3    1000     0.00006         0.00215         0.00005      0.00007
' 4   10000     0.00006         0.03449         0.00005      0.00007
' 5  100000     0.00008         0.35003         0.00005      0.00008

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
End Enum

Private Const Top As Long = 1

Private Type TPRIVATE
    Stack As Collection
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    
    Set this.Stack = New Collection

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Enumerator
'@Description "Enumerate stack from top to bottom."
Public Property Get Enumerate() As IEnumVARIANT
Attribute Enumerate.VB_Description = "Enumerate stack from top to bottom."
Attribute Enumerate.VB_UserMemId = -4
' Called by For Each by setting Attribute Enumerate.VB_UserMemId = -4.
' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
'
' Don't mutate the Stack object during the For Each loop. This will result in
' unpredictable behavior which is reflecting VB's native semantics.
    
    Set Enumerate = this.Stack.[_NewEnum]

End Property


'@Description "Removes all items from the stack."
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all items from the stack."

    Set this.Stack = New Collection

End Sub


'@Description "Returns the number of items contained in the stack."
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items contained in the stack."
    
    Count = this.Stack.Count

End Property


'@Description "Returns True if the stack is empty or False if not."
Public Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Returns True if the stack is empty or False if not."
    
    IsEmpty = (this.Stack.Count = 0)

End Property


'@DefaultMember
'@Description "Returns the item at the top of the stack."
Public Property Get Peek() As Variant
Attribute Peek.VB_Description = "Returns the item at the top of the stack."
Attribute Peek.VB_UserMemId = 0
    
    If this.Stack.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, TypeName(Me) & ".Peek", "Stack is empty"
    AssignVariant Peek, this.Stack.Item(Top)

End Property


'@Description "Returns and removes the item at the top of the stack."
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Returns and removes the item at the top of the stack."
    
    If this.Stack.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, TypeName(Me) & ".Pop", "Stack is empty"
    AssignVariant Pop, this.Stack.Item(Top)
    this.Stack.Remove Top

End Function


'@Description "Adds an item at the top of the stack."
Public Sub Push(ByVal Item As Variant)
Attribute Push.VB_Description = "Adds an item at the top of the stack."
    
    If this.Stack.Count = 0 Then
        this.Stack.Add Item
    Else
        this.Stack.Add Item, before:=Top
    End If

End Sub


'@Description "Exports all stack items to an array."
Public Function Items(Optional ByVal base As Long) As Variant()
Attribute Items.VB_Description = "Exports all stack items to an array."

    If this.Stack.Count = 0 Then
        Items = Array()
        Exit Function
    End If
    
    Dim arr() As Variant: ReDim arr(base To base + this.Stack.Count - 1)
    Dim i As Long: i = base
    Dim var As Variant
    For Each var In this.Stack
        If VBA.IsObject(var) Then
            Set arr(i) = var
        Else
            arr(i) = var
        End If
        i = i + 1
    Next
    
    Items = arr

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Assigns a value or an object reference to a variable."
'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub AssignVariant(ByRef variable As Variant, ByVal value As Variant)
Attribute AssignVariant.VB_Description = "Assigns a value or an object reference to a variable."
' Prevents getting a Collection item twice at the expense of a stack frame.
    
    If VBA.IsObject(value) Then
        Set variable = value
    Else
        variable = value
    End If

End Sub
