<div align="center">

## TreeView \- Check parent nodes when all child nodes are checked


</div>

### Description

Will check the parent node of a Treeview control when all of its child nodes are checked. will uncheck parent node when any one of it's child nodes are unchecked
 
### More Info
 
Assumes: Treeview control named TreeView1

none that I am aware of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Gillham](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-gillham.md)
**Level**          |Advanced
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-gillham-treeview-check-parent-nodes-when-all-child-nodes-are-checked__1-48560/archive/master.zip)





### Source Code

```
'==================================='
' This code has been provided for  '
' use as FREEWARE - you may edit  '
' or change any part of it for your '
' own needs. But please give credit'
' to myself, and do not sell this  '
' unmodified code          '
'    By Steve Gillham      '
'==================================='
Private Sub TreeView1_NodeCheck(ByVal node As MSComctlLib.node)
    Call CheckChild(TreeView1, node)  'perform check on child nodes
End Sub
Private Sub CheckChild(Tree As TreeView, CurrentNode As node)
  Dim ParentIndex   'used to find out the index of the parent node from the child node that was clicked
  Dim CheckChecked As Integer   'Used to decide whether or not parent is to be checked
  Dim j As Integer  'Counter
  'This code works by finding the parent node of the node
  'that you clicked on and then looking to see if all of
  'the other same level nodes (as the one you clicked)
  'are checked. If so it then checks the parent node and Vice Versa
  If CurrentNode.Checked = True Then 'If node is checked then check parent node ONLY if ALL child nodes are checked
    If Tree.Nodes.Item(CurrentNode.Index).Checked = True Then  'If "My Node"(My Node's Index) is checked then
        ParentIndex = Tree.Nodes.Item(CurrentNode.Index).Parent.Index  'locate index of parent node
      For j = 1 To Tree.Nodes.Item(ParentIndex).Children         'run loop to find out which child nodes
        If Tree.Nodes.Item(ParentIndex + j).Checked = False Then    'are checked and which are not...
          Me.BackColor = vbRed                    'Store value of check/uncheck as 0 for
          CheckChecked = CheckChecked + 0               'unchecked and 1 for checked, add to
        ElseIf Tree.Nodes.Item(ParentIndex + j).Checked = True Then   'previous value
          Me.BackColor = vbGreen
          CheckChecked = CheckChecked + 1
        End If
      Next j
      If CheckChecked = Tree.Nodes.Item(ParentIndex).Children Then    'if the number of checked nodes is equal
        Tree.Nodes.Item(ParentIndex).Checked = True           'to number of child nodes then all child
      Else                                'nodes are checked so check parent node
        Tree.Nodes.Item(ParentIndex).Checked = False
      End If
    End If
  ElseIf CurrentNode.Checked = False Then                     'if the current node is unchecked then
        ParentIndex = Tree.Nodes.Item(CurrentNode.Index).Parent.Index    'uncheck the parent node as ALL child nodes
      Tree.Nodes.Item(ParentIndex).Checked = False              'must be checked to before the parent node is checked
  End If
End Sub
```

