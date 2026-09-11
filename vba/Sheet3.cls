Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' the SEARCH row drives the MATCHED LOTS strip at the top: open it while a
    ' box is filled, close it again the moment the row is cleared
    BuyerGroupStrip.OnSearchRowChange Target
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' double-click the MATCHED LOTS bar (row 6) to lift the matching buyers'
    ' blocks to the top of the table - double-click it again to put them back
    BuyerGroupStrip.OnBarDoubleClick Target, Cancel
End Sub

Private Sub Worksheet_Activate()
    BuyerGroupStrip.SyncStrip Me
    BuyerGroupStrip.ClearSay
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    BuyerGroupStrip.ClearSay
End Sub
