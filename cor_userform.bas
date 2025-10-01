Option Explicit

Private Sub UserForm_Initialize()
    UpdateUI
End Sub

Private Sub cmdCopy_Click()
    If G_BatchIndex < 1 Or G_BatchIndex > G_BatchCount Then Exit Sub
    CopyTextToClipboard txtBatch.Text
End Sub

Private Sub cmdPrev_Click()
    If G_BatchIndex > 1 Then
        G_BatchIndex = G_BatchIndex - 1
        UpdateUI
    End If
End Sub

Private Sub cmdNext_Click()
    If G_BatchIndex < G_BatchCount Then
        G_BatchIndex = G_BatchIndex + 1
        UpdateUI
    End If
End Sub

Private Sub cmdDone_Click()
    ' Close form and continue pipeline
    Me.Hide
    G_WorkflowReadyToContinue = True
    Continue_After_Batches
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub UpdateUI()
    If G_BatchCount = 0 Then
        lblCounter.Caption = "No batches."
        txtBatch.Text = ""
        cmdCopy.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdDone.Enabled = False
        Exit Sub
    End If

    lblCounter.Caption = "Batch " & G_BatchIndex & " of " & G_BatchCount & _
                         "  |  Paste into Credit Studio, export, SAVE the opened file(s), then proceed."

    txtBatch.Text = G_Batches(G_BatchIndex)

    cmdCopy.Enabled = True
    cmdPrev.Enabled = (G_BatchIndex > 1)
    cmdNext.Enabled = (G_BatchIndex < G_BatchCount)
    cmdDone.Enabled = (G_BatchIndex = G_BatchCount)

    ' Convenience: auto-copy on batch navigation
    CopyTextToClipboard txtBatch.Text
End Sub