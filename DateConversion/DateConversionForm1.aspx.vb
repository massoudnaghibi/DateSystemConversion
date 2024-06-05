Public Class DateConversionForm1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim stTodayDate As String = ""
        stTodayDate = stDateDesc(stDateConvert(Today.ToString("yyyy/MM/dd"), DateConvType2.Julian_Persian), "P", "WDMY")
        stTodayDate = stTodayDate & stDateDesc(stDateConvert(Today.ToString("yyyy/MM/dd"), DateConvType2.Julian_Islamic), "I", "DMY")
        stTodayDate = stTodayDate & stDateDesc(Today.ToString("yyyy/MM/dd"), "J", "DMY")
        TodayDateLabel.Text = stTodayDate
    End Sub

    Protected Sub DateConvButton_Click(sender As Object, e As EventArgs) Handles DateConvButton.Click
        OutDateLabel.Text = stDateConvert(InDateTextBox.Text.ToString, RadioButtonList1.SelectedItem.Value)
        TodayDateLabel.Text = Today.ToString("yyyy/MM/dd")
    End Sub

    Protected Sub DateDescButton_Click(sender As Object, e As EventArgs) Handles DateDescButton.Click
        TodayDateLabel.Text = InDateTextBox.Text
        OutDateLabel.Text = stDateDesc(InDateTextBox.Text.ToString, RadioButtonList2.SelectedItem.Value, "WDMY")
    End Sub
End Class