Public Class kkk
    Inherits System.Web.UI.Page

    Sub Page_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim db_Time As Date = Now() 'CDate("2021/8/24 09:49") '資料庫時間(從資料庫取得) 

        Dim sys_Time As Date = Now()

        Me.lb_db_time.Text = db_Time.ToString("yyyy/MM/dd HH:mm:ss")
        Me.lb_sys_time.Text = sys_Time.ToString("yyyy/MM/dd HH:mm:ss")



        If Not IsPostBack Then
            Dim timeCheck As Integer = DateDiff(DateInterval.Minute, db_Time, sys_Time)

            If timeCheck >= -5 AndAlso timeCheck <= 5 Then 'DB時間與系統時間相差不到5分鐘的話,,,  
                Me.lb_memo.Text = "彈出視窗..."
                '輸出javascript指令 
                'Me.custom_script.Text = "<scr" & "ipt> window.onload = function(){ $('#myModal').modal('show'); } </scr" & "ipt>"
                Me.custom_script.Text = "<script> window.onload = function(){ $('#myModal').modal('show'); } </script>"
            Else
                Me.lb_memo.Text = "不做任何事..."
            End If

        End If

    End Sub
End Class