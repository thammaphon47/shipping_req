Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.IO
Imports Oracle.DataAccess.Client
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Partial Class About
    Inherits System.Web.UI.Page
    Public Class ShippingItem
        Public Property ItemNo As String
        Public Property ItemName As String
        Public Property Qty As Decimal
        Public Property Unit As String
        Public Property UnitPrice As Decimal
        Public Property Amount As Decimal
        Public Property BoxSize As String
        Public Property Curr As String
    End Class


    Dim ConStr As String = "Data Source=192.168.1.7;Initial Catalog=CS_DATA;User Id=sa;Password=p@ssw0rd;"
    Dim oradb As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.6)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=MCF)));User Id=MCF380;Password=MCF380;"
    Dim oraconn As New OracleConnection(oradb)

    Private Function CreateShippingItemsTable() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("INV_D_MARK", GetType(String))
        dt.Columns.Add("DESCRIPTION", GetType(String))
        dt.Columns.Add("QTY", GetType(Integer))
        dt.Columns.Add("UNIT", GetType(String))
        dt.Columns.Add("CURRENCY", GetType(String))
        dt.Columns.Add("UNIT_PRICE", GetType(Decimal))
        dt.Columns.Add("AMOUNT", GetType(Decimal))
        Return dt
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Try

                LoadCompanyList()
                txtToCompany.Text = ""
                txtAddress.Text = ""
            Catch ex As Exception
                lblMessage.Text = "เกิดข้อผิดพลาด: " & ex.Message
                lblMessage.Visible = True
            Finally
                If oraconn.State = ConnectionState.Open Then oraconn.Close()
            End Try
        End If

        If Session("Username") Is Nothing Then
            Response.Redirect("Home.aspx")
        End If
    End Sub

    Private Sub LoadCompanyList()
        Dim Myselect As String = "SELECT COMPANY_CD, OFCL_NM, ADDR1 || ',' || ZIP_CD AS ADDR1 FROM CM_BP_ALL WHERE COMPANY_CD LIKE 'C%' AND ADDR1 IS NOT NULL ORDER BY OFCL_NM ASC"
        Dim MyDataSet As New DataSet()

        Try
            Using cmd As New OracleCommand(Myselect, oraconn)
                Using adapter As New OracleDataAdapter(cmd)
                    oraconn.Open()
                    adapter.Fill(MyDataSet)
                End Using
            End Using

            ddlCompanyList.DataSource = MyDataSet.Tables(0)
            ddlCompanyList.DataTextField = "OFCL_NM"
            ddlCompanyList.DataValueField = "COMPANY_CD"
            ddlCompanyList.DataBind()

            ddlCompanyList.Items.Insert(0, New ListItem("-- Select Company --", ""))
        Catch ex As Exception
            Throw
        Finally
            If oraconn.State = ConnectionState.Open Then oraconn.Close()
        End Try
    End Sub

    Protected Sub ddlCompanyList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCompanyList.SelectedIndexChanged
        Try
            Dim selectedCompanyCode As String = ddlCompanyList.SelectedValue
            If String.IsNullOrEmpty(selectedCompanyCode) Then
                txtToCompany.Text = ""
                txtAddress.Text = ""
                Return
            End If

            Dim query As String = "SELECT COMPANY_CD, OFCL_NM, ADDR1 || ',' || ZIP_CD AS ADDR1 FROM CM_BP_ALL WHERE COMPANY_CD = :CompanyCode"
            Using cmd As New OracleCommand(query, oraconn)
                cmd.Parameters.Add(":CompanyCode", OracleDbType.Varchar2).Value = selectedCompanyCode
                oraconn.Open()

                Using reader As OracleDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        txtToCompany.Text = reader("COMPANY_CD").ToString()
                        txtAddress.Text = reader("ADDR1").ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            lblMessage.Text = "เกิดข้อผิดพลาด: " & ex.Message
            lblMessage.Visible = True
        Finally
            If oraconn.State = ConnectionState.Open Then oraconn.Close()
        End Try
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim SP_NO As String = txtShippingNo.Text.Trim()
        If SP_NO = "" Then Exit Sub

        Dim dt As New DataTable()
        Dim sql As String = "SELECT * FROM [CS_DATA].[dbo].[q_SHIPP_REQ] WHERE SPP_NO = @SPP_NO"

        Try
            Using conn As New SqlConnection(ConStr)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@SPP_NO", SP_NO)
                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using

            If dt.Rows.Count > 0 Then
                Dim row = dt.Rows(0)

                txtShippingNo.Text = row("SPP_NO").ToString()
                txtDate.Text = Convert.ToDateTime(row("SPP_DATE")).ToString("yyyy-MM-dd")
                txtShippingReqDate.Text = Convert.ToDateTime(row("SPP_REQ_DATE")).ToString("yyyy-MM-dd")
                txtAttention.Text = row("ATTENTION").ToString()
                txtToCompany.Text = row("CS_NO").ToString()

                Dim companyCode As String = row("TO_COMPANY").ToString()
                If ddlCompanyList.Items.FindByValue(companyCode) IsNot Nothing Then
                    ddlCompanyList.SelectedValue = companyCode
                End If

                txtAddress.Text = row("ADDRESS").ToString()
                ddlDeliveryBy.SelectedValue = row("DELIVERY_TYPE").ToString()
                ddlValue.SelectedValue = row("VALUE_TYPE").ToString()
                ddlPaidBy.SelectedValue = row("PAID_TYPE").ToString()
                txtRecipientAC.Text = row("AC_NO1").ToString()
                txtThirdPartyAC.Text = row("AC_NO2").ToString()
                txtTotalAmount.Text = row("TOTAL_AMOUNT").ToString()

                LoadCompanyNameFromOracle(companyCode)

                ' Load detail rows
                Dim dtDetail As New DataTable()
                dtDetail.Columns.Add("ID", GetType(String))
                dtDetail.Columns.Add("MARK_NO", GetType(String))
                dtDetail.Columns.Add("DESCRIPTION", GetType(String))
                dtDetail.Columns.Add("QTY", GetType(String))
                dtDetail.Columns.Add("UNIT", GetType(String))
                dtDetail.Columns.Add("CURRENCY", GetType(String))
                dtDetail.Columns.Add("UNIT_PRICE", GetType(Decimal))
                dtDetail.Columns.Add("AMOUNT", GetType(Decimal))

                Dim totalNW As Double = 0
                Dim totalGW As Double = 0

                For Each dr As DataRow In dt.Rows
                    dtDetail.Rows.Add(
                        Guid.NewGuid().ToString(),
                        "",
                        dr("ITEM_NO").ToString() & " " & dr("ITEM_NAME").ToString(),
                        dr("QTY_NUMBER").ToString(),
                        dr("QTY_UNIT").ToString(),
                        dr("PRICE_CURR").ToString(),
                        Convert.ToDecimal(dr("PRICE_UNIT_PX")),
                        Convert.ToDecimal(dr("PRICE_AMOUNT"))
                    )
                    totalNW += Convert.ToDouble(dr("NET_WT"))
                    totalGW += Convert.ToDouble(dr("GW"))
                Next

                ViewState("ShippingItems") = dtDetail
                gvShippingItems.DataSource = dtDetail
                gvShippingItems.DataBind()

                txtTotalNW.Text = totalNW.ToString("N2")
                txtTotalGW.Text = totalGW.ToString("N2")
                lblMessage.Visible = False
            Else
                lblMessage.Text = "Shipping No not found."
                lblMessage.Visible = True
            End If

        Catch ex As Exception
            lblMessage.Text = "Error: " & ex.Message
            lblMessage.Visible = True
        End Try
    End Sub

    Private Sub LoadCompanyNameFromOracle(companyCode As String)
        If String.IsNullOrEmpty(companyCode) Then Return

        Dim query As String = "SELECT OFCL_NM FROM CM_BP_ALL WHERE COMPANY_CD = :CompanyCode"
        Try
            Using cmd As New OracleCommand(query, oraconn)
                cmd.Parameters.Add(":CompanyCode", OracleDbType.Varchar2).Value = companyCode
                oraconn.Open()
                Using reader As OracleDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        ddlCompanyList.SelectedItem.Text = reader("OFCL_NM").ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            lblMessage.Text = "Error loading company name: " & ex.Message
            lblMessage.Visible = True
        Finally
            If oraconn.State = ConnectionState.Open Then oraconn.Close()
        End Try
    End Sub

    Protected Sub ddlPaidBy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPaidBy.SelectedIndexChanged
        txtRecipientAC.Enabled = (ddlPaidBy.SelectedValue = "Recipient")
        txtThirdPartyAC.Enabled = (ddlPaidBy.SelectedValue = "Third Party")
    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Session.Clear()
        Session.Abandon()
        Response.Redirect("Home.aspx")
    End Sub

    Protected Sub btnInv_Click(sender As Object, e As EventArgs) Handles btnInv.Click
        Dim INV_NO As String = txtInvoiceNo.Text.Trim()
        If INV_NO = "" Then Return

        Using conn As New SqlConnection(ConStr)
            Dim cmd As New SqlCommand("SELECT * FROM [CS_DATA].[dbo].[q_INVOICE] WHERE INV_NO = @INV_NO ORDER BY INV_D_NO", conn)
            cmd.Parameters.AddWithValue("@INV_NO", INV_NO)
            Dim dt As New DataTable()
            Dim adapter As New SqlDataAdapter(cmd)
            adapter.Fill(dt)

            If dt.Rows.Count > 0 Then
                ViewState("ShippingItems") = dt
                gvShippingItems.DataSource = dt
                gvShippingItems.DataBind()
            End If
        End Using
    End Sub

Protected Sub gvShippingItems_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles gvShippingItems.RowEditing
        gvShippingItems.EditIndex = e.NewEditIndex
        BindShippingItemsGrid()
    End Sub

  Protected Sub gvShippingItems_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) Handles gvShippingItems.RowUpdating
        Dim dt As DataTable = TryCast(ViewState("ShippingItems"), DataTable)
        If dt IsNot Nothing Then
            Dim index As Integer = e.RowIndex
            Dim row As GridViewRow = gvShippingItems.Rows(index)

            ' ดึงค่าใหม่จาก TextBox ใน GridView
            Dim txtMarkNo As TextBox = CType(row.FindControl("txtMarkNo"), TextBox)
            Dim txtDesc As TextBox = CType(row.FindControl("txtDesc"), TextBox)
            Dim txtQty As TextBox = CType(row.FindControl("txtQty"), TextBox)
            Dim txtUnit As TextBox = CType(row.FindControl("txtUnit"), TextBox)
            Dim txtCurrency As TextBox = CType(row.FindControl("txtCurrency"), TextBox)
            Dim txtUnitPrice As TextBox = CType(row.FindControl("txtUnitPrice"), TextBox)
            Dim txtAmount As TextBox = CType(row.FindControl("txtAmount"), TextBox)

            ' อัพเดตค่าใน DataTable
            dt.Rows(index)("MARK_NO") = txtMarkNo.Text
            dt.Rows(index)("DESCRIPTION") = txtDesc.Text
            dt.Rows(index)("QTY") = Convert.ToDecimal(txtQty.Text)
            dt.Rows(index)("UNIT") = txtUnit.Text
            dt.Rows(index)("CURRENCY") = txtCurrency.Text
            dt.Rows(index)("UNIT_PRICE") = Convert.ToDecimal(txtUnitPrice.Text)
            dt.Rows(index)("AMOUNT") = Convert.ToDecimal(txtAmount.Text)

            ' บันทึกกลับ ViewState
            ViewState("ShippingItems") = dt

            ' ออกจากโหมดแก้ไข
            gvShippingItems.EditIndex = -1

            ' รีบิ๊น GridView และอัพเดตยอดรวม
            BindShippingItemsGrid()
            UpdateTotals()
        End If
    End Sub


    Protected Sub gvShippingItems_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs)
        gvShippingItems.EditIndex = -1
        gvShippingItems.DataSource = ViewState("ShippingItems")
        gvShippingItems.DataBind()
    End Sub
    Protected Sub gvShippingItems_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles gvShippingItems.RowCommand
        If e.CommandName = "Delete" Then
            Dim idToDelete As String = e.CommandArgument.ToString()
            Dim dt As DataTable = TryCast(ViewState("ShippingItems"), DataTable)
            If dt IsNot Nothing Then
                Dim rows() As DataRow = dt.Select("ID = '" & idToDelete & "'")
                If rows.Length > 0 Then
                    dt.Rows.Remove(rows(0))
                    dt.AcceptChanges()
                    ViewState("ShippingItems") = dt
                    BindShippingItemsGrid()
                    UpdateTotals()
                End If
            End If
        End If
    End Sub

    Private Sub UpdateTotals()
        Dim dt As DataTable = TryCast(ViewState("ShippingItems"), DataTable)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            txtTotalCartons.Text = "0"
            txtTotalNW.Text = "0"
            txtTotalGW.Text = "0"
            txtTotalAmount.Text = "0"
            Return
        End If

        Dim totalCartons As Integer = dt.Rows.Count
        Dim totalAmount As Decimal = dt.AsEnumerable().Sum(Function(dr) Convert.ToDecimal(dr("AMOUNT")))

        txtTotalCartons.Text = totalCartons.ToString()
        txtTotalAmount.Text = totalAmount.ToString("N2")
    End Sub

    Private Sub BindShippingItemsGrid()
        Dim dt As DataTable = TryCast(ViewState("ShippingItems"), DataTable)
        gvShippingItems.DataSource = dt
        gvShippingItems.DataBind()
    End Sub


    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            ' ตรวจสอบ controls ต่างๆ ว่าไม่เป็น Nothing
            If txtInvoiceNo Is Nothing Then
                ShowMessage("txtInvoiceNo is not initialized.")
                Return
            End If

            If txtDate Is Nothing OrElse String.IsNullOrWhiteSpace(txtDate.Text) Then
                ShowMessage("Please enter Invoice Date.")
                Return
            End If

            If txtSailingDate Is Nothing OrElse String.IsNullOrWhiteSpace(txtSailingDate.Text) Then
                ShowMessage("Please enter Sailing Date.")
                Return
            End If

            If ddlCompanyList Is Nothing OrElse String.IsNullOrWhiteSpace(ddlCompanyList.SelectedValue) Then
                ShowMessage("Please select Company.")
                Return
            End If

            If ddlValue Is Nothing Then
                ShowMessage("Value dropdown not initialized.")
                Return
            End If

            If ddlDeliveryBy Is Nothing Then
                ShowMessage("Delivery dropdown not initialized.")
                Return
            End If

            If txtShippingNo Is Nothing OrElse String.IsNullOrWhiteSpace(txtShippingNo.Text) Then
                ShowMessage("Please enter Shipping No.")
                Return
            End If

            If gvShippingItems Is Nothing OrElse gvShippingItems.Rows.Count = 0 Then
                ShowMessage("Shipping items are missing.")
                Return
            End If

            If ddlReport Is Nothing OrElse String.IsNullOrEmpty(ddlReport.SelectedValue) Then
                ShowMessage("Please select a report type.")
                Return
            End If

            ' สร้างเลข INV_NO อัตโนมัติถ้าไม่มีการกรอก
            If String.IsNullOrWhiteSpace(txtInvoiceNo.Text) Then
                txtInvoiceNo.Text = GenerateInvoiceNo()
            End If

            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConStr").ConnectionString)
                conn.Open()
                Using trans = conn.BeginTransaction()
                    Try
                        If InvoiceExists(txtInvoiceNo.Text, conn, trans) Then
                            UpdateInvoiceHeader(conn, trans)
                            DeleteInvoiceDetails(conn, trans)
                            InsertInvoiceDetails(conn, trans)
                            UpdateShipping(conn, trans)
                            ShowMessage("Update Invoice Successfully!")
                        Else
                            InsertInvoiceHeader(conn, trans)
                            InsertInvoiceDetails(conn, trans)
                            UpdateShipping(conn, trans)
                            ShowMessage("Save New Invoice Successfully!")
                        End If
                        trans.Commit()
                    Catch ex As Exception
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            RedirectToReportAsPDF()

        Catch ex As Exception
            ShowMessage("Error: " & ex.Message.Replace("'", "\'"))
        End Try
    End Sub



    Private Function InvoiceExists(invNo As String, conn As SqlConnection, trans As SqlTransaction) As Boolean
        Using cmd As New SqlCommand("SELECT 1 FROM [CS_DATA].[dbo].[INVOICE_H] WHERE INV_NO = @INV_NO", conn, trans)
            cmd.Parameters.AddWithValue("@INV_NO", invNo)
            Return cmd.ExecuteScalar() IsNot Nothing
        End Using
    End Function

    Private Sub UpdateInvoiceHeader(conn As SqlConnection, trans As SqlTransaction)
        Dim query As String = "UPDATE [CS_DATA].[dbo].[INVOICE_H] SET " & _
            "INV_DATE = @INV_DATE, INV_TO = @INV_TO, INV_ADDRESS = @INV_ADDRESS, " & _
            "INV_ATTENTION = @INV_ATTENTION, INV_VALUES = @INV_VALUES, INV_DELIVERY = @INV_DELIVERY, " & _
            "INV_SAILING_DT = @INV_SAILING_DT, INV_LOCATION_FROM = @INV_LOCATION_FROM, " & _
            "INV_LOCATION_TO = @INV_LOCATION_TO, INV_FREIGHT = @INV_FREIGHT, INV_AWB = @INV_AWB, " & _
            "INV_REMARK = @INV_REMARK, INV_TOTAL_CARTONS = @INV_TOTAL_CARTONS, " & _
            "INV_TOTAL_NW = @INV_TOTAL_NW, INV_TOTAL_GW = @INV_TOTAL_GW, INV_TOTAL_AMOUNT = @INV_TOTAL_AMOUNT, " & _
            "INV_SPP_NO = @INV_SPP_NO, INV_MEASUREMENT = @INV_MEASUREMENT, INV_NOTIFY_PARTY = @INV_NOTIFY_PARTY, " & _
            "INV_TERM = @INV_TERM, INV_CS_CODE = @INV_CS_CODE " & _
            "WHERE INV_NO = @INV_NO"
        Using cmd As New SqlCommand(query, conn, trans)
            AddInvoiceParameters(cmd)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub InsertInvoiceHeader(conn As SqlConnection, trans As SqlTransaction)
        Dim query As String = "INSERT INTO [CS_DATA].[dbo].[INVOICE_H] VALUES " & _
            "(@INV_NO, @INV_DATE, @INV_TO, @INV_ADDRESS, @INV_ATTENTION, @INV_VALUES, " & _
            "@INV_DELIVERY, @INV_SAILING_DT, @INV_LOCATION_FROM, @INV_LOCATION_TO, " & _
            "@INV_FREIGHT, @INV_AWB, @INV_REMARK, @INV_TOTAL_CARTONS, @INV_TOTAL_NW, " & _
            "@INV_TOTAL_GW, @INV_TOTAL_AMOUNT, @INV_SPP_NO, @INV_MEASUREMENT, " & _
            "@INV_NOTIFY_PARTY, @INV_TERM, @INV_CS_CODE)"
        Using cmd As New SqlCommand(query, conn, trans)
            AddInvoiceParameters(cmd)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub AddInvoiceParameters(cmd As SqlCommand)
        ' ตรวจสอบและแปลงวันที่อย่างปลอดภัย
        Dim invDate As Object = DBNull.Value
        Dim dt As DateTime
        If txtDate IsNot Nothing AndAlso DateTime.TryParse(txtDate.Text, dt) Then
            invDate = dt
        End If

        Dim sailingDate As Object = DBNull.Value
        If txtSailingDate IsNot Nothing AndAlso DateTime.TryParse(txtSailingDate.Text, dt) Then
            sailingDate = dt
        End If

        ' ตรวจสอบ DropDownList และ TextBox ก่อนใช้
        Dim invTo As String = If(ddlCompanyList IsNot Nothing AndAlso ddlCompanyList.SelectedIndex >= 0, ddlCompanyList.SelectedValue, "")
        Dim invValues As String = If(ddlValue IsNot Nothing AndAlso ddlValue.SelectedIndex >= 0, ddlValue.SelectedValue, "")
        Dim invDelivery As String = If(ddlDeliveryBy IsNot Nothing AndAlso ddlDeliveryBy.SelectedIndex >= 0, ddlDeliveryBy.SelectedValue, "")
        Dim csCode As String = If(txtTo IsNot Nothing, txtTo.Text, "")

        cmd.Parameters.AddWithValue("@INV_NO", If(txtInvoiceNo IsNot Nothing, txtInvoiceNo.Text, ""))
        cmd.Parameters.AddWithValue("@INV_DATE", invDate)
        cmd.Parameters.AddWithValue("@INV_SAILING_DT", sailingDate)
        cmd.Parameters.AddWithValue("@INV_TO", invTo)
        cmd.Parameters.AddWithValue("@INV_ADDRESS", If(txtAddress IsNot Nothing, txtAddress.Text, ""))
        cmd.Parameters.AddWithValue("@INV_ATTENTION", If(txtAttention IsNot Nothing, txtAttention.Text, ""))
        cmd.Parameters.AddWithValue("@INV_VALUES", invValues)
        cmd.Parameters.AddWithValue("@INV_DELIVERY", invDelivery)
        cmd.Parameters.AddWithValue("@INV_LOCATION_FROM", If(txtFrom IsNot Nothing, txtFrom.Text, ""))
        cmd.Parameters.AddWithValue("@INV_LOCATION_TO", If(txtTo IsNot Nothing, txtTo.Text, ""))
        cmd.Parameters.AddWithValue("@INV_FREIGHT", If(txtFreight IsNot Nothing, txtFreight.Text, ""))
        cmd.Parameters.AddWithValue("@INV_AWB", If(txtAWB IsNot Nothing, txtAWB.Text, ""))
        cmd.Parameters.AddWithValue("@INV_REMARK", If(txtRemark IsNot Nothing, txtRemark.Text, ""))
        cmd.Parameters.AddWithValue("@INV_TOTAL_CARTONS", ParseDouble(If(txtTotalCartons IsNot Nothing, txtTotalCartons.Text, "")))
        cmd.Parameters.AddWithValue("@INV_TOTAL_NW", ParseDouble(If(txtTotalNW IsNot Nothing, txtTotalNW.Text, "")))
        cmd.Parameters.AddWithValue("@INV_TOTAL_GW", ParseDouble(If(txtTotalGW IsNot Nothing, txtTotalGW.Text, "")))
        cmd.Parameters.AddWithValue("@INV_TOTAL_AMOUNT", ParseDouble(If(txtTotalAmount IsNot Nothing, txtTotalAmount.Text, "")))
        cmd.Parameters.AddWithValue("@INV_SPP_NO", If(txtShippingNo IsNot Nothing, txtShippingNo.Text, ""))
        cmd.Parameters.AddWithValue("@INV_MEASUREMENT", If(txtMeasurement IsNot Nothing, txtMeasurement.Text, ""))
        cmd.Parameters.AddWithValue("@INV_NOTIFY_PARTY", If(txtNotifyParty IsNot Nothing, txtNotifyParty.Text, ""))
        cmd.Parameters.AddWithValue("@INV_TERM", If(txtTerm IsNot Nothing, txtTerm.Text, ""))
        cmd.Parameters.AddWithValue("@INV_CS_CODE", csCode)
    End Sub

    Private Function ParseDouble(text As String) As Object
        Dim result As Double
        If Double.TryParse(text, result) Then
            Return result
        Else
            Return DBNull.Value
        End If
    End Function

    Private Sub DeleteInvoiceDetails(conn As SqlConnection, trans As SqlTransaction)
        Using cmd As New SqlCommand("DELETE FROM [CS_DATA].[dbo].[INVOICE_D] WHERE INV_NO = @INV_NO", conn, trans)
            cmd.Parameters.AddWithValue("@INV_NO", If(txtInvoiceNo IsNot Nothing, txtInvoiceNo.Text, ""))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub InsertInvoiceDetails(conn As SqlConnection, trans As SqlTransaction)
        If gvShippingItems Is Nothing OrElse gvShippingItems.Rows.Count = 0 Then
            Throw New Exception("Shipping items are missing.")
        End If

        For i As Integer = 0 To gvShippingItems.Rows.Count - 1
            Dim row = gvShippingItems.Rows(i)
            Dim query As String = "INSERT INTO [CS_DATA].[dbo].[INVOICE_D] " & _
                "VALUES (@DetailNo, @INV_NO, @Col1, @Col2, @Col3, @Col4, @Col5, @Col6, @Col7, @Col8)"
            Using cmd As New SqlCommand(query, conn, trans)
                cmd.Parameters.AddWithValue("@DetailNo", txtInvoiceNo.Text & i.ToString("0000"))
                cmd.Parameters.AddWithValue("@INV_NO", txtInvoiceNo.Text)
                For col As Integer = 0 To 7
                    Dim val As String = Server.HtmlDecode(row.Cells(col).Text).Trim()
                    cmd.Parameters.AddWithValue("@Col" & (col + 1).ToString(), val)
                Next
                cmd.ExecuteNonQuery()
            End Using
        Next
    End Sub

    Private Sub UpdateShipping(conn As SqlConnection, trans As SqlTransaction)
        Dim query As String = "UPDATE [CS_DATA].[dbo].[SPPING_H] " & _
                              "SET INVOICE_NO = @INV_NO, AWB_NO = @AWB_NO " & _
                              "WHERE SPP_NO = @SPP_NO"
        Using cmd As New SqlCommand(query, conn, trans)
            cmd.Parameters.AddWithValue("@INV_NO", If(txtInvoiceNo IsNot Nothing, txtInvoiceNo.Text, ""))
            cmd.Parameters.AddWithValue("@AWB_NO", If(txtAWB IsNot Nothing, txtAWB.Text, ""))
            cmd.Parameters.AddWithValue("@SPP_NO", If(txtShippingNo IsNot Nothing, txtShippingNo.Text, ""))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub ShowMessage(msg As String)
        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('" & msg.Replace("'", "\'") & "');", True)
    End Sub

    Private Sub RedirectToReportAsPDF()
        Try
            If ddlReport Is Nothing OrElse String.IsNullOrEmpty(ddlReport.SelectedValue) Then
                ShowMessage("Please select a report type.")
                Return
            End If

            Dim report As New ReportDocument()
            Dim reportPath As String = Server.MapPath("~/Reports/")
            Dim reportFile As String = ""

            Select Case ddlReport.SelectedValue
                Case "TKR"
                    reportFile = "TKRReport.rpt"
                Case "INTECH"
                    reportFile = "INTECHReport.rpt"
                Case "CHICAGO"
                    reportFile = "ChicagoReport.rpt"
                Case Else
                    ShowMessage("Invalid report type selected.")
                    Return
            End Select

            report.Load(reportPath & reportFile)
            report.SetParameterValue("InvoiceNo", If(txtInvoiceNo IsNot Nothing, txtInvoiceNo.Text, ""))

            Using stream As Stream = report.ExportToStream(ExportFormatType.PortableDocFormat)
                Using ms As New MemoryStream()
                    stream.CopyTo(ms)
                    Dim bytes As Byte() = ms.ToArray()

                    Response.Clear()
                    Response.ContentType = "application/pdf"
                    Response.AddHeader("content-disposition", "inline;filename=Invoice_" & txtInvoiceNo.Text & ".pdf")
                    Response.BinaryWrite(bytes)
                    Response.End()
                End Using
            End Using
        Catch ex As Exception
            ShowMessage("Error generating report: " & ex.Message.Replace("'", "\'"))
        End Try
    End Sub

    Private Function GenerateInvoiceNo() As String
        Dim today As String = DateTime.Now.ToString("yyyyMMdd")
        Dim prefix As String = "INV" & today
        Dim lastNo As Integer = 0

        Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConStr").ConnectionString)
            conn.Open()
            Dim query As String = "SELECT MAX(INV_NO) FROM [CS_DATA].[dbo].[INVOICE_H] WHERE INV_NO LIKE @Prefix + '%'"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@Prefix", prefix)
                Dim result = cmd.ExecuteScalar()
                If result IsNot DBNull.Value AndAlso result IsNot Nothing Then
                    Dim lastInv As String = result.ToString()
                    Integer.TryParse(lastInv.Substring(prefix.Length), lastNo)
                End If
            End Using
        End Using

        Dim newNo As String = prefix & (lastNo + 1).ToString("D4")
        Return newNo
    End Function


End Class
