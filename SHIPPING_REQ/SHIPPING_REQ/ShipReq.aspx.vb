Imports System.Data
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.IO
Imports Oracle.DataAccess.Client
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class _Default
    Inherits System.Web.UI.Page

    Public Class ShippingItem
        Public Property ItemNo As String
        Public Property ItemName As String
        Public Property Qty As Decimal
        Public Property Unit As String
        Public Property Curr As String
        Public Property UnitPrice As Decimal
        Public Property Amount As Decimal
        Public Property BoxSize As String
        Public Property NW As Decimal
        Public Property GW As Decimal
        Public Property BoxCount As String
    End Class

    Dim ConStr As String = "Data Source=192.168.1.7;Initial Catalog=CS_DATA;User Id=sa;Password=p@ssw0rd;"
    Dim sqlConn As New SqlConnection(ConStr)

    Dim oradb As String = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.6)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=MCF)));User Id=MCF380;Password=MCF380;"
    Dim oraconn As New OracleConnection(oradb)

    Dim dt As New DataTable
    Dim rowIndex As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Try
                Dim Myselect As String = "SELECT COMPANY_CD, OFCL_NM, ADDR1 || ',' || ZIP_CD AS ADDR1 " &
                                         "FROM CM_BP_ALL WHERE COMPANY_CD LIKE 'C%' AND ADDR1 IS NOT NULL " &
                                         "ORDER BY OFCL_NM ASC"

                Dim cmd As New OracleCommand(Myselect, oraconn)
                Dim adapter As New OracleDataAdapter(cmd)
                Dim MyDataSet As New DataSet()

                oraconn.Open()
                adapter.Fill(MyDataSet)
                oraconn.Close()

                ddlCompanyList.DataSource = MyDataSet.Tables(0)
                ddlCompanyList.DataTextField = "OFCL_NM"
                ddlCompanyList.DataValueField = "COMPANY_CD"
                ddlCompanyList.DataBind()

                ddlCompanyList.Items.Insert(0, New ListItem("-- Select Company --", ""))

                txtToCompany.Text = ""
                txtAddress.Text = ""
            Catch ex As Exception
                lblMessage.Text = "เกิดข้อผิดพลาด: " & ex.Message
                lblMessage.Visible = True
            Finally
                If oraconn.State = ConnectionState.Open Then
                    oraconn.Close()
                End If
            End Try

            ' เช็คว่า query string มี parameter download หรือไม่
            Dim downloadNo As String = Request.QueryString("download")
            If Not String.IsNullOrEmpty(downloadNo) Then
                ExportPDF(downloadNo)
            End If
        End If

        If Session("Username") Is Nothing Then
            Response.Redirect("Home.aspx")
        End If
    End Sub


    Protected Sub ddlCompanyList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCompanyList.SelectedIndexChanged
        Try
            Dim selectedCompanyCode As String = ddlCompanyList.SelectedValue

            If Not String.IsNullOrEmpty(selectedCompanyCode) Then
                Dim query As String = "SELECT COMPANY_CD, OFCL_NM, ADDR1 || ',' || ZIP_CD AS ADDR1 " &
                                      "FROM CM_BP_ALL WHERE COMPANY_CD = :CompanyCode"

                Dim cmd As New OracleCommand(query, oraconn)
                cmd.Parameters.Add(":CompanyCode", OracleDbType.Varchar2).Value = selectedCompanyCode

                oraconn.Open()
                Dim reader As OracleDataReader = cmd.ExecuteReader()

                If reader.Read() Then
                    txtToCompany.Text = reader("COMPANY_CD").ToString()
                    txtAddress.Text = reader("ADDR1").ToString()
                End If

                reader.Close()
                oraconn.Close()
            Else
                txtToCompany.Text = ""
                txtAddress.Text = ""
            End If
        Catch ex As Exception
            lblMessage.Text = "เกิดข้อผิดพลาด: " & ex.Message
            lblMessage.Visible = True
        Finally
            If oraconn.State = ConnectionState.Open Then
                oraconn.Close()
            End If
        End Try
    End Sub

    ' ฟังก์ชันที่ถูกเรียกเมื่อคลิกปุ่ม Search
    Protected Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim SP_NO As String = txtSearch.Text.Trim()
        Dim dt As New DataTable()

        If String.IsNullOrEmpty(SP_NO) Then
            lblMessage.Text = "Please input Shipping No."
            lblMessage.CssClass = "message-label"
            lblMessage.Visible = True
            Return
        End If

        Using conn As New SqlConnection(ConStr)
            Dim query As String = "SELECT * FROM [CS_DATA].[dbo].[q_SHIPP_REQ] WHERE SPP_NO = @SP_NO ORDER BY SPP_D_NO"
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@SP_NO", SP_NO)
                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using

        If dt.Rows.Count > 0 Then
            ' Header fields
            txtShippingNo.Text = dt.Rows(0)("SPP_NO").ToString()
            txtDate.Text = If(IsDBNull(dt.Rows(0)("SPP_DATE")), "", Convert.ToDateTime(dt.Rows(0)("SPP_DATE")).ToString("yyyy-MM-dd"))
            txtShippingReqDate.Text = If(IsDBNull(dt.Rows(0)("SPP_REQ_DATE")), "", Convert.ToDateTime(dt.Rows(0)("SPP_REQ_DATE")).ToString("yyyy-MM-dd"))
            txtAttention.Text = dt.Rows(0)("ATTENTION").ToString()
            txtToCompany.Text = dt.Rows(0)("CS_NO").ToString()
            ddlCompanyList.SelectedValue = dt.Rows(0)("TO_COMPANY").ToString()
            txtAddress.Text = dt.Rows(0)("ADDRESS").ToString()
            ddlDeliveryBy.SelectedValue = dt.Rows(0)("DELIVERY_TYPE").ToString()
            ddlValue.SelectedValue = dt.Rows(0)("VALUE_TYPE").ToString()
            ddlPaidBy.SelectedValue = dt.Rows(0)("PAID_TYPE").ToString()
            txtRecipientAC.Text = dt.Rows(0)("AC_NO1").ToString()
            txtThirdPartyAC.Text = dt.Rows(0)("AC_NO2").ToString()
            txtTotalAmount.Text = If(IsDBNull(dt.Rows(0)("TOTAL_AMOUNT")), "0", dt.Rows(0)("TOTAL_AMOUNT").ToString())

            ' Prepare detail table
            Dim detailTable As New DataTable()
            detailTable.Columns.AddRange({
                New DataColumn("ItemNo"),
                New DataColumn("ItemName"),
                New DataColumn("Qty", GetType(Decimal)),
                New DataColumn("Unit"),
                New DataColumn("UnitPrice", GetType(Decimal)),
                New DataColumn("Amount", GetType(Decimal)),
                New DataColumn("Curr"),
                New DataColumn("BoxSize"),
                New DataColumn("NW", GetType(Decimal)),
                New DataColumn("GW", GetType(Decimal)),
                New DataColumn("BoxCount")
            })

            For Each row As DataRow In dt.Rows
                detailTable.Rows.Add(
                    row("ITEM_NO"),
                    row("ITEM_NAME"),
                    If(IsDBNull(row("QTY_NUMBER")), 0D, Convert.ToDecimal(row("QTY_NUMBER"))),
                    row("QTY_UNIT"),
                    If(IsDBNull(row("PRICE_UNIT_PX")), 0D, Convert.ToDecimal(row("PRICE_UNIT_PX"))),
                    If(IsDBNull(row("PRICE_AMOUNT")), 0D, Convert.ToDecimal(row("PRICE_AMOUNT"))),
                    row("PRICE_CURR"),
                    row("BOX_SIZE"),
                    If(IsDBNull(row("NET_WT")), 0D, Convert.ToDecimal(row("NET_WT"))),
                    If(IsDBNull(row("GW")), 0D, Convert.ToDecimal(row("GW"))),
                    row("BOX_COUNT")
                )
            Next

            ' Bind to GridView
            gvShippingItems.DataSource = detailTable
            gvShippingItems.DataBind()

            ' Save to Session
            Session("ShippingItems") = detailTable

            lblMessage.Text = "Shipping No. found"
            lblMessage.CssClass = "message-label message-success"
            lblMessage.Visible = True
        Else
            lblMessage.Text = "Shipping No. not found"
            lblMessage.CssClass = "message-label"
            lblMessage.Visible = True
        End If
    End Sub

    Private Function ConvertDataTableToShippingItemList(dt As DataTable) As List(Of ShippingItem)
        Dim list As New List(Of ShippingItem)()

        For Each row As DataRow In dt.Rows
            Dim item As New ShippingItem() With {
                .ItemNo = row("ItemNo").ToString(),
                .ItemName = row("ItemName").ToString(),
                .Qty = If(IsNumeric(row("Qty")), Convert.ToInt32(row("Qty")), 0),
                .Unit = row("Unit").ToString(),
                .Curr = row("Curr").ToString(),
                .UnitPrice = If(IsNumeric(row("UnitPrice")), Convert.ToDouble(row("UnitPrice")), 0),
                .Amount = If(IsNumeric(row("Amount")), Convert.ToDouble(row("Amount")), 0),
                .BoxSize = row("BoxSize").ToString(),
                .NW = If(IsNumeric(row("NW")), Convert.ToDouble(row("NW")), 0),
                .GW = If(IsNumeric(row("GW")), Convert.ToDouble(row("GW")), 0),
                .BoxCount = If(IsNumeric(row("BoxCount")), Convert.ToInt32(row("BoxCount")), 0)
            }
            list.Add(item)
        Next

        Return list
    End Function

    ' btnSave_Click
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If sqlConn.State = ConnectionState.Closed Then
            sqlConn.Open()
        End If

        Dim transaction As SqlTransaction = sqlConn.BeginTransaction()

        Try
            Dim shippingNo As String = txtShippingNo.Text.Trim()
            Dim detailTable As DataTable = TryCast(Session("ShippingItems"), DataTable)

            ' คำนวณยอดรวม
            Dim totalAmount As Decimal = 0
            For Each row As DataRow In detailTable.Rows
                If Not IsDBNull(row("Amount")) Then
                    totalAmount += Convert.ToDecimal(row("Amount"))
                End If
            Next

            ' เตรียมวันที่
            Dim sppDate As Object = DBNull.Value
            If Not String.IsNullOrEmpty(txtDate.Text) Then
                Dim parsedDate As Date
                If Date.TryParse(txtDate.Text, parsedDate) Then
                    sppDate = parsedDate
                End If
            End If

            Dim reqDate As Object = DBNull.Value
            If Not String.IsNullOrEmpty(txtShippingReqDate.Text) Then
                Dim parsedDate As Date
                If Date.TryParse(txtShippingReqDate.Text, parsedDate) Then
                    reqDate = parsedDate
                End If
            End If

            ' Insert หรือ Update Header
            If String.IsNullOrEmpty(shippingNo) Then
                shippingNo = GenerateShippingNo(sqlConn, transaction)
                txtShippingNo.Text = shippingNo

                Dim insertHeader As String = "INSERT INTO [CS_DATA].[dbo].[SPPING_H] " &
                    "(SPP_NO, SPP_DATE, TO_COMPANY, ATTENTION, ADDRESS, AC_NO1, AC_NO2, DELIVERY_TYPE, PAID_TYPE, VALUE_TYPE, TOTAL_AMOUNT, SPP_REQ_DATE, CS_NO) " &
                    "VALUES (@SPP_NO, @SPP_DATE, @TO_COMPANY, @ATTENTION, @ADDRESS, @AC_NO1, @AC_NO2, @DELIVERY_TYPE, @PAID_TYPE, @VALUE_TYPE, @TOTAL_AMOUNT, @SPP_REQ_DATE, @CS_NO)"

                Using cmd As New SqlCommand(insertHeader, sqlConn, transaction)
                    cmd.Parameters.AddWithValue("@SPP_NO", shippingNo)
                    cmd.Parameters.AddWithValue("@SPP_DATE", sppDate)
                    cmd.Parameters.AddWithValue("@TO_COMPANY", ddlCompanyList.SelectedValue)
                    cmd.Parameters.AddWithValue("@ATTENTION", txtAttention.Text)
                    cmd.Parameters.AddWithValue("@ADDRESS", txtAddress.Text)
                    cmd.Parameters.AddWithValue("@AC_NO1", txtRecipientAC.Text)
                    cmd.Parameters.AddWithValue("@AC_NO2", txtThirdPartyAC.Text)
                    cmd.Parameters.AddWithValue("@DELIVERY_TYPE", ddlDeliveryBy.SelectedValue)
                    cmd.Parameters.AddWithValue("@PAID_TYPE", ddlPaidBy.SelectedValue)
                    cmd.Parameters.AddWithValue("@VALUE_TYPE", ddlValue.SelectedValue)
                    cmd.Parameters.AddWithValue("@TOTAL_AMOUNT", totalAmount)
                    cmd.Parameters.AddWithValue("@SPP_REQ_DATE", reqDate)
                    cmd.Parameters.AddWithValue("@CS_NO", txtToCompany.Text)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Dim updateHeader As String = "UPDATE [CS_DATA].[dbo].[SPPING_H] SET " &
                    "SPP_DATE = @SPP_DATE, TO_COMPANY = @TO_COMPANY, ATTENTION = @ATTENTION, ADDRESS = @ADDRESS, " &
                    "AC_NO1 = @AC_NO1, AC_NO2 = @AC_NO2, DELIVERY_TYPE = @DELIVERY_TYPE, PAID_TYPE = @PAID_TYPE, " &
                    "VALUE_TYPE = @VALUE_TYPE, TOTAL_AMOUNT = @TOTAL_AMOUNT, SPP_REQ_DATE = @SPP_REQ_DATE, CS_NO = @CS_NO " &
                    "WHERE SPP_NO = @SPP_NO"

                Using cmd As New SqlCommand(updateHeader, sqlConn, transaction)
                    cmd.Parameters.AddWithValue("@SPP_NO", shippingNo)
                    cmd.Parameters.AddWithValue("@SPP_DATE", sppDate)
                    cmd.Parameters.AddWithValue("@TO_COMPANY", ddlCompanyList.SelectedValue)
                    cmd.Parameters.AddWithValue("@ATTENTION", txtAttention.Text)
                    cmd.Parameters.AddWithValue("@ADDRESS", txtAddress.Text)
                    cmd.Parameters.AddWithValue("@AC_NO1", txtRecipientAC.Text)
                    cmd.Parameters.AddWithValue("@AC_NO2", txtThirdPartyAC.Text)
                    cmd.Parameters.AddWithValue("@DELIVERY_TYPE", ddlDeliveryBy.SelectedValue)
                    cmd.Parameters.AddWithValue("@PAID_TYPE", ddlPaidBy.SelectedValue)
                    cmd.Parameters.AddWithValue("@VALUE_TYPE", ddlValue.SelectedValue)
                    cmd.Parameters.AddWithValue("@TOTAL_AMOUNT", totalAmount)
                    cmd.Parameters.AddWithValue("@SPP_REQ_DATE", reqDate)
                    cmd.Parameters.AddWithValue("@CS_NO", txtToCompany.Text)
                    cmd.ExecuteNonQuery()
                End Using

                ' ลบรายการเดิมก่อนเพิ่มใหม่
                Using cmd As New SqlCommand("DELETE FROM [CS_DATA].[dbo].[SPPING_D] WHERE SPP_NO = @SPP_NO", sqlConn, transaction)
                    cmd.Parameters.AddWithValue("@SPP_NO", shippingNo)
                    cmd.ExecuteNonQuery()
                End Using
            End If

            ' Insert รายการ Detail
            For i As Integer = 0 To detailTable.Rows.Count - 1
                Dim row = detailTable.Rows(i)

                Dim insertDetail As String = "INSERT INTO [CS_DATA].[dbo].[SPPING_D] " &
                    "(SPP_NO, ITEM_NO, ITEM_NAME, QTY_NUMBER, QTY_UNIT, PRICE_CURR, PRICE_UNIT_PX, PRICE_AMOUNT, BOX_SIZE, NET_WT, GW, BOX_COUNT, SPP_D_NO) " &
                    "VALUES (@SPP_NO, @ITEM_NO, @ITEM_NAME, @QTY_NUMBER, @QTY_UNIT, @PRICE_CURR, @PRICE_UNIT_PX, @PRICE_AMOUNT, @BOX_SIZE, @NET_WT, @GW, @BOX_COUNT, @SPP_D_NO)"

                Using cmd As New SqlCommand(insertDetail, sqlConn, transaction)
                    cmd.Parameters.AddWithValue("@SPP_NO", shippingNo)
                    cmd.Parameters.AddWithValue("@ITEM_NO", row("ItemNo"))
                    cmd.Parameters.AddWithValue("@ITEM_NAME", row("ItemName"))
                    cmd.Parameters.AddWithValue("@QTY_NUMBER", If(IsDBNull(row("Qty")), 0, row("Qty")))
                    cmd.Parameters.AddWithValue("@QTY_UNIT", row("Unit"))
                    cmd.Parameters.AddWithValue("@PRICE_CURR", row("Curr"))
                    cmd.Parameters.AddWithValue("@PRICE_UNIT_PX", If(IsDBNull(row("UnitPrice")), 0, row("UnitPrice")))
                    cmd.Parameters.AddWithValue("@PRICE_AMOUNT", If(IsDBNull(row("Amount")), 0, row("Amount")))
                    cmd.Parameters.AddWithValue("@BOX_SIZE", row("BoxSize"))
                    cmd.Parameters.AddWithValue("@NET_WT", If(IsDBNull(row("NW")), 0, row("NW")))
                    cmd.Parameters.AddWithValue("@GW", If(IsDBNull(row("GW")), 0, row("GW")))
                    cmd.Parameters.AddWithValue("@BOX_COUNT", row("BoxCount"))
                    cmd.Parameters.AddWithValue("@SPP_D_NO", shippingNo & Format(i, "000"))
                    cmd.ExecuteNonQuery()
                End Using
            Next

            transaction.Commit()
            sqlConn.Close()

            ' แจ้งเตือนสำเร็จ พร้อม redirect ไปดาวน์โหลด PDF
            ClientScript.RegisterStartupScript(Me.GetType(), "alert",
                "alert('บันทึกสำเร็จ! หมายเลข Shipping No: " & shippingNo & "'); " &
                "window.location.href='ShipReq.aspx?download=" & shippingNo & "';", True)

            ClearForm()
            Session("ShippingItems") = Nothing
        Catch ex As Exception
            If sqlConn.State = ConnectionState.Open Then
                Try
                    transaction.Rollback()
                Catch
                End Try
                sqlConn.Close()
            End If
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('เกิดข้อผิดพลาด: " & ex.Message.Replace("'", "\'") & "');", True)
        End Try
    End Sub

    Private Sub ExportPDF(sppNo As String)
        Dim rpt As New ReportDocument()
        Dim reportPath As String = Server.MapPath("~/Reports/ShipReport.rpt")
        rpt.Load(reportPath)

        Dim logonInfo As New TableLogOnInfo()
        logonInfo.ConnectionInfo.ServerName = "192.168.1.7"
        logonInfo.ConnectionInfo.DatabaseName = "CS_DATA"
        logonInfo.ConnectionInfo.UserID = "sa"
        logonInfo.ConnectionInfo.Password = "p@ssw0rd"

        For Each table As Table In rpt.Database.Tables
            table.ApplyLogOnInfo(logonInfo)
        Next

        rpt.SetParameterValue("SPP_NO", sppNo)

        Dim stream As Stream = rpt.ExportToStream(ExportFormatType.PortableDocFormat)
        Response.Clear()
        Response.Buffer = True
        Response.ContentType = "application/pdf"
        Response.AddHeader("Content-Disposition", "attachment;filename=Shipping_" & sppNo & ".pdf")

        stream.Seek(0, SeekOrigin.Begin)
        stream.CopyTo(Response.OutputStream)

        Response.Flush()
        Response.End()
    End Sub

   Private Function GenerateShippingNo(conn As SqlConnection, trans As SqlTransaction) As String
        Dim currentDate As String = DateTime.Now.ToString("yyyyMMdd")
        Dim prefix As String = "SP" & currentDate
        Dim query As String = "SELECT TOP 1 SPP_NO FROM [CS_DATA].[dbo].[SPPING_H] WHERE SPP_NO LIKE @Prefix ORDER BY SPP_NO DESC"

        Using cmd As New SqlCommand(query, conn, trans)
            cmd.Parameters.AddWithValue("@Prefix", prefix & "%")
            Dim lastShippingNo As Object = cmd.ExecuteScalar()

            If lastShippingNo IsNot Nothing Then
                Dim lastNumber As Integer = Integer.Parse(lastShippingNo.ToString().Substring(10))
                Return prefix & (lastNumber + 1).ToString("D5")
            Else
                Return prefix & "00001"
            End If
        End Using
    End Function

    ' ฟังก์ชันที่ถูกเรียกเมื่อคลิกปุ่ม Cancel
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Session.Clear()       ' เคลียร์ค่าทั้งหมดใน Session
        Session.Abandon()     ' ยกเลิก Session
        Response.Redirect("Home.aspx") ' กลับไปหน้า Login
    End Sub

    Protected Sub ddlPaidBy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPaidBy.SelectedIndexChanged
        Select Case ddlPaidBy.Text
            Case "Recipient"
                txtRecipientAC.Enabled = True
                txtThirdPartyAC.Enabled = False

            Case "Third Party"
                txtRecipientAC.Enabled = False
                txtThirdPartyAC.Enabled = True

            Case "TKR", ""
                txtRecipientAC.Enabled = False
                txtThirdPartyAC.Enabled = False
        End Select
    End Sub

    Protected Sub btnAddBox_Click(sender As Object, e As EventArgs) Handles btnAddBox.Click
        Dim qty As Decimal
        Dim unitPrice As Decimal
        Dim netWeight As Decimal
        Dim gw As Decimal

        Decimal.TryParse(txtQty.Text, qty)
        Decimal.TryParse(txtUnitPrice.Text, unitPrice)
        Decimal.TryParse(txtNW.Text, netWeight)
        Decimal.TryParse(txtGW.Text, gw)

        Dim amount As Decimal = qty * unitPrice
        txtAmount.Text = amount.ToString("N2")

        ' Retrieve or create DataTable
        Dim dt As DataTable = TryCast(Session("ShippingItems"), DataTable)
        If dt Is Nothing Then
            dt = New DataTable()
            dt.Columns.AddRange({
                New DataColumn("ItemNo"), New DataColumn("ItemName"),
                New DataColumn("Qty", GetType(Decimal)), New DataColumn("Unit"),
                New DataColumn("UnitPrice", GetType(Decimal)), New DataColumn("Amount", GetType(Decimal)),
                New DataColumn("Curr"), New DataColumn("BoxSize"),
                New DataColumn("NW", GetType(Decimal)), New DataColumn("GW", GetType(Decimal)),
                New DataColumn("BoxCount")
            })
        End If

        ' Add new row
        Dim newRow As DataRow = dt.NewRow()
        newRow("ItemNo") = txtItemNo.Text.Trim()
        newRow("ItemName") = txtItemName.Text.Trim()
        newRow("Qty") = qty
        newRow("Unit") = txtUnit.Text.Trim()
        newRow("UnitPrice") = unitPrice
        newRow("Amount") = amount
        newRow("Curr") = ddlCurr.SelectedValue
        newRow("BoxSize") = txtBoxSize.Text.Trim()
        newRow("NW") = netWeight
        newRow("GW") = gw
        newRow("BoxCount") = txtBoxCount.Text.Trim()

        dt.Rows.Add(newRow)

        ' Update Session and rebind
        Session("ShippingItems") = dt
        gvShippingItems.DataSource = dt
        gvShippingItems.DataBind()

        ' Update total
        Dim totalAmount As Decimal = dt.AsEnumerable().Sum(Function(r) r.Field(Of Decimal)("Amount"))
        txtTotalAmount.Text = totalAmount.ToString("N2")

        ' Clear inputs
        txtItemNo.Text = ""
        txtItemName.Text = ""
        txtQty.Text = ""
        txtUnit.Text = ""
        ddlCurr.SelectedIndex = -1
        txtUnitPrice.Text = ""
        txtAmount.Text = ""
        txtBoxSize.Text = ""
        txtNW.Text = ""
        txtGW.Text = ""
        txtBoxCount.Text = ""
    End Sub

    Protected Sub gvShippingItems_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)
        ' ดึงรายการจาก Session
        ' แปลงจาก DataTable เป็น List(Of ShippingItem)
        Dim dtItems As DataTable = TryCast(Session("ShippingItems"), DataTable)
        Dim itemList As New List(Of ShippingItem)

        If dtItems IsNot Nothing Then
            For Each se As DataRow In dtItems.Rows
                Dim item As New ShippingItem() With {
                    .ItemNo = se("ItemNo").ToString(),
                    .ItemName = se("ItemName").ToString(),
                    .Qty = Convert.ToDecimal(se("Qty")),
                    .Unit = se("Unit").ToString(),
                    .Curr = se("Curr").ToString(),
                    .UnitPrice = Convert.ToDecimal(se("UnitPrice")),
                    .Amount = Convert.ToDecimal(se("Amount")),
                    .BoxSize = se("BoxSize").ToString(),
                    .NW = Convert.ToDecimal(se("NW")),
                    .GW = Convert.ToDecimal(se("GW")),
                    .BoxCount = se("BoxCount").ToString()
                }
                itemList.Add(item)
            Next
        End If

        ' ตรวจสอบว่ามี CommandArgument ที่เป็น ItemNo หรือไม่
        Dim row As GridViewRow = gvShippingItems.Rows(e.RowIndex)
        Dim btnDelete As LinkButton = TryCast(row.FindControl("lnkDelete"), LinkButton)

        If btnDelete IsNot Nothing AndAlso Not String.IsNullOrEmpty(btnDelete.CommandArgument) Then
            ' Case 1: ลบโดยใช้ ItemNo
            Dim itemNoToDelete As String = btnDelete.CommandArgument
            Dim itemToRemove = itemList.FirstOrDefault(Function(x) x.ItemNo = itemNoToDelete)
            If itemToRemove IsNot Nothing Then
                itemList.Remove(itemToRemove)
            End If
        Else
            ' Case 2: ลบโดยใช้ index ปกติ
            If e.RowIndex >= 0 AndAlso e.RowIndex < itemList.Count Then
                itemList.RemoveAt(e.RowIndex)
            End If
        End If

        ' เก็บกลับเข้า Session
        Session("ShippingItems") = itemList

        ' Bind ใหม่
        gvShippingItems.DataSource = itemList
        gvShippingItems.DataBind()

        ' อัปเดตยอดรวม
        Dim totalAmount As Decimal = itemList.Sum(Function(x) x.Amount)
        txtTotalAmount.Text = totalAmount.ToString("N2")
    End Sub

    Private Sub ClearForm()
        txtShippingNo.Text = ""
        txtDate.Text = ""
        ddlCompanyList.SelectedIndex = -1
        txtAttention.Text = ""
        txtAddress.Text = ""
        txtRecipientAC.Text = ""
        txtThirdPartyAC.Text = ""
        ddlDeliveryBy.SelectedIndex = -1
        ddlPaidBy.SelectedIndex = -1
        ddlValue.SelectedIndex = -1
        txtShippingReqDate.Text = ""
        txtToCompany.Text = ""

        Session("ShippingItems") = Nothing
        gvShippingItems.DataSource = Nothing
        gvShippingItems.DataBind()

        lblMessage.Text = ""
        lblMessage.Visible = False
    End Sub

End Class
