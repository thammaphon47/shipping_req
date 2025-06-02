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
        If Session("Username") Is Nothing Then
            Response.Redirect("Home.aspx")
            Exit Sub
        End If

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

        If Request.QueryString("download") IsNot Nothing AndAlso Request.QueryString("report") IsNot Nothing Then
            Dim shippingNo As String = Request.QueryString("download")
            Dim reportType As String = Request.QueryString("report")
            ExportShippingReport(shippingNo, reportType)
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

    Protected Sub ddlPaidBy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPaidBy.SelectedIndexChanged
        txtRecipientAC.Enabled = (ddlPaidBy.SelectedValue = "Recipient")
        txtThirdPartyAC.Enabled = (ddlPaidBy.SelectedValue = "Third Party")
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

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Session.Clear()
        Session.Abandon()
        Response.Redirect("Home.aspx")
    End Sub

    Protected Sub btnInv_Click(sender As Object, e As EventArgs) Handles btnInv.Click
        Dim INV_NO As String = txtInvSearch.Text.Trim()
        Dim deliveryMethod As String = ddlDeliveryBy.SelectedValue

        If String.IsNullOrEmpty(INV_NO) Then
            lblMessage.Text = "Please input INVOICE No."
            lblMessage.CssClass = "message-label"
            lblMessage.Visible = True
            Exit Sub
        End If

        Dim connString As String = ConfigurationManager.ConnectionStrings("ConStr").ConnectionString
        Dim query As String = "SELECT " &
            "INV_NO, INV_DATE, INV_TO, INV_ADDRESS, INV_ATTENTION, INV_VALUES, INV_DELIVERY, " &
            "INV_SAILING_DT, INV_LOCATION_FROM, INV_LOCATION_TO, INV_FREIGHT, INV_AWB, INV_REMARK, " &
            "INV_TOTAL_CARTONS, INV_TOTAL_NW, INV_TOTAL_GW, INV_TOTAL_AMOUNT, INV_SPP_NO, " &
            "INV_MEASUREMENT, INV_NOTIFY_PARTY, INV_TERM, INV_CS_CODE, INV_D_NO, INV_D_MARK, " &
            "INV_D_DESCRIPTION, INV_D_QTY, INV_D_UNIT, INV_D_UNIT_PRICE, INV_D_AMOUNT, BOX_SIZE, " &
            "INV_PRICE_CURR FROM CS_DATA.dbo.q_INVOICE WHERE INV_NO = @INV_NO"

        ' ถ้าเลือก delivery method เพิ่มเงื่อนไขการค้นหา
        If Not String.IsNullOrEmpty(deliveryMethod) Then
            query &= " AND INV_DELIVERY = @INV_DELIVERY"
        End If

        query &= " ORDER BY INV_D_NO"

        Dim dt As New DataTable()

        Try
            Using conn As New SqlConnection(connString)
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@INV_NO", INV_NO)
                    If Not String.IsNullOrEmpty(deliveryMethod) Then
                        cmd.Parameters.AddWithValue("@INV_DELIVERY", deliveryMethod)
                    End If

                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using

            If dt.Rows.Count = 0 Then
                lblMessage.Text = "Invoice not found."
                lblMessage.CssClass = "message-label"
                lblMessage.Visible = True
                Exit Sub
            End If

            ' ==== SET HEADER ====
            Dim row As DataRow = dt.Rows(0)
            txtInvoiceNo.Text = row("INV_NO").ToString()
            txtInvoiceDate.Text = Convert.ToDateTime(row("INV_DATE")).ToString("yyyy-MM-dd")
            txtTo.Text = row("INV_LOCATION_TO").ToString()
            txtAddress.Text = row("INV_ADDRESS").ToString()
            txtAttention.Text = row("INV_ATTENTION").ToString()

            If ddlValue.Items.FindByValue(row("INV_VALUES").ToString()) IsNot Nothing Then
                ddlValue.SelectedValue = row("INV_VALUES").ToString()
            End If

            If ddlDeliveryBy.Items.FindByValue(row("INV_DELIVERY").ToString()) IsNot Nothing Then
                ddlDeliveryBy.SelectedValue = row("INV_DELIVERY").ToString()
            End If

            txtSailingDate.Text = Convert.ToDateTime(row("INV_SAILING_DT")).ToString("yyyy-MM-dd")
            txtFrom.Text = row("INV_LOCATION_FROM").ToString()

            If ddlCompanyList.Items.FindByValue(row("INV_TO").ToString()) IsNot Nothing Then
                ddlCompanyList.SelectedValue = row("INV_TO").ToString()
            End If

            txtFreight.Text = row("INV_FREIGHT").ToString()
            txtAWB.Text = row("INV_AWB").ToString()
            txtRemark.Text = row("INV_REMARK").ToString()
            txtTotalCartons.Text = row("INV_TOTAL_CARTONS").ToString()
            txtTotalNW.Text = row("INV_TOTAL_NW").ToString()
            txtTotalGW.Text = row("INV_TOTAL_GW").ToString()
            txtTotalAmount.Text = row("INV_TOTAL_AMOUNT").ToString()
            txtShippingNo.Text = row("INV_SPP_NO").ToString()
            txtMeasurement.Text = row("INV_MEASUREMENT").ToString()
            txtNotifyParty.Text = row("INV_NOTIFY_PARTY").ToString()
            txtTerm.Text = row("INV_TERM").ToString()
            txtToCompany.Text = row("INV_CS_CODE").ToString()

            txtShippingReqDate.Text = txtInvoiceDate.Text
            txtDate.Text = DateTime.Today.ToString("yyyy-MM-dd")

            ' ==== SET DETAIL TABLE ====
            Dim detailTable As New DataTable()
            detailTable.Columns.AddRange({
                New DataColumn("ItemNo"),
                New DataColumn("ItemName"),
                New DataColumn("Qty", GetType(Decimal)),
                New DataColumn("Unit"),
                New DataColumn("UnitPrice", GetType(Decimal)),
                New DataColumn("Amount", GetType(Decimal)),
                New DataColumn("Curr"),
                New DataColumn("BoxSize")
            })

            For Each r As DataRow In dt.Rows
                detailTable.Rows.Add(
                    r("INV_D_MARK").ToString(),
                    r("INV_D_DESCRIPTION").ToString(),
                    If(IsDBNull(r("INV_D_QTY")), 0D, Convert.ToDecimal(r("INV_D_QTY"))),
                    r("INV_D_UNIT").ToString(),
                    If(IsDBNull(r("INV_D_UNIT_PRICE")), 0D, Convert.ToDecimal(r("INV_D_UNIT_PRICE"))),
                    If(IsDBNull(r("INV_D_AMOUNT")), 0D, Convert.ToDecimal(r("INV_D_AMOUNT"))),
                    r("INV_PRICE_CURR").ToString(),
                    r("BOX_SIZE").ToString()
                )
            Next

            gvShippingItems.DataSource = detailTable
            gvShippingItems.DataBind()

            Session("ShippingItems") = detailTable

            lblMessage.Text = "Invoice data loaded successfully"
            lblMessage.CssClass = "message-label message-success"
            lblMessage.Visible = True

        Catch ex As Exception
            lblMessage.Text = "Error: " & ex.Message
            lblMessage.CssClass = "message-label"
            lblMessage.Visible = True
        End Try
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

        If dt IsNot Nothing Then
            gvShippingItems.DataSource = dt
            gvShippingItems.DataBind()
        Else
            gvShippingItems.DataSource = Nothing
            gvShippingItems.DataBind()
        End If
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim config = ConfigurationManager.ConnectionStrings("ConStr")
        If config Is Nothing Then
            Throw New Exception("Connection string 'ConStr' not found.")
        End If

        Dim conn As New SqlConnection(config.ConnectionString)
        Dim transaction As SqlTransaction = Nothing

        Try
            conn.Open()
            transaction = conn.BeginTransaction()

            ' สร้าง INV_NO ใหม่
            Dim invoiceNo As String = GenerateInvoiceNo()
            txtInvoiceNo.Text = invoiceNo ' แสดงผลบนหน้าเว็บ

            ' ตรวจสอบว่ามีอยู่แล้วหรือยัง
            Dim sqlCheck As String = "SELECT INV_NO FROM [CS_DATA].[dbo].[INVOICE_H] WHERE INV_NO = @INV_NO"
            Using cmdCheck As New SqlCommand(sqlCheck, conn, transaction)
                cmdCheck.Parameters.AddWithValue("@INV_NO", invoiceNo)
                Dim dtCheck As New DataTable()
                Dim adapter As New SqlDataAdapter(cmdCheck)
                adapter.Fill(dtCheck)
                If dtCheck.Rows.Count > 0 Then
                    Throw New Exception("Invoice number already exists.")
                End If
            End Using

            ' INSERT Header
            Dim insertSql As String = "INSERT INTO [CS_DATA].[dbo].[INVOICE_H] " &
                "(INV_NO, INV_DATE, INV_TO, INV_ADDRESS, INV_ATTENTION, INV_VALUES, INV_DELIVERY, INV_SAILING_DT, " &
                "INV_LOCATION_FROM, INV_LOCATION_TO, INV_FREIGHT, INV_AWB, INV_REMARK, INV_TOTAL_CARTONS, INV_TOTAL_NW, " &
                "INV_TOTAL_GW, INV_TOTAL_AMOUNT, INV_SPP_NO, INV_MEASUREMENT, INV_NOTIFY_PARTY, INV_TERM, INV_CS_CODE) " &
                "VALUES (@INV_NO, @INV_DATE, @INV_TO, @INV_ADDRESS, @INV_ATTENTION, @INV_VALUES, @INV_DELIVERY, @INV_SAILING_DT, " &
                "@INV_LOCATION_FROM, @INV_LOCATION_TO, @INV_FREIGHT, @INV_AWB, @INV_REMARK, @INV_TOTAL_CARTONS, @INV_TOTAL_NW, " &
                "@INV_TOTAL_GW, @INV_TOTAL_AMOUNT, @INV_SPP_NO, @INV_MEASUREMENT, @INV_NOTIFY_PARTY, @INV_TERM, @INV_CS_CODE)"

            Using cmdInsert As New SqlCommand(insertSql, conn, transaction)
                cmdInsert.Parameters.AddWithValue("@INV_NO", invoiceNo)
                cmdInsert.Parameters.AddWithValue("@INV_DATE", txtDate.Text)
                cmdInsert.Parameters.AddWithValue("@INV_TO", ddlCompanyList.SelectedItem.Text)
                cmdInsert.Parameters.AddWithValue("@INV_ADDRESS", txtAddress.Text)
                cmdInsert.Parameters.AddWithValue("@INV_ATTENTION", txtAttention.Text)
                cmdInsert.Parameters.AddWithValue("@INV_VALUES", ddlValue.SelectedItem.Text)
                cmdInsert.Parameters.AddWithValue("@INV_DELIVERY", ddlDeliveryBy.SelectedItem.Text)
                cmdInsert.Parameters.AddWithValue("@INV_SAILING_DT", txtSailingDate.Text)
                cmdInsert.Parameters.AddWithValue("@INV_LOCATION_FROM", txtFrom.Text)
                cmdInsert.Parameters.AddWithValue("@INV_LOCATION_TO", txtTo.Text)
                cmdInsert.Parameters.AddWithValue("@INV_FREIGHT", txtFreight.Text)
                cmdInsert.Parameters.AddWithValue("@INV_AWB", txtAWB.Text)
                cmdInsert.Parameters.AddWithValue("@INV_REMARK", txtRemark.Text)
                cmdInsert.Parameters.AddWithValue("@INV_TOTAL_CARTONS", ToNullableDouble(txtTotalCartons.Text))
                cmdInsert.Parameters.AddWithValue("@INV_TOTAL_NW", ToNullableDouble(txtTotalNW.Text))
                cmdInsert.Parameters.AddWithValue("@INV_TOTAL_GW", ToNullableDouble(txtTotalGW.Text))
                cmdInsert.Parameters.AddWithValue("@INV_TOTAL_AMOUNT", ToNullableDouble(txtTotalAmount.Text))
                cmdInsert.Parameters.AddWithValue("@INV_SPP_NO", txtShippingNo.Text)
                cmdInsert.Parameters.AddWithValue("@INV_MEASUREMENT", txtMeasurement.Text)
                cmdInsert.Parameters.AddWithValue("@INV_NOTIFY_PARTY", txtNotifyParty.Text)
                cmdInsert.Parameters.AddWithValue("@INV_TERM", txtTerm.Text)
                cmdInsert.Parameters.AddWithValue("@INV_CS_CODE", txtToCompany.Text)
                cmdInsert.ExecuteNonQuery()
            End Using

            ' INSERT รายการ INVOICE_D จาก GridView
            Dim dtDetail As DataTable = CType(ViewState("ShippingItems"), DataTable)
            If dtDetail IsNot Nothing AndAlso dtDetail.Rows.Count > 0 Then
                Dim index As Integer = 1
                For Each row As DataRow In dtDetail.Rows
                    Dim mark As String = row("MARK_NO").ToString().Trim()
                    Dim description As String = row("DESCRIPTION").ToString().Trim()
                    Dim qty As Double = If(IsNumeric(row("QTY")), Convert.ToDouble(row("QTY")), 0)
                    Dim unit As String = row("UNIT").ToString().Trim()
                    Dim currency As String = row("CURRENCY").ToString().Trim()
                    Dim unitPrice As Double = If(IsNumeric(row("UNIT_PRICE")), Convert.ToDouble(row("UNIT_PRICE")), 0)
                    Dim amount As Double = If(IsNumeric(row("AMOUNT")), Convert.ToDouble(row("AMOUNT")), 0)

                    Dim sqlDetail As String = "INSERT INTO [CS_DATA].[dbo].[INVOICE_D] " &
                        "(INV_D_NO, INV_NO, INV_D_MARK, INV_D_DESCRIPTION, INV_D_QTY, INV_D_UNIT, " &
                        "INV_D_UNIT_PRICE, INV_D_AMOUNT, BOX_SIZE, INV_PRICE_CURR) " &
                        "VALUES (@INV_D_NO, @INV_NO, @MARK, @DESC, @QTY, @UNIT, @PRICE, @AMOUNT, @BOX_SIZE, @CURR)"

                    Using cmdDetail As New SqlCommand(sqlDetail, conn, transaction)
                        cmdDetail.Parameters.AddWithValue("@INV_D_NO", invoiceNo & "-" & index.ToString("D3"))
                        cmdDetail.Parameters.AddWithValue("@INV_NO", invoiceNo)
                        cmdDetail.Parameters.AddWithValue("@MARK", If(mark <> "", mark, DBNull.Value))
                        cmdDetail.Parameters.AddWithValue("@DESC", If(description <> "", description, DBNull.Value))
                        cmdDetail.Parameters.AddWithValue("@QTY", qty)
                        cmdDetail.Parameters.AddWithValue("@UNIT", If(unit <> "", unit, DBNull.Value))
                        cmdDetail.Parameters.AddWithValue("@PRICE", unitPrice)
                        cmdDetail.Parameters.AddWithValue("@AMOUNT", amount)
                        cmdDetail.Parameters.AddWithValue("@BOX_SIZE", DBNull.Value)
                        cmdDetail.Parameters.AddWithValue("@CURR", If(currency <> "", currency, DBNull.Value))
                        cmdDetail.ExecuteNonQuery()
                    End Using
                    index += 1
                Next
            End If

            ' UPDATE ตาราง Shipping
            Dim updateShipping As String = "UPDATE [CS_DATA].[dbo].[SPPING_H] SET INVOICE_NO=@INV_NO, AWB_NO=@AWB WHERE SPP_NO=@SPP_NO"
            Using cmdUpdate As New SqlCommand(updateShipping, conn, transaction)
                cmdUpdate.Parameters.AddWithValue("@INV_NO", invoiceNo)
                cmdUpdate.Parameters.AddWithValue("@AWB", txtAWB.Text)
                cmdUpdate.Parameters.AddWithValue("@SPP_NO", txtShippingNo.Text)
                cmdUpdate.ExecuteNonQuery()
            End Using

            ' Commit transaction
            transaction.Commit()

            ' Show success alert and redirect to PDF report
            ClientScript.RegisterStartupScript(Me.GetType(), "alert",
                "alert('บันทึกสำเร็จ! หมายเลข Invoice No: " & invoiceNo & "');" &
                "window.location.href='Invoce.aspx?download=" & invoiceNo & "&report=" & ddlReport.SelectedValue & "';", True)

            ' Clear form and session
            ClearForm()
            Session("ShippingItems") = Nothing

        Catch ex As Exception
            ' Rollback in case of error
            If transaction IsNot Nothing Then transaction.Rollback()
            ClientScript.RegisterStartupScript(Me.GetType(), "error", "alert('เกิดข้อผิดพลาด: " & ex.Message & "');", True)

        Finally
            If conn.State = ConnectionState.Open Then conn.Close()
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
        Dim sql As String = "" &
            "INSERT INTO [CS_DATA].[dbo].[INVOICE_H] (" &
            "INV_NO, INV_DATE, INV_TO, INV_ADDRESS, INV_ATTENTION, INV_VALUES, " &
            "INV_DELIVERY, INV_SAILING_DT, INV_LOCATION_FROM, INV_LOCATION_TO, " &
            "INV_FREIGHT, INV_AWB, INV_REMARK, INV_TOTAL_CARTONS, INV_TOTAL_NW, " &
            "INV_TOTAL_GW, INV_TOTAL_AMOUNT, INV_SPP_NO, INV_MEASUREMENT, " &
            "INV_NOTIFY_PARTY, INV_TERM, INV_CS_CODE) " &
            "VALUES (" &
            "@INV_NO, @INV_DATE, @INV_TO, @INV_ADDRESS, @INV_ATTENTION, @INV_VALUES, " &
            "@INV_DELIVERY, @INV_SAILING_DT, @INV_LOCATION_FROM, @INV_LOCATION_TO, " &
            "@INV_FREIGHT, @INV_AWB, @INV_REMARK, @INV_TOTAL_CARTONS, @INV_TOTAL_NW, " &
            "@INV_TOTAL_GW, @INV_TOTAL_AMOUNT, @INV_SPP_NO, @INV_MEASUREMENT, " &
            "@INV_NOTIFY_PARTY, @INV_TERM, @INV_CS_CODE)"


        Using cmd As New SqlCommand(sql, conn, trans)
            cmd.Parameters.AddWithValue("@INV_NO", txtInvoiceNo.Text.Trim())
            cmd.Parameters.AddWithValue("@INV_DATE", DateTime.Parse(txtDate.Text))
            cmd.Parameters.AddWithValue("@INV_TO", ddlCompanyList.SelectedValue)
            cmd.Parameters.AddWithValue("@INV_ADDRESS", txtAddress.Text)
            cmd.Parameters.AddWithValue("@INV_ATTENTION", txtAttention.Text)
            cmd.Parameters.AddWithValue("@INV_VALUES", ddlValue.SelectedValue)
            cmd.Parameters.AddWithValue("@INV_DELIVERY", ddlDeliveryBy.SelectedValue)
            cmd.Parameters.AddWithValue("@INV_SAILING_DT", DateTime.Parse(txtSailingDate.Text))
            cmd.Parameters.AddWithValue("@INV_LOCATION_FROM", txtFrom.Text)
            cmd.Parameters.AddWithValue("@INV_LOCATION_TO", txtTo.Text)
            cmd.Parameters.AddWithValue("@INV_FREIGHT", txtFreight.Text)
            cmd.Parameters.AddWithValue("@INV_AWB", txtAWB.Text)
            cmd.Parameters.AddWithValue("@INV_REMARK", txtRemark.Text)
            cmd.Parameters.AddWithValue("@INV_TOTAL_CARTONS", CDbl(txtTotalCartons.Text))
            cmd.Parameters.AddWithValue("@INV_TOTAL_NW", CDbl(txtTotalNW.Text))
            cmd.Parameters.AddWithValue("@INV_TOTAL_GW", CDbl(txtTotalGW.Text))
            cmd.Parameters.AddWithValue("@INV_TOTAL_AMOUNT", CDbl(txtTotalAmount.Text))
            cmd.Parameters.AddWithValue("@INV_SPP_NO", txtShippingNo.Text)
            cmd.Parameters.AddWithValue("@INV_MEASUREMENT", txtMeasurement.Text)
            cmd.Parameters.AddWithValue("@INV_NOTIFY_PARTY", txtNotifyParty.Text)
            cmd.Parameters.AddWithValue("@INV_TERM", txtTerm.Text)
            cmd.Parameters.AddWithValue("@INV_CS_CODE", txtTo.Text)
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
        Dim sql As String = "DELETE FROM [CS_DATA].[dbo].[INVOICE_D] WHERE INV_NO = @INV_NO"
        Using cmd As New SqlCommand(sql, conn, trans)
            cmd.Parameters.AddWithValue("@INV_NO", txtInvoiceNo.Text.Trim())
            cmd.ExecuteNonQuery()
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

            Dim txtMarkNo As TextBox = CType(row.FindControl("txtMarkNo"), TextBox)
            Dim txtDesc As TextBox = CType(row.FindControl("txtDesc"), TextBox)
            Dim txtQty As TextBox = CType(row.FindControl("txtQty"), TextBox)
            Dim txtUnit As TextBox = CType(row.FindControl("txtUnit"), TextBox)
            Dim txtCurrency As TextBox = CType(row.FindControl("txtCurrency"), TextBox)
            Dim txtUnitPrice As TextBox = CType(row.FindControl("txtUnitPrice"), TextBox)
            Dim txtAmount As TextBox = CType(row.FindControl("txtAmount"), TextBox)

            If txtMarkNo IsNot Nothing Then dt.Rows(index)("MARK_NO") = txtMarkNo.Text.Trim()
            If txtDesc IsNot Nothing Then dt.Rows(index)("DESCRIPTION") = txtDesc.Text.Trim()

            Dim qty As Decimal = 0D
            Decimal.TryParse(txtQty.Text, qty)
            dt.Rows(index)("QTY") = qty

            If txtUnit IsNot Nothing Then dt.Rows(index)("UNIT") = txtUnit.Text.Trim()
            If txtCurrency IsNot Nothing Then dt.Rows(index)("CURRENCY") = txtCurrency.Text.Trim()

            Dim unitPrice As Decimal = 0D
            Decimal.TryParse(txtUnitPrice.Text, unitPrice)
            dt.Rows(index)("UNIT_PRICE") = unitPrice

            ' คำนวณ Amount ใหม่ (หรือจะใช้ค่าที่ผู้ใช้ใส่มาก็ได้)
            Dim amount As Decimal = qty * unitPrice
            dt.Rows(index)("AMOUNT") = amount

            ViewState("ShippingItems") = dt
            gvShippingItems.EditIndex = -1
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
                ' ใช้ Replace ป้องกัน quote error หาก ID เป็น string
                Dim rows() As DataRow = dt.Select("ID = '" & idToDelete.Replace("'", "''") & "'")

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

    Private Function ToNullableDouble(value As String) As Object
        If String.IsNullOrWhiteSpace(value) Then
            Return DBNull.Value
        Else
            Dim result As Double
            If Double.TryParse(value, result) Then
                Return result
            Else
                Return DBNull.Value
            End If
        End If
    End Function

    Private Function ToDBNullIfEmpty(value As Object) As Object
        If value Is Nothing OrElse String.IsNullOrWhiteSpace(value.ToString()) Then
            Return DBNull.Value
        Else
            Return value
        End If
    End Function

    Private Sub ExportShippingReport(shippingNo As String, reportType As String)
        Dim reportFile As String = ""
        Dim parameterName As String = ""

        ' เลือกรายงานตามประเภท
        Select Case reportType.ToUpper()
            Case "TKR"
                reportFile = "TKRReport.rpt"
                parameterName = "TKR_NO"
            Case "INTECH"
                reportFile = "INTECHReport.rpt"
                parameterName = "INTECH_NO"
            Case "CHICAGO"
                reportFile = "ChicagoReport.rpt"
                parameterName = "CHICAGO_NO"
            Case Else
                ShowMessage("Invalid report type.")
                Return
        End Select

        Dim rpt As New ReportDocument()
        Dim reportPath As String = Server.MapPath("~/Reports/" & reportFile)
        rpt.Load(reportPath)

        ' เชื่อมต่อฐานข้อมูล
        Dim connInfo As New CrystalDecisions.Shared.ConnectionInfo()
        connInfo.ServerName = "192.168.1.7"
        connInfo.DatabaseName = "CS_DATA"
        connInfo.UserID = "sa"
        connInfo.Password = "p@ssw0rd"
        connInfo.IntegratedSecurity = False ' เปลี่ยนเป็น True หากใช้ Windows Authentication

        ' Apply ข้อมูลเข้าสู่ระบบกับทุก table
        For Each table As CrystalDecisions.CrystalReports.Engine.Table In rpt.Database.Tables
            Dim logonInfo As CrystalDecisions.Shared.TableLogOnInfo = table.LogOnInfo
            logonInfo.ConnectionInfo = connInfo
            table.ApplyLogOnInfo(logonInfo)
        Next

        ' ตรวจสอบและเชื่อมต่อ Subreports (ถ้ามี)
        For Each subreport As ReportDocument In rpt.Subreports
            For Each table As CrystalDecisions.CrystalReports.Engine.Table In subreport.Database.Tables
                Dim logonInfo As CrystalDecisions.Shared.TableLogOnInfo = table.LogOnInfo
                logonInfo.ConnectionInfo = connInfo
                table.ApplyLogOnInfo(logonInfo)
            Next
        Next

        ' ตั้งค่าพารามิเตอร์
        rpt.SetParameterValue(parameterName, shippingNo)

        ' สร้าง PDF และส่งออก
        Using stream As Stream = rpt.ExportToStream(ExportFormatType.PortableDocFormat)
            Response.Clear()
            Response.Buffer = True
            Response.ContentType = "application/pdf"
            Response.AddHeader("content-disposition", "attachment;filename=" & reportType & "_Shipping_" & shippingNo & ".pdf")
            stream.Seek(0, SeekOrigin.Begin)
            stream.CopyTo(Response.OutputStream)
            Response.Flush()
            Response.End()
        End Using
    End Sub

    Private Function GenerateInvoiceNo() As String
        Dim today As String = DateTime.Now.ToString("yyyyMMdd")
        Dim prefix As String = "INV" & today
        Dim lastNo As Integer = 0

        Try
            Using conn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConStr").ConnectionString)
                conn.Open()
                Dim query As String = "SELECT MAX(INV_NO) FROM [CS_DATA].[dbo].[INVOICE_H] WHERE INV_NO LIKE @Prefix + '%'"
                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@Prefix", prefix)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                        Dim lastInv As String = result.ToString()
                        Integer.TryParse(lastInv.Substring(prefix.Length), lastNo)
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Handle error logging here
            Throw New ApplicationException("Error generating invoice number.", ex)
        End Try

        Dim newNo As String = prefix & (lastNo + 1).ToString("D4")
        Return newNo
    End Function

    Private Sub ClearForm()
        txtInvoiceNo.Text = ""
        txtDate.Text = ""
        ddlCompanyList.ClearSelection()
        txtAddress.Text = ""
        txtAttention.Text = ""
        ddlValue.ClearSelection()
        ddlDeliveryBy.ClearSelection()
        txtSailingDate.Text = ""
        txtFrom.Text = ""
        txtTo.Text = ""
        txtFreight.Text = ""
        txtAWB.Text = ""
        txtRemark.Text = ""
        txtTotalCartons.Text = ""
        txtTotalNW.Text = ""
        txtTotalGW.Text = ""
        txtTotalAmount.Text = ""
        txtShippingNo.Text = ""
        txtMeasurement.Text = ""
        txtNotifyParty.Text = ""
        txtTerm.Text = ""

        ' ล้าง GridView และ ViewState
        ViewState("ShippingItems") = Nothing
        gvShippingItems.DataSource = Nothing
        gvShippingItems.DataBind()
    End Sub

End Class
