<%@ Page Title="Shipping Request" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false"
    CodeBehind="Invoce.aspx.vb" Inherits="SHIPPING_REQ.About" %>

<%@ Import Namespace="System.Data" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css" rel="stylesheet" />
    <style>
        .form-control, .form-input {
            width: 100%;
            padding: 6px 8px;
            box-sizing: border-box;
            border: 1px solid #ccc;
            border-radius: 4px;
            font-size: 13px;
        }

        .green { background-color: #e8f5e9; }
        .red { background-color: #ffebee; }

        .btn {
            padding: 6px 12px;
            background-color: #007acc;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            margin-left: 5px;
        }

        .btn:hover { background-color: #005fa3; }
        .btn-save { background-color: #28a745; }
        .btn-save:hover { background-color: #1e7e34; }
        .btn-update { background-color: #F2C078; }
        .btn-update:hover { background-color: #FE5D26; }
        .btn-print { background-color: #ffc107; color: black; }
        .btn-print:hover { background-color: #e0a800; }
        .btn-cancel { background-color: #dc3545; }
        .btn-cancel:hover { background-color: #c82333; }
        .btn-invoice {
            float: right;
            background-color: #6c757d;
            color: white;
            padding: 6px 12px;
            border-radius: 4px;
            margin-top: -25px;
        }

        .box-section { margin-top: 20px; }

        table {
            width: 100%;
            border-spacing: 5px;
        }

        th, td {
            padding: 5px;
            text-align: left;
        }

        .message-label {
            display: block;
            margin-top: 10px;
            padding: 8px 12px;
            border-radius: 4px;
            font-weight: bold;
            font-size: 13px;
        }

        .message-success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }

        .message-error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
    </style>
</asp:Content>

<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div class="container">
        <h1 style="font-size: 40px; color: #0078D7;"><strong>INVOICE EXPORT</strong></h1>

        <!-- Invoice Detail -->
        <div class="box-section">
            <div class="section-title"><strong>Invoice Detail</strong></div>
            <div style="margin-top: 20px; text-align: right;">
    <strong>Search: </strong>
    <asp:TextBox ID="txtInvSearch" runat="server" CssClass="form-control green" 
                 style="width: 300px; display: inline-block;" placeholder="Invoice No"></asp:TextBox>
    <asp:Button ID="btnInv" runat="server" CssClass="btn" Text="🔍INV" OnClick="btnInv_Click" />
</div>
            <table>
                <tr>
                    <td>Invoice No.</td>
                    <td><asp:TextBox ID="txtInvoiceNo" runat="server" CssClass="form-input" /></td>
                    <td>From</td>
                    <td><asp:TextBox ID="txtFrom" runat="server" CssClass="form-input" /></td>
                    <td>To</td>
                    <td><asp:TextBox ID="txtTo" runat="server" CssClass="form-input" /></td>
                    
                    
                </tr>
<tr>
   <td><label for="txtSailingDate">Sailing Date:</label></td>
   <td>
      <asp:TextBox ID="txtSailingDate" runat="server" TextMode="Date" CssClass="form-input" />
   </td>
   <td><label for="txtInvoiceDate">Invoice Date:</label></td>
   <td>
      <asp:TextBox ID="txtInvoiceDate" runat="server" TextMode="Date" CssClass="form-input" />
   </td>
</tr>

                    <td>Freight</td>
                    <td><asp:TextBox ID="txtFreight" runat="server" CssClass="form-input" /></td>
                    <td>AWB</td>
                    <td><asp:TextBox ID="txtAWB" runat="server" CssClass="form-input" /></td>
                    <td>Term</td>
                    <td><asp:TextBox ID="txtTerm" runat="server" CssClass="form-input" /></td>
                </tr>
                <tr>
                    <td>Notify Party</td>
                    <td colspan="5"><asp:TextBox ID="txtNotifyParty" runat="server" CssClass="form-input" /></td>
                    <td>Report</td>
                    <td>
                        <asp:DropDownList ID="ddlReport" runat="server" CssClass="form-input">
                            <asp:ListItem Text="-- Please Select --" Value="" />
                            <asp:ListItem Text="TKR" Value="tkr" />
                            <asp:ListItem Text="INTECH" Value="intech" />
                            <asp:ListItem Text="CHICAGO" Value="chicago" />
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td>Remark</td>
                    <td colspan="4"><asp:TextBox ID="txtRemark" runat="server" CssClass="form-input" /></td>
                    <td>Measurement</td>
                    <td colspan="2"><asp:TextBox ID="txtMeasurement" runat="server" CssClass="form-input" /></td>
                </tr>
            </table>
        </div>

        <!-- Shipping Header -->
        <div class="box-section">
            <div class="section-title"><strong>Shipping Header</strong></div>
            <table>
                <tr>
                    <td><strong>Shipping No.</strong></td>
                    <td colspan="5"><asp:TextBox ID="txtShippingNo" runat="server" CssClass="form-control" /></td>
                    <td><asp:Button ID="btnSearch" runat="server" Text="Search" CssClass="btn" OnClick="btnSearch_Click" /></td>
                </tr>
                <tr>
                    <td><strong>Date</strong></td>
                    <td colspan="3"><asp:TextBox ID="txtDate" runat="server" TextMode="Date" CssClass="form-control" /></td>
                    <td><strong>Shipping REQ.</strong></td>
                    <td colspan="3"><asp:TextBox ID="txtShippingReqDate" runat="server" TextMode="Date" CssClass="form-control" /></td>
                </tr>
                <tr>
                    <td><strong>Attention</strong></td>
                    <td colspan="8"><asp:TextBox ID="txtAttention" runat="server" CssClass="form-control" Width="100%" /></td>
                </tr>
                <tr>
                    <td><strong>To Company Code:</strong></td>
                    <td><asp:TextBox ID="txtToCompany" runat="server" CssClass="form-control" ReadOnly="true" /></td>
                    <td colspan="7">
                        <asp:DropDownList ID="ddlCompanyList" runat="server" CssClass="form-control" AutoPostBack="true" OnSelectedIndexChanged="ddlCompanyList_SelectedIndexChanged" />
                        <asp:Label ID="lblMessage" runat="server" CssClass="message-label message-error" Visible="false" />
                    </td>
                </tr>
                <tr>
                    <td><strong>Address</strong></td>
                    <td colspan="8"><asp:TextBox ID="txtAddress" runat="server" CssClass="form-control" TextMode="MultiLine" Rows="2" Width="100%" /></td>
                </tr>
                <tr>
                    <td><strong>Delivery by:</strong></td>
                    <td colspan="4">
                        <asp:DropDownList ID="ddlDeliveryBy" runat="server" CssClass="form-control">
                            <asp:ListItem Text="-- Select Delivery Method --" Value="" />
                            <asp:ListItem Text="OCS" Value="1" />
                            <asp:ListItem Text="Fedex" Value="2" />
                            <asp:ListItem Text="DHL" Value="3" />
                            <asp:ListItem Text="No Specified" Value="4" />
                        </asp:DropDownList>
                    </td>
                    <td><strong>Value:</strong></td>
                    <td colspan="4">
                        <asp:DropDownList ID="ddlValue" runat="server" CssClass="form-control">
                            <asp:ListItem Text="-- Select Value --" Value="" />
                            <asp:ListItem Text="Commercail Value" Value="1" />
                            <asp:ListItem Text="No Commercail Value" Value="2" />
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td><strong>Paid by:</strong></td>
                    <td colspan="4">
                        <asp:DropDownList ID="ddlPaidBy" runat="server" CssClass="form-control" AutoPostBack="True" OnSelectedIndexChanged="ddlPaidBy_SelectedIndexChanged">
                            <asp:ListItem Text="-- Select Paid By --" Value="" />
                            <asp:ListItem Text="TKR" Value="TKR" />
                            <asp:ListItem Text="Recipient" Value="Recipient" />
                            <asp:ListItem Text="Third Party" Value="Third Party" />
                        </asp:DropDownList>
                    </td>
                    <td><strong>Recipient A/C No.:</strong></td>
                    <td><asp:TextBox ID="txtRecipientAC" runat="server" CssClass="form-control" /></td>
                    <td><strong>Third Party A/C No.:</strong></td>
                    <td colspan="3"><asp:TextBox ID="txtThirdPartyAC" runat="server" CssClass="form-control" /></td>
                </tr>
            </table>
        </div>

        <!-- Shipping Detail -->
        <div class="box-section">
<div class="section-title"><strong>Shipping Detail</strong></div>

<asp:GridView ID="gvShippingItems" runat="server" AutoGenerateColumns="False"
    OnRowEditing="gvShippingItems_RowEditing"
    OnRowUpdating="gvShippingItems_RowUpdating"
    OnRowCancelingEdit="gvShippingItems_RowCancelingEdit"
    DataKeyNames="ID"
    AutoGenerateEditButton="True">
    
    <Columns>
        <asp:TemplateField HeaderText="No" ItemStyle-Width="40px" ItemStyle-HorizontalAlign="Center">
            <ItemTemplate>
                <%# Container.DataItemIndex + 1 %>
            </ItemTemplate>
        </asp:TemplateField>

<asp:TemplateField HeaderText="MARK&NUMBER">
    <ItemTemplate>
        &nbsp;  
    </ItemTemplate>
    <EditItemTemplate>
        <asp:TextBox ID="txtMarkNo" runat="server" Text=''></asp:TextBox>
    </EditItemTemplate>
</asp:TemplateField>

        <asp:TemplateField HeaderText="DESCRIPTION">
            <ItemTemplate>
                <%# Eval("DESCRIPTION") %>
            </ItemTemplate>
            <EditItemTemplate>
                <asp:TextBox ID="txtDesc" runat="server" Text='<%# Bind("DESCRIPTION") %>'></asp:TextBox>
            </EditItemTemplate>
        </asp:TemplateField>

        <asp:TemplateField HeaderText="Qty">
            <ItemTemplate>
                <%# Eval("QTY") %>
            </ItemTemplate>
            <EditItemTemplate>
                <asp:TextBox ID="txtQty" runat="server" Text='<%# Bind("QTY") %>'></asp:TextBox>
            </EditItemTemplate>
        </asp:TemplateField>

        <asp:TemplateField HeaderText="Unit">
            <ItemTemplate>
                <%# Eval("UNIT") %>
            </ItemTemplate>
            <EditItemTemplate>
                <asp:TextBox ID="txtUnit" runat="server" Text='<%# Bind("UNIT") %>'></asp:TextBox>
            </EditItemTemplate>
        </asp:TemplateField>

        <asp:TemplateField HeaderText="Currency">
            <ItemTemplate>
                <%# Eval("CURRENCY") %>
            </ItemTemplate>
            <EditItemTemplate>
                <asp:TextBox ID="txtCurrency" runat="server" Text='<%# Bind("CURRENCY") %>'></asp:TextBox>
            </EditItemTemplate>
        </asp:TemplateField>

        <asp:TemplateField HeaderText="Unit Price">
            <ItemTemplate>
                <%# Eval("UNIT_PRICE", "{0:N2}") %>
            </ItemTemplate>
            <EditItemTemplate>
                <asp:TextBox ID="txtUnitPrice" runat="server" Text='<%# Bind("UNIT_PRICE") %>'></asp:TextBox>
            </EditItemTemplate>
        </asp:TemplateField>

        <asp:TemplateField HeaderText="Amount">
            <ItemTemplate>
                <%# Eval("AMOUNT", "{0:N2}") %>
            </ItemTemplate>
            <EditItemTemplate>
                <asp:TextBox ID="txtAmount" runat="server" Text='<%# Bind("AMOUNT") %>'></asp:TextBox>
            </EditItemTemplate>
        </asp:TemplateField>
    </Columns>
</asp:GridView>

<asp:Label ID="Label1" runat="server" CssClass="text-danger" Visible="false"></asp:Label>


            <!-- Totals -->
            <table style="margin-top: 15px;">
                <tr>
                    <td>Total Carton(s): <asp:TextBox ID="txtTotalCartons" runat="server" CssClass="form-input" /></td>
                    <td>Total NW: <asp:TextBox ID="txtTotalNW" runat="server" CssClass="form-input" /></td>
                    <td>Total GW: <asp:TextBox ID="txtTotalGW" runat="server" CssClass="form-input" /></td>
                    <td>Total Amount: <asp:TextBox ID="txtTotalAmount" runat="server" CssClass="form-input" Text="0" /></td>
                </tr>
            </table>
        </div>

        <!-- Buttons -->
        <div style="margin-top: 20px; text-align: right;">
            <asp:Button ID="btnSave" runat="server" CssClass="btn btn-save" Text="Save" OnClick="btnSave_Click" />
            <asp:Button ID="btnCancel" runat="server" CssClass="btn btn-cancel" Text="Logout" OnClick="btnCancel_Click" />
        </div>
    </div>
</asp:Content>
