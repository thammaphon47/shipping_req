<%@ Page Title="Home Page" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false"
    CodeBehind="ShipReq.aspx.vb" Inherits="SHIPPING_REQ._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <style>
       .form-control {
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

        .section { margin-top: 20px; }

        table {
            width: 100%;
            border-spacing: 5px;
        }
          .section1 { margin-top: 20px; }

        table {
            width: 100%;
            border-spacing: 5px;
        }

        th, td {
            padding: 5px;
            text-align: left;
        }

        .style1 { width: 24px; }

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

        .style2 { height: 42px; }
        
    </style>
</asp:Content>

<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div class="container">
        <asp:HiddenField ID="hfEditIndex" runat="server" Value="-1" />

        <div style="text-align: center;">
            <h1 style="font-size: 40px;"><strong><span style="color: #0078D7;">Shipping Request</span></strong></h1>
        </div>

        <div class="sub-header">
            <strong>Please submit this form to the coordinator 2 days before shipping</strong>
        </div>
        <div class="section1">             
<div style="margin-top: 20px; text-align: right;">
    <strong>Search: </strong>
    <asp:TextBox ID="txtSearch" runat="server" CssClass="form-control green" 
                 style="width: 300px; display: inline-block;" placeholder="Shipping No"></asp:TextBox>
    <asp:Button ID="btnEdit" runat="server" CssClass="btn" Text="Edit" OnClick="btnEdit_Click" />
</div>
  </div>
        <div class="section">
            <table>
                <tr>
                    <td>Shipping No.:</td>
                    <td>
                        <asp:TextBox ID="txtShippingNo" runat="server" ReadOnly="true" CssClass="form-control green" placeholder="Enter Shipping No." />
                    </td>
                    <td>Date:</td>
                    <td>
                        <asp:TextBox ID="txtDate" runat="server" TextMode="Date" CssClass="form-control"  />
                    </td>
                    <td>Shipping REQ:</td>
                    <td class="style1">
                        <asp:TextBox ID="txtShippingReqDate" runat="server" TextMode="Date" CssClass="form-control" />
                    </td>
                </tr>

                <tr>
                    <td>Attention:</td>
                    <td colspan="5">
                        <asp:TextBox ID="txtAttention" runat="server" CssClass="form-control" placeholder="Enter Attention Name" />
                    </td>
                </tr>

                <tr>
                    <td>To Company Code:</td>
                    <td>
                        <asp:TextBox ID="txtToCompany" runat="server" CssClass="form-control" ReadOnly="true" />
                    </td>
                    <td colspan="4">
                        <asp:DropDownList ID="ddlCompanyList" runat="server" CssClass="form-control" AutoPostBack="True" OnSelectedIndexChanged="ddlCompanyList_SelectedIndexChanged">
                            <asp:ListItem Text="---- Select Company ----" Value="" />
                        </asp:DropDownList>
                        <asp:Label ID="lblMessage" runat="server" CssClass="message-label message-error" Visible="false" />
                    </td>
                </tr>

                <tr>
                    <td class="style2">Address:</td>
                    <td colspan="5" class="style2">
                        <asp:TextBox ID="txtAddress" runat="server" CssClass="form-control" placeholder="Enter Shipping Address" />
                    </td>
                </tr>

                <tr>
                    <td>Delivery by:</td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddlDeliveryBy" runat="server" CssClass="form-control">
                            <asp:ListItem Text="-- Select Delivery Method --" Value="" />
                            <asp:ListItem Text="OCS" Value="1" />
                            <asp:ListItem Text="Fedex" Value="2" />
                            <asp:ListItem Text="DHL" Value="3" />
                            <asp:ListItem Text="No Specified" Value="4" />
                        </asp:DropDownList>
                    </td>
                    <td>Value:</td>
                    <td colspan="2">
                        <asp:DropDownList ID="ddlValue" runat="server" CssClass="form-control">
                            <asp:ListItem Text="-- Select Value  --" Value="" />
                            <asp:ListItem Text="Commercail Value" Value="1" />
                            <asp:ListItem Text="No Commercail Value" Value="2" />
                        </asp:DropDownList>
                    </td>
                </tr>

              <tr>
    <td>Paid by:</td>
    <td colspan="3">
        <asp:DropDownList ID="ddlPaidBy" runat="server" CssClass="form-control"
            AutoPostBack="True" OnSelectedIndexChanged="ddlPaidBy_SelectedIndexChanged">
            <asp:ListItem Text="-- Select Paid By --" Value="" />
            <asp:ListItem Text="TKR" Value="TKR" />
            <asp:ListItem Text="Recipient" Value="Recipient" />
            <asp:ListItem Text="Third Party" Value="Third Party" />
        </asp:DropDownList>
    </td>
</tr>
<tr>
    <td>Recipient A/C No.:</td>
    <td colspan="2">
        <asp:TextBox ID="txtRecipientAC" runat="server" CssClass="form-control" 
                     placeholder="Enter Recipient A/C No." Enabled="false" />
    </td>
    <td>Third Party A/C No.:</td>
    <td colspan="2">
        <asp:TextBox ID="txtThirdPartyAC" runat="server" CssClass="form-control" 
                     ForeColor="Purple" placeholder="Enter Third Party A/C No." Enabled="false" />
    </td>
</tr>
            </table>
        </div>

        <div class="section">
            <h3>SHIPPING DETAIL</h3>
            <table>
                <tr>
                    <th>Item No.</th>
                    <th>Item Name</th>
                    <th>QTY</th>
                    <th>Unit</th>
                    <th>Unit Price</th>
                    <th>CURR</th>
                    <th>Amount</th>
                    <th>Box Size</th>
                    <th>NW</th>
                    <th>GW</th>
                    <th>Box Count</th>
                </tr>
                <tr>
                    <td><asp:TextBox ID="txtItemNo" runat="server" CssClass="form-control" placeholder="Item No." /></td>
                    <td><asp:TextBox ID="txtItemName" runat="server" CssClass="form-control" placeholder="Item Name" /></td>
                    <td><asp:TextBox ID="txtQty" runat="server" CssClass="form-control green" Text="" placeholder="Quantity" /></td>
                    <td><asp:TextBox ID="txtUnit" runat="server" CssClass="form-control" placeholder="Unit" /></td>
                    <td><asp:TextBox ID="txtUnitPrice" runat="server" CssClass="form-control green" Text="" placeholder="Unit Price" /></td>
                    <td>
    <asp:DropDownList ID="ddlCurr" runat="server" CssClass="form-control">
        <asp:ListItem Text="-- Select Currency --" Value="" />
        <asp:ListItem Text="THB" Value="THB" />
        <asp:ListItem Text="USD" Value="USD" />
        <asp:ListItem Text="IDR" Value="IDR" />
        <asp:ListItem Text="JPY" Value="JPY" />
        <asp:ListItem Text="INR" Value="INR" />
                <asp:ListItem Text="SGD" Value="SGD" />
    </asp:DropDownList>
</td>
                    <td><asp:TextBox ID="txtAmount" runat="server" ReadOnly="true" CssClass="form-control red" Text="" placeholder="Amount" /></td>
                    <td><asp:TextBox ID="txtBoxSize" runat="server" CssClass="form-control green" placeholder="BS" /></td>
                    <td><asp:TextBox ID="txtNW" runat="server" CssClass="form-control green" Text="" placeholder="NW" /></td>
                    <td><asp:TextBox ID="txtGW" runat="server" CssClass="form-control green" Text="" placeholder="GW" /></td>
                    <td><asp:TextBox ID="txtBoxCount" runat="server" CssClass="form-control green" Text="" placeholder="BC" /></td>
                    <td><asp:Button ID="btnAddBox" runat="server" CssClass="btn green" Text="+" /></td>
                </tr>
            </table>

            <table>
                <tr>
<asp:GridView ID="gvShippingItems" runat="server"
    AutoGenerateColumns="False"
    DataKeyNames="ItemNo"
    OnRowDeleting="gvShippingItems_RowDeleting">

    <Columns>
        <asp:TemplateField HeaderText="No">
            <ItemTemplate>
                <%# Container.DataItemIndex + 1 %>
            </ItemTemplate>
        </asp:TemplateField>

        <asp:BoundField DataField="ItemNo" HeaderText="Item No" />
        <asp:BoundField DataField="ItemName" HeaderText="Item Name" />
        <asp:BoundField DataField="Qty" HeaderText="Qty" DataFormatString="{0:N2}" />
        <asp:BoundField DataField="Unit" HeaderText="Unit" />
        <asp:BoundField DataField="Curr" HeaderText="Currency" />
        <asp:BoundField DataField="UnitPrice" HeaderText="Unit Price" DataFormatString="{0:N2}" />
        <asp:BoundField DataField="Amount" HeaderText="Amount" DataFormatString="{0:N2}" />
        <asp:BoundField DataField="BoxSize" HeaderText="Box Size" />
        <asp:BoundField DataField="NW" HeaderText="NW" DataFormatString="{0:N2}" />
        <asp:BoundField DataField="GW" HeaderText="GW" DataFormatString="{0:N2}" />
        <asp:BoundField DataField="BoxCount" HeaderText="Box Count" />
        
        <asp:TemplateField HeaderText="">
            <ItemTemplate>
                <asp:LinkButton ID="lnkDelete" runat="server" Text="Delete" 
                    CommandName="Delete"
                    CommandArgument='<%# Eval("ItemNo") %>'
                    OnClientClick="return confirm('Are you sure you want to delete this item?');" />
            </ItemTemplate>
        </asp:TemplateField>
    </Columns>
</asp:GridView>

                </tr>
                <tr>
                    <td style="height: 100px;"></td>

                </tr>
            </table>

            <div style="text-align: right; margin-top: 10px;">
                <strong>Total Amount: </strong>
                <asp:TextBox ID="txtTotalAmount" runat="server" CssClass="form-control green" style="width: 100px; display: inline-block;" Text="0" />
            </div>
        </div>

        <div style="margin-top: 20px; text-align: right;">
            <asp:Button ID="btnSave" runat="server" CssClass="btn btn-save" Text="Save" OnClick="btnSave_Click" />
            <asp:Button ID="btnCancel" runat="server" CssClass="btn btn-cancel" Text="logout" OnClick="btnCancel_Click" />
        </div>
    </div>
</asp:Content>