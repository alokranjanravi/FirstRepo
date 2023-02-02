<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Extract_Mail.aspx.cs" Inherits="MsgE.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Honeywell</title>
    <link href="Resource/css/custom.css" rel="stylesheet" />
    <link href="Resource/css/bootstrap.min.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
   
        <div class="container-fluid px-0">
            <header>
                <div class="shape"></div>
                <div class="container">
                    <div class="row">
                        <div class="col-md-6 left-side">
                        </div>
                        <div class="col-md-6 right-side">
                            <h1>HONEYWELL PSC</h1>
                            <br />
                            <br />
                            <div class="text-center">
                                <asp:Button ID="btnExtract" runat="server" Text="EXTRACT EMAIL" OnClick="btnExtract_Click" CssClass="btn extract-button" />
                                  <asp:Button ID="btnReset" runat="server" Text="RESET" OnClick="btnReset_Click" CssClass="btn extract-button" style="margin-left:65px" visible="false"/>
                                </div>
                            <br />

                            <div class="text-center">
                                <h6 style="color: white">
                                    <asp:Label ID="lblMessage" Text="&nbsp;" runat="server" />
                                    <asp:Label ID="lblMail" Text="&nbsp;" runat="server" Visible="false" />
                                    <asp:Label ID="lblExcelErr" Text="&nbsp;" runat="server" Visible="false" />
                                    
                               </h6>
                            </div>
                            <br />
                            <div class="row">
                                <div class="col-md-4">
                                    <h6 style="color: white">
                                        <asp:Label ID="lblTotal" Text="" runat="server" />
                                    </h6>
                                </div>
                                <div class="col-md-4">
                                    <h6 style="color: white">
                                        <asp:Label ID="lblExtract" Text="&nbsp;" runat="server" />
                                    </h6>
                                </div>
                                <div class="col-md-4">
                                    <h6 style="color: white">
                                        <asp:Label ID="lblFailed" Text="&nbsp;" runat="server" />
                                    </h6>
                                </div>
                            </div>
                          </div>
                    </div>
                </div>
            </header>
        </div>
    </form>
</body>
</html>
