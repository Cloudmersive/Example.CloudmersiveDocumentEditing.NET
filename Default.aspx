<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="CloudmersiveDocumentEditDemo._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Cloudmersive Document Editing Demo</h1>
        <p class="lead">Insert table, image and customer headers into an existing document.  Convert to PDF.</p>
        
        <asp:FileUpload ID="fileDocumentEditing" runat="server" /><asp:Button ID="btnUpload" runat="server" Text="Convert" OnClick="btnUpload_Click" />

        <div runat="server" id="resultDiv"></div>
    </div>

    <div class="row">
        <div class="col-md-4">
        </div>
    </div>

</asp:Content>
