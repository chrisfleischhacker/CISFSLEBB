<%@ Page trace="false" Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" StylesheetTheme="SkinFile" MasterPageFile="~/Site.master"  %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="PageTitle" runat="server" ContentPlaceHolderID="PageTitle">
    <h1>CISF SLEBB Management</h1>
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">

    <asp:Panel ID="pnlHeader" CssClass="pnlHeader" runat="server">
        <p>The SLEBB Management application assists with upload and distribution of PowerPoint presentations to the CISF Portal and Space Logistics Electronic Bulletin Boards located throughout the building.</p>
        <p>
            <asp:Label runat="server" ID="lblError" /></p>
            </asp:Panel>
    <asp:Panel ID="pnlUpload" CssClass="pnlUpload" runat="server" Visible="true">
        <asp:FileUpload ID="fileupload1" runat="server" />
        <asp:Button ID="Button1" runat="server" Text="  Upload  " OnClick="UploadPPTX" Height="22px" />
    </asp:Panel>
    <asp:Panel ID="pnlSLEBBs" CssClass="pnlSLEBBs" runat="server" Visible="false">
        <p>You may choose to update the CISF Home Portal and the following SLEBBs located in this building.</p>
        <asp:CheckBoxList ID="cbxDirStat" runat="server" Visible="false" RepeatDirection="Horizontal" OnDataBound="CheckAllBoxes">
        </asp:CheckBoxList>
        <br />
        <asp:Button ID="btnFileProcessing" runat="server" Text="Update selected SLEBBs and Portal" OnClick="UpdateBBs_OnClick" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClick="Cancel_OnClick" />

    </asp:Panel>
    <asp:Panel ID="pnlImages" CssClass="pnlImages" runat="server" ScrollBars="Vertical" Height="600px" Visible="false">
        <p>The following slides were created from your upload.</p>
        <asp:DataList ID="dlImages" runat="server" RepeatColumns="4" CellPadding="3" Visible="false">
            <SelectedItemStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />

            <ItemTemplate>
                <asp:Image ID="ibImage" runat="server" ImageUrl='<%# Eval("Name", "jpg/{0}") %>' Width="165px" Height="155px" />
            </ItemTemplate>
            <FooterStyle />
            <ItemStyle />
        </asp:DataList>

    </asp:Panel>
</asp:Content>

