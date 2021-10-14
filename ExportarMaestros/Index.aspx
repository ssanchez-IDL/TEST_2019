﻿<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Index.aspx.vb" Inherits="ExportarMaestros._Default" %>

<%@ Register assembly="CrystalDecisions.Web, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" namespace="CrystalDecisions.Web" tagprefix="CR" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>Exportar Maestros</h1>
    </div>

    <div class="row">
        <div class="col-md-4">
            <table style="width:100%;">
                <tr>
                    <td rowspan="3" style="vertical-align: top; padding: 10px; text-align: center;">
                        <asp:Button ID="cmdCargarGrids" runat="server" Text="Cargar datos" />
                        <br />
                        <br />
                        <asp:Button ID="cmdExportarExcel" runat="server" Text="Exportar Excel" Visible="False" />
                        <br />
                    </td>
                    <td colspan="2" style="vertical-align: top">
                        <asp:Label ID="lblInfo" runat="server" ForeColor="#FF9900"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top; padding-right: 10px">
                        <asp:Panel ID="PNLSIGMA" runat="server" Visible="False">
                            <asp:Label ID="Label2" runat="server" Text="Maestros SIGMA - Filas:"></asp:Label>
                            <asp:Label ID="LBLTotalSigma" runat="server"></asp:Label>
                            <asp:GridView ID="GVSIGMA" runat="server" AutoGenerateColumns="False" DataKeyNames="Referencia">
                                <Columns>
                                    <asp:BoundField DataField="Referencia" HeaderText="Referencia" ReadOnly="True" SortExpression="Referencia" />
                                    <asp:BoundField DataField="ReferenciaPSA" HeaderText="ReferenciaPSA" SortExpression="ReferenciaPSA" />
                                    <asp:BoundField DataField="Descripcion" HeaderText="Descripcion" SortExpression="Descripcion" />
                                    <asp:BoundField DataField="Familia" HeaderText="Familia" SortExpression="Familia" />
                                    <asp:BoundField DataField="AlmacenTeorico" HeaderText="AlmacenTeorico" SortExpression="AlmacenTeorico" />
                                </Columns>
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:GridView>
                            <asp:SqlDataSource ID="SDSCargarSIGMA" runat="server" ConnectionString="<%$ ConnectionStrings:CNSBDSYSJIT %>" ProviderName="<%$ ConnectionStrings:CNSBDSYSJIT.ProviderName %>" SelectCommand="SP_IIS_MaestroSIGMASILS" SelectCommandType="StoredProcedure"></asp:SqlDataSource>
                        </asp:Panel>
                    </td>
                    <td style="vertical-align: top; padding-left: 10px"">
                        <asp:Panel ID="PNLSILES" runat="server" Visible="False">
                            <asp:Label ID="Label3" runat="server" Text="Maestros SILES - Filas:"></asp:Label>
                            <asp:Label ID="LBLTotalSiles" runat="server"></asp:Label>
                            <asp:GridView ID="GVSILES" runat="server" AutoGenerateColumns="False" DataKeyNames="IdReferencia,IdProveedor">
                                <Columns>
                                    <asp:BoundField DataField="IdReferencia" HeaderText="IdReferencia" ReadOnly="True" SortExpression="IdReferencia" />
                                    <asp:BoundField DataField="IdProveedor" HeaderText="IdProveedor" ReadOnly="True" SortExpression="IdProveedor" />
                                    <asp:BoundField DataField="Denominacion" HeaderText="Denominacion" SortExpression="Denominacion" />
                                    <asp:BoundField DataField="IdFamilia" HeaderText="IdFamilia" SortExpression="IdFamilia" />
                                </Columns>
                            </asp:GridView>
                            <asp:SqlDataSource ID="SDSCargarSILES" runat="server" ConnectionString="<%$ ConnectionStrings:CNSBDSYSJIT %>" ProviderName="<%$ ConnectionStrings:CNSBDSYSJIT.ProviderName %>" SelectCommand="SP_IIS_MaestroSILES" SelectCommandType="StoredProcedure"></asp:SqlDataSource>
                        </asp:Panel>
                        <br />
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top; padding-right: 10px" colspan="2">
                        <CR:CrystalReportViewer ID="CrystalReportViewer1" runat="server" AutoDataBind="true" />
                    </td>
                </tr>
                </table>
            <br />
        </div>
    </div>

</asp:Content>