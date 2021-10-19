Imports System.IO
Imports ClosedXML.Excel
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Try

            Catch ex As Exception

            End Try
        End If
    End Sub

    Protected Sub cmdCargarGrids_Click(sender As Object, e As EventArgs) Handles cmdCargarGrids.Click
        Try
            GVSIGMA.DataSource = SDSCargarSIGMA
            GVSIGMA.DataBind()
            LBLTotalSigma.Text = GVSIGMA.Rows.Count
            PNLSIGMA.Visible = True

            GVSILES.DataSource = SDSCargarSILES
            GVSILES.DataBind()
            LBLTotalSiles.Text = GVSILES.Rows.Count
            PNLSILES.Visible = True

            cmdExportarExcel.Visible = True
            lblInfo.Visible = False
        Catch ex As Exception
            lblInfo.Text = "Error al cargar los GridViews. Descripción del error : " & ex.Message
            lblInfo.Visible = True
            PNLSILES.Visible = False
            PNLSIGMA.Visible = False
            cmdExportarExcel.Visible = False
        End Try
    End Sub

    Protected Sub cmdExportarExcel_Click(sender As Object, e As EventArgs) Handles cmdExportarExcel.Click
        Try
            lblInfo.Visible = False
            'Dim dt As DataTable

            Dim dt As DataTable = CType(SDSCargarSIGMA.Select(DataSourceSelectArguments.Empty), DataView).Table
            Dim fecha As String = Date.Today.ToString("dd-MM-yyyy")

            Using wb As New XLWorkbook()
                wb.Worksheets.Add(dt, "SIGMA")
                wb.Worksheets(0).Columns.AdjustToContents()

                dt = CType(SDSCargarSILES.Select(DataSourceSelectArguments.Empty), DataView).Table
                wb.Worksheets.Add(dt, "SILES")
                wb.Worksheets(1).Columns.AdjustToContents()

                Response.Clear()
                Response.Buffer = True
                Response.Charset = ""
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                Response.AddHeader("content-disposition", "attachment;filename=Maestros_SIGMA-SILES_" & fecha & ".xlsx")
                Using MyMemoryStream As New MemoryStream()
                    wb.SaveAs(MyMemoryStream)
                    MyMemoryStream.WriteTo(Response.OutputStream)
                    Response.Flush()
                    Response.[End]()
                End Using
            End Using
            lblInfo.Visible = False
        Catch exthread As System.Threading.ThreadAbortException
            'El response.end rompe la ejecución de la página, pero no es un error propiamente.
        Catch ex As Exception
            lblInfo.Text = "Error al generar el fichero Excel. Descripción del error : " & ex.Message
            lblInfo.Visible = True
        End Try
    End Sub

    Protected Sub cmdCargarCR_Click(sender As Object, e As EventArgs) Handles cmdCargarCR.Click
        Dim rd As New ReportDocument
        Dim crTableLogonInfos As New TableLogOnInfos
        Dim crTableLogonInfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim crTables As Tables
        Dim crTable As Table

        Try
            rd.Load(Server.MapPath("\\Informes\\InformeMaestros.rpt"))

            With crConnectionInfo
                .ServerName = System.Configuration.ConfigurationManager.AppSettings("srv").ToString()
                .DatabaseName = System.Configuration.ConfigurationManager.AppSettings("bd").ToString()
                .UserID = System.Configuration.ConfigurationManager.AppSettings("usr").ToString()
                .Password = System.Configuration.ConfigurationManager.AppSettings("pwd").ToString()
            End With

            crTables = rd.Database.Tables

            For Each crTable In crTables
                crTableLogonInfo = crTable.LogOnInfo
                crTableLogonInfo.ConnectionInfo = crConnectionInfo
                crTable.ApplyLogOnInfo(crTableLogonInfo)
            Next

            CRV_PA_Informe.ReportSource = rd
            CRV_PA_Informe.RefreshReport()

            CRV_PA_Informe.Visible = True
        Catch ex As Exception
            lblInfo.Text = "Error al intentar mostrar el informe: " & ex.Message
        End Try
    End Sub
End Class