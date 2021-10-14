Imports System.IO
Imports ClosedXML.Excel

Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

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

                'dt = CType(SDSCargarSILES.Select(DataSourceSelectArguments.Empty), DataView).Table
                'dt = TryCast(GVSILES.DataSource, DataTable)
                'wb.Worksheets.Add(dt, "SILES")
                'wb.Worksheets(1).Columns.AdjustToContents()

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
End Class