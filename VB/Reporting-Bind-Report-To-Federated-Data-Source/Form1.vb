Imports DevExpress.DataAccess.ConnectionParameters
Imports DevExpress.DataAccess.DataFederation
Imports DevExpress.DataAccess.Excel
Imports DevExpress.DataAccess.Sql
Imports DevExpress.XtraReports.Configuration
Imports DevExpress.XtraReports.UI
Imports System
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Namespace BindReportToFederatedDataSource
    Partial Public Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub Button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
            Dim designTool As New ReportDesignTool(CreateReport())
            designTool.ShowRibbonDesignerDialog()
        End Sub

        Private Shared Function CreateFederationDataSource(ByVal sql As SqlDataSource, ByVal excel As ExcelDataSource) As FederationDataSource
            Dim sqlSource As New Source(sql.Name, sql, "Categories")
            Dim excelSource As New Source(excel.Name, excel, "")

            Dim selectNode = sqlSource.From().Select("CategoryName").Join(excelSource, "[Excel_Products.CategoryID] = [Sql_Categories.CategoryID]").Select("CategoryID", "ProductName", "UnitPrice").Build("CategoriesProducts")
                ' Select the "CategoryName" column from 
                ' the SQL Source for the Federation query result
                ' Join an Excel Source using the "[Excel_Products.CategoryID] = [Sql_Categories.CategoryID]" condition
                ' Select the required columns from the Excel Source for the Federation query result
                ' Name a Federation query
            Dim federationDataSource = New FederationDataSource()
            federationDataSource.Queries.Add(selectNode)
            federationDataSource.RebuildResultSchema()
            Return federationDataSource
        End Function

        Public Shared Function CreateReport() As XtraReport
            Dim report = New XtraReport()
            Dim detailBand = New DetailBand() With {.HeightF = 50}
            report.Bands.Add(detailBand)
            Dim categoryLabel = New XRLabel() With {.WidthF = 150}
            Dim productLabel = New XRLabel() With { _
                .WidthF = 300, _
                .LocationF = New PointF(200, 0) _
            }
            categoryLabel.ExpressionBindings.Add(New ExpressionBinding("BeforePrint", "Text", "[CategoryName]"))
            productLabel.ExpressionBindings.Add(New ExpressionBinding("BeforePrint", "Text", "[ProductName]"))
            detailBand.Controls.AddRange( { categoryLabel, productLabel })

            Dim sqlDataSource = CreateSqlDataSource()
            Dim excelDataSource = CreateExcelDataSource()
            Dim federationDataSource = CreateFederationDataSource(sqlDataSource, excelDataSource)
            report.ComponentStorage.AddRange(New IComponent() { sqlDataSource, excelDataSource, federationDataSource })
            report.DataSource = federationDataSource
            report.DataMember = "CategoriesProducts"

            Return report
        End Function

        Private Shared Function CreateSqlDataSource() As SqlDataSource
            Dim connectionParameters = New Access97ConnectionParameters("Data/nwind.mdb", "", "")
            Dim sqlDataSource = New SqlDataSource(connectionParameters) With {.Name = "Sql_Categories"}
            Dim categoriesQuery = SelectQueryFluentBuilder.AddTable("Categories").SelectAllColumnsFromTable().Build("Categories")
            sqlDataSource.Queries.Add(categoriesQuery)
            sqlDataSource.RebuildResultSchema()
            Return sqlDataSource
        End Function

        Private Shared Function CreateExcelDataSource() As ExcelDataSource
            Dim excelDataSource = New ExcelDataSource() With {.Name = "Excel_Products"}
            excelDataSource.FileName = "Data/Products.xlsx"
            excelDataSource.SourceOptions = New ExcelSourceOptions() With {.ImportSettings = New ExcelWorksheetSettings("Sheet")}
            excelDataSource.RebuildResultSchema()
            Return excelDataSource
        End Function
    End Class
End Namespace
