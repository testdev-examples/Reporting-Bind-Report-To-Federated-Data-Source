Imports System
Imports System.Windows.Forms
Imports System.IO
Imports DevExpress.DataAccess.ConnectionParameters
Imports DevExpress.DataAccess.Sql
Imports DevExpress.DataAccess.Excel
Imports DevExpress.DataAccess.DataFederation
Imports System.ComponentModel
Imports System.Drawing
Imports DevExpress.XtraReports.UI

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
            ' Create a federated query's SQL and Excel sources.
            Dim sqlSource As New Source(sql.Name, sql, "Categories")
            Dim excelSource As New Source(excel.Name, excel, "")

            ' Create a federated query.
            Dim selectNode = sqlSource.From().Select("CategoryName").Join(excelSource, "[Excel_Products.CategoryID] = [Sql_Categories.CategoryID]").Select("CategoryID", "ProductName", "UnitPrice").Build("CategoriesProducts")
                ' Select the "CategoryName" column from the SQL source.
                ' Join the SQL source with the Excel source based on the "CategoryID" key field.
                ' Select columns from the Excel source.
                ' Specify the query's name and build it. 

            ' Create a federated data source and add the federated query to the collection.
            Dim federationDataSource = New FederationDataSource()
            federationDataSource.Queries.Add(selectNode)
            ' Build the data source schema to display it in the Field List.
            federationDataSource.RebuildResultSchema()

            Return federationDataSource
        End Function
        Public Shared Function CreateReport() As XtraReport
            ' Create a new report.
            Dim report = New XtraReport()

            ' Create data sources. 
            Dim sqlDataSource = CreateSqlDataSource()
            Dim excelDataSource = CreateExcelDataSource()
            Dim federationDataSource = CreateFederationDataSource(sqlDataSource, excelDataSource)
            ' Add all data sources to the report to avoid serialization issues. 
            report.ComponentStorage.AddRange(New IComponent() { sqlDataSource, excelDataSource, federationDataSource })
            ' Assign a federated data source to the report.
            report.DataSource = federationDataSource
            report.DataMember = "CategoriesProducts"

            ' Add the Detail band and two labels bound to the federated data source's fields.
            Dim detailBand = New DetailBand() With {.HeightF = 50}
            report.Bands.Add(detailBand)
            Dim categoryLabel = New XRLabel() With {.WidthF = 150}
            Dim productLabel = New XRLabel() With {
                .WidthF = 300,
                .LocationF = New PointF(200, 0)
            }
            categoryLabel.ExpressionBindings.Add(New ExpressionBinding("BeforePrint", "Text", "[CategoryName]"))
            productLabel.ExpressionBindings.Add(New ExpressionBinding("BeforePrint", "Text", "[ProductName]"))
            detailBand.Controls.AddRange( { categoryLabel, productLabel })

            Return report
        End Function
        Private Shared Function CreateSqlDataSource() As SqlDataSource
            Dim connectionParameters = New Access97ConnectionParameters(Path.Combine(Path.GetDirectoryName(GetType(Form1).Assembly.Location), "Data/nwind.mdb"), "", "")
            Dim sqlDataSource = New SqlDataSource(connectionParameters) With {.Name = "Sql_Categories"}
            Dim categoriesQuery = SelectQueryFluentBuilder.AddTable("Categories").SelectAllColumnsFromTable().Build("Categories")
            sqlDataSource.Queries.Add(categoriesQuery)
            sqlDataSource.RebuildResultSchema()
            Return sqlDataSource
        End Function
        Private Shared Function CreateExcelDataSource() As ExcelDataSource
            Dim excelDataSource = New ExcelDataSource() With {.Name = "Excel_Products"}
            excelDataSource.FileName = Path.Combine(Path.GetDirectoryName(GetType(Form1).Assembly.Location), "Data/Products.xlsx")
            excelDataSource.SourceOptions = New ExcelSourceOptions() With {.ImportSettings = New ExcelWorksheetSettings("Sheet")}
            excelDataSource.RebuildResultSchema()
            Return excelDataSource
        End Function
    End Class
End Namespace
