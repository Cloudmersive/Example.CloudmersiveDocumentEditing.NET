using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Diagnostics;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Api;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Client;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Model;
using System.IO;

namespace CloudmersiveDocumentEditDemo
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            // Configure API key authorization: Apikey
            Configuration.Default.AddApiKey("Apikey", "71b067cf-1d07-474d-9403-1d0e53ca3da4");



            var apiInstance = new EditDocumentApi();
            var inputFile = new MemoryStream(fileDocumentEditing.FileBytes); // System.IO.Stream | Input file to perform the operation on.

            // Begin editing a document
            string result = apiInstance.EditDocumentBeginEditing(inputFile);

            // Add a header

            var reqConfig = new DocxSetFooterRequest(); // DocxSetHeaderRequest | 

            reqConfig.InputFileUrl = result.Replace("\"", "");

            DocxParagraph p = new DocxParagraph();
            DocxRun run = new DocxRun();
            DocxText text = new DocxText();
            text.TextContent = "Hello, World!";
            run.TextItems = new List<DocxText> { text };

            DocxRun run2 = new DocxRun();
            DocxText text2 = new DocxText();
            text2.TextContent = "- From Cloudmersive";
            run2.Bold = true;
            run2.TextItems = new List<DocxText> { text2 };
                       
            p.ContentRuns = new List<DocxRun> { run, run2 };

            DocxFooter footer = new DocxFooter();
            footer.Paragraphs = new List<DocxParagraph> { p };

            reqConfig.FooterToApply = footer;

            DocxSetFooterResponse footerResult = apiInstance.EditDocumentDocxSetFooter(reqConfig);

            // Add a table

            var addTable = new InsertDocxTablesRequest(); // InsertDocxTablesRequest | 

            addTable.InputFileUrl = footerResult.EditedDocumentURL;

            DocxTable table = new DocxTable();

            table.TopBorderType = "Single";
            table.TopBorderSize = 4;
            table.TopBorderSpace = 0;
            table.TopBorderColor = "000000";

            table.LeftBorderType = "Single";
            table.LeftBorderSize = 4;
            table.LeftBorderColor = "000000";

            table.RightBorderType = "Single";
            table.RightBorderSize = 4;
            table.RightBorderColor = "000000";

            table.CellHorizontalBorderType = "Single";
            table.CellHorizontalBorderSize = 4;
            table.CellHorizontalBorderColor = "000000";
            table.CellHorizontalBorderSpace = 0;

            table.CellVerticalBorderType = "Single";
            table.CellVerticalBorderSize = 4;
            table.CellVerticalBorderColor = "000000";
            table.CellVerticalBorderSpace = 0;

            DocxTableRow headerRow = new DocxTableRow();
            DocxTableCell col1 = new DocxTableCell();
            col1.CellShadingFill = "FFFFFF";

            p = new DocxParagraph();
            run = new DocxRun();
            text = new DocxText();
            text.TextContent = "First Header";
            run.TextItems = new List<DocxText>() { text };
            run.Bold = true;
            p.ContentRuns = new List<DocxRun>() { run };
            col1.Paragraphs = new List<DocxParagraph>() { p };
            headerRow.RowCells = new List<DocxTableCell>() { col1 };

            table.TableRows = new List<DocxTableRow> { headerRow };


            col1 = new DocxTableCell();
            col1.CellShadingFill = "FFFFFF";

            p = new DocxParagraph();
            run = new DocxRun();
            text = new DocxText();
            text.TextContent = "Second Header";
            run.TextItems = new List<DocxText>() { text };
            run.Bold = true;
            p.ContentRuns = new List<DocxRun>() { run };
            col1.Paragraphs = new List<DocxParagraph>() { p };
            headerRow.RowCells = new List<DocxTableCell>() { col1 };

            headerRow.RowCells.Add(col1);

            table.TableRows = new List<DocxTableRow> { headerRow };

            for (int i = 1; i <= 10; i++)
            {
                // Column 1

                DocxTableRow dataRow = new DocxTableRow();
                DocxTableCell dataCol1 = new DocxTableCell();
                dataCol1.CellShadingFill = "FFFFFF";

                p = new DocxParagraph();
                run = new DocxRun();
                text = new DocxText();
                text.TextContent = "Data Row " + i.ToString();
                run.TextItems = new List<DocxText>() { text };
                run.Bold = false;
                p.ContentRuns = new List<DocxRun>() { run };
                dataCol1.Paragraphs = new List<DocxParagraph>() { p };
                dataRow.RowCells = new List<DocxTableCell>() { dataCol1 };

                table.TableRows.Add(dataRow);

                // Column 2

                dataCol1 = new DocxTableCell();
                dataCol1.CellShadingFill = "FFFFFF";

                p = new DocxParagraph();
                run = new DocxRun();
                text = new DocxText();
                text.TextContent = "Data Row " + i.ToString();
                run.TextItems = new List<DocxText>() { text };
                run.Bold = false;
                p.ContentRuns = new List<DocxRun>() { run };
                dataCol1.Paragraphs = new List<DocxParagraph>() { p };
                dataRow.RowCells = new List<DocxTableCell>() { dataCol1 };

                dataRow.RowCells.Add(dataCol1);

                table.TableRows.Add(dataRow);
            }

            addTable.TableToInsert = table;

            InsertDocxTablesResponse tableAdded = apiInstance.EditDocumentDocxInsertTable(addTable);

            // Finish editing

            var finEditing = new FinishEditingRequest(); // FinishEditingRequest | 
            finEditing.InputFileUrl = tableAdded.EditedDocumentURL;

            byte[] finishedDocument = apiInstance.EditDocumentFinishEditing(finEditing);

            // Convert to PDF

            var convertApi = new ConvertDocumentApi();
            var finishedDoc = new MemoryStream(finishedDocument); // System.IO.Stream | Input file to perform the operation on.

            byte[] pdf = convertApi.ConvertDocumentDocxToPdf(finishedDoc);

            // Display the result

            var viewerApi = new ViewerToolsApi();

            var finalResult = new MemoryStream(pdf);
            ViewerResponse viewerSnippet = viewerApi.ViewerToolsCreateSimple(finalResult);

            resultDiv.InnerHtml = viewerSnippet.HtmlEmbed.Replace("iframe", "iframe width='800' height='1000'");
        }
    }
}