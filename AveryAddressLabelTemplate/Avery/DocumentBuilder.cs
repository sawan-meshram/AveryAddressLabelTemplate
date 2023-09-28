using System;
using System.Collections.Generic;
using Xceed.Words.NET;
using Xceed.Document.NET;

namespace AveryAddressLabelTemplate.Avery
{
    public class DocumentBuilder 
    {
        private const double FONT_SIZE = 10;

        private readonly AveryNumber _averyNumber;
        private readonly AveryAddressLabel _averyAddressLabel;
        private readonly List<Mailing> _mailings;
        private readonly int _totalLabelPerSheet;
        private readonly string _averyDocumentDirectoryPath;

        private int rows = 0;
        private int cols = 0;

        public DocumentBuilder(AveryDocumentBuilder documentBuilder)
        {
            _averyNumber = documentBuilder.AveryNumber;
            _averyAddressLabel = documentBuilder.AddressLabel;
            _mailings = documentBuilder.Mailings;
            _totalLabelPerSheet = documentBuilder.TotalLabelPerSheet;
            _averyDocumentDirectoryPath = documentBuilder.AveryDocumentDirectoryPath;

            //Prepare runtime document file
            PreparedDocumentFile();
            BuildDocumentTemplateDesign();
        }

        private void PreparedDocumentFile()
        {
            DocumentFileName = "MailMerge_"+DateTime.Now.ToString("yyyy.MM.dd_hh.mm.ss.tt")+".docx";
            DocumentFilePath = System.IO.Path.Combine(_averyDocumentDirectoryPath, DocumentFileName);
        }


        public string DocumentFileName { get; private set; }
        public string DocumentFilePath { get; private set; }


        private void BuildDocumentTemplateDesign()
        {
            try
            {
                // Create a new document.
                using (DocX document = DocX.Create(DocumentFilePath))
                {
                    //Prepare address label templates
                    var table = GetAddressLabelTemplate(document);

                    document.PageWidth = _averyAddressLabel.PaperSizeWidth;
                    document.PageHeight = _averyAddressLabel.PaperSizeHeight;

                    document.MarginTop = document.MarginBottom = _averyAddressLabel.TopMargin;
                    document.MarginLeft = document.MarginRight = _averyAddressLabel.SideMargin;

                    document.InsertTable(table);

                    //Add mailing records to document
                    AddMailingRecordToDocument(document);


                    // Save all changes made to this document.
                    document.Save();
                    Console.WriteLine("\tCreated: " + DocumentFilePath);
                }
            }

            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }//eof BuildDocumentTemplateDesign()


        private void AddMailingRecordToDocument(Document document)
        {
            try
            {
                int index = 0;
                bool allWritten = false;

                var table = document.Tables[0];

                for (int row = 0; row < rows; row++)
                {
                    Row _row = table.Rows[row];

                    for (int col = 0; col < cols; col++)
                    {

                        if (col % 2 == 0)
                        {
                            //Reading all records, therefore break this iteration
                            if (index == _mailings.Count)
                            {
                                allWritten = true;
                                break;
                            }
                            //Get Mailing Details from List Collection
                            Mailing mail = _mailings[index++];

                            var cell = _row.Cells[col];

                            //Adding mailing details to cells
                            Paragraph p = cell.InsertParagraph(mail.OwnerName).FontSize(FONT_SIZE)
                                .AppendLine(mail.Street).FontSize(FONT_SIZE)
                                .AppendLine(string.Join(" ", mail.City, mail.State, mail.ZipCode)).FontSize(FONT_SIZE);
                        }
                    }
                    //All data are added to rows, therefore break this iteration
                    if (allWritten) break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }//eof AddMailingRecordToDocument()


        private Table GetAddressLabelTemplate(DocX doc)
        {

            rows = (_averyAddressLabel.LabelsPerSheet / _averyAddressLabel.ColsPerSheet) * _totalLabelPerSheet;
            cols = (_averyAddressLabel.ColsPerSheet * 2) - 1; // to add spacing between each column, adding extra cols for giving rowspan on table
            //Console.WriteLine("rows : {0}, cols : {1}", rows, cols);

            var table = doc.AddTable(rows, cols);
            table.Design = TableDesign.TableGrid;


            //Set Label height
            for (int row = 0; row < rows; row++)
            {
                Row _row = table.Rows[row];
                _row.Height = _averyAddressLabel.LabelHeight;
            }

            //Set Label weidth & horizontal gap
            for (int col = 0; col < cols; col++)
            {
                if (col % 2 == 0)
                {
                    table.SetColumnWidth(col, _averyAddressLabel.LabelWidth);
                }
                else
                {
                    table.SetColumnWidth(col, _averyAddressLabel.HorizontalGap);
                }
            }
            //Set Row span on a table
            for (int col = 0; col < cols; col++)
            {
                if (col % 2 != 0)
                {
                    table.MergeCellsInColumn(col, 0, table.Rows.Count - 1);
                }
            }
            return table;
        }//eof GetAddressLabelTemplate()
    }
}
