using System;
using System.Collections.Generic;
using Xceed.Words.NET;

using Xceed.Document.NET;
using AveryAddressLabelTemplate.Avery;

namespace AveryAddressLabelTemplate
{
    class MainClass
    {
        private const double FONT_SIZE = 10;
        static string csvFilePath = "Give address file full path here...";

        public static void Main(string[] args)
        {

            List<Mailing> mailings = ReadMailingAddress(csvFilePath);
            Console.WriteLine("Total Records ::" + mailings.Count);

            DocumentBuilder builder = new AveryDocumentBuilder(AveryNumber.AVERY_5663)
                .PreparedMailingRecord(mailings)
                .PreparedAveryDocumentDirectoryPath("Give directory path for store result here....")
                .BuildAveryDocument();

            Console.WriteLine("DocxFilePath :" + builder.DocumentFilePath);
            Console.WriteLine("DocxFileName :" + builder.DocumentFileName);
        }



        public static List<Mailing> ReadMailingAddress(string csvFilePath)
        {
            List<Mailing> mails = new List<Mailing>();


            try
            {
                using (var csvReader = new CsvReader(csvFilePath, ignoreHeader: true))
                {
                    foreach (var row in csvReader.ReadAll())
                    {
                        //Console.WriteLine(string.Join(", ", row));

                        string[] address = new string[4];



                        string[] vals = row[4].Split(',');


                        address[0] = vals[0].Trim();
                        address[3] = "";
                        if (vals.Length == 3)
                        {
                            address[1] = vals[1].Trim();
                            address[2] = vals[2].Trim();
                        }
                        else
                        {
                            address[1] = address[2] = "";
                        }
                        mails.Add(new Mailing(row[3], address[0], address[1], address[2], address[3]));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            return mails;
        }//eof ReadMailingAddress()

    }
}
