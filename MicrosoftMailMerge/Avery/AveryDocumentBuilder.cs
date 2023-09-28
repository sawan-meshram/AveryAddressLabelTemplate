using System;
using System.Collections.Generic;

namespace AveryAddressLabelTemplate.Avery
{
    public class AveryDocumentBuilder : AbstractAveryTemplate
    {
        public AveryDocumentBuilder(AveryNumber averyNumber)
        {
            AveryNumber = averyNumber;
            AddressLabel = GetAveryAddressLabel(averyNumber);
        }

        public AveryDocumentBuilder PreparedMailingRecord(List<Mailing> mailings)
        {
            Mailings = mailings;
            TotalLabelPerSheet = GetRequiredLabelPerSheet(mailings.Count, AddressLabel.LabelsPerSheet);
            return this;
        }

        public AveryDocumentBuilder PreparedAveryDocumentDirectoryPath(string directoryPath)
        {
            AveryDocumentDirectoryPath = directoryPath;
            return this;
        }

        private int GetRequiredLabelPerSheet(int totalRecords, int labelPerSheet)
        {
            return (int)Math.Ceiling((float)totalRecords / labelPerSheet);
        }

        public DocumentBuilder BuildAveryDocument()
        {
            return new DocumentBuilder(this);
        }

        public AveryNumber AveryNumber { get; private set; }
        public List<Mailing> Mailings { get; private set; }
        public AveryAddressLabel AddressLabel { get; private set; }
        public int TotalLabelPerSheet { get; private set; }
        public string AveryDocumentDirectoryPath { get; private set; }
    }
}
