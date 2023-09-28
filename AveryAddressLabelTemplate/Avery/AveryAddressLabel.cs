using System;
namespace AveryAddressLabelTemplate.Avery
{
    public class AveryAddressLabel
    {

        private const float INCHES_TO_POINT = 72;


        /// <summary>
        /// Initializes a Avery Template dimension with help of this class. 
        /// All dimension must be initialize in inches, so it auto convert to DOCx library dimension.
        /// </summary>
        /// <param name="labelsPerSheet">Labels per sheet.</param>
        /// <param name="paperSizeWidth">Paper size width in inches.</param>
        /// <param name="paperSizeHeight">Paper size height in inches.</param>
        /// <param name="labelWidth">Label width in inches.</param>
        /// <param name="labelHeight">Label height in inches.</param>
        /// <param name="sideMargin">Side margin in inches.</param>
        /// <param name="topMargin">Top margin in inches.</param>
        /// <param name="horizontalGap">Horizontal gap in inches.</param>
        public AveryAddressLabel(int labelsPerSheet, float paperSizeWidth, float paperSizeHeight,
            float labelWidth, float labelHeight, float sideMargin, float topMargin, float horizontalGap)
        {
            LabelsPerSheet = labelsPerSheet;
            PaperSizeWidth = paperSizeWidth * INCHES_TO_POINT;
            PaperSizeHeight = paperSizeHeight * INCHES_TO_POINT;
            LabelWidth = labelWidth * INCHES_TO_POINT;
            LabelHeight = labelHeight * INCHES_TO_POINT;
            SideMargin = sideMargin * INCHES_TO_POINT;
            TopMargin = topMargin * INCHES_TO_POINT;
            HorizontalGap = horizontalGap * INCHES_TO_POINT;
        }

        public int LabelsPerSheet { get; private set; }
        public float PaperSizeWidth { get; private set; }
        public float PaperSizeHeight { get; private set; }
        public float LabelWidth { get; private set; }
        public float LabelHeight { get; private set; }
        public float SideMargin { get; private set; }
        public float TopMargin { get; private set; }
        public float HorizontalGap { get; private set; }

        public string TemplateNumber { get; set; }
        public int ColsPerSheet { get; set; }
    }
}
