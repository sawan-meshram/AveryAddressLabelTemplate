using System;
namespace AveryAddressLabelTemplate.Avery
{
    public abstract class AbstractAveryTemplate
    {
        //public abstract 
        protected AveryAddressLabel GetAveryAddressLabel(AveryNumber averyNumber)
        {
            AveryAddressLabel template = null;

            if (averyNumber == AveryNumber.AVERY_5160)
            {
                template = new AveryAddressLabel(30, 8.5f, 11, 2.625f, 1, 0.1875f, 0.49f, 0.125f)
                {
                    ColsPerSheet = 3,
                    TemplateNumber = "Avery 5160"
                };
            }
            else if (averyNumber == AveryNumber.AVERY_5660)
            {
                template = new AveryAddressLabel(30, 8.5f, 11, 2.625f, 1, 0.1875f, 0.49f, 0.125f)
                {
                    ColsPerSheet = 3,
                    TemplateNumber = "Avery 5660"
                };
            }
            else if (averyNumber == AveryNumber.AVERY_5661)
            {
                template = new AveryAddressLabel(20, 8.5f, 11, 4f, 1, 0.1875f, 0.49f, 0.125f)
                {
                    ColsPerSheet = 2,
                    TemplateNumber = "Avery 5661"
                };
            }
            else if (averyNumber == AveryNumber.AVERY_5662)
            {
                template = new AveryAddressLabel(14, 8.5f, 11, 4f, 1.33f, 0.1875f, 0.83f, 0.125f)
                {
                    ColsPerSheet = 2,
                    TemplateNumber = "Avery 5662"
                };
            }
            else if (averyNumber == AveryNumber.AVERY_5663)
            {
                template = new AveryAddressLabel(10, 8.5f, 11, 4f, 2f, 0.1875f, 0.49f, 0.125f)
                {
                    ColsPerSheet = 2,
                    TemplateNumber = "Avery 5663"
                };
            }
            return template;
        }
    }
}
