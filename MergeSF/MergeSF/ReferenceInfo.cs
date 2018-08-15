using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ujihara.Chemistry.MSOffice;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ujihara.Chemistry.MergeSF
{
    public class ReferenceInfo
    {
        public ReferenceInfo()
        {
        }

        internal string _AccessionNumber;
        public string AccessionNumber
        {
            get { return _AccessionNumber; }
            set { _AccessionNumber = value; }
        }

        internal string _Title;
        public string Title
        {
            get { return _Title; }
            set { _Title = value; }
        }
        
        internal string _By;
        public string By
        {
            get { return _By; }
            set { _By = value; }
        }
        
        internal string _PatentAssignee;
        public string PatentAssignee
        {
            get { return _PatentAssignee; }
            set { _PatentAssignee = value; }
        }

        internal string _CorporateSource;
        public string CorporateSource
        {
            get { return _CorporateSource; }
            set { _CorporateSource = value; }
        }

        internal string _Source;
        public string Source
        {
            get { return _Source; }
            set { _Source = value; }
        }

        internal string _Publisher;
        public string Publisher
        {
            get { return _Publisher; }
            set { _Publisher = value; }
        }

        internal string _DocumentType;
        public string DocumentType
        {
            get { return _DocumentType; }
            set { _DocumentType = value; }
        }

        internal string _Language;
        public string Language
        {
            get { return _Language; }
            set { _Language = value; }
        }

        internal string _PatentInfomation;
        public string PatentInfomation
        {
            get { return _PatentInfomation; }
            set { _PatentInfomation = value; }
        }

        internal string _Abstract;
        public string Abstract
        {
            get { return _Abstract; }
            set { _Abstract = value; }
        }

        public virtual byte[] AbstractImage { get; set; }

        public IEnumerable<SubstanceInfo> SubstancesInfo
        {
            get;
            set;
        }

        internal string _Copyright;
        public string Copyright
        {
            get { return _Copyright; }
            set { _Copyright = value; }
        }
    }

    public class AReferenceInfo
        : ReferenceInfo
    {
        internal A.Graphic Graphic { get; set; }
        internal OpenXmlPartContainer Container { get; set; }

        public override byte[] AbstractImage
        {
            get
            {
                if (base.AbstractImage == null)
                {
                    base.AbstractImage = WordUtility.ExtractGraphicPart(Container, Graphic);
                }
                return base.AbstractImage;
            }

            set
            {
                base.AbstractImage = value;
            }
        }
    }
}
