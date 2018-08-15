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
    public class SubstanceInfo
    {
        /// <summary>
        /// Order in master.
        /// </summary>
        public int? Order { get; set; }

        /// <summary>
        /// CAS registry number.
        /// </summary>
        internal string _CASRN = null;
        public string CASRN
        {
            get { return _CASRN; }
            set { _CASRN = value; }
        }
        /// <summary>
        /// Name described in a document.
        /// </summary>
        internal string _Name = null;
        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }
        
        internal string _Keywords = null;
        /// <summary>
        /// Keywords separated by <value>';'.</value>
        /// </summary>
        public string Keywords
        {
            get { return _Keywords; }
            set { _Keywords = value; }
        }
        
        internal string _CAIndexName = null;
        /// <summary>
        /// Name indexed in CAS
        /// </summary>
        public string CAIndexName
        {
            get { return _CAIndexName; }
            set { _CAIndexName = value; }
        }
        
        internal string _MolecularFormula = null;
        public string MolecularFormula
        {
            get { return _MolecularFormula; }
            set { _MolecularFormula = value; }
        }

        internal string _ClassIdentifier;
        public string ClassIdentifier
        {
            get { return _ClassIdentifier; }
            set { _ClassIdentifier = value; }
        }

        internal ICollection<string> _OtherNames;
        public ICollection<string> OtherNames
        {
            get { return _OtherNames; }
            set { _OtherNames = value; }
        }

        public virtual byte[] Bitmap { get; set; }

        internal string _Copyright;
        public string Copyright
        {
            get { return _Copyright; }
            set { _Copyright = value; }
        }
    }

    public class ASubstanceInfo
        : SubstanceInfo
    {
        internal A.Graphic Graphic { get; set; }
        internal OpenXmlPartContainer Container { get; set; }

        public override byte[] Bitmap
        {
            get
            {
                if (base.Bitmap == null)
                {
                    base.Bitmap = WordUtility.ExtractGraphicPart(Container, Graphic);
                }
                return base.Bitmap;
            }
            set
            {
                base.Bitmap = value;
            }
        }
    }
}
