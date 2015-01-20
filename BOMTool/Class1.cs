using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Inventor;
//using Autodesk.iLogic;
//using Inventor;

namespace BOMTool
{
    public class Class1
    {
        public static Inventor.Application m_inventorApplication;
        public List<BomRowItem> BomList = new List<BomRowItem>();
        public static Inventor.Application InventorApplication 
        { 
            get
            {
                return Class1.m_inventorApplication;
            }
            
            set
            {
                Class1.m_inventorApplication = value;
            }
        }

        /// <summary>
        /// Begins the reformatting of the Inventor BOM
        /// </summary>
        /// <param name="InventorBomList"></param>
        public void BeginReformatBomForExcel(ref List<BomRowItem> InventorBomList)
        {
            //MessageBox.Show("Inventor Bom list count =" + InventorBomList.Count);

            var grouped = InventorBomList.OrderBy(x => x.BomRowType).GroupBy(x => x.BomRowType);
            //InventorBomList.RemoveRange(0, InventorBomList.Count);
            // InventorBomList.RemoveAll(NotEmpty);
            //MessageBox.Show("InventorBomList.Count= " + InventorBomList.Count);
            int SubAssemblyInt = 1;
            int DetailedPartsInt = 200;
            int COTSContentImportedInt = 500;
            foreach (var group in grouped)
            {
                foreach (BomRowItem item in group)
                {
                    switch (item.BomRowType)
                    {   
                        case 1: //Specifications = no item number
                            item.ItemNo = 9999;
                            BomList.Add(item);
                            break;
                        case 2: // Sub assemblies = 1 to 199
                            item.ItemNo = SubAssemblyInt;
                            SubAssemblyInt++;
                            BomList.Add(item);
                            break;
                        case 3: // Detailed Parts = 200 to 500
                            item.ItemNo = DetailedPartsInt;
                            DetailedPartsInt++;
                            BomList.Add(item);
                            break;
                        case 4: // COTS Parts/Content Centre/Imported Components = 500 to 999
                            item.ItemNo = COTSContentImportedInt;
                            COTSContentImportedInt++;
                            BomList.Add(item);
                            break;
                        case 5: //Parent Assembly
                            MessageBox.Show("Should only be one of these!");
                            item.ItemNo = 0;
                            BomList.Add(item);
                            break;
                        default:
                            break;
                    }
                }
            }
            //hopefully sort by ItemNo
            BomList.OrderBy(x => x.ItemNo);
            //MessageBox.Show("BomList.Count= " + BomList.Count);
            InventorBomList = BomList;
        }

        /// <summary>
        /// Updates the Inventor item number.
        /// </summary>
        /// <param name="oBOMROWs">the BOMROWs collection</param>
        /// <param name="oSortedPartsList"></param>
        public void UpdateInventorPartsList(BOMRowsEnumerator oBOMROWs, List<BomRowItem> oSortedPartsList)
        {
            //MessageBox.Show("Reached UpdateInventorPartsList Sub");
            ComponentDefinition oCompdef;
            foreach (BOMRow oRow in oBOMROWs)
            {
                oCompdef = oRow.ComponentDefinitions[1];
                long itemNo = (from BomRowItem a in oSortedPartsList where a.FileName == oCompdef.Document.FullFileName select a.ItemNo).FirstOrDefault();
                if (itemNo == 0 || itemNo == 9999)
                {
                    oRow.ItemNumber = "";
                }
                else
                {
                    oRow.ItemNumber = itemNo.ToString();
                }
            }
        }
        /// <summary>
        /// Splits a string into lines based on max length
        /// </summary>
        /// <param name="stringToSplit">the string we want to split</param>
        /// <param name="maximumLineLength">int to determine max line length</param>
        /// <returns>an IEnumerable containing string values</returns>
        /// <remarks>copied from this page: https://stackoverflow.com/questions/22368434/best-way-to-split-string-into-lines-with-maximum-length-without-breaking-words 
        /// - had to modify it to be a List<> instead of an Enumerable<> as List<> allows the Count() </remarks>
        public List<string> SplitToLines(string stringToSplit, int maximumLineLength)
        {
            var words = stringToSplit.Split(' ').Concat(new[] { "" });
            return
                words
                    .Skip(1)
                    .Aggregate(
                        words.Take(1).ToList(),
                        (a, w) =>
                        {
                            var last = a.Last();
                            while (last.Length > maximumLineLength)
                            {
                                a[a.Count() - 1] = last.Substring(0, maximumLineLength);
                                last = last.Substring(maximumLineLength);
                                a.Add(last);
                            }
                            var test = last + " " + w;
                            if (test.Length > maximumLineLength)
                            {
                                a.Add(w);
                            }
                            else
                            {
                                a[a.Count() - 1] = test;
                            }
                            return a;
                        });
        }
    }
    public class BomRowItem
    {
        public string FileName { get; set; }
        public string PartNo { get; set; }
        public string Descr { get; set; }
        public string Rev { get; set; }
        public long ItemNo { get; set; }
        public string Classification { get; set; }
        public string Material { get; set; }
        public long Qty { get; set; }
        public string Vendor { get; set; }
        public string Comments { get; set; }
        public long BomRowType { get; set; }
        public string status { get; set; }
    }
}
