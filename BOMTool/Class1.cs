using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;

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
        public void BeginReformatBomForExcel(ref List<BomRowItem> InventorBomList)
        {
            //MessageBox.Show("Inventor Bom list count =" + InventorBomList.Count);

            var grouped = InventorBomList.OrderBy(x => x.BomRowType).GroupBy(x => x.BomRowType);
            //InventorBomList.RemoveRange(0, InventorBomList.Count);
            // InventorBomList.RemoveAll(NotEmpty);
            MessageBox.Show("InventorBomList.Count= " + InventorBomList.Count);
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
                            item.ItemNo = 0;
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
                        default:
                            break;
                    }
                }
            }
            //hopefully sort by ItemNo
            BomList.OrderBy(x => x.ItemNo);
            MessageBox.Show("BomList.Count= " + BomList.Count);
            InventorBomList = BomList;
        }

        private bool NotEmpty(BomRowItem obj)
        {
            return obj.BomRowType > 0;
        }
    }
    public class BomRowItem :IComparable<BomRowItem>
    {
        public string PartNo;
        public string Descr;
        public string Rev;
        public long ItemNo;
        public string Classification;
        public string Material;
        public long Qty;
        public string Vendor;
        public string Comments;
        public long BomRowType;
        public List<BomRowItem> Children;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m_partNo">Part Number</param>
        /// <param name="m_descr">Description</param>
        /// <param name="m_rev">Revision</param>
        /// <param name="m_itemno">Item Number</param>
        /// <param name="m_classification">Classification</param>
        /// <param name="m_material">Material</param>
        /// <param name="m_quantity">Quantity</param>
        /// <param name="m_vendor">Vendor</param>
        /// <param name="m_comments">Comments</param>
        /// <param name="m_bomrowtype">Bom Row Type</param>
        /// <param name="m_children">Children?</param>
        public BomRowItem(
            string m_partNo, 
            string m_descr, 
            string m_rev, 
            long m_itemno,
            string m_classification,
            string m_material,
            long m_quantity,
            string m_vendor,
            string m_comments,
            long m_bomrowtype,
            List<BomRowItem> m_children = null)
        {
            this.PartNo = m_partNo;
            this.Descr = m_descr;
            this.Rev = m_rev;
            this.ItemNo = m_itemno;
            this.Classification = m_classification;
            this.Material = m_material;
            this.Qty = m_quantity;
            this.Vendor = m_vendor;
            this.Comments = m_comments;
            this.BomRowType = m_bomrowtype;
            this.Children = m_children;
        }
        public int CompareTo(BomRowItem other)
        {
            return this.CompareTo(other);
        }

        /// <summary>
        /// Split the collection into groups based on bomrowtype
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static List<List<BomRowItem>> Split(List<BomRowItem> source)
        {
            return source.Select((x, i) => new {
                Index = i,
                Value = x
            }).GroupBy(x => x.Index).Select(x => x.Select(v => v.Value).ToList()).ToList();
        }

        /// <summary>
        /// Allows for sorting of the BomRowItem class
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="KeySelector"></param>
        public void Sort<T>(Func<BomRowItem, T> KeySelector) where T : IComparable
        {
            BowRowItemComparer ic = new BowRowItemComparer();
            this.Children.Sort(ic);
        }

    }
    public class BowRowItemComparer : IComparer<BomRowItem>
    {
        //public BomRowItemComparer()
        //{
        
        //}
        public int Compare(BomRowItem x, BomRowItem y)
        {
            return x.BomRowType.CompareTo(y.BomRowType);
        }
    }
}
