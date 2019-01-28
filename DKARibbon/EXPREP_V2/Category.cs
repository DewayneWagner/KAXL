using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace EXPREP_V2
{
    public class Category
    {
        Master M;
        public Category(Master m) => M = m;

        public Category() { }

        public Category(string cat, Item item, Master M)
        {
            if(item.Cat != null)
            {
                CleanCategory = item.Cat;
            }
            else
            {
                DirtyCategory = cat;
                try
                {
                    CleanCategory = M.CategoryReferenceDictionary[cat].CleanCategory;
                }
                catch
                {
                    CleanCategory = "Fix";
                }
            }            
        }
        public Category(string dirty, string clean)
        {
            DirtyCategory = dirty;
            CleanCategory = clean; 
        }

        public string DirtyCategory { get; set; }
        public string CleanCategory { get; set; }
        public CategoryReferenceDictionary CRD { get; set; }
    }
    public class CategoryReferenceDictionary
    {
        private const int RowQ = 43;
        private const int ColQ = 2;

        public CategoryReferenceDictionary() { }

        private readonly Dictionary<string, Category> categoryReferenceDictionary;

        private Master M;

        public CategoryReferenceDictionary(Master m)
        {
            M = m;
            Category cat = new Category(M);

            string[,] CatRefA = new string[RowQ, ColQ]
            {
            {"Anode","Anode"},
            {"Battery","Battery"},
            {"Cable","Cable"},
            {"Cases","Cases"},
            {"Coke","Coke" },
            {"AFE Laptops and Accessories","Computers" },
            {"Computer","Computers" },
            {"Computer Cost Expense","Computers"},
            {"Electronics","Electronics"},
            {"Fasteners","Fasteners"},
            {"Fixed Asset Purchase","Fixed Asset Purchase"},
            {"Foam","Miscellaneous"},
            {"Intercompany","ICO" },
            { "Machined Parts","Machined Parts" },
            { "Material-Project","Material-Project" },
            { "Material-R&D","Material-R&D" },
            { "Adhesive","Miscellaneous" },
            { "Advertising Expense","Miscellaneous" },
            { "Assembly","Miscellaneous" },
            { "Label","Miscellaneous" },
            { "Miscellaneous","Miscellaneous" },
            { "Net","Miscellaneous" },
            { "NoSpendItems","Miscellaneous" },
            { "Product Demo Expenses","Miscellaneous" },
            { "Promotional Materials Expense","Miscellaneous" },
            { "PT Procurement","Miscellaneous" },
            { "Repair Fixed Asset","Miscellaneous" },
            { "Shipping /Freight","Miscellaneous" },
            { "Tradeshows /Conventions Expense","Miscellaneous" },
            { "Urethane","Miscellaneous" },
            { "Office Supplies","Office Supplies" },
            { "O -Ring","O-Ring" },
            { "Personal Protection Equipment Expense","PPE" },
            { "PPE","PPE" },
            { "Pure Parts","Custom Parts" },
            { "Custom","Custom Parts" },
            { "Pure","Custom Parts" },
            { "PVF", "PVF" },
            { "Service","Service" },
            { "Shop Supply Expense","Shop Supply Expense" },
            { "Subcontractor","Subcontractor" },
            { "Small Tools Expense","Tools" },
            { "Tools","Tools" },
            };

            categoryReferenceDictionary = new Dictionary<string, Category>();

            int length = CatRefA.GetLength(0);

            for (int i = 0; i < length; i++)
            {
                Category c = new Category(CatRefA[i, 0], CatRefA[i, 1]);
                categoryReferenceDictionary.Add(c.DirtyCategory, c);
            }
        }
        public Category this[string key] => key != null && categoryReferenceDictionary.ContainsKey(key)? 
            categoryReferenceDictionary[key] : null;
    }
}
