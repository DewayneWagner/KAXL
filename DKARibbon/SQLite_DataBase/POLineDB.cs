using EXPREP_V2;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EXPREP_V2.Status;

namespace DKARibbon.SQLite_DataBase
{
    class POLineDB
    {
        public POLineDB(string po, double num) { Key = po + Convert.ToString(num); }
        public string Key { get; set; }
        public string PONumber { get; set; }
        public double LineNumber { get; set; }
        public CleanStatusE Status { get; set; }
        public DateTime MostRecentlyScheduledDeliveryDate { get; set; }
        public DateTime POCreatedDate { get; set; }
        public string VendorName { get; set; }
        public double Quantity { get; set; }
        public string ItemNum { get; set; }
        public double UnitPrice { get; set; }
        public bool IsICO { get; set; }

        public string GetKey(string po, double num) => po + Convert.ToString(num);

    }
    class POLinesDBConfig : IEntityTypeConfiguration<POLineDB>
    {
        public void Configure(EntityTypeBuilder<POLineDB> builder)
        {
            builder.HasKey(p => p.Key);

            builder.Property(p => p.Status)
                .HasConversion(new EnumToStringConverter<CleanStatusE>());
        }
    }
    class ItemDBConfig : IEntityTypeConfiguration<Item>
    {
        public void Configure(EntityTypeBuilder<Item> builder)
        {
            builder.HasKey(i => i.Num);
        }
    }
}
