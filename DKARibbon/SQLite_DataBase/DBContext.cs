using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using EXPREP_V2;

namespace DKARibbon.SQLite_DataBase
{
    class DBContext : DbContext
    {
        private const string _sqLiteDBPath = @"R:\Supply Chain\ProcurementDB";
        private const string _fileName = @"Procurement_DB.sqlite";
        private bool _isCreated;
        public DBContext()
        {
            if (!_isCreated)
            {
                _isCreated = true;
                Database.EnsureDeleted();
                Database.EnsureCreated();
            }
        }
        public virtual DbSet<POLineDB> POlines { get; set; }
        public virtual DbSet<Item> Items { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            string dbPath = Path.Combine(_sqLiteDBPath, _fileName);
            optionsBuilder.UseSqlite($"Filename={dbPath}");
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyConfiguration(new POLinesDBConfig());
            modelBuilder.ApplyConfiguration(new ItemDBConfig());
        }        
    }
}
