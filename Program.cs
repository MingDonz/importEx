using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using ExcelImport.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelImport
{
    class Program
    {
        static void Main(string[] args)
        {
            var cs = @"Data Source=(LocalDb);Initial Catalog=dbo;Integrated Security=True";
            SqlConnection conn = new SqlConnection(cs);

            List<Phone> phones = new List<Phone>();
            FileStream fs = new FileStream(@"E:\testwb.xlsx", FileMode.Open);
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            ISheet sheet = wb.GetSheetAt(0);
            int rowIndex = 2;
            foreach (var p in phones)
            {
                // lấy row hiện tại
                var nowRow = sheet.GetRow(rowIndex);

                var a_install = Convert.ToInt32(nowRow.GetCell(1).StringCellValue);
                var a_uninstall = Convert.ToInt32(nowRow.GetCell(2).StringCellValue);
                var i_installDevices = Convert.ToInt32(nowRow.GetCell(3).StringCellValue);
                var i_uninstallDevices = Convert.ToInt32(nowRow.GetCell(4).StringCellValue);

                phones.Add(new Phone()
                {
                    android_Install = a_install,
                    android_Uninstall = a_uninstall,
                    iOS_Install = i_installDevices,
                    iOS_Uninstall = i_uninstallDevices
                });

                // tăng index khi lấy xong
                rowIndex++;
            }
           
                SqlCommand cmd = new SqlCommand(@"CREATE TABLE [dbo].[f_device_install](
	            [Id] [int] IDENTITY(1,1) NOT NULL,
	            [date] [datetime] NOT NULL,
	            [date_key] [nvarchar](8) NOT NULL,
	            [android_install] [int] NULL,
	            [iOS_install] [int] NULL,
	            [android_uninstall] [int] NULL,
	            [iOS_uninstall] [int] NULL,
	            [CreatedAt] [datetime2](7) NOT NULL,
	            [CreatedBy] [nvarchar](50) NULL,
	            [UpdatedAt] [datetime2](7) NOT NULL,
	            [UpdatedBy] [nvarchar](50) NULL,
                PRIMARY KEY CLUSTERED ([Id] ASC)
                WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]) ON [PRIMARY];", conn);
            conn.Open();
            cmd.ExecuteNonQuery();

            string insertSt = @"INSERT INTO dbo.f_device_install (Id, date, date_key, android_install, iOS_install, android_uninstall, iOS_uninstall) 
                                    VALUES (@Id, @date, @date_key, @android_install, @iOS_install, @android_uninstall, @iOS_uninstall);";
            var command = new SqlCommand(insertSt);
            foreach (var items in phones)
            {

                command.Parameters.AddWithValue("@Id", (SqlDbType)items.Id);
                command.Parameters.AddWithValue("@date_key", items.DateKey);
                command.Parameters.AddWithValue("@android_install", (SqlDbType)items.android_Install);
                command.Parameters.AddWithValue("@iOS_install", (SqlDbType)items.iOS_Install);
                command.Parameters.AddWithValue("@android_uninstall", (SqlDbType)items.android_Uninstall);
                command.Parameters.AddWithValue("@iOS_uninstall", (SqlDbType)items.iOS_Uninstall);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            using (var context = new PhoneContext())
            {
                context.BulkInsert(phones);
            }
        }





        public class Phone
        {
            public int Id { get; set; }
            public DateTime Date { get; set; }
            public string DateKey { get; set; }
            public int android_Install { get; set; }
            public int iOS_Install { get; set; }
            public int android_Uninstall { get; set; }
            public int iOS_Uninstall { get; set; }
            public DateTime CreatedAt { get; set; }
            public string CreatedBy { get; set; }
            public DateTime UpdatedAt { get; set; }
            public string UpdateBy { get; set; }

        }
        class PhoneContext : DbContext
        {
            public DbSet<Phone> Phones { get; set; }

            protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
            {
                optionsBuilder.UseSqlServer(@"Data Source=LocalDb;Initial Catalog=Db;Integrated Security=True");
                base.OnConfiguring(optionsBuilder);
            }

            internal void BulkInsert(List<Phone> phones)
            {
                throw new NotImplementedException();
            }
        }
    }
}
