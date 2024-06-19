using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace Project_Demo_Excel.Models
{
    public partial class DatabasePrjContext : DbContext
    {
        public DatabasePrjContext()
        {
        }

        public DatabasePrjContext(DbContextOptions<DatabasePrjContext> options)
            : base(options)
        {
        }

        public virtual DbSet<DtlObjectdatum> DtlObjectdata { get; set; }
        public virtual DbSet<NnDapTran> NnDapTrans { get; set; }
        public virtual DbSet<NnDienTichTuoi> NnDienTichTuois { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                var configuration = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                var connectionString = configuration.GetConnectionString("DBContext");
                optionsBuilder.UseSqlServer(connectionString);
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<DtlObjectdatum>(entity =>
            {
                entity.HasKey(e => e.Objectid).HasName("OBJECTDATA_P");

                entity.ToTable("DTL_OBJECTDATA");

                entity.Property(e => e.Objectid)
                    .ValueGeneratedNever()
                    .HasColumnName("OBJECTID");
                entity.Property(e => e.Catalogpath)
                    .HasMaxLength(1000)
                    .HasColumnName("CATALOGPATH");
                entity.Property(e => e.Cateid).HasColumnName("CATEID");
                entity.Property(e => e.Checkout)
                    .HasMaxLength(50)
                    .HasColumnName("CHECKOUT");
                entity.Property(e => e.Classname)
                    .HasMaxLength(100)
                    .HasColumnName("CLASSNAME");
                entity.Property(e => e.Collectionpath)
                    .HasMaxLength(1000)
                    .HasColumnName("COLLECTIONPATH");
                entity.Property(e => e.Createdby)
                    .HasMaxLength(200)
                    .HasColumnName("CREATEDBY");
                entity.Property(e => e.Createddate)
                    .HasColumnType("datetime")
                    .HasColumnName("CREATEDDATE");
                entity.Property(e => e.Drawid).HasColumnName("DRAWID");
                entity.Property(e => e.Indoorid).HasColumnName("INDOORID");
                entity.Property(e => e.Keywordcloud).HasColumnName("KEYWORDCLOUD");
                entity.Property(e => e.Langcode)
                    .HasMaxLength(5)
                    .HasColumnName("LANGCODE");
                entity.Property(e => e.Mainimage)
                    .HasMaxLength(500)
                    .HasColumnName("MAINIMAGE");
                entity.Property(e => e.Mapid).HasColumnName("MAPID");
                entity.Property(e => e.Modifiedby)
                    .HasMaxLength(200)
                    .HasColumnName("MODIFIEDBY");
                entity.Property(e => e.Modifieddate)
                    .HasColumnType("datetime")
                    .HasColumnName("MODIFIEDDATE");
                entity.Property(e => e.Note).HasColumnName("NOTE");
                entity.Property(e => e.Publishedby)
                    .HasMaxLength(50)
                    .HasColumnName("PUBLISHEDBY");
                entity.Property(e => e.Publisheddate)
                    .HasColumnType("datetime")
                    .HasColumnName("PUBLISHEDDATE");
                entity.Property(e => e.Publishedyn).HasColumnName("PUBLISHEDYN");
                entity.Property(e => e.Regioncode)
                    .HasMaxLength(50)
                    .HasColumnName("REGIONCODE");
                entity.Property(e => e.Regionfullname).HasColumnName("REGIONFULLNAME");
                entity.Property(e => e.Regionname)
                    .HasMaxLength(100)
                    .HasColumnName("REGIONNAME");
                entity.Property(e => e.Regionpathcode)
                    .HasMaxLength(100)
                    .HasColumnName("REGIONPATHCODE");
                entity.Property(e => e.Siteid)
                    .HasMaxLength(50)
                    .HasColumnName("SITEID");
                entity.Property(e => e.Title).HasColumnName("TITLE");
                entity.Property(e => e.Urlname)
                    .HasMaxLength(150)
                    .HasColumnName("URLNAME");
                entity.Property(e => e.Vtid).HasColumnName("VTID");
                entity.Property(e => e.Workflowid).HasColumnName("WORKFLOWID");
                entity.Property(e => e.Workflowstate).HasColumnName("WORKFLOWSTATE");
                entity.Property(e => e.Workflowundoyn).HasColumnName("WORKFLOWUNDOYN");
            });

            modelBuilder.Entity<NnDapTran>(entity =>
            {
                entity.HasKey(e => e.Id).HasName("PK__NN_DapTr__3214EC27893540BC");

                entity.ToTable("NN_DapTran");

                entity.Property(e => e.Id)
                    .ValueGeneratedNever()
                    .HasColumnName("ID");
                entity.Property(e => e.CapCongTrinhThuyLoi)
                    .HasMaxLength(500)
                    .HasColumnName("capCongTrinhThuyLoi");
                entity.Property(e => e.ChieuCao)
                    .HasMaxLength(50)
                    .HasColumnName("chieuCao");
                entity.Property(e => e.ChieuDaiDaKienCo)
                    .HasMaxLength(50)
                    .HasColumnName("chieuDaiDaKienCo");
                entity.Property(e => e.ChieuDaiMuongDat)
                    .HasMaxLength(50)
                    .HasColumnName("chieuDaiMuongDat");
                entity.Property(e => e.DiaChi)
                    .HasMaxLength(255)
                    .HasColumnName("diaChi");
                entity.Property(e => e.DienTichLuuVuc)
                    .HasMaxLength(50)
                    .HasColumnName("dienTichLuuVuc");
                entity.Property(e => e.DonViQuanLy)
                    .HasMaxLength(255)
                    .HasColumnName("donViQuanLy");
                entity.Property(e => e.GhiChu)
                    .HasMaxLength(100)
                    .HasColumnName("ghiChu");
                entity.Property(e => e.LoaiDap)
                    .HasMaxLength(1000)
                    .HasColumnName("loaiDap");
                entity.Property(e => e.MaDapTran)
                    .HasMaxLength(25)
                    .HasColumnName("maDapTran");
                entity.Property(e => e.MucDichSuDung)
                    .HasMaxLength(255)
                    .HasColumnName("mucDichSuDung");
                entity.Property(e => e.NamBatDauHoatDong).HasColumnName("namBatDauHoatDong");
                entity.Property(e => e.NamXayDung)
                    .HasMaxLength(50)
                    .HasColumnName("namXayDung");
                entity.Property(e => e.Objectid).HasColumnName("OBJECTID");
                entity.Property(e => e.Regioncode)
                    .HasMaxLength(50)
                    .HasColumnName("REGIONCODE");
                entity.Property(e => e.Regionfullname)
                    .HasMaxLength(1000)
                    .HasColumnName("REGIONFULLNAME");
                entity.Property(e => e.Regionname)
                    .HasMaxLength(100)
                    .HasColumnName("REGIONNAME");
                entity.Property(e => e.Regionpathcode)
                    .HasMaxLength(100)
                    .HasColumnName("REGIONPATHCODE");
                entity.Property(e => e.Siteid)
                    .HasMaxLength(10)
                    .HasColumnName("SITEID");
                entity.Property(e => e.TenDapTran)
                    .HasMaxLength(255)
                    .HasColumnName("tenDapTran");
                entity.Property(e => e.TinhTrang)
                    .HasMaxLength(1000)
                    .HasColumnName("tinhTrang");
            });

            modelBuilder.Entity<NnDienTichTuoi>(entity =>
            {
                entity.HasKey(e => e.Id).HasName("PK__NN_DienT__3214EC27D3649923");

                entity.ToTable("NN_DienTichTuoi");

                entity.Property(e => e.Id)
                    .ValueGeneratedNever()
                    .HasColumnName("ID");
                entity.Property(e => e.DauMoiDapTran)
                    .HasMaxLength(500)
                    .HasColumnName("dauMoiDapTran");
                entity.Property(e => e.DauMoiDuongOng)
                    .HasMaxLength(500)
                    .HasColumnName("dauMoiDuongOng");
                entity.Property(e => e.DauMoiHoChua)
                    .HasMaxLength(500)
                    .HasColumnName("dauMoiHoChua");
                entity.Property(e => e.DauMoiKeSong)
                    .HasMaxLength(500)
                    .HasColumnName("dauMoiKeSong");
                entity.Property(e => e.DauMoiKenhMuong)
                    .HasMaxLength(500)
                    .HasColumnName("dauMoiKenhMuong");
                entity.Property(e => e.DauMoiTramBom)
                    .HasMaxLength(500)
                    .HasColumnName("dauMoiTramBom");
                entity.Property(e => e.DienTichTuoiThietKe)
                    .HasMaxLength(50)
                    .HasColumnName("dienTichTuoiThietKe");
                entity.Property(e => e.DienTichTuoiThucTe)
                    .HasMaxLength(50)
                    .HasColumnName("dienTichTuoiThucTe");
                entity.Property(e => e.Nam).HasColumnName("nam");
                entity.Property(e => e.Objectid).HasColumnName("OBJECTID");
                entity.Property(e => e.Regioncode)
                    .HasMaxLength(50)
                    .HasColumnName("REGIONCODE");
                entity.Property(e => e.Regionfullname)
                    .HasMaxLength(1000)
                    .HasColumnName("REGIONFULLNAME");
                entity.Property(e => e.Regionname)
                    .HasMaxLength(100)
                    .HasColumnName("REGIONNAME");
                entity.Property(e => e.Regionpathcode)
                    .HasMaxLength(100)
                    .HasColumnName("REGIONPATHCODE");
                entity.Property(e => e.Siteid)
                    .HasMaxLength(10)
                    .HasColumnName("SITEID");
                entity.Property(e => e.TuyenDanNuoc)
                    .HasMaxLength(255)
                    .HasColumnName("tuyenDanNuoc");
                entity.Property(e => e.TuyenOng)
                    .HasMaxLength(100)
                    .HasColumnName("tuyenOng");
                entity.Property(e => e.VuTuoi)
                    .HasMaxLength(255)
                    .HasColumnName("vuTuoi");
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);

        public List<DapTranSummary> GetDapTranSummary()
        {
            var query = from nn_dt in NnDapTrans
                        join dtl_od in DtlObjectdata on nn_dt.Objectid equals dtl_od.Objectid into details
                        from dtl_od in details.DefaultIfEmpty()
                        join nn_dtt in NnDienTichTuois on dtl_od.Regioncode equals nn_dtt.Regioncode into areas
                        from nn_dtt in areas.DefaultIfEmpty()
                        select new
                        {
                            nn_dt.Regionname,
                            nn_dt.DienTichLuuVuc,
                            nn_dt.ChieuDaiDaKienCo,
                            nn_dt.ChieuDaiMuongDat,
                            nn_dt.DonViQuanLy,
                            nn_dt.LoaiDap,
                            nn_dtt.DienTichTuoiThietKe,
                            nn_dtt.DienTichTuoiThucTe
                        };

            // Step 2: Perform parsing and aggregation in memory
            var groupedData = query.AsEnumerable().GroupBy(x => x.Regionname).Select(g => new
            {
                RegionName = g.Key,
                TotalArea = (double)g.Sum(x => decimal.TryParse(x.DienTichLuuVuc, out decimal area) ? area : 0m),
                TotalDams = g.Count(),
                UBNDTown = g.Count(x => !string.IsNullOrWhiteSpace(x.DonViQuanLy) && x.DonViQuanLy.Trim().Contains("huyện", StringComparison.OrdinalIgnoreCase)),
                UBNDVillage = g.Count(x => !string.IsNullOrWhiteSpace(x.DonViQuanLy) && x.DonViQuanLy.Trim().Contains("xã", StringComparison.OrdinalIgnoreCase)),
                Company = g.Count(x => !string.IsNullOrWhiteSpace(x.DonViQuanLy) && x.DonViQuanLy.Trim().Contains("TNHH MTV", StringComparison.OrdinalIgnoreCase)),
                Others = g.Count(x => !string.IsNullOrWhiteSpace(x.DonViQuanLy) && !x.DonViQuanLy.Contains("UBND", StringComparison.OrdinalIgnoreCase) && !x.DonViQuanLy.Contains("TNHH MTV", StringComparison.OrdinalIgnoreCase)),
                SpecialDams = g.Count(x => x.LoaiDap == "Đập quan trọng đặc biệt"),
                LargeDams = g.Count(x => !string.IsNullOrWhiteSpace(x.LoaiDap) && x.LoaiDap.Contains("lớn", StringComparison.OrdinalIgnoreCase)),
                MediumDams = g.Count(x => !string.IsNullOrWhiteSpace(x.LoaiDap) && x.LoaiDap.Contains("vừa", StringComparison.OrdinalIgnoreCase)),
                SmallDams = g.Count(x => !string.IsNullOrWhiteSpace(x.LoaiDap) && x.LoaiDap.Contains("nhỏ", StringComparison.OrdinalIgnoreCase)),
                TotalLength = (double)g.Sum(x => (decimal.TryParse(x.ChieuDaiDaKienCo, out decimal length1) ? length1 : 0m) + (decimal.TryParse(x.ChieuDaiMuongDat, out decimal length2) ? length2 : 0m)),
                SolidBuilt = (double)g.Sum(x => decimal.TryParse(x.ChieuDaiDaKienCo, out decimal solid) ? solid : 0m),
                EarthChannels = (double)g.Sum(x => decimal.TryParse(x.ChieuDaiMuongDat, out decimal earth) ? earth : 0m),
                Plan = (double)g.Sum(x => x.DienTichTuoiThietKe != null && decimal.TryParse(x.DienTichTuoiThietKe, out decimal plan) ? plan : 0m),
                Actual = (double)g.Sum(x => x.DienTichTuoiThucTe != null && decimal.TryParse(x.DienTichTuoiThucTe, out decimal actual) ? actual : 0m)
            }).ToList();

            // Step 3: Add row numbers and map to DapTranSummary
            var dataWithRowNumber = groupedData.Select((item, index) => new DapTranSummary
            {
                STT = index + 1,
                RegionName = item.RegionName,
                TotalArea = item.TotalArea,
                TotalDams = item.TotalDams,
                UBNDTown = item.UBNDTown,
                UBNDVillage = item.UBNDVillage,
                Company = item.Company,
                Others = item.Others,
                SpecialDams = item.SpecialDams,
                LargeDams = item.LargeDams,
                MediumDams = item.MediumDams,
                SmallDams = item.SmallDams,
                TotalLength = item.TotalLength,
                SolidBuilt = item.SolidBuilt,
                EarthChannels = item.EarthChannels,
                Plan = item.Plan,
                Actual = item.Actual
            }).ToList();

            return dataWithRowNumber;

        }


        public class DapTranSummary
        {
            public int STT { get; set; }
            public string RegionName { get; set; }
            public double TotalArea { get; set; }
            public int TotalDams { get; set; }
            public int UBNDTown { get; set; }
            public int UBNDVillage { get; set; }
            public int Company { get; set; }
            public int Others { get; set; }
            public int SpecialDams { get; set; }
            public int LargeDams { get; set; }
            public int MediumDams { get; set; }
            public int SmallDams { get; set; }
            public double TotalLength { get; set; }
            public double SolidBuilt { get; set; }
            public double EarthChannels { get; set; }
            public double Plan { get; set; }
            public double Actual { get; set; }
        }
    }
}
