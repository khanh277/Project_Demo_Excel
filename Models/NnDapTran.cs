using System;
using System.Collections.Generic;

namespace Project_Demo_Excel.Models;

public partial class NnDapTran
{
    public int Id { get; set; }

    public int? Objectid { get; set; }

    public string? Siteid { get; set; }

    public string? MaDapTran { get; set; }

    public string? TenDapTran { get; set; }

    public string? MucDichSuDung { get; set; }

    public string? DiaChi { get; set; }

    public string? ChieuCao { get; set; }

    public string? ChieuDaiDaKienCo { get; set; }

    public string? DonViQuanLy { get; set; }

    public int? NamBatDauHoatDong { get; set; }

    public string? Regioncode { get; set; }

    public string? Regionpathcode { get; set; }

    public string? Regionname { get; set; }

    public string? Regionfullname { get; set; }

    public string? LoaiDap { get; set; }

    public string? TinhTrang { get; set; }

    public string? DienTichLuuVuc { get; set; }

    public string? ChieuDaiMuongDat { get; set; }

    public string? NamXayDung { get; set; }

    public string? GhiChu { get; set; }

    public string? CapCongTrinhThuyLoi { get; set; }
}
