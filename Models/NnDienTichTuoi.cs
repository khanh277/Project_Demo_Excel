using System;
using System.Collections.Generic;

namespace Project_Demo_Excel.Models;

public partial class NnDienTichTuoi
{
    public int Id { get; set; }

    public int? Objectid { get; set; }

    public string? Siteid { get; set; }

    public string? VuTuoi { get; set; }

    public int? Nam { get; set; }

    public string? DienTichTuoiThietKe { get; set; }

    public string? DienTichTuoiThucTe { get; set; }

    public string? TuyenDanNuoc { get; set; }

    public string? TuyenOng { get; set; }

    public string? DauMoiDapTran { get; set; }

    public string? DauMoiHoChua { get; set; }

    public string? DauMoiTramBom { get; set; }

    public string? DauMoiKenhMuong { get; set; }

    public string? DauMoiDuongOng { get; set; }

    public string? DauMoiKeSong { get; set; }

    public string? Regioncode { get; set; }

    public string? Regionpathcode { get; set; }

    public string? Regionname { get; set; }

    public string? Regionfullname { get; set; }
}
