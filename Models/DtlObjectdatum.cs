using System;
using System.Collections.Generic;

namespace Project_Demo_Excel.Models;

public partial class DtlObjectdatum
{
    public int Objectid { get; set; }

    public int? Cateid { get; set; }

    public string? Createdby { get; set; }

    public DateTime? Createddate { get; set; }

    public string? Modifiedby { get; set; }

    public DateTime? Modifieddate { get; set; }

    public string? Checkout { get; set; }

    public string? Keywordcloud { get; set; }

    public string? Title { get; set; }

    public string? Note { get; set; }

    public string? Classname { get; set; }

    public string? Mainimage { get; set; }

    public string? Catalogpath { get; set; }

    public string? Urlname { get; set; }

    public int? Mapid { get; set; }

    public int? Drawid { get; set; }

    public int? Indoorid { get; set; }

    public int? Vtid { get; set; }

    public string? Langcode { get; set; }

    public string? Siteid { get; set; }

    public string? Regioncode { get; set; }

    public string? Regionpathcode { get; set; }

    public string? Regionname { get; set; }

    public string? Regionfullname { get; set; }

    public int? Workflowid { get; set; }

    public int? Workflowstate { get; set; }

    public bool? Workflowundoyn { get; set; }

    public DateTime? Publisheddate { get; set; }

    public string? Publishedby { get; set; }

    public bool? Publishedyn { get; set; }

    public string? Collectionpath { get; set; }
}
