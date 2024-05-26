using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Import_file_Excel.Data;
using OfficeOpenXml;

public class ImportController : Controller
{
    private Db_ThucTapEntities db = new Db_ThucTapEntities();

    // GET: Import
    public ActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public ActionResult Index(HttpPostedFileBase excelFile, string dataType)
    {
        string message = string.Empty;
        int count = 0;

        if (excelFile != null && excelFile.ContentLength > 0 &&
            (excelFile.FileName.EndsWith(".xls") || excelFile.FileName.EndsWith(".xlsx")))
        {
            try
            {
                string uploadsDir = Server.MapPath("~/App_Data/uploads");

                if (!Directory.Exists(uploadsDir))
                {
                    Directory.CreateDirectory(uploadsDir);
                }

                string path = Path.Combine(uploadsDir, Path.GetFileName(excelFile.FileName));
                excelFile.SaveAs(path);

                bool importResult = ImportData(path, dataType, out count);
                message = importResult ? $"Successfully imported {count} records." : "Failed to import data.";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in file upload: {ex.Message}");
                message = "An error occurred while processing the file.";
            }
        }
        else
        {
            message = "Invalid file format. Please upload an Excel file.";
        }

        ViewBag.Message = message;
        return View();
    }

    public bool ImportData(string filePath, string dataType, out int count)
    {
        var result = false;
        count = 0;
        try
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                int startColumn = 1;
                int startRow = 2;
                ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
                object data = null;

                switch (dataType)
                {
                    case "Vung":
                        ImportVung(workSheet, startRow, startColumn, out count);
                        break;
                    case "Tinh":
                        ImportTinh(workSheet, startRow, startColumn, out count);
                        break;
                    case "Huyen":
                        ImportHuyen(workSheet, startRow, startColumn, out count);
                        break;
                    case "Xa":
                        ImportXa(workSheet, startRow, startColumn, out count);
                        break;
                }

                result = true;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Exception in ImportData: {ex.Message}");
        }
        return result;
    }

    private void ImportVung(ExcelWorksheet workSheet, int startRow, int startColumn, out int count)
    {
        count = 0;
        var vungs = new List<VUNG>();
        try
        {
            while (true)
            {
                var maVung = workSheet.Cells[startRow, startColumn].Value?.ToString();
                var tenVung = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();

                if (string.IsNullOrEmpty(maVung) || string.IsNullOrEmpty(tenVung))
                    break;

                if (!db.VUNGs.Any(v => v.MaVung == maVung))
                {
                    var vung = new VUNG { MaVung = maVung, TenVung = tenVung };
                    db.VUNGs.Add(vung);
                    count++;
                }

                startRow++;
            }
            if (vungs.Count > 0)
            {
                db.VUNGs.AddRange(vungs);
                db.SaveChanges();
            }
        }
        catch (DbEntityValidationException ex)
        {
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    System.Diagnostics.Debug.WriteLine($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
        }
    }

    private void ImportTinh(ExcelWorksheet workSheet, int startRow, int startColumn, out int count)
    {
        count = 0;
        var tinhs = new List<TINH>();
        try
        {
            while (true)
            {
                var maTinh = workSheet.Cells[startRow, startColumn].Value?.ToString();
                var maVung = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();
                var tenTinh = workSheet.Cells[startRow, startColumn + 2].Value?.ToString();

                if (string.IsNullOrEmpty(maTinh) || string.IsNullOrEmpty(maVung) || string.IsNullOrEmpty(tenTinh))
                    break;

                if (!db.TINHs.Any(t => t.MaTinh == maTinh))
                {
                    var tinh = new TINH { MaTinh = maTinh, MaVung = maVung, TenTinh = tenTinh };
                    db.TINHs.Add(tinh);
                    count++;
                }

                startRow++;
            }
            if (tinhs.Count > 0)
            {
                db.TINHs.AddRange(tinhs);
                db.SaveChanges();
            }
        }
        catch (DbEntityValidationException ex)
        {
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    System.Diagnostics.Debug.WriteLine($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
        }
    }

    private void ImportHuyen(ExcelWorksheet workSheet, int startRow, int startColumn, out int count)
    {
        count = 0;
        var huyens = new List<HUYEN>();
        try
        {
            while (true)
            {
                var maHuyen = workSheet.Cells[startRow, startColumn].Value?.ToString();
                var maTinh = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();
                var tenHuyen = workSheet.Cells[startRow, startColumn + 2].Value?.ToString();

                if (string.IsNullOrEmpty(maHuyen) || string.IsNullOrEmpty(maTinh) || string.IsNullOrEmpty(tenHuyen))
                    break;

                if (!db.HUYENs.Any(h => h.MaHuyen == maHuyen))
                {
                    var huyen = new HUYEN { MaHuyen = maHuyen, MaTinh = maTinh, TenHuyen = tenHuyen };
                    db.HUYENs.Add(huyen);
                    count++;
                }

                startRow++;
            }
            if (huyens.Count > 0)
            {
                db.HUYENs.AddRange(huyens);
                db.SaveChanges();
            }
        }
        catch (DbEntityValidationException ex)
        {
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    System.Diagnostics.Debug.WriteLine($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
        }
    }

    private void ImportXa(ExcelWorksheet workSheet, int startRow, int startColumn, out int count)
    {
        count = 0;
        var xas = new List<XA>();
        try
        {
            while (true)
            {
                var maXa = workSheet.Cells[startRow, startColumn].Value?.ToString();
                var maHuyen = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();
                var maTinh = workSheet.Cells[startRow, startColumn + 2].Value?.ToString();
                var tenXa = workSheet.Cells[startRow, startColumn + 3].Value?.ToString();

                if (string.IsNullOrEmpty(maXa) || string.IsNullOrEmpty(maTinh) || string.IsNullOrEmpty(tenXa))
                    break;

                if (!db.XAs.Any(x => x.MaXa == maXa))
                {
                    var xa = new XA { MaXa = maXa, MaHuyen = maHuyen, MaTinh = maTinh, TenXa = tenXa };
                    db.XAs.Add(xa);
                    count++;
                }

                startRow++;
            }
            if (xas.Count > 0)
            {
                db.XAs.AddRange(xas);
                db.SaveChanges();
            }
        }
        catch (DbEntityValidationException ex)
        {
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    System.Diagnostics.Debug.WriteLine($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
        }
    }
}
