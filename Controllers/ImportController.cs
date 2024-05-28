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
        List<string> errorMessages = new List<string>();

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

                bool importResult = ImportData(path, dataType, out count, errorMessages);
                if (importResult)
                {
                    message = $"Successfully imported {count} records.";
                    if (errorMessages.Count > 0)
                    {
                        message += " However, there were some errors:\n" + string.Join("\n", errorMessages);
                    }
                }
                else
                {
                    message = "Failed to import data.";
                }
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

    public bool ImportData(string filePath, string dataType, out int count, List<string> errorMessages)
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
                        ImportVung(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "Tinh":
                        ImportTinh(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "Huyen":
                        ImportHuyen(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "Xa":
                        ImportXa(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "DiaBan":
                        ImportDiaBan(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "ThongTinHo":
                        ImportTTinHo(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "ThanhVienTrongHo":
                        ImportThanhVienTrongHo(workSheet, startRow, startColumn, out count, errorMessages);
                        break;
                    case "ThongTinTuVong":
                        ImportThongTinTuVong(workSheet, startRow, startColumn, out count, errorMessages);
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

    private void ImportVung(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
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
                    vungs.Add(vung);
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

    private void ImportTinh(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
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
                    tinhs.Add(tinh);
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

    private void ImportHuyen(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
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
                    huyens.Add(huyen);
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

    private void ImportXa(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
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
                    xas.Add(xa);
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

    private void ImportDiaBan(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
    {
        count = 0;
        var diaBans = new List<DIABAN>();
        try
        {
            while (true)
            {
                var maTinh = workSheet.Cells[startRow, startColumn].Value?.ToString();
                var maHuyen = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();
                var maXa = workSheet.Cells[startRow, startColumn + 2].Value?.ToString();
                var tenDiaBan = workSheet.Cells[startRow, startColumn + 3].Value?.ToString();

                if (string.IsNullOrEmpty(maXa) || string.IsNullOrEmpty(maTinh) || string.IsNullOrEmpty(tenDiaBan))
                    break;

                if (!db.DIABANs.Any(x => x.TenDB == tenDiaBan))
                {
                    var dBan = new DIABAN { MaTinh = maTinh, MaHuyen = maHuyen, MaXa = maXa, TenDB = tenDiaBan };
                    diaBans.Add(dBan);
                    count++;
                }

                startRow++;
            }
            if (diaBans.Count > 0)
            {
                db.DIABANs.AddRange(diaBans);
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

    private void ImportTTinHo(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
    {
        count = 0;
        var hos = new List<THONGTINHO>();

        while (true)
        {
            var maHo = workSheet.Cells[startRow, startColumn].Value?.ToString();
            var maTinh = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();
            var maHuyen = workSheet.Cells[startRow, startColumn + 2].Value?.ToString();
            var maXa = workSheet.Cells[startRow, startColumn + 3].Value?.ToString();
            var maVung = workSheet.Cells[startRow, startColumn + 4].Value?.ToString();
            var tenDiaBan = workSheet.Cells[startRow, startColumn + 5].Value?.ToString();
            var hoSo = workSheet.Cells[startRow, startColumn + 6].Value?.ToString();
            var nam = workSheet.Cells[startRow, startColumn + 7].Value?.ToString();
            var hoTenChuHo = workSheet.Cells[startRow, startColumn + 8].Value?.ToString();
            var diaChi = workSheet.Cells[startRow, startColumn + 9].Value?.ToString();
            var maNhanVien = workSheet.Cells[startRow, startColumn + 10].Value?.ToString();
            var ngayKThuc = workSheet.Cells[startRow, startColumn + 11].Value?.ToString();
            var ngayPhVan = workSheet.Cells[startRow, startColumn + 12].Value?.ToString();
            var kinhDo = workSheet.Cells[startRow, startColumn + 13].Value?.ToString();
            var viDo = workSheet.Cells[startRow, startColumn + 14].Value?.ToString();
            var sdt = workSheet.Cells[startRow, startColumn + 15].Value?.ToString();
            var tsnk = workSheet.Cells[startRow, startColumn + 16].Value?.ToString();
            var tsnam = workSheet.Cells[startRow, startColumn + 17].Value?.ToString();
            var tsnu = workSheet.Cells[startRow, startColumn + 18].Value?.ToString();
            var kt9 = workSheet.Cells[startRow, startColumn + 19].Value?.ToString();
            var c45 = workSheet.Cells[startRow, startColumn + 20].Value?.ToString();
            var kt14 = workSheet.Cells[startRow, startColumn + 21].Value?.ToString();
            var nguoiXN = workSheet.Cells[startRow, startColumn + 22].Value?.ToString();
            var nguoiTao = workSheet.Cells[startRow, startColumn + 23].Value?.ToString();
            var ngayTao = workSheet.Cells[startRow, startColumn + 24].Value?.ToString();
            var phienBan = workSheet.Cells[startRow, startColumn + 25].Value?.ToString();

            if (string.IsNullOrEmpty(maHo) || string.IsNullOrEmpty(maTinh) || string.IsNullOrEmpty(maHuyen) || string.IsNullOrEmpty(maXa) || string.IsNullOrEmpty(tenDiaBan) || string.IsNullOrEmpty(hoTenChuHo))
                break;

            if (!db.THONGTINHOes.Any(h => h.MaHo == maHo))
            {
                var ho = new THONGTINHO
                {
                    MaHo = maHo,
                    MaTinh = maTinh,
                    MaHuyen = maHuyen,
                    MaXa = maXa,
                    MaVung = maVung,
                    TenDB = tenDiaBan,
                    HoSo = hoSo,
                    Nam = nam,
                    HoTenChuHo = hoTenChuHo,
                    DiaChi = diaChi,
                    MaNV = maNhanVien,
                    NgayKThuc = !string.IsNullOrEmpty(ngayKThuc) ? (DateTime?)DateTime.Parse(ngayKThuc) : null,
                    NgayPVan = !string.IsNullOrEmpty(ngayPhVan) ? (DateTime?)DateTime.Parse(ngayPhVan) : null,
                    KinhDo = !string.IsNullOrEmpty(kinhDo) ? (decimal?)decimal.Parse(kinhDo) : null,
                    ViDo = !string.IsNullOrEmpty(viDo) ? (decimal?)decimal.Parse(viDo) : null,
                    SDT = sdt,
                    TSNK = tsnk,
                    TSNAM = tsnam,
                    TSNU = tsnu,
                    KT9 = kt9,
                    C45 = c45,
                    KT14 = kt14,
                    NguoiXN = nguoiXN,
                    NguoiTao = nguoiTao,
                    NgayTao = !string.IsNullOrEmpty(ngayTao) ? (TimeSpan?)TimeSpan.Parse(ngayTao) : null,
                    PhienBan = phienBan
                };
                hos.Add(ho);
                count++;
            }
            else
            {
                errorMessages.Add($"Row {startRow}: MaHo {maHo} already exists.");
            }

            startRow++;
        }

        try
        {
            if (hos.Count > 0)
            {
                db.THONGTINHOes.AddRange(hos);
                db.SaveChanges();
            }
        }
        catch (DbEntityValidationException ex)
        {
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    errorMessages.Add($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
        }
    }

    private void ImportThanhVienTrongHo(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
    {
        count = 0;
        var thanhViens = new List<THANHVIENTRONGHO>();

        // Danh sách các thuộc tính của THANHVIENTRONGHO
        var properties = typeof(THANHVIENTRONGHO).GetProperties();
        var propertyNames = properties.Select(p => p.Name).ToArray();

        while (true)
        {
            var thanhVien = new THANHVIENTRONGHO();
            bool isEmptyRow = true;

            for (int i = 0; i < propertyNames.Length; i++)
            {
                var propertyName = propertyNames[i];
                var cellValue = workSheet.Cells[startRow, startColumn + i].Value?.ToString();

                if (!string.IsNullOrEmpty(cellValue))
                {
                    isEmptyRow = false;
                }

                var property = properties.FirstOrDefault(p => p.Name == propertyName);
                if (property != null)
                {
                    // Chuyển đổi giá trị từ chuỗi sang kiểu dữ liệu của thuộc tính
                    object convertedValue = Convert.ChangeType(cellValue, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
                    property.SetValue(thanhVien, convertedValue);
                }
            }

            if (isEmptyRow)
            {
                break; // Nếu hàng trống, kết thúc vòng lặp
            }

            thanhViens.Add(thanhVien);
            count++;
            startRow++;
        }

        // Thêm danh sách các thành viên vào cơ sở dữ liệu
        using (var context = new Db_ThucTapEntities())
        {
            context.THANHVIENTRONGHOes.AddRange(thanhViens);
            context.SaveChanges();
        }
    }

    private void ImportThongTinTuVong(ExcelWorksheet workSheet, int startRow, int startColumn, out int count, List<string> errorMessages)
    {
        count = 0;
        var thongTinTuVongs = new List<THONGTINTUVONG>();

        while (true)
        {
            var maHo = workSheet.Cells[startRow, startColumn].Value?.ToString();
            var maTVong = workSheet.Cells[startRow, startColumn + 1].Value?.ToString();
            var sttTV = workSheet.Cells[startRow, startColumn + 2].Value?.ToString();
            var hoTenTV = workSheet.Cells[startRow, startColumn + 3].Value?.ToString();
            var gioiTinh = workSheet.Cells[startRow, startColumn + 4].Value?.ToString() == "1"; // Assuming 1 for true, 0 for false
            var thangTV = workSheet.Cells[startRow, startColumn + 5].Value?.ToString();
            var namTV = workSheet.Cells[startRow, startColumn + 6].Value?.ToString();
            var thangSinh = workSheet.Cells[startRow, startColumn + 7].Value?.ToString();
            var namSinh = workSheet.Cells[startRow, startColumn + 8].Value?.ToString();
            var tuoi = workSheet.Cells[startRow, startColumn + 9].Value?.ToString();
            var nguyenNhan = workSheet.Cells[startRow, startColumn + 10].Value?.ToString();

            if (string.IsNullOrEmpty(maHo) || string.IsNullOrEmpty(maTVong))
                break;

            if (!db.THONGTINTUVONGs.Any(t => t.MaHo == maHo && t.MaTVong == maTVong))
            {
                var thongTin = new THONGTINTUVONG
                {
                    MaHo = maHo,
                    MaTVong = maTVong,
                    STTTV = sttTV,
                    HoTenTV = hoTenTV,
                    GioiTinh = gioiTinh,
                    ThangTV = thangTV,
                    NamTV = namTV,
                    ThangSinh = thangSinh,
                    NamSinh = namSinh,
                    Tuoi = tuoi,
                    NguyenNhan = nguyenNhan
                };
                thongTinTuVongs.Add(thongTin);
                count++;
            }
            else
            {
                errorMessages.Add($"Row {startRow}: MaHo {maHo} and MaTVong {maTVong} already exists.");
            }

            startRow++;
        }

        try
        {
            if (thongTinTuVongs.Count > 0)
            {
                db.THONGTINTUVONGs.AddRange(thongTinTuVongs);
                db.SaveChanges();
            }
        }
        catch (DbEntityValidationException ex)
        {
            foreach (var validationErrors in ex.EntityValidationErrors)
            {
                foreach (var validationError in validationErrors.ValidationErrors)
                {
                    errorMessages.Add($"Property: {validationError.PropertyName} Error: {validationError.ErrorMessage}");
                }
            }
        }
    }
}

