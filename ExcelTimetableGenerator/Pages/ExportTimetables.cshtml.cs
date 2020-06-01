using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTimetableGenerator.Data;
using ExcelTimetableGenerator.Models;
using ExcelTimetableGenerator.Shared;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using Microsoft.Extensions.Hosting;
using System.Security;
using System.Security.Permissions;
using Microsoft.AspNetCore.Authorization;

namespace ExcelTimetableGenerator.Pages
{
    [Authorize(Roles = "ALLSTAFF")]
    public class ExportTimetablesModel : PageModel
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _configuration;
        private IWebHostEnvironment _hostingEnvironment;
        public ExportTimetablesModel(
            ApplicationDbContext context,
            IConfiguration configuration,
            IWebHostEnvironment hostingEnvironment
            )
        {
            _context = context;
            _configuration = configuration;
            _hostingEnvironment = hostingEnvironment;
        }

        public string AcademicYear { get; set; }
        public string UserDetails { get; set; }
        public string UserGreeting { get; set; }
        public string SystemVersion { get; set; }
        public string PlanningSystem { get; set; }

        public string ExportParentPath { get; set; }
        public string ZipPath { get; set; }
        public string ExportPath { get; set; }
        

        public int NumFilesExported { get; set; }

        public IList<Programme> Programme { get; set; }
        public IList<Course> Course { get; set; }
        public IList<Group> Group { get; set; }

        public IList<Day> Day { get; set; }
        public IList<Time> Time { get; set; }
        public IList<TimetableSection> TimetableSection { get; set; }

        public IList<Week> Week { get; set; }
        public IList<TermDate> TermDate { get; set; }
        public IList<BankHoliday> BankHoliday { get; set; }

        public async Task<IActionResult> OnGetAsync(Guid sessionID, string academicYear, int plan, string course)
        {
            int planRevisionID = 0;

            if (plan >= 1)
            {
                planRevisionID = plan;
            }
            else
            {
                planRevisionID = int.Parse(_configuration.GetSection("ProResource")["PlanRevisionID"]);
            }

            AcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);

            UserDetails = await Identity.GetFullName(academicYear, User.Identity.Name.Split('\\').Last(), _context);

            UserGreeting = Identity.GetGreeting();

            SystemVersion = _configuration["Version"];

            PlanningSystem = _configuration.GetSection("SystemSettings")["PlanningSystem"];

            string CurrentAcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);

            //Data from Curriculum Planning
            Course = await _context.Course
                .FromSqlInterpolated($"EXEC SPR_ETG_CourseData @AcademicYear={CurrentAcademicYear}, @PlanRevisionID={planRevisionID}, @Course={course}")
                .ToListAsync();

            Group = await _context.Group
                .FromSqlInterpolated($"EXEC SPR_ETG_GroupData @AcademicYear={CurrentAcademicYear}, @PlanRevisionID={planRevisionID}, @Course={course}")
                .ToListAsync();

            Programme = await _context.Programme
                .FromSqlInterpolated($"EXEC SPR_ETG_ProgrammeData @AcademicYear={CurrentAcademicYear}, @PlanRevisionID={planRevisionID}, @Course={course}")
                .ToListAsync();

            //Data for Tiemtable Grid
            Day = await _context.Day
                .FromSqlInterpolated($"EXEC SPR_ETG_Day")
                .ToListAsync();

            Time = await _context.Time
                .FromSqlInterpolated($"EXEC SPR_ETG_Time")
                .ToListAsync();

            TimetableSection = await _context.TimetableSection
                .FromSqlInterpolated($"EXEC SPR_ETG_TimetableSection")
                .ToListAsync();

            Week = await _context.Week
                .FromSqlInterpolated($"EXEC SPR_ETG_Week")
                .ToListAsync();

            TermDate = await _context.TermDate
                .FromSqlInterpolated($"EXEC SPR_ETG_TermDate")
                .ToListAsync();

            BankHoliday = await _context.BankHoliday
                .FromSqlInterpolated($"EXEC SPR_ETG_BankHoliday")
                .ToListAsync();

            ExportParentPath = _hostingEnvironment.WebRootPath + @"\Exports";
            ZipPath = _hostingEnvironment.WebRootPath + @"\Exports\" + sessionID.ToString();
            ExportPath = _hostingEnvironment.WebRootPath + @"\Exports\" + sessionID.ToString() + @"\TTB";
            

            NumFilesExported = 0;

            bool haveReadPermission = true;
            bool haveWritePermission = true;
            string outputPath = "";

            //Delete previous exports older than a certain number of days
            bool deletedOldFolders = DeleteOldFolders();

            //Check permissions for export path
            try
            {
                string filePath = ZipPath;

                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);  
                }

                //Delete directory again as we will get a file exists error when exporting later
                if (Directory.Exists(filePath))
                {
                    Directory.Delete(filePath, true);
                }
            }
            catch (UnauthorizedAccessException)
            {
                //Can't write to save path
                haveWritePermission = false;
                outputPath = ZipPath;
            }

            //College Logo
            string collegeLogoPath = _hostingEnvironment.WebRootPath + @"\images\CollegeLogo.png";
            try
            {
                var collegeLogoStream = new System.IO.FileStream(collegeLogoPath, System.IO.FileMode.Open);
                collegeLogoStream.Close();
                collegeLogoStream.Dispose();
            }
            catch (UnauthorizedAccessException)
            {
                //Can't read logo
                haveReadPermission = false;
                outputPath = collegeLogoPath;
            }

            //Only continue if web app has write permissions to the target or temp folders
            if (haveWritePermission == true)
            {
                NumFilesExported = await WriteExcelFilesAsync(CurrentAcademicYear, Programme, collegeLogoPath, haveReadPermission, haveWritePermission);
            }
            
            string result = 
                "{\"timetables\":{\"savePath\":\"" + ZipPath.Replace("\\", "\\\\") + @"\\\\Timetables.zip" + "\",\"haveReadPermission\":\"" + haveReadPermission + "\",\"haveWritePermission\":\"" + haveWritePermission + "\",\"outputPath\":\"" + outputPath.Replace("\\", "\\\\") + "\",\"numFilesExported\":" + NumFilesExported + "}}";

            //return Page();
            return Content(result);
        }

        public bool DeleteOldFolders()
        {
            int minutesToKeepFolders = int.Parse(_configuration.GetSection("SystemSettings")["MinutesToKeepFolders"]);

            try
            {
                string[] folders = Directory.GetDirectories(ExportParentPath);

                foreach (string folder in folders)
                {
                    DirectoryInfo directoryInfo = new DirectoryInfo(folder);
                    if (directoryInfo.CreationTime < DateTime.Now.AddHours(-minutesToKeepFolders))
                        directoryInfo.Delete(true);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<int> WriteExcelFilesAsync(string CurrentAcademicYear, IList<Programme> Programme, string collegeLogoPath, bool haveReadPermission, bool haveWritePermission)
        {
            PlanningSystem = _configuration.GetSection("SystemSettings")["PlanningSystem"];

            // If cannot write to folder then no point continuing
            if (!haveWritePermission)
            {
                return 0;
            }
            
            string filePath = null;
            string fileName = null;
            string fileURL = null;
            int rowNum = 0;
            int startAtRowNum = 0;
            int colNum = 0;
            int numFilesSaved = 0;
            int programmeNameMaxLength = int.Parse(_configuration.GetSection("SystemSettings")["ProgrammeNameMaxLength"]);

            byte[] bytes = null;
            //Only include college logo if it could be accessed
            if (haveReadPermission == true)
            {
                var collegeLogoStream = new System.IO.FileStream(collegeLogoPath, System.IO.FileMode.Open);
                bytes = IOUtils.ToByteArray(collegeLogoStream);
                collegeLogoStream.Close();
                collegeLogoStream.Dispose();
            }

            //Create folder with session id
            filePath = ZipPath;
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }

            //Create Master Excel Programme List
            if (Programme != null && Programme.Count > 0)
            {
                filePath = ExportPath + @"\";

                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }

                //Generate each Excel workbook
                fileName = @"Master List.xlsx";
                fileName = FileFunctions.MakeValidFileName(fileName); //Sanitize
                fileURL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, fileName);
                FileInfo file = new FileInfo(Path.Combine(filePath, fileName));
                var memory = new MemoryStream();
                using (var fs = new FileStream(Path.Combine(filePath, fileName), FileMode.Create, FileAccess.Write))
                {
                    IWorkbook workbook = new XSSFWorkbook();
                    ISheet sheet;
                    IRow row;
                    ICell cell;
                    CellRangeAddress region;
                    IDrawing drawing;
                    IClientAnchor anchor;
                    IPicture picture;

                    //Only include college logo if it could be accessed
                    int collegeLogo = 0;
                    if (haveReadPermission == true)
                    {
                        //Add college logo
                        collegeLogo = workbook.AddPicture(bytes, PictureType.PNG);
                    }

                    //For date formatting
                    ICreationHelper createHelper = workbook.GetCreationHelper();
                    IHyperlink link = createHelper.CreateHyperlink(HyperlinkType.Url);

                    //Cell formats
                    CellType ctString = CellType.String;
                    CellType ctNumber = CellType.Numeric;
                    CellType ctFormula = CellType.Formula;
                    CellType ctBlank = CellType.Blank;

                    //Cell Colours
                    XSSFColor cLightBlue = new XSSFColor(Color.LightSteelBlue);
                    XSSFColor cReturned = new XSSFColor(Color.Yellow);

                    //Cell Borders
                    BorderStyle bLight = BorderStyle.Thin;
                    BorderStyle bMedium = BorderStyle.Medium;

                    //Fonts
                    //Header
                    IFont fHeader = workbook.CreateFont();
                    fHeader.FontHeight = 20;
                    fHeader.FontName = "Arial";
                    fHeader.IsBold = true;

                    //Sub-header
                    IFont fSubHeader = workbook.CreateFont();
                    fSubHeader.FontHeight = 14;
                    fSubHeader.FontName = "Arial";
                    fSubHeader.IsBold = true;

                    //Table Header
                    IFont fBold = workbook.CreateFont();
                    fBold.IsBold = true;

                    //Cell formats
                    //Page Header
                    XSSFCellStyle sHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                    sHeader.SetFont(fHeader);

                    //Page Subheader
                    XSSFCellStyle sSubHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                    sHeader.SetFont(fSubHeader);
                    sHeader.SetFont(fBold);

                    //Table Header
                    XSSFCellStyle sTableHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                    sTableHeader.SetFont(fBold);
                    sTableHeader.SetFillForegroundColor(cLightBlue);
                    sTableHeader.FillPattern = FillPattern.SolidForeground;
                    sTableHeader.BorderTop = bMedium;
                    sTableHeader.BorderBottom = bMedium;
                    sTableHeader.BorderLeft = bMedium;
                    sTableHeader.BorderRight = bMedium;
                    sTableHeader.WrapText = true;

                    XSSFCellStyle sReturned = (XSSFCellStyle)workbook.CreateCellStyle();
                    sReturned.SetFont(fBold);
                    sReturned.SetFillForegroundColor(cReturned);
                    sReturned.FillPattern = FillPattern.SolidForeground;
                    sReturned.BorderTop = bMedium;
                    sReturned.BorderBottom = bMedium;
                    sReturned.BorderLeft = bMedium;
                    sReturned.BorderRight = bMedium;

                    //Border only
                    XSSFCellStyle sBorderMedium = (XSSFCellStyle)workbook.CreateCellStyle();
                    sBorderMedium.BorderTop = bMedium;
                    sBorderMedium.BorderBottom = bMedium;
                    sBorderMedium.BorderLeft = bMedium;
                    sBorderMedium.BorderRight = bMedium;

                    XSSFCellStyle sBorderDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sBorderDate.BorderTop = bMedium;
                    sBorderDate.BorderBottom = bMedium;
                    sBorderDate.BorderLeft = bMedium;
                    sBorderDate.BorderRight = bMedium;
                    sBorderDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    //Create Index Sheet
                    sheet = workbook.CreateSheet("Master List");
                    row = sheet.CreateRow(0);
                    row.Height = 1000;

                    //Insert College Logo (to right) - second col/row must be greater otherwise nothing appears
                    drawing = sheet.CreateDrawingPatriarch();
                    anchor = createHelper.CreateClientAnchor();
                    anchor.Col1 = 9;
                    anchor.Row1 = 0;
                    anchor.Col2 = 11;
                    anchor.Row2 = 1;

                    if (haveReadPermission == true)
                    {
                        picture = drawing.CreatePicture(anchor, collegeLogo);
                    }

                    cell = row.CreateCell(0);

                    cell.SetCellValue("Master List of Programmes from " + PlanningSystem + " for " + CurrentAcademicYear);
                    cell.CellStyle = sHeader;

                    //Merge header row
                    region = CellRangeAddress.ValueOf("A1:I1");
                    sheet.AddMergedRegion(region);

                    row = sheet.CreateRow(1);

                    row = sheet.CreateRow(2);

                    cell = row.CreateCell(0, ctString);
                    cell.SetCellValue("Site");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(1, ctString);
                    cell.SetCellValue("Fac Code");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(2, ctString);
                    cell.SetCellValue("Fac Name");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(3, ctString);
                    cell.SetCellValue("Team Code");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(4, ctString);
                    cell.SetCellValue("Team Name");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(5, ctString);
                    cell.SetCellValue("Prog Code");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(6, ctString);
                    cell.SetCellValue("Prog Title");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(7, ctString);
                    cell.SetCellValue("Mode of Attendance");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(8, ctString);
                    cell.SetCellValue("Status");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(9, ctString);
                    cell.SetCellValue("Num Courses");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(10, ctString);
                    cell.SetCellValue("Returned");
                    cell.CellStyle = sTableHeader;

                    //Column widths
                    sheet.SetColumnWidth(0, 16 * 256);
                    sheet.SetColumnWidth(1, 8 * 256);
                    sheet.SetColumnWidth(2, 20 * 256);
                    sheet.SetColumnWidth(3, 8 * 256);
                    sheet.SetColumnWidth(4, 40 * 256);
                    sheet.SetColumnWidth(5, 16 * 256);
                    sheet.SetColumnWidth(6, 60 * 256);
                    sheet.SetColumnWidth(7, 16 * 256);
                    sheet.SetColumnWidth(8, 20 * 256);
                    sheet.SetColumnWidth(9, 8 * 256);
                    sheet.SetColumnWidth(10, 10 * 256);

                    //The current row in the worksheet
                    rowNum = 2;
                    string progFileName = null;
                    string progFolder = null;
                    string progFileURL = null;

                    foreach (var programme in Programme)
                    {
                        rowNum += 1;
                        row = sheet.CreateRow(rowNum);
 
                        progFileName = FileFunctions.MakeValidFileName(FileFunctions.ShortenString(programme.ProgCode + @" - " + programme.ProgTitle, programmeNameMaxLength) + @".xlsx"); //Sanitize
                        progFolder = programme.FacCode + @"/" + programme.TeamCode + @"/";
                        //progFileURL = FileFunctions.FormatHyperlink(string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, progFolder + progFileName));
                        progFileURL = FileFunctions.FormatHyperlink(progFolder + progFileName);
                        link.Address = progFileURL;

                        cell = row.CreateCell(0, ctString);
                        cell.SetCellValue(programme.SiteName);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(1, ctString);
                        cell.SetCellValue(programme.FacCode);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(2, ctString);
                        cell.SetCellValue(programme.FacName);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(3, ctString);
                        cell.SetCellValue(programme.TeamCode);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(4, ctString);
                        cell.SetCellValue(programme.TeamName);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(5, ctString);
                        cell.SetCellValue(programme.ProgCode);
                        cell.CellStyle = sBorderMedium;
                        cell.Hyperlink = link;

                        cell = row.CreateCell(6, ctString);
                        cell.SetCellValue(programme.ProgTitle);
                        cell.CellStyle = sBorderMedium;
                        cell.Hyperlink = link;

                        cell = row.CreateCell(7, ctString);
                        cell.SetCellValue(programme.ModeOfAttendanceName);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(8, ctString);
                        cell.SetCellValue(programme.ProgStatus);
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(9, ctNumber);
                        if (programme.Course != null && programme.Course.Count > 0)
                        {
                            cell.SetCellValue(programme.Course.Count);
                        }
                        else
                        {
                            cell.SetCellValue(0);
                        }
                        cell.CellStyle = sBorderMedium;

                        cell = row.CreateCell(10, ctBlank);
                        cell.CellStyle = sReturned;
                    }

                    workbook.Write(fs);

                    using (var stream = new FileStream(Path.Combine(filePath, fileName), FileMode.Open))
                    {
                        await stream.CopyToAsync(memory);
                    }
                    memory.Position = 0;
                }
            }

            //Create Excel file for each programme
            if (Programme != null && Programme.Count > 0)
            {
                foreach (var programme in Programme)
                {
                    numFilesSaved += 1;

                    //Save timetables into departmental and team folders (creates a folder if it doesn't exist)
                    filePath = ExportPath + @"\" + programme.FacCode + @"\" + programme.TeamCode + @"\";

                    if (!Directory.Exists(filePath))
                    {
                        Directory.CreateDirectory(filePath);
                    }

                    //Generate each Excel workbook
                    fileName = FileFunctions.MakeValidFileName(FileFunctions.ShortenString(programme.ProgCode + @" - " + programme.ProgTitle, programmeNameMaxLength) + @".xlsx"); //Sanitize
                    fileURL = FileFunctions.FormatHyperlink(string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, fileName));
                    FileInfo file = new FileInfo(Path.Combine(filePath, fileName));
                    var memory = new MemoryStream();
                    using (var fs = new FileStream(Path.Combine(filePath, fileName), FileMode.Create, FileAccess.Write))
                    {
                        IWorkbook workbook = new XSSFWorkbook();
                        ISheet sheet;
                        IRow row;
                        ICell cell;
                        CellRangeAddress region;
                        IDrawing drawing;
                        IClientAnchor anchor;
                        IPicture picture;

                        int collegeLogo = 0;
                        if (haveReadPermission == true)
                        {
                            //Add college logo
                            collegeLogo = workbook.AddPicture(bytes, PictureType.PNG);
                        }

                        //For date formatting
                        ICreationHelper createHelper = workbook.GetCreationHelper();

                        //Cell formats
                        CellType ctString = CellType.String;
                        CellType ctNumber = CellType.Numeric;
                        CellType ctFormula = CellType.Formula;
                        CellType ctBlank = CellType.Blank;

                        //Cell Colours
                        XSSFColor cLightBlue = new XSSFColor(Color.LightSteelBlue);
                        XSSFColor cBlue = new XSSFColor(Color.RoyalBlue);
                        XSSFColor cMainAim = new XSSFColor(Color.Yellow);
                        XSSFColor cEngMaths = new XSSFColor(Color.PaleGreen);
                        XSSFColor cWex = new XSSFColor(Color.Plum);
                        XSSFColor cDSS = new XSSFColor(Color.Lavender);
                        XSSFColor cTut = new XSSFColor(Color.LightCyan);
                        XSSFColor cOther = new XSSFColor(Color.NavajoWhite);
                        XSSFColor cWhite = new XSSFColor(Color.White);

                        XSSFColor cEarlySlots = new XSSFColor(Color.LavenderBlush);
                        XSSFColor cLateSlots = new XSSFColor(Color.Lavender);
                        XSSFColor cWeeklyHours = new XSSFColor(Color.Honeydew);
                        XSSFColor cDaySectionName = new XSSFColor(Color.DeepSkyBlue);

                        //Cell Borders
                        BorderStyle bLight = BorderStyle.Thin;
                        BorderStyle bMedium = BorderStyle.Medium;

                        //Fonts
                        //Header
                        IFont fHeader = workbook.CreateFont();
                        fHeader.FontHeight = 20;
                        fHeader.FontName = "Arial";
                        fHeader.IsBold = true;

                        //Sub-header
                        IFont fSubHeader = workbook.CreateFont();
                        fSubHeader.FontHeight = 14;
                        fSubHeader.FontName = "Arial";
                        fSubHeader.IsBold = true;

                        //Table Header
                        IFont fBold = workbook.CreateFont();
                        fBold.IsBold = true;

                        IFont fWhite = workbook.CreateFont();
                        fWhite.Color = IndexedColors.White.Index;

                        //Cell formats
                        //Page Header
                        XSSFCellStyle sHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                        sHeader.SetFont(fHeader);

                        //Page Subheader
                        XSSFCellStyle sSubHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                        sHeader.SetFont(fSubHeader);
                        sHeader.SetFont(fBold);

                        //Table Header
                        XSSFCellStyle sTableHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTableHeader.SetFont(fBold);
                        sTableHeader.SetFillForegroundColor(cLightBlue);
                        sTableHeader.FillPattern = FillPattern.SolidForeground;
                        sTableHeader.BorderTop = bMedium;
                        sTableHeader.BorderBottom = bMedium;
                        sTableHeader.BorderLeft = bMedium;
                        sTableHeader.BorderRight = bMedium;
                        sTableHeader.WrapText = true;

                        XSSFCellStyle sTableHeaderCenter = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTableHeaderCenter.SetFont(fBold);
                        sTableHeaderCenter.SetFillForegroundColor(cLightBlue);
                        sTableHeaderCenter.FillPattern = FillPattern.SolidForeground;
                        sTableHeaderCenter.Alignment = HorizontalAlignment.Center;
                        sTableHeaderCenter.BorderTop = bMedium;
                        sTableHeaderCenter.BorderBottom = bMedium;
                        sTableHeaderCenter.BorderLeft = bMedium;
                        sTableHeaderCenter.BorderRight = bMedium;
                        sTableHeaderCenter.WrapText = true;

                        //Date
                        XSSFCellStyle sDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));
                        sDate.BorderTop = bMedium;
                        sDate.BorderBottom = bMedium;
                        sDate.BorderLeft = bMedium;
                        sDate.BorderRight = bMedium;

                        //Rotated
                        XSSFCellStyle sRotated = (XSSFCellStyle)workbook.CreateCellStyle();
                        sRotated.Rotation = 90;
                        sRotated.BorderTop = bMedium;
                        sRotated.BorderBottom = bMedium;
                        sRotated.BorderLeft = bMedium;
                        sRotated.BorderRight = bMedium;

                        //Merged Centred
                        XSSFCellStyle sMergedCentredTotal = (XSSFCellStyle)workbook.CreateCellStyle();
                        sMergedCentredTotal.SetFont(fBold);
                        sMergedCentredTotal.Alignment = HorizontalAlignment.Center;
                        sMergedCentredTotal.VerticalAlignment = VerticalAlignment.Center;
                        sMergedCentredTotal.BorderTop = bLight;
                        sMergedCentredTotal.BorderBottom = bLight;
                        sMergedCentredTotal.BorderLeft = bLight;
                        sMergedCentredTotal.BorderRight = bLight;

                        XSSFCellStyle sWeeklyHours = (XSSFCellStyle)workbook.CreateCellStyle();
                        sWeeklyHours.SetFont(fBold);
                        sWeeklyHours.SetFillForegroundColor(cWeeklyHours);
                        sWeeklyHours.FillPattern = FillPattern.SolidForeground;
                        sWeeklyHours.Alignment = HorizontalAlignment.Center;
                        sWeeklyHours.VerticalAlignment = VerticalAlignment.Center;
                        sWeeklyHours.BorderTop = bLight;
                        sWeeklyHours.BorderBottom = bLight;
                        sWeeklyHours.BorderLeft = bLight;
                        sWeeklyHours.BorderRight = bLight;

                        //Timetable Grid DaySection
                        XSSFCellStyle sDaySectionNameLast = (XSSFCellStyle)workbook.CreateCellStyle();
                        sDaySectionNameLast.SetFont(fWhite);
                        sDaySectionNameLast.SetFillForegroundColor(cDaySectionName);
                        sDaySectionNameLast.FillPattern = FillPattern.SolidForeground;
                        sDaySectionNameLast.BorderBottom = bMedium;

                        XSSFCellStyle sDaySectionName = (XSSFCellStyle)workbook.CreateCellStyle();
                        sDaySectionName.SetFont(fWhite);
                        sDaySectionName.SetFillForegroundColor(cDaySectionName);
                        sDaySectionName.FillPattern = FillPattern.SolidForeground;

                        //Merged Right
                        XSSFCellStyle sMergedRightTotal = (XSSFCellStyle)workbook.CreateCellStyle();
                        sMergedRightTotal.SetFont(fBold);
                        sMergedRightTotal.Alignment = HorizontalAlignment.Right;
                        sMergedRightTotal.VerticalAlignment = VerticalAlignment.Center;
                        sMergedRightTotal.BorderTop = bLight;
                        sMergedRightTotal.BorderBottom = bLight;
                        sMergedRightTotal.BorderLeft = bLight;
                        sMergedRightTotal.BorderRight = bLight;

                        //Table Header (Rotated)
                        XSSFCellStyle sTableHeaderLightRotated = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTableHeaderLightRotated.SetFont(fBold);
                        sTableHeaderLightRotated.SetFillForegroundColor(cLightBlue);
                        sTableHeaderLightRotated.FillPattern = FillPattern.SolidForeground;
                        sTableHeaderLightRotated.Rotation = 90;
                        sTableHeaderLightRotated.Alignment = HorizontalAlignment.Center;
                        sTableHeaderLightRotated.BorderTop = bLight;
                        sTableHeaderLightRotated.BorderBottom = bLight;
                        sTableHeaderLightRotated.BorderLeft = bLight;
                        sTableHeaderLightRotated.BorderRight = bLight;

                        XSSFCellStyle sTableHeaderRotated = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTableHeaderRotated.SetFont(fBold);
                        sTableHeaderRotated.SetFillForegroundColor(cLightBlue);
                        sTableHeaderRotated.FillPattern = FillPattern.SolidForeground;
                        sTableHeaderRotated.Rotation = 90;
                        sTableHeaderRotated.BorderTop = bMedium;
                        sTableHeaderRotated.BorderBottom = bMedium;
                        sTableHeaderRotated.BorderLeft = bMedium;
                        sTableHeaderRotated.BorderRight = bMedium;

                        XSSFCellStyle sTableHeaderCenterRotated = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTableHeaderCenterRotated.SetFont(fBold);
                        sTableHeaderCenterRotated.SetFillForegroundColor(cLightBlue);
                        sTableHeaderCenterRotated.FillPattern = FillPattern.SolidForeground;
                        sTableHeaderCenterRotated.Rotation = 90;
                        sTableHeaderCenterRotated.Alignment = HorizontalAlignment.Center;
                        sTableHeaderCenterRotated.BorderTop = bMedium;
                        sTableHeaderCenterRotated.BorderBottom = bMedium;
                        sTableHeaderCenterRotated.BorderLeft = bMedium;
                        sTableHeaderCenterRotated.BorderRight = bMedium;

                        XSSFCellStyle sTotalRight = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTotalRight.SetFont(fBold);
                        sTotalRight.Alignment = HorizontalAlignment.Right;

                        //Types of courses
                        XSSFCellStyle sMainAim = (XSSFCellStyle)workbook.CreateCellStyle();
                        sMainAim.SetFillForegroundColor(cMainAim);
                        sMainAim.FillPattern = FillPattern.SolidForeground;
                        sMainAim.BorderTop = bMedium;
                        sMainAim.BorderBottom = bMedium;
                        sMainAim.BorderLeft = bMedium;
                        sMainAim.BorderRight = bMedium;
                        sMainAim.SetFont(fBold);

                        XSSFCellStyle sEngMaths = (XSSFCellStyle)workbook.CreateCellStyle();
                        sEngMaths.SetFillForegroundColor(cEngMaths);
                        sEngMaths.FillPattern = FillPattern.SolidForeground;
                        sEngMaths.BorderTop = bMedium;
                        sEngMaths.BorderBottom = bMedium;
                        sEngMaths.BorderLeft = bMedium;
                        sEngMaths.BorderRight = bMedium;

                        XSSFCellStyle sWex = (XSSFCellStyle)workbook.CreateCellStyle();
                        sWex.SetFillForegroundColor(cWex);
                        sWex.FillPattern = FillPattern.SolidForeground;
                        sWex.BorderTop = bMedium;
                        sWex.BorderBottom = bMedium;
                        sWex.BorderLeft = bMedium;
                        sWex.BorderRight = bMedium;

                        XSSFCellStyle sDSS = (XSSFCellStyle)workbook.CreateCellStyle();
                        sDSS.SetFillForegroundColor(cDSS);
                        sDSS.FillPattern = FillPattern.SolidForeground;
                        sDSS.BorderTop = bMedium;
                        sDSS.BorderBottom = bMedium;
                        sDSS.BorderLeft = bMedium;
                        sDSS.BorderRight = bMedium;

                        XSSFCellStyle sTut = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTut.SetFillForegroundColor(cTut);
                        sTut.FillPattern = FillPattern.SolidForeground;
                        sTut.BorderTop = bMedium;
                        sTut.BorderBottom = bMedium;
                        sTut.BorderLeft = bMedium;
                        sTut.BorderRight = bMedium;

                        XSSFCellStyle sOther = (XSSFCellStyle)workbook.CreateCellStyle();
                        sOther.SetFillForegroundColor(cOther);
                        sOther.FillPattern = FillPattern.SolidForeground;
                        sOther.BorderTop = bMedium;
                        sOther.BorderBottom = bMedium;
                        sOther.BorderLeft = bMedium;
                        sOther.BorderRight = bMedium;

                        XSSFCellStyle sTotal = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTotal.SetFont(fBold);
                        sTotal.BorderTop = bMedium;
                        sTotal.BorderBottom = bMedium;
                        sTotal.BorderLeft = bMedium;
                        sTotal.BorderRight = bMedium;

                        //Border Only
                        XSSFCellStyle sBorderLight = (XSSFCellStyle)workbook.CreateCellStyle();
                        sBorderLight.BorderTop = bLight;
                        sBorderLight.BorderBottom = bLight;
                        sBorderLight.BorderLeft = bLight;
                        sBorderLight.BorderRight = bLight;

                        XSSFCellStyle sBorderMedium = (XSSFCellStyle)workbook.CreateCellStyle();
                        sBorderMedium.BorderTop = bMedium;
                        sBorderMedium.BorderBottom = bMedium;
                        sBorderMedium.BorderLeft = bMedium;
                        sBorderMedium.BorderRight = bMedium;

                        XSSFCellStyle sEarlySlots = (XSSFCellStyle)workbook.CreateCellStyle();
                        sEarlySlots.SetFillForegroundColor(cEarlySlots);
                        sEarlySlots.FillPattern = FillPattern.SolidForeground;
                        sEarlySlots.BorderTop = bLight;
                        sEarlySlots.BorderBottom = bLight;
                        sEarlySlots.BorderLeft = bLight;
                        sEarlySlots.BorderRight = bLight;

                        XSSFCellStyle sLateSlots = (XSSFCellStyle)workbook.CreateCellStyle();
                        sLateSlots.SetFillForegroundColor(cLateSlots);
                        sLateSlots.FillPattern = FillPattern.SolidForeground;
                        sLateSlots.BorderTop = bLight;
                        sLateSlots.BorderBottom = bLight;
                        sLateSlots.BorderLeft = bLight;
                        sLateSlots.BorderRight = bLight;

                        //Underlined
                        XSSFCellStyle sUnderlined = (XSSFCellStyle)workbook.CreateCellStyle();
                        sUnderlined.BorderBottom = bMedium;

                        //Types of courses (date fields
                        XSSFCellStyle sMainAimDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sMainAimDate.SetFillForegroundColor(cMainAim);
                        sMainAimDate.FillPattern = FillPattern.SolidForeground;
                        sMainAimDate.BorderTop = bMedium;
                        sMainAimDate.BorderBottom = bMedium;
                        sMainAimDate.BorderLeft = bMedium;
                        sMainAimDate.BorderRight = bMedium;
                        sMainAimDate.SetFont(fBold);
                        sMainAimDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        XSSFCellStyle sEngMathsDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sEngMathsDate.SetFillForegroundColor(cEngMaths);
                        sEngMathsDate.FillPattern = FillPattern.SolidForeground;
                        sEngMathsDate.BorderTop = bMedium;
                        sEngMathsDate.BorderBottom = bMedium;
                        sEngMathsDate.BorderLeft = bMedium;
                        sEngMathsDate.BorderRight = bMedium;
                        sEngMathsDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        XSSFCellStyle sWexDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sWexDate.SetFillForegroundColor(cWex);
                        sWexDate.FillPattern = FillPattern.SolidForeground;
                        sWexDate.BorderTop = bMedium;
                        sWexDate.BorderBottom = bMedium;
                        sWexDate.BorderLeft = bMedium;
                        sWexDate.BorderRight = bMedium;
                        sWexDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        XSSFCellStyle sDSSDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sDSSDate.SetFillForegroundColor(cDSS);
                        sDSSDate.FillPattern = FillPattern.SolidForeground;
                        sDSSDate.BorderTop = bMedium;
                        sDSSDate.BorderBottom = bMedium;
                        sDSSDate.BorderLeft = bMedium;
                        sDSSDate.BorderRight = bMedium;
                        sDSSDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        XSSFCellStyle sTutDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sTutDate.SetFillForegroundColor(cTut);
                        sTutDate.FillPattern = FillPattern.SolidForeground;
                        sTutDate.BorderTop = bMedium;
                        sTutDate.BorderBottom = bMedium;
                        sTutDate.BorderLeft = bMedium;
                        sTutDate.BorderRight = bMedium;
                        sTutDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        XSSFCellStyle sOtherDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sOtherDate.SetFillForegroundColor(cOther);
                        sOtherDate.FillPattern = FillPattern.SolidForeground;
                        sOtherDate.BorderTop = bMedium;
                        sOtherDate.BorderBottom = bMedium;
                        sOtherDate.BorderLeft = bMedium;
                        sOtherDate.BorderRight = bMedium;
                        sOtherDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        //Border Only
                        XSSFCellStyle sBorderDate = (XSSFCellStyle)workbook.CreateCellStyle();
                        sBorderDate.BorderTop = bMedium;
                        sBorderDate.BorderBottom = bMedium;
                        sBorderDate.BorderLeft = bMedium;
                        sBorderDate.BorderRight = bMedium;
                        sBorderDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                        //Create Index Sheet
                        sheet = workbook.CreateSheet("Index");
                        row = sheet.CreateRow(0);
                        row.Height = 1000;

                        //Insert College Logo (to right) - second col/row must be greater otherwise nothing appears
                        drawing = sheet.CreateDrawingPatriarch();
                        anchor = createHelper.CreateClientAnchor();
                        anchor.Col1 = 11;
                        anchor.Row1 = 0;
                        anchor.Col2 = 13;
                        anchor.Row2 = 1;

                        if (haveReadPermission == true)
                        {
                            picture = drawing.CreatePicture(anchor, collegeLogo);
                        }

                        cell = row.CreateCell(0);

                        cell.SetCellValue(programme.ProgCode + " - " + programme.ProgTitle);
                        cell.CellStyle = sHeader;

                        //Merge header row
                        region = CellRangeAddress.ValueOf("A1:K1");
                        sheet.AddMergedRegion(region);

                        row = sheet.CreateRow(1);

                        row = sheet.CreateRow(2);

                        cell = row.CreateCell(0, ctString);
                        cell.SetCellValue("Course Code");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(1, ctString);
                        cell.SetCellValue("Course Title");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(2, ctString);
                        cell.SetCellValue("Qual");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(3, ctString);
                        cell.SetCellValue("Award Body");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(4, ctString);
                        cell.SetCellValue("Hours Per Week");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(5, ctString);
                        cell.SetCellValue("Length in Weeks");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(6, ctString);
                        cell.SetCellValue("Planned Learning Hours");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(7, ctString);
                        cell.SetCellValue("Planned EEP Hours");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(8, ctString);
                        cell.SetCellValue("Start Date");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(9, ctString);
                        cell.SetCellValue("End Date");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(10, ctString);
                        cell.SetCellValue("Site");
                        cell.CellStyle = sTableHeader;

                        cell = row.CreateCell(11, ctString);
                        cell.SetCellValue("Notes");
                        cell.CellStyle = sTableHeader;

                        //Column widths
                        sheet.SetColumnWidth(0, 16 * 256);
                        sheet.SetColumnWidth(1, 40 * 256);
                        sheet.SetColumnWidth(2, 14 * 256);
                        sheet.SetColumnWidth(3, 30 * 256);
                        sheet.SetColumnWidth(4, 8 * 256);
                        sheet.SetColumnWidth(5, 8 * 256);
                        sheet.SetColumnWidth(6, 10 * 256);
                        sheet.SetColumnWidth(7, 10 * 256);
                        sheet.SetColumnWidth(8, 12 * 256);
                        sheet.SetColumnWidth(9, 12 * 256);
                        sheet.SetColumnWidth(10, 20 * 256);
                        sheet.SetColumnWidth(11, 20 * 256);

                        //The current row in the worksheet
                        rowNum = 2;

                        //Generate each Excel worksheet
                        if (programme.Course != null && programme.Course.Count > 0)
                        {
                            foreach (var crs in programme.Course)
                            {
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);
                                XSSFCellStyle cellStyle;
                                XSSFCellStyle cellStyleDate;

                                switch (crs.CourseOrder)
                                {
                                    case 0:
                                        cellStyle = sOther;
                                        cellStyleDate = sOtherDate;
                                        break;
                                    case 1:
                                        cellStyle = sMainAim;
                                        cellStyleDate = sMainAimDate;
                                        break;
                                    case 2:
                                        cellStyle = sOther;
                                        cellStyleDate = sOtherDate;
                                        break;
                                    case 3:
                                        cellStyle = sEngMaths;
                                        cellStyleDate = sEngMathsDate;
                                        break;
                                    case 4:
                                        cellStyle = sEngMaths;
                                        cellStyleDate = sEngMathsDate;
                                        break;
                                    case 5:
                                        cellStyle = sTut;
                                        cellStyleDate = sTutDate;
                                        break;
                                    case 6:
                                        cellStyle = sWex;
                                        cellStyleDate = sWexDate;
                                        break;
                                    case 7:
                                        cellStyle = sDSS;
                                        cellStyleDate = sDSSDate;
                                        break;
                                    case 8:
                                        cellStyle = sOther;
                                        cellStyleDate = sOtherDate;
                                        break;
                                    default:
                                        cellStyle = sMainAim;
                                        cellStyleDate = sMainAimDate;
                                        break;
                                }

                                cell = row.CreateCell(0, ctString);
                                cell.SetCellValue(crs.CourseCode);
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue(crs.CourseTitle);
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(2, ctString);
                                cell.SetCellValue(crs.AimCode);
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(3, ctString);
                                cell.SetCellValue(crs.AwardBody);
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(4, ctNumber);
                                if (crs.HoursPerWeek >= 0)
                                {
                                    cell.SetCellValue((double)crs.HoursPerWeek);
                                }
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(5, ctNumber);
                                if (crs.Weeks >= 0)
                                {
                                    cell.SetCellValue((double)crs.Weeks);
                                }
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(6, ctNumber);
                                if (crs.PLHMax >= 0)
                                {
                                    cell.SetCellValue((double)crs.PLHMax);
                                }
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(7, ctNumber);
                                if (crs.EEPMax >= 0)
                                {
                                    cell.SetCellValue((double)crs.EEPMax);
                                }
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(8, ctNumber);
                                cell.CellStyle = cellStyleDate;
                                cell.SetCellValue((DateTime)crs.StartDate);

                                cell = row.CreateCell(9, ctNumber);
                                cell.CellStyle = cellStyleDate;
                                cell.SetCellValue((DateTime)crs.EndDate);

                                cell = row.CreateCell(10, ctString);
                                cell.SetCellValue(crs.SiteName);
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(11, ctString);
                                cell.SetCellValue(crs.Notes);
                                cell.CellStyle = cellStyle;
                            }

                            ////Set widths(does not seem to work very well)
                            //int numberOfColumns = indexSheet.GetRow(4).PhysicalNumberOfCells;
                            //for (int i = 1; i <= numberOfColumns; i++)
                            //{
                            //    indexSheet.AutoSizeColumn(i);
                            //    GC.Collect(); // Add this line
                            //}
                        }
                        else
                        {
                            rowNum += 1;
                            row = sheet.CreateRow(rowNum);

                            cell = row.CreateCell(0, ctString);
                            cell.CellStyle = sHeader;
                            cell.SetCellValue("Error - No courses could be loaded. Please check CourseData Stored Procedure");
                        }



                        if (programme.Group != null && programme.Group.Count > 0)
                        {
                            foreach (var group in programme.Group)
                            {
                                sheet = workbook.CreateSheet(group.ProgCodeWithGroup);

                                //Need narrow columns
                                sheet.DefaultColumnWidth = 2;

                                row = sheet.CreateRow(0);
                                row.Height = 1000;

                                //Insert College Logo (to right) - second col/row must be greater otherwise nothing appears
                                drawing = sheet.CreateDrawingPatriarch();
                                anchor = createHelper.CreateClientAnchor();
                                anchor.Col1 = 48;
                                anchor.Row1 = 0;
                                anchor.Col2 = 53;
                                anchor.Row2 = 1;

                                if (haveReadPermission == true)
                                {
                                    picture = drawing.CreatePicture(anchor, collegeLogo);
                                }

                                cell = row.CreateCell(0);

                                cell.SetCellValue("Group " + group.GroupCode + " for " + programme.ProgCode + " - " + programme.ProgTitle);
                                cell.CellStyle = sHeader;

                                row = sheet.CreateRow(1);
                                row = sheet.CreateRow(2);
                                row = sheet.CreateRow(3);

                                cell = row.CreateCell(0, ctBlank);

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue("Programme Details");
                                cell.CellStyle = sSubHeader;

                                //Merge Cells for Title
                                region = CellRangeAddress.ValueOf("A1:AJ1");
                                sheet.AddMergedRegion(region);

                                row = sheet.CreateRow(4);

                                cell = row.CreateCell(0, ctBlank);

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue("Faculty:");

                                cell = row.CreateCell(2, ctString);
                                cell.SetCellValue(programme.FacName);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(3, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(4, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(5, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(6, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(7, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(8, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(9, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(10, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(11, ctString);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(12, ctString);

                                cell = row.CreateCell(13, ctString);
                                cell.SetCellValue("Mode:");

                                cell = row.CreateCell(20, ctString);
                                cell.SetCellValue(programme.ModeOfAttendanceName);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(21, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(22, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(23, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(24, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(25, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(26, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(27, ctString);
                                cell.CellStyle = sUnderlined;

                                row = sheet.CreateRow(5);

                                cell = row.CreateCell(0, ctBlank);

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue("Team:");

                                cell = row.CreateCell(2, ctString);
                                cell.SetCellValue(programme.TeamName);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(3, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(4, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(5, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(6, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(7, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(8, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(9, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(10, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(11, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(12, ctString);

                                cell = row.CreateCell(13, ctString);
                                cell.SetCellValue("Site:");

                                cell = row.CreateCell(20, ctString);
                                cell.SetCellValue(programme.SiteName);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(21, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(22, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(23, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(24, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(25, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(26, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(27, ctString);
                                cell.CellStyle = sUnderlined;

                                row = sheet.CreateRow(6);

                                cell = row.CreateCell(0, ctBlank);

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue("Parent Code:");

                                cell = row.CreateCell(2, ctString);
                                cell.SetCellValue(programme.ProgCode);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(3, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(4, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(5, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(6, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(7, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(8, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(9, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(10, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(11, ctString);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(12, ctString);

                                cell = row.CreateCell(13, ctString);
                                cell.SetCellValue("Prog Planned Hours:");

                                cell = row.CreateCell(20, ctNumber);
                                cell.SetCellValue((double)(programme.PLHMax + programme.EEPMax));
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(21, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(22, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(23, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(24, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(25, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(26, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(27, ctString);
                                cell.CellStyle = sUnderlined;

                                row = sheet.CreateRow(7);

                                cell = row.CreateCell(0, ctBlank);

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue("Title:");

                                cell = row.CreateCell(2, ctString);
                                cell.SetCellValue(programme.ProgTitle);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(3, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(4, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(5, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(6, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(7, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(8, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(9, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(10, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(11, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(12, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(13, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(14, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(15, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(16, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(17, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(18, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(19, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(20, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(21, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(22, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(23, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(24, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(25, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(26, ctString);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(27, ctString);
                                cell.CellStyle = sUnderlined;

                                //Merge Cells
                                region = CellRangeAddress.ValueOf("A3:AC3");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("A4:A8");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("AC4:AC8");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("M5:M7");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("B4:AB4");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("C5:L5");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("N5:T5");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("U5:AB5");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("C6:L6");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("N6:T6");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("U6:AB6");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("C7:L7");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("N7:T7");
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("U7:AB7");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("C8:AB8");
                                sheet.AddMergedRegion(region);

                                region = CellRangeAddress.ValueOf("A9:AC9");
                                sheet.AddMergedRegion(region);

                                row = sheet.CreateRow(8);
                                cell = row.CreateCell(0, ctBlank);

                                row = sheet.CreateRow(9);

                                //Put a border around top section
                                region = CellRangeAddress.ValueOf("A3:AC9");
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                row = sheet.CreateRow(10);

                                if (TimetableSection != null && TimetableSection.Count > 0)
                                {
                                    foreach (var section in TimetableSection)
                                    {
                                        cell = row.CreateCell(0, ctBlank);
                                        cell.CellStyle = sTableHeader;

                                        cell = row.CreateCell(1, ctBlank);
                                        cell.CellStyle = sTableHeader;

                                        colNum = 1;

                                        if (Time != null && Time.Count > 0)
                                        {
                                            foreach (var time in Time)
                                            {
                                                colNum += 1;

                                                cell = row.CreateCell(colNum, ctString);
                                                cell.SetCellValue(time.TimeName.ToString());
                                                cell.CellStyle = sTableHeaderLightRotated;
                                            }
                                        }

                                        colNum += 1;
                                        cell = row.CreateCell(colNum, ctString);
                                        cell.SetCellValue("Weekly Hrs");
                                        cell.CellStyle = sTableHeader;

                                        colNum += 1;
                                        cell = row.CreateCell(colNum, ctString);
                                        cell.SetCellValue("Number of slots available in the FE academic year");
                                        cell.CellStyle = sTableHeader;
                                    }
                                }
                                //Draw timetable grid
                                //The current row in the worksheet
                                rowNum = 10;
                                if (Day != null && Day.Count > 0)
                                {
                                    foreach (var day in Day)
                                    {
                                        startAtRowNum = rowNum + 1;

                                        if (TimetableSection != null && TimetableSection.Count > 0)
                                        {
                                            foreach (var section in TimetableSection)
                                            {
                                                rowNum += 1;

                                                row = sheet.CreateRow(rowNum);

                                                cell = row.CreateCell(0, ctString);
                                                cell.SetCellValue(day.DayName);
                                                cell.CellStyle = sTableHeaderCenterRotated;

                                                cell = row.CreateCell(1, ctString);
                                                cell.SetCellValue(section.SectionName);

                                                if(section.SectionName == "Full Year/Weeks numbers")
                                                {
                                                    cell.CellStyle = sDaySectionNameLast;
                                                }
                                                else if (section.SectionName == "Staff Name")
                                                {
                                                    //No Style
                                                }
                                                else
                                                {
                                                    cell.CellStyle = sDaySectionName;
                                                }

                                                colNum = 1;

                                                XSSFCellStyle cellStyle;

                                                if (Time != null && Time.Count > 0)
                                                {
                                                    foreach (var time in Time)
                                                    {
                                                        colNum += 1;

                                                        if (time.Hours < 9)
                                                        {
                                                            cellStyle = sEarlySlots;
                                                        }
                                                        else if (time.Hours >= 17)
                                                        {
                                                            cellStyle = sLateSlots;
                                                        }
                                                        else
                                                        {
                                                            cellStyle = sBorderLight;
                                                        }

                                                        cell = row.CreateCell(colNum, ctBlank);
                                                        cell.CellStyle = cellStyle;
                                                    }
                                                }

                                                //Add additional rows
                                                colNum += 1;
                                                cell = row.CreateCell(colNum, ctNumber);
                                                cell.CellStyle = sWeeklyHours;

                                                colNum += 1;
                                                cell = row.CreateCell(colNum, ctNumber);
                                                cell.SetCellValue(day.Slots);
                                                cell.CellStyle = sMergedCentredTotal;
                                            }
                                        }

                                        //Merge Day rows and end rows
                                        region = CellRangeAddress.ValueOf("A" + (startAtRowNum + 1) + ":A" + (rowNum + 1));
                                        sheet.AddMergedRegion(region);

                                        region = CellRangeAddress.ValueOf("BA" + (startAtRowNum + 1) + ":BA" + (rowNum + 1));
                                        sheet.AddMergedRegion(region);
                                        region = CellRangeAddress.ValueOf("BB" + (startAtRowNum + 1) + ":BB" + (rowNum + 1));
                                        sheet.AddMergedRegion(region);
                                    }
                                }

                                //Total row
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);
                                cell = row.CreateCell(0, ctString);
                                cell.SetCellValue("Total Weekly Hours");
                                cell.CellStyle = sMergedRightTotal;
                                cell = row.CreateCell(52, ctFormula);
                                cell.CellStyle = sTotal;
                                cell.CellFormula = "BA12+BA17+BA22+BA27+BA32";

                                region = CellRangeAddress.ValueOf("A37:AZ37");
                                sheet.AddMergedRegion(region);

                                sheet.SetColumnWidth(0, 4 * 256);
                                sheet.SetColumnWidth(1, 32 * 256);

                                sheet.SetColumnWidth(52, 8 * 256);
                                sheet.SetColumnWidth(53, 16 * 256);

                                //Signature rows at bottom
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);
                                cell = row.CreateCell(0, ctBlank);

                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);

                                cell = row.CreateCell(0, ctString);
                                cell.SetCellValue("Position:");

                                cell = row.CreateCell(1, ctBlank);

                                cell = row.CreateCell(2, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(3, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(4, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(5, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(6, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(7, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(8, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(9, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(10, ctBlank);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(11, ctBlank);

                                cell = row.CreateCell(12, ctString);
                                cell.SetCellValue("Print Name:");

                                cell = row.CreateCell(16, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(17, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(18, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(19, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(20, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(21, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(22, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(23, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(24, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(25, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(26, ctBlank);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(27, ctBlank);

                                cell = row.CreateCell(28, ctString);
                                cell.SetCellValue("Signed:");

                                cell = row.CreateCell(32, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(33, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(34, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(35, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(36, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(37, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(38, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(39, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(40, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(41, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(42, ctBlank);
                                cell.CellStyle = sUnderlined;

                                //Merge Cells
                                region = CellRangeAddress.ValueOf("A" + (rowNum + 1) + ":B" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("C" + (rowNum + 1) + ":K" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("M" + (rowNum + 1) + ":P" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("Q" + (rowNum + 1) + ":AA" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("AC" + (rowNum + 1) + ":AF" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("AG" + (rowNum + 1) + ":AQ" + (rowNum + 1));
                                sheet.AddMergedRegion(region);

                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);
                                cell = row.CreateCell(0, ctBlank);

                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);

                                cell = row.CreateCell(0, ctString);
                                cell.SetCellValue("Actioned By:");

                                cell = row.CreateCell(1, ctBlank);

                                cell = row.CreateCell(2, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(3, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(4, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(5, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(6, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(7, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(8, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(9, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(10, ctBlank);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(11, ctBlank);

                                cell = row.CreateCell(12, ctString);
                                cell.SetCellValue("Date:");

                                cell = row.CreateCell(16, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(17, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(18, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(19, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(20, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(21, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(22, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(23, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(24, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(25, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(26, ctBlank);
                                cell.CellStyle = sUnderlined;

                                cell = row.CreateCell(27, ctBlank);

                                cell = row.CreateCell(28, ctString);
                                cell.SetCellValue("Checked:");

                                cell = row.CreateCell(32, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(33, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(34, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(35, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(36, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(37, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(38, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(39, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(40, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(41, ctBlank);
                                cell.CellStyle = sUnderlined;
                                cell = row.CreateCell(42, ctBlank);
                                cell.CellStyle = sUnderlined;

                                //Merge Cells
                                region = CellRangeAddress.ValueOf("A" + (rowNum + 1) + ":B" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("C" + (rowNum + 1) + ":K" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("M" + (rowNum + 1) + ":P" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("Q" + (rowNum + 1) + ":AA" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("AC" + (rowNum + 1) + ":AF" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("AG" + (rowNum + 1) + ":AQ" + (rowNum + 1));
                                sheet.AddMergedRegion(region);

                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);
                                cell = row.CreateCell(0, ctBlank);

                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);

                                cell = row.CreateCell(0, ctString);
                                cell.SetCellValue("Comments/ Specialist Room Request:");

                                for (int i = 2; i <= 42; i++)
                                {
                                    cell = row.CreateCell(i, ctBlank);
                                    cell.CellStyle = sUnderlined;
                                }

                                region = CellRangeAddress.ValueOf("A" + (rowNum + 1) + ":B" + (rowNum + 1));
                                sheet.AddMergedRegion(region);
                                region = CellRangeAddress.ValueOf("C" + (rowNum + 1) + ":AQ" + (rowNum + 1));
                                sheet.AddMergedRegion(region);

                                //This should be coded into loop to ensure it is applied to correct ranges
                                int startRow;
                                int endRow;

                                //Header - does not work when cells have backgrounds
                                //startRow = 11;
                                //endRow = 11;

                                //region = CellRangeAddress.ValueOf("C" + startRow + ":D" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("E" + startRow + ":H" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("I" + startRow + ":L" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("M" + startRow + ":P" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("Q" + startRow + ":T" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("U" + startRow + ":X" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("Y" + startRow + ":AB" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AC" + startRow + ":AF" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AG" + startRow + ":AJ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AK" + startRow + ":AN" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AO" + startRow + ":AR" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AS" + startRow + ":AV" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AW" + startRow + ":AZ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //Monday
                                startRow = 12;
                                endRow = 16;

                                //region = CellRangeAddress.ValueOf("C" + startRow + ":D" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("E" + startRow + ":H" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("I" + startRow + ":L" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("M" + startRow + ":P" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Q" + startRow + ":T" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("U" + startRow + ":X" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Y" + startRow + ":AB" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AC" + startRow + ":AF" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AG" + startRow + ":AJ" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AK" + startRow + ":AN" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AO" + startRow + ":AR" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AS" + startRow + ":AV" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AW" + startRow + ":AZ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //Tuesday
                                startRow = 17;
                                endRow = 21;

                                //region = CellRangeAddress.ValueOf("C" + startRow + ":D" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("E" + startRow + ":H" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("I" + startRow + ":L" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("M" + startRow + ":P" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Q" + startRow + ":T" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("U" + startRow + ":X" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Y" + startRow + ":AB" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AC" + startRow + ":AF" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AG" + startRow + ":AJ" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AK" + startRow + ":AN" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AO" + startRow + ":AR" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AS" + startRow + ":AV" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AW" + startRow + ":AZ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //Wednesday
                                startRow = 22;
                                endRow = 26;

                                //region = CellRangeAddress.ValueOf("C" + startRow + ":D" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("E" + startRow + ":H" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("I" + startRow + ":L" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("M" + startRow + ":P" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Q" + startRow + ":T" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("U" + startRow + ":X" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Y" + startRow + ":AB" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AC" + startRow + ":AF" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AG" + startRow + ":AJ" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AK" + startRow + ":AN" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AO" + startRow + ":AR" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AS" + startRow + ":AV" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AW" + startRow + ":AZ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //Thursday
                                startRow = 27;
                                endRow = 31;

                                //region = CellRangeAddress.ValueOf("C" + startRow + ":D" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("E" + startRow + ":H" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("I" + startRow + ":L" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("M" + startRow + ":P" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Q" + startRow + ":T" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("U" + startRow + ":X" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Y" + startRow + ":AB" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AC" + startRow + ":AF" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AG" + startRow + ":AJ" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AK" + startRow + ":AN" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AO" + startRow + ":AR" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AS" + startRow + ":AV" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AW" + startRow + ":AZ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //Friday
                                startRow = 32;
                                endRow = 36;

                                //region = CellRangeAddress.ValueOf("C" + startRow + ":D" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("E" + startRow + ":H" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("I" + startRow + ":L" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("M" + startRow + ":P" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Q" + startRow + ":T" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("U" + startRow + ":X" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("Y" + startRow + ":AB" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AC" + startRow + ":AF" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                region = CellRangeAddress.ValueOf("AG" + startRow + ":AJ" + endRow);
                                RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AK" + startRow + ":AN" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AO" + startRow + ":AR" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AS" + startRow + ":AV" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);

                                //region = CellRangeAddress.ValueOf("AW" + startRow + ":AZ" + endRow);
                                //RegionUtil.SetBorderTop(2, region, sheet, workbook);
                                //RegionUtil.SetBorderBottom(2, region, sheet, workbook);
                                //RegionUtil.SetBorderLeft(2, region, sheet, workbook);
                                //RegionUtil.SetBorderRight(2, region, sheet, workbook);
                            }
                        }
                        else
                        {
                            sheet = workbook.CreateSheet("ERROR No Groups");
                            row = sheet.CreateRow(0);
                            cell = row.CreateCell(0, ctString);
                            cell.CellStyle = sHeader;
                            cell.SetCellValue("Error - No groups could be loaded. Please check GroupData Stored Procedure");
                        }

                        if (Week != null && Week.Count > 0)
                        {
                            //Create Weeks Sheet
                            sheet = workbook.CreateSheet("Weeks");
                            row = sheet.CreateRow(0);
                            cell = row.CreateCell(0);

                            cell.SetCellValue("Weeks");
                            cell.CellStyle = sHeader;

                            row = sheet.CreateRow(1);
                            row = sheet.CreateRow(2);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Week Num");
                            cell.CellStyle = sTableHeader;

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue("Week Desc");
                            cell.CellStyle = sTableHeader;

                            cell = row.CreateCell(2, ctString);
                            cell.SetCellValue("Notes");
                            cell.CellStyle = sTableHeader;

                            cell = row.CreateCell(3, ctString);
                            cell.SetCellValue("Other Details");
                            cell.CellStyle = sTableHeader;

                            rowNum = 2;

                            foreach (var week in Week)
                            {
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);

                                cell = row.CreateCell(0, ctNumber);
                                if (week.WeekNum >= 0)
                                {
                                    cell.SetCellValue((double)week.WeekNum);
                                }
                                cell.CellStyle = sBorderMedium;

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue(week.WeekDesc);
                                cell.CellStyle = sBorderMedium;

                                cell = row.CreateCell(2, ctString);
                                cell.SetCellValue(week.Notes);
                                cell.CellStyle = sBorderMedium;

                                cell = row.CreateCell(3, ctString);
                                cell.SetCellValue(week.Notes2);
                                cell.CellStyle = sBorderMedium;
                            }

                            //Column widths
                            sheet.SetColumnWidth(0, 8 * 256);
                            sheet.SetColumnWidth(1, 20 * 256);
                            sheet.SetColumnWidth(2, 20 * 256);
                            sheet.SetColumnWidth(3, 40 * 256);
                        }

                        if (TermDate != null && TermDate.Count > 0)
                        {
                            //Create TermDates Sheet
                            sheet = workbook.CreateSheet("Term Dates");
                            row = sheet.CreateRow(0);
                            cell = row.CreateCell(0);

                            cell.SetCellValue("Term Dates");
                            cell.CellStyle = sHeader;

                            row = sheet.CreateRow(1);

                            row = sheet.CreateRow(2);
                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Academic Year " + CurrentAcademicYear);
                            cell.CellStyle = sTableHeader;

                            cell = row.CreateCell(1, ctBlank);
                            cell.CellStyle = sTableHeader;

                            //Merge top table header row
                            region = CellRangeAddress.ValueOf("A3:B3");
                            sheet.AddMergedRegion(region);

                            row = sheet.CreateRow(3);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Term Date");
                            cell.CellStyle = sTableHeader;

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue("Dates");
                            cell.CellStyle = sTableHeader;

                            rowNum = 3;

                            foreach (var termDate in TermDate)
                            {
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);

                                XSSFCellStyle cellStyle;

                                if (termDate.IsTerm == true)
                                {
                                    cellStyle = sTotal;
                                }
                                else
                                {
                                    cellStyle = sBorderMedium;
                                }

                                cell = row.CreateCell(0, ctNumber);
                                cell.SetCellValue(termDate.TermDateName);
                                cell.CellStyle = cellStyle;

                                cell = row.CreateCell(1, ctString);
                                cell.SetCellValue(termDate.Dates);
                                cell.CellStyle = sBorderMedium;
                            }

                            //Add extra rows
                            rowNum += 1;
                            row = sheet.CreateRow(rowNum);

                            //Merge bottom row
                            region = CellRangeAddress.ValueOf("A" + (rowNum + 1) + ":B" + (rowNum + 1));
                            sheet.AddMergedRegion(region);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("*End of year dates will vary by course; some students may finish earlier or later");
                            cell.CellStyle = sTotalRight;



                            //Column widths
                            sheet.SetColumnWidth(0, 20 * 256);
                            sheet.SetColumnWidth(1, 90 * 256);

                            if (BankHoliday != null && BankHoliday.Count > 0)
                            {
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);
                                rowNum += 1;
                                row = sheet.CreateRow(rowNum);

                                cell = row.CreateCell(0);
                                cell.SetCellValue("Bank Holidays");
                                cell.CellStyle = sTableHeader;

                                cell = row.CreateCell(1);
                                cell.CellStyle = sTableHeader;

                                region = CellRangeAddress.ValueOf("A" + (rowNum + 1) + ":B" + (rowNum + 1));
                                sheet.AddMergedRegion(region);

                                foreach (var bankHoliday in BankHoliday)
                                {
                                    rowNum += 1;
                                    row = sheet.CreateRow(rowNum);

                                    cell = row.CreateCell(0, ctString);
                                    cell.SetCellValue(bankHoliday.BankHolidayDesc);
                                    cell.CellStyle = sBorderMedium;

                                    cell = row.CreateCell(1, ctBlank);
                                    cell.CellStyle = sBorderMedium;

                                    region = CellRangeAddress.ValueOf("A" + (rowNum + 1) + ":B" + (rowNum + 1));
                                    sheet.AddMergedRegion(region);
                                }
                            }
                        }

                        workbook.Write(fs);
                    }
                    using (var stream = new FileStream(Path.Combine(filePath, fileName), FileMode.Open))
                    {
                        await stream.CopyToAsync(memory);
                    }
                    memory.Position = 0;
                }
            }

            //Now create a compressed zip file with all timetables in deleting any previous zips
            ZipFile.CreateFromDirectory(ExportPath, ZipPath + @"\Timetables.zip");

            //Now delete the folder where the timetables were generated leaving just the zip file in the folder
            if (Directory.Exists(ExportPath))
            {
                Directory.Delete(ExportPath, true);
            }

            return numFilesSaved;
        }
    }
}