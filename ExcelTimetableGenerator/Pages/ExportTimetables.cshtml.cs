using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTimetableGenerator.Models;
using ExcelTimetableGenerator.Shared;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace ExcelTimetableGenerator.Pages
{
    public class ExportTimetablesModel : PageModel
    {
        private readonly ExcelTimetableGenerator.Data.ApplicationDbContext _context;
        private IHostingEnvironment _hostingEnvironment;
        public ExportTimetablesModel(ExcelTimetableGenerator.Data.ApplicationDbContext context, IHostingEnvironment hostingEnvironment)
        {
            _context = context;
            _hostingEnvironment = hostingEnvironment;
        }

        public string SavePath { get; set; }

        public int NumFilesExported { get; set; }

        public IList<Programme> Programme { get; set; }
        public IList<Course> Course { get; set; }
        public IList<Group> Group { get; set; }

        public IList<Day> Day { get; set; }
        public IList<Time> Time { get; set; }
        public IList<TimetableSection> TimetableSection { get; set; }

        public async Task<IActionResult> OnGetAsync(string academicYear, int planRevisionID)
        {
            planRevisionID = 66;

            string CurrentAcademicYear = await AcademicYearFunctions.GetAcademicYear(academicYear, _context);
            var academicYearParam = new SqlParameter("@AcademicYear", CurrentAcademicYear);
            var planRevisionIDParam = new SqlParameter("@PlanRevisionID", planRevisionID);

            //Data from Curriculum Planning
            Course = await _context.Course
                .FromSql("EXEC SPR_ETG_CourseData @AcademicYear, @PlanRevisionID", academicYearParam, planRevisionIDParam)
                .ToListAsync();

            Group = await _context.Group
                .FromSql("EXEC SPR_ETG_GroupData @AcademicYear, @PlanRevisionID", academicYearParam, planRevisionIDParam)
                .ToListAsync();

            Programme = await _context.Programme
                .FromSql("EXEC SPR_ETG_ProgrammeData @AcademicYear, @PlanRevisionID", academicYearParam, planRevisionIDParam)
                .ToListAsync();

            //Data for Tiemtable Grid
            Day = await _context.Day
                .FromSql("EXEC SPR_ETG_Day")
                .ToListAsync();

            Time = await _context.Time
                .FromSql("EXEC SPR_ETG_Time")
                .ToListAsync();

            TimetableSection = await _context.TimetableSection
                .FromSql("EXEC SPR_ETG_TimetableSection")
                .ToListAsync();

            SavePath = _hostingEnvironment.WebRootPath + @"\ExportedTimetables";
            string filePath = null;
            string fileName = null;
            string fileURL = null;
            int rowNum = 0;
            int colNum = 0;
            NumFilesExported = 0;

            //College Logo
            var collegeLogoStream = new System.IO.FileStream(_hostingEnvironment.WebRootPath + @"\images\CollegeLogo.png", System.IO.FileMode.Open);
            byte[] bytes = IOUtils.ToByteArray(collegeLogoStream);
            collegeLogoStream.Close();
            collegeLogoStream.Dispose();

            foreach (var programme in Programme)
            {
                NumFilesExported += 1;

                //Save timetables into departmental and team folders (creates a folder if it doesn't exist)
                filePath = SavePath + @"\" + programme.FacCode + @"\" + programme.TeamCode + @"\";
                Directory.CreateDirectory(filePath);

                //Generate each Excel workbook
                fileName = @"" + programme.ProgCode + " - " + programme.ProgTitle + ".xlsx";
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
                    IDrawing drawing;
                    IClientAnchor anchor;
                    IPicture picture;

                    //Add college logo
                    int collegeLogo = workbook.AddPicture(bytes, PictureType.PNG);

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

                    //Cell Borders
                    BorderStyle border = BorderStyle.Medium;

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

                    //Table Header
                    XSSFCellStyle sTableHeader = (XSSFCellStyle)workbook.CreateCellStyle();
                    sTableHeader.SetFont(fBold);
                    sTableHeader.SetFillForegroundColor(cLightBlue);
                    sTableHeader.FillPattern = FillPattern.SolidForeground;
                    sTableHeader.BorderTop = border;
                    sTableHeader.BorderBottom = border;
                    sTableHeader.BorderLeft = border;
                    sTableHeader.BorderRight = border;
                    sTableHeader.WrapText = true;

                    //Date
                    XSSFCellStyle sDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));
                    sDate.BorderTop = border;
                    sDate.BorderBottom = border;
                    sDate.BorderLeft = border;
                    sDate.BorderRight = border;

                    //Types of courses
                    XSSFCellStyle sMainAim = (XSSFCellStyle)workbook.CreateCellStyle();
                    sMainAim.SetFillForegroundColor(cMainAim);
                    sMainAim.FillPattern = FillPattern.SolidForeground;
                    sMainAim.BorderTop = border;
                    sMainAim.BorderBottom = border;
                    sMainAim.BorderLeft = border;
                    sMainAim.BorderRight = border;
                    sMainAim.SetFont(fBold);

                    XSSFCellStyle sEngMaths = (XSSFCellStyle)workbook.CreateCellStyle();
                    sEngMaths.SetFillForegroundColor(cEngMaths);
                    sEngMaths.FillPattern = FillPattern.SolidForeground;
                    sEngMaths.BorderTop = border;
                    sEngMaths.BorderBottom = border;
                    sEngMaths.BorderLeft = border;
                    sEngMaths.BorderRight = border;

                    XSSFCellStyle sWex = (XSSFCellStyle)workbook.CreateCellStyle();
                    sWex.SetFillForegroundColor(cWex);
                    sWex.FillPattern = FillPattern.SolidForeground;
                    sWex.BorderTop = border;
                    sWex.BorderBottom = border;
                    sWex.BorderLeft = border;
                    sWex.BorderRight = border;

                    XSSFCellStyle sDSS = (XSSFCellStyle)workbook.CreateCellStyle();
                    sDSS.SetFillForegroundColor(cDSS);
                    sDSS.FillPattern = FillPattern.SolidForeground;
                    sDSS.BorderTop = border;
                    sDSS.BorderBottom = border;
                    sDSS.BorderLeft = border;
                    sDSS.BorderRight = border;

                    XSSFCellStyle sTut = (XSSFCellStyle)workbook.CreateCellStyle();
                    sTut.SetFillForegroundColor(cTut);
                    sTut.FillPattern = FillPattern.SolidForeground;
                    sTut.BorderTop = border;
                    sTut.BorderBottom = border;
                    sTut.BorderLeft = border;
                    sTut.BorderRight = border;

                    XSSFCellStyle sOther = (XSSFCellStyle)workbook.CreateCellStyle();
                    sOther.SetFillForegroundColor(cOther);
                    sOther.FillPattern = FillPattern.SolidForeground;
                    sOther.BorderTop = border;
                    sOther.BorderBottom = border;
                    sOther.BorderLeft = border;
                    sOther.BorderRight = border;

                    //Border Only
                    XSSFCellStyle sBorder = (XSSFCellStyle)workbook.CreateCellStyle();
                    sBorder.BorderTop = border;
                    sBorder.BorderBottom = border;
                    sBorder.BorderLeft = border;
                    sBorder.BorderRight = border;

                    //Types of courses (date fields
                    XSSFCellStyle sMainAimDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sMainAimDate.SetFillForegroundColor(cMainAim);
                    sMainAimDate.FillPattern = FillPattern.SolidForeground;
                    sMainAimDate.BorderTop = border;
                    sMainAimDate.BorderBottom = border;
                    sMainAimDate.BorderLeft = border;
                    sMainAimDate.BorderRight = border;
                    sMainAimDate.SetFont(fBold);
                    sMainAimDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    XSSFCellStyle sEngMathsDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sEngMathsDate.SetFillForegroundColor(cEngMaths);
                    sEngMathsDate.FillPattern = FillPattern.SolidForeground;
                    sEngMathsDate.BorderTop = border;
                    sEngMathsDate.BorderBottom = border;
                    sEngMathsDate.BorderLeft = border;
                    sEngMathsDate.BorderRight = border;
                    sEngMathsDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    XSSFCellStyle sWexDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sWexDate.SetFillForegroundColor(cWex);
                    sWexDate.FillPattern = FillPattern.SolidForeground;
                    sWexDate.BorderTop = border;
                    sWexDate.BorderBottom = border;
                    sWexDate.BorderLeft = border;
                    sWexDate.BorderRight = border;
                    sWexDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    XSSFCellStyle sDSSDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sDSSDate.SetFillForegroundColor(cDSS);
                    sDSSDate.FillPattern = FillPattern.SolidForeground;
                    sDSSDate.BorderTop = border;
                    sDSSDate.BorderBottom = border;
                    sDSSDate.BorderLeft = border;
                    sDSSDate.BorderRight = border;
                    sDSSDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    XSSFCellStyle sTutDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sTutDate.SetFillForegroundColor(cTut);
                    sTutDate.FillPattern = FillPattern.SolidForeground;
                    sTutDate.BorderTop = border;
                    sTutDate.BorderBottom = border;
                    sTutDate.BorderLeft = border;
                    sTutDate.BorderRight = border;
                    sTutDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    XSSFCellStyle sOtherDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sOtherDate.SetFillForegroundColor(cOther);
                    sOtherDate.FillPattern = FillPattern.SolidForeground;
                    sOtherDate.BorderTop = border;
                    sOtherDate.BorderBottom = border;
                    sOtherDate.BorderLeft = border;
                    sOtherDate.BorderRight = border;
                    sOtherDate.SetDataFormat(createHelper.CreateDataFormat().GetFormat("dd/mm/yyyy"));

                    //Border Only
                    XSSFCellStyle sBorderDate = (XSSFCellStyle)workbook.CreateCellStyle();
                    sBorderDate.BorderTop = border;
                    sBorderDate.BorderBottom = border;
                    sBorderDate.BorderLeft = border;
                    sBorderDate.BorderRight = border;
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
                    picture = drawing.CreatePicture(anchor, collegeLogo);

                    cell = row.CreateCell(0);
                    
                    cell.SetCellValue(programme.ProgCode + " - " + programme.ProgTitle);
                    cell.CellStyle = sHeader;

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
                    cell.SetCellValue("Planned Learning Hours 16-18");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(7, ctString);
                    cell.SetCellValue("Planned EEP Hours 16-18");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(8, ctString);
                    cell.SetCellValue("Planned Learning Hours 19+");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(9, ctString);
                    cell.SetCellValue("Planned EEP Hours 19+");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(10, ctString);
                    cell.SetCellValue("Start Date");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(11, ctString);
                    cell.SetCellValue("End Date");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(12, ctString);
                    cell.SetCellValue("Site");
                    cell.CellStyle = sTableHeader;

                    cell = row.CreateCell(13, ctString);
                    cell.SetCellValue("Notes");
                    cell.CellStyle = sTableHeader;

                    //Column widths
                    sheet.SetColumnWidth(0, 16 * 256);
                    sheet.SetColumnWidth(1, 40 * 256);
                    sheet.SetColumnWidth(2, 14 * 256);
                    sheet.SetColumnWidth(3, 12 * 256);
                    sheet.SetColumnWidth(4, 8 * 256);
                    sheet.SetColumnWidth(5, 8 * 256);
                    sheet.SetColumnWidth(6, 10 * 256);
                    sheet.SetColumnWidth(7, 10 * 256);
                    sheet.SetColumnWidth(8, 10 * 256);
                    sheet.SetColumnWidth(9, 10 * 256);
                    sheet.SetColumnWidth(10, 12 * 256);
                    sheet.SetColumnWidth(11, 12 * 256);
                    sheet.SetColumnWidth(12, 20 * 256);
                    sheet.SetColumnWidth(13, 20 * 256);
                    sheet.SetColumnWidth(14, 20 * 256);

                    //The current row in the worksheet
                    rowNum = 2;

                    //Generate each Excel worksheet
                    if (programme.Course != null && programme.Course.Count > 0)
                    {
                        foreach (var course in programme.Course)
                        {
                            rowNum += 1;
                            row = sheet.CreateRow(rowNum);
                            XSSFCellStyle cellStyle;
                            XSSFCellStyle cellStyleDate;

                            switch(course.CourseOrder)
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
                            cell.SetCellValue(course.CourseCode);
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue(course.CourseTitle);
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(2, ctString);
                            cell.SetCellValue(course.AimCode);
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(3, ctString);
                            cell.SetCellValue(course.AwardBody);
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(4, ctNumber);
                            if(course.HoursPerWeek >= 0)
                            {
                                cell.SetCellValue((double)course.HoursPerWeek);
                            }
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(5, ctNumber);
                            if (course.Weeks >= 0)
                            {
                                cell.SetCellValue((double)course.Weeks);
                            }
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(6, ctNumber);
                            if (course.PLH1618 >= 0)
                            {
                                cell.SetCellValue((double)course.PLH1618);
                            }
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(7, ctNumber);
                            if (course.EEP1618 >= 0)
                            {
                                cell.SetCellValue((double)course.EEP1618);
                            }
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(8, ctNumber);
                            if (course.PLH19 >= 0)
                            {
                                cell.SetCellValue((double)course.PLH19);
                            }
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(9, ctNumber);
                            if (course.EEP19 >= 0)
                            {
                                cell.SetCellValue((double)course.EEP19);
                            }
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(10, ctNumber);
                            cell.CellStyle = cellStyleDate;
                            cell.SetCellValue((DateTime)course.StartDate);

                            cell = row.CreateCell(11, ctNumber);
                            cell.CellStyle = cellStyleDate;
                            cell.SetCellValue((DateTime)course.EndDate);

                            cell = row.CreateCell(12, ctString);
                            cell.SetCellValue(course.SiteName);
                            cell.CellStyle = cellStyle;

                            cell = row.CreateCell(13, ctString);
                            cell.SetCellValue(course.Notes);
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
                            anchor.Col1 = 11;
                            anchor.Row1 = 0;
                            anchor.Col2 = 13;
                            anchor.Row2 = 1;
                            picture = drawing.CreatePicture(anchor, collegeLogo);

                            cell = row.CreateCell(0);

                            cell.SetCellValue("Group " + group.GroupCode + " for " + programme.ProgCode + " - " + programme.ProgTitle);
                            cell.CellStyle = sHeader;

                            row = sheet.CreateRow(1);
                            row = sheet.CreateRow(2);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Programme Details");
                            cell.CellStyle = sSubHeader;

                            row = sheet.CreateRow(3);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Faculty:");

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue(programme.FacName);
                            cell.CellStyle = sBorder;

                            cell = row.CreateCell(2, ctString);

                            cell = row.CreateCell(3, ctString);
                            cell.SetCellValue("Mode:");

                            cell = row.CreateCell(4, ctString);
                            cell.SetCellValue(programme.ModeOfAttendanceName);
                            cell.CellStyle = sBorder;

                            row = sheet.CreateRow(4);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Team:");

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue(programme.TeamName);
                            cell.CellStyle = sBorder;

                            cell = row.CreateCell(2, ctString);

                            cell = row.CreateCell(3, ctString);
                            cell.SetCellValue("Site:");

                            cell = row.CreateCell(4, ctString);
                            cell.SetCellValue(programme.SiteName);
                            cell.CellStyle = sBorder;

                            row = sheet.CreateRow(5);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Parent Code:");

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue(programme.ProgCode);
                            cell.CellStyle = sBorder;

                            cell = row.CreateCell(2, ctString);

                            cell = row.CreateCell(3, ctString);
                            cell.SetCellValue("Prog Planned Hours:");

                            cell = row.CreateCell(4, ctNumber);
                            cell.SetCellValue((double)(programme.PLHMax + programme.EEPMax));
                            cell.CellStyle = sBorder;

                            row = sheet.CreateRow(6);

                            cell = row.CreateCell(0, ctString);
                            cell.SetCellValue("Title:");

                            cell = row.CreateCell(1, ctString);
                            cell.SetCellValue(programme.ProgTitle);
                            cell.CellStyle = sBorder;

                            row = sheet.CreateRow(7);
                            row = sheet.CreateRow(8);

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
                                            cell.CellStyle = sTableHeader;
                                        }
                                    }
                                }
                            }
                            //Draw timetable grid
                            //The current row in the worksheet
                            rowNum = 8;
                            if (Day != null && Day.Count > 0)
                            {
                                foreach (var day in Day)
                                {
                                    if (TimetableSection != null && TimetableSection.Count > 0)
                                    {
                                        foreach (var section in TimetableSection)
                                        {
                                            rowNum += 1;

                                            row = sheet.CreateRow(rowNum);

                                            cell = row.CreateCell(0, ctString);
                                            cell.SetCellValue(day.DayName);

                                            cell = row.CreateCell(1, ctString);
                                            cell.SetCellValue(section.SectionName);

                                            colNum = 1;

                                            if (Time != null && Time.Count > 0)
                                            {
                                                foreach (var time in Time)
                                                {
                                                    colNum += 1;

                                                    cell = row.CreateCell(colNum, ctBlank);
                                                    cell.CellStyle = sBorder;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            sheet.SetColumnWidth(0, 16 * 256);
                            sheet.SetColumnWidth(1, 40 * 256);
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

                    workbook.Write(fs);
                }
                using (var stream = new FileStream(Path.Combine(filePath, fileName), FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
            }

            return Page();
        }
    }
}