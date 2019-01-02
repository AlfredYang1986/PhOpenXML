using OMT01.phXlsx;
using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OMT01 {
    class Program {
        static void Main(string[] args) {
            phExcelFacade.getInstance().CreateNewExcel("alfred.xlsx");
            //phExcelFacade.getInstance().PushValueInCell("alfred.xlsx", "yang", "A1");
            //phExcelFacade.getInstance().PushValueInCell("alfred.xlsx", "yuan", "D1");
            //phExcelFacade.getInstance().MergeCell("alfred.xlsx", "A1", "B2");
            //phExcelFacade.getInstance().AddFont("alfred.xlsx", "Calibri", 16, "FF0000", true, "A1");
            //phExcelFacade.getInstance().AddFill("alfred.xlsx", "FFFF00", "D1");
            //phExcelFacade.getInstance().AddAlignment("alfred.xlsx", HorizontalAlignmentValues.Center, VerticalAlignmentValues.Center, "D1");
            Console.ReadKey();
        }
    }
}
