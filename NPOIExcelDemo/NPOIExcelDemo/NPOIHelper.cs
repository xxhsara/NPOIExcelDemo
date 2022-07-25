using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIExcelDemo
{
    public class NPOIHelper
    {
        public IWorkbook GenerateExcel()
        {
            HSSFWorkbook hSSFWorkbook = new HSSFWorkbook();
            ISheet sheet1 = hSSFWorkbook.CreateSheet("Sheet1");

            IRow row0 = sheet1.CreateRow(0); //第一行
            ICell row0Cell0 = row0.CreateCell(0);
            row0Cell0.SetCellValue("第一列");

            ICell row0Cell1 = row0.CreateCell(1);
            row0Cell1.SetCellValue("第二列");

            IRow row1 = sheet1.CreateRow(1); //第二行
            ICell row1Cell0 = row1.CreateCell(0);
            row1Cell0.SetCellValue("111");

            ICell row1Cell1 = row1.CreateCell(1);
            row1Cell1.SetCellValue("222");

            return hSSFWorkbook;
        }
    }
}
