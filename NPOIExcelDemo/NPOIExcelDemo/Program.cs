// See https://aka.ms/new-console-template for more information
using NPOIExcelDemo;

var workbook=NPOIHelper.GenerateExcel();
using (FileStream stream = new FileStream("E:\\09项目实战\\NPOIExcelDemo\\NPOIExcelDemo\\NPOIExcelDemo\\NPOITest.xls", FileMode.Create))
{
    workbook.Write(stream);
}
