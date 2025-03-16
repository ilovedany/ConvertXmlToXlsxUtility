using Spire.Xls;
using System.Xml;
using System.Text;
using ConvertXmlToXlsxUtility;
using System.Diagnostics;
class Program{
    static void Main(string[] args){

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        XmlDocument xDoc = new XmlDocument();
        string[] columns = {"ТабНом", "ФИО", "УчебноеЗаведение", "Специальность", "ПрежняяДолжность", "ДатаНазначенияНаПрежнююДолжность", "НоваяДолжность", "ДатаНазначения","ВидРезерва","Уровень","Должность"};
        xDoc.Load(args[0]); //первый параметр - путь к xml файлу

        XmlElement? xRoot = xDoc.DocumentElement;
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        WorksheetColumns worksheetColumns = new WorksheetColumns();
        FileInfo xmlFile = new FileInfo(args[0]);

        const int startPosition = 4;

        string[] columns2 = new string[columns.Length - startPosition];
        Array.Copy(columns, startPosition, columns2, 0, columns2.Length);

        int count_G_List = 0;
        int countLine = 0;

        countLine++;
        worksheetColumns.WorkingWithXml(workbook,countLine,columns);
        
        foreach (XmlNode xnode in xRoot)
        {  
            foreach (XmlNode userNode in xnode.ChildNodes)
            {
                countLine++;
                worksheetColumns.WorkingWithXml(userNode, workbook, countLine, columns, 0);
                
                foreach (XmlNode profList in userNode.ChildNodes)
                {
                    foreach (XmlNode userList in profList.ChildNodes)
                    {
                        if (userList.NodeType == XmlNodeType.Element)
                        {           
                            worksheetColumns.WorkingWithXml(userList, workbook, countLine, columns2, startPosition);
                            countLine++;
                            count_G_List++;
                        }
                    }
                }
                countLine--;
                worksheetColumns.WorkingWithXml(workbook,countLine-count_G_List+1,countLine,startPosition);
                count_G_List = 0;
            }
            
        }
        worksheet.AllocatedRange.AutoFitColumns();
        workbook.SaveToFile(args[1], ExcelVersion.Version2016);
        Process.Start(new ProcessStartInfo(args[1]) { UseShellExecute = true });
        xmlFile.Delete();

    }
}