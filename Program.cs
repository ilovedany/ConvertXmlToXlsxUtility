using Spire.Xls;
using System.Xml;
using System.Text;
using ConvertXmlToXlsxUtility;
class Program{
    static void Main(string[] args){

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        XmlDocument xDoc = new XmlDocument();
        
        xDoc.Load(args[0]); //первый параметр - путь к xml файлу

        XmlElement? xRoot = xDoc.DocumentElement;

        Workbook workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets[0];

        WorksheetColumns worksheetColumns = new WorksheetColumns();

        string[] columns = new string[args.Length-2];
        Array.Copy(args, 2, columns, 0, columns.Length);

        int lineCount = 0;
        int countLine = 1;
        string[] columns2 = new string[columns.Length - 4];
        Array.Copy(columns, 4, columns2, 0, columns2.Length); 

        worksheetColumns.WorkingWithXml(workbook,countLine,columns);
       
        if (xRoot != null)
        {
            foreach (XmlElement xnode in xRoot)
            {
                if (xnode.Name == "LIST_G_KADR")
                {
                    foreach (XmlElement userNode in xnode.GetElementsByTagName("G_KADR"))
                    {
                        countLine++;
                        worksheetColumns.WorkingWithXml(userNode, workbook, countLine, columns,0);

                        foreach (XmlElement profList in userNode.GetElementsByTagName("LIST_G_PROF"))
                        {
                            foreach (XmlElement userList in profList.GetElementsByTagName("G_PROF"))
                            {
                                worksheetColumns.WorkingWithXml(userList, workbook, countLine, columns2,4);
                                countLine++;
                                lineCount++;

                            }
                        }
                        countLine--;
                        worksheetColumns.WorkingWithXml(workbook,countLine-lineCount+1,countLine,4);
                        lineCount = 0;
                    }     
                }
            }
        }
        worksheet.AllocatedRange.AutoFitColumns();
        workbook.SaveToFile(args[1], ExcelVersion.Version2016);
    }
}