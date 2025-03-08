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

        //string[] columns = new string[args.Length-2];
        //Array.Copy(args, 2, columns, 0, columns.Length);

        int maxLine = 0;

        int count_G_List = 0;
        int count_G_Reserv = 0;
        
        int countLine = 1;

        string[] columns2 = new string[columns.Length - 4];
        Array.Copy(columns, 4, columns2, 0, columns2.Length);
        
        string[] columns3 = new string[columns.Length - 8];
        Array.Copy(columns, 8, columns3, 0, columns3.Length);

        worksheetColumns.WorkingWithXml(workbook,countLine,columns);
        
        foreach (XmlElement xnode in xRoot)
        {
            if (xnode.Name == "LIST_G_KADR")
            {
                foreach (XmlElement userNode in xnode.GetElementsByTagName("G_KADR"))
                {
                    countLine++;

                    worksheetColumns.WorkingWithXml(userNode, workbook, countLine, columns, 0);

                    foreach (XmlElement profList in userNode.GetElementsByTagName("LIST_G_PROF"))
                    {
                        foreach (XmlElement userList in profList.GetElementsByTagName("G_PROF"))
                        {
                            worksheetColumns.WorkingWithXml(userList, workbook, countLine, columns2, 4);

                            countLine++;
                            count_G_List++;

                        }
                    }
                    if (columns.Count() == 11)
                    {
                        countLine -= count_G_List;

                        foreach (XmlElement reservList in userNode.GetElementsByTagName("LIST_G_RESERV"))
                        {
                            foreach (XmlElement greserv in reservList.GetElementsByTagName("G_RESERV"))
                            {
                                worksheetColumns.WorkingWithXml(greserv, workbook, countLine, columns3, 8);
                                countLine++;
                                count_G_Reserv++;
                            }
                        }
                        if(count_G_List > count_G_Reserv){
                            countLine += count_G_List-count_G_Reserv;
                            maxLine = count_G_List;
                        }
                        else{
                            maxLine = count_G_Reserv;
                        }
                        count_G_Reserv=0;
                    }

                    countLine--;

                    if (columns.Count() == 11){

                        worksheetColumns.WorkingWithXml(workbook,countLine-maxLine+1,countLine,4);

                    }
                    else{
                        worksheetColumns.WorkingWithXml(workbook,countLine-count_G_List+1,countLine,4);
                    }

                    count_G_List=0;
                    
                }
                
            }
        }
        worksheet.AllocatedRange.AutoFitColumns();
        workbook.SaveToFile(args[1], ExcelVersion.Version2016);
        
        Process.Start(new ProcessStartInfo(args[1]) { UseShellExecute = true });
        xmlFile.Delete();

    }
}
