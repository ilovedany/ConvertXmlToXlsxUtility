using Spire.Xls;
using System.Xml;
using System.Text;
class Program{
    static void Main(string[] args){

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        XmlDocument xDoc = new XmlDocument();
        
        xDoc.Load(args[0]); //первый параметр - путь к xml файлу




        XmlElement? xRoot = xDoc.DocumentElement;

        Workbook workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets[0];

        

        worksheet.Range[1,1].Value = "ТабНом";
        worksheet.Range[1,2].Value = "ФИО";
        worksheet.Range[1,3].Value = "УчебноеЗаведение";
        worksheet.Range[1,4].Value = "Специальность";
        worksheet.Range[1,5].Value = "ПрежняяДолжность";
        worksheet.Range[1,6].Value = "ДатаНазначенияНаПрежнююДолжность";
        worksheet.Range[1,7].Value = "НоваяДолжность";
        worksheet.Range[1,8].Value = "ДатаНазначения";


        CellStyle style = workbook.Styles.Add("newStyle");
        style.Font.IsBold = true;
        worksheet.Range[1, 1, 1, 8].Style = style;

        int countLine=1;
        int lineCount = 0;
        if (xRoot != null)
        {

            foreach (XmlElement xnode in xRoot)
            {
                if (xnode.Name == "LIST_G_KADR")
                {

                    foreach (XmlElement userNode in xnode.GetElementsByTagName("G_KADR"))
                    {
                        countLine++;

                        worksheet.Range[countLine,1].Value =  userNode["ТабНом"]?.InnerText;
                        
                        worksheet.Range[countLine,2].Value = userNode["ФИО"]?.InnerText;
                        worksheet.Range[countLine,3].Value = userNode["УчебноеЗаведение"]?.InnerText;
                        worksheet.Range[countLine,4].Value = userNode["Специальность"]?.InnerText;

                        foreach (XmlElement profList in userNode.GetElementsByTagName("LIST_G_PROF"))
                        {
                            foreach (XmlElement userList in profList.GetElementsByTagName("G_PROF"))
                            {
                                worksheet.Range[countLine,5].Value =  userList["ПрежняяДолжность"]?.InnerText;
                                worksheet.Range[countLine,6].Value =  userList["ДатаНазначенияНаПрежнююДолжность"]?.InnerText;
                                worksheet.Range[countLine,7].Value =  userList["НоваяДолжность"]?.InnerText;
                                worksheet.Range[countLine,8].Value =  userList["ДатаНазначения"]?.InnerText;

                                lineCount++;
                                countLine++; 
                            }
                        }
                        countLine--;
                        worksheet.Range[countLine-lineCount+1,1,countLine,1].Merge();
                        worksheet.Range[countLine-lineCount+1,2,countLine,2].Merge(); 
                        worksheet.Range[countLine-lineCount+1,3,countLine,3].Merge(); 
                        worksheet.Range[countLine-lineCount+1,4,countLine,4].Merge();  
                        worksheet.AllocatedRange.AutoFitColumns();
                        workbook.SaveToFile(args[1], ExcelVersion.Version2016); //второй параметр - место сохранения xlsx файла
                        lineCount = 0;
                        
                    }
                
                }
            }
        }

    }
}

    
