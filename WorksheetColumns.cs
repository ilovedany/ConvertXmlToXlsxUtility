using Spire.Xls;
using System.Xml;

namespace ConvertXmlToXlsxUtility
{
    public class WorksheetColumns
    {
        //метод для отображения заголовка
        public void WorkingWithXml(Workbook workbook,int countLine, string[] columns){
            int position = 0;
            foreach(var c in columns){
                position++;
                workbook.Worksheets[0].Range[countLine,position].Value = c;
            }
        }
        //метод для заполнения данных
        public void WorkingWithXml(XmlNode xmlNode,Workbook workbook,int countLine, string[] columns, int position){
            foreach(var c in columns){
                position++;
                workbook.Worksheets[0].Range[countLine,position].Value = xmlNode[c]?.InnerText;
            }   
        }
        //метод для объединения строк
        public void WorkingWithXml(Workbook workbook,int a,int c, int d){
            for(int i=1; i<d+1; i++){
                workbook.Worksheets[0].Range[a,i,c,i].Merge();
            }   
        }
        



    }
}