# HTML-Table-To-Excel
Convert HTML Table to .xlsx file using EPPlus. EPPlus library only supports .xlsx(Office 2007 and above).  
The TableToExcel class uses EPPlus and Windows.Form. It can convert HTML to bytes array.  

Usage:  
1. Stream  
TableToExcel temp = new TableToExcel();  
Response.BinaryWrite(temp.process(html));  
2. File  
TableToExcel temp = new TableToExcel();  
using (StreamWriter file = new StreamWriter("C:\\temp.xlsx"))  
{  
    file.Write(temp.process(html));  
}  

P.S. Please solved the problems of WebBrowser by yourself.  
