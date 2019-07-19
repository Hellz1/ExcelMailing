using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using ExcelGenerator.DTO;
using ExcelGenerator.Model;

using ExcelPasswordAdder;

using System.Globalization;

namespace ExcelGenerator
{
   class Program
   {
      static void Main(string[] args)
      {
         using (ExcelPackage excel = new ExcelPackage())
         {

            //consumimos la api de reportes

            /****Bloque de obtención de datos***/
            string URL = "http://localhost:62706/api/Master/reporte";

            List<ReportData> reportData = getContentFromURL(URL);

            //Filtramos la lista:  "ID AGENTE" , "POLIZA", "ASEGURADO", "ULTIMA VISITA A DENTALIA", "FIN DE VIGENCIA", "PRECIO", "PAGO","AHORRO"
            List<ReportDataDTO> filteredData = reportData
               .Select(
                  s => new ReportDataDTO
                  {
                     id_agente = s.id_agente,
                     poliza = s.poliza,
                     asegurado = s.nombre + " " + s.apellido_paterno + " " + s.apellido_materno,
                     ultima_visita = s.ultima_visita,
                     fin_vigencia = s.fin_vigencia,
                     precio = s.precio,
                     pago = s.pago,
                     ahorro = s.ahorro
                  }).ToList();

            excel.Workbook.Worksheets.Add("Reporte");

            var currentExcelWorksheet = excel.Workbook.Worksheets["Reporte"];


            /********Agregamos una imagen******/
            string path = @"C:\Users\chava\source\repos\ExcelGenerator\ExcelGenerator\img\1200x100.png";
            Bitmap logo = new Bitmap(Image.FromFile(path));

            if (logo.HorizontalResolution == 0 || logo.VerticalResolution == 0)
               logo.SetResolution(96, 96);

            var picture = currentExcelWorksheet.Drawings.AddPicture("Logo",logo);
            
            //Agregamos los encabezados
            List<string[]> headerRow = new List<string[]>()
            {
               new string[] { "ID AGENTE" , "POLIZA", "ASEGURADO", "ULTIMA VISITA A DENTALIA", "FIN DE VIGENCIA", "PRECIO", "PAGO","AHORRO"}
            };

            //Determinamos el rango de la cabecera
            string headerRange = "A6:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "6";

            //Por si queremos estilos en los encabezados
            Color colFromHex = Color.FromArgb(255,217,102);

            currentExcelWorksheet.Cells[headerRange].Style.Font.SetFromFont(new Font("Century Gothic",10));
            currentExcelWorksheet.Cells[headerRange].Style.Font.Bold = true;
            currentExcelWorksheet.Cells[headerRange].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            currentExcelWorksheet.Cells[headerRange].Style.Fill.BackgroundColor.SetColor(colFromHex);
            
            //Cargamos la información al excel
            currentExcelWorksheet.Cells[headerRange].LoadFromArrays(headerRow);
            
            /**
             * Tambien podemos agregar datos a valores especificos de la celda:
             * currentExcelWorksheet.Cells["A1"].Value = "Hello World!";
             * */

            //Cargamos información al excel (a partir de A7)
            //"ID AGENTE" , "POLIZA", "ASEGURADO", "ULTIMA VISITA A DENTALIA", "FIN DE VIGENCIA", "PRECIO", "PAGO","AHORRO"

            currentExcelWorksheet.Cells[7,1].LoadFromCollection(filteredData);

            currentExcelWorksheet.Cells[currentExcelWorksheet.Dimension.Address].AutoFitColumns();

            //Hacemos el autofit
            for (int i = 1; i <= currentExcelWorksheet.Dimension.End.Column; i++)
            {
               currentExcelWorksheet.Column(i).AutoFit();
               currentExcelWorksheet.Column(i).BestFit = true;
            }

            picture.SetPosition(0, 0,0,currentExcelWorksheet.Dimension.End.Column);

            currentExcelWorksheet.PrinterSettings.FitToPage = true;

            //FileInfo excelFile = new FileInfo(@"C:\Users\chava\source\repos\ExcelGenerator\ExcelGenerator\excel.xlsx");

            //Configuracion
            //excel.Encryption.Password = "1";

            byte[] stream = excel.GetAsByteArray();

            ExcelPasswordAdder.Program.initExcelProcess(stream);
         }

      }


     static private List<ReportData> getContentFromURL(string URL)
      {

         var client = new WebClient();
         var content = client.DownloadString(URL);

         List<ReportData> reportData = JsonConvert.DeserializeObject<List<ReportData>>(content);

         return reportData;
      }
    
   }
}
