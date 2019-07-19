using System;

namespace ExcelGenerator.DTO
{
   public class ReportDataDTO
   {
      public int id_agente { get; set; }
      public string poliza { get; set; }
      public string asegurado { get; set; }
      public string ultima_visita { get; set; }
      public string fin_vigencia { get; set; }
      public double precio { get; set; }
      public double pago { get; set; }
      public double ahorro { get; set; }
   }

}
