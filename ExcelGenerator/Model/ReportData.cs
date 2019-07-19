using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelGenerator.Model
{

   public class ReportData
   {
      public int aseguradora { get; set; }
      public int asegurado { get; set; }
      public int territorio { get; set; }
      public int zona { get; set; }
      public string oficina { get; set; }
      public int estado { get; set; }
      public int municipio { get; set; }
      public string da { get; set; }
      public int id_agente { get; set; }
      public string nombre_agente { get; set; }
      public string poliza { get; set; }
      public string apellido_paterno { get; set; }
      public string apellido_materno { get; set; }
      public string nombre { get; set; }
      public string inicio_vigencia { get; set; }
      public string fin_vigencia { get; set; }
      public string ultima_visita { get; set; }
      public double precio { get; set; }
      public double pago { get; set; }
      public double ahorro { get; set; }
      public string plan_convenio { get; set; }
   }
}
