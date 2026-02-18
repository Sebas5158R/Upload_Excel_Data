using Microsoft.Win32;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Windows;
using tabla_excel.models;

namespace tabla_excel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        List<Empleado> empleados = new List<Empleado>();
        List<string> errores = new List<string>();

        private bool filtroActivo = false;
        int sueldosInvalidos = 0;

        private void BtnSeleccionar_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog
                {
                    Filter = "Archivos Excel (*.xlsx)|*.xlsx"
                };

                if (dialog.ShowDialog() == true)
                {
                    var lista = LeerExcel(dialog.FileName);
                    dgDatos.ItemsSource = lista;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR GENERAL:\n" + ex.ToString());
            }
        }


        private List<Empleado> LeerExcel(string ruta)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(ruta)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int filas = worksheet.Dimension.Rows;

                for (int fila = 2; fila <= filas; fila++)
                {
                    string tipo = worksheet.Cells[fila, 1].Text.Trim().ToUpper();

                    if (tipo != "CC" && tipo != "CE")
                    {
                        MessageBox.Show($"Tipo de documento inválido en fila {fila}. Solo se permite CC o CE.");
                        continue;
                    }


                    //carga todos
                    Empleado emp = new Empleado
                    {
                        TipoDoc = worksheet.Cells[fila, 1].Text,
                        NroDoc = worksheet.Cells[fila, 2].Text,
                        Sueldo = decimal.Parse(worksheet.Cells[fila, 3].Text)
                    };

                    empleados.Add(emp);

                    //carga solo los tipos de documentos válidos
                    empleados.Add(new Empleado
                    {
                        TipoDoc = tipo,
                        NroDoc = worksheet.Cells[fila, 2].Text,
                        Sueldo = decimal.Parse(worksheet.Cells[fila, 3].Text)
                    });

                    decimal sueldo = 0;

                    bool esNumeroValido = decimal.TryParse(
                       worksheet.Cells[fila, 3].Text,
                       System.Globalization.NumberStyles.Any,
                        new System.Globalization.CultureInfo("es-CO"),
                        out sueldo
                    );

                    if (!esNumeroValido)
                    {
                        errores.Add($"Fila {fila}: Sueldo inválido");
                        continue;
                    }

                    if (sueldo < 0)
                    {
                        errores.Add($"Fila {fila}: Sueldo negativo");
                        continue;
                    }
                }

                if (errores.Any())
                {
                    MessageBox.Show(string.Join("\n", errores));
                }

                decimal totalNomina = empleados.Sum(e => e.Sueldo);

                MessageBox.Show(totalNomina.ToString("C0", new System.Globalization.CultureInfo("es-CO")));

                var agrupado = empleados
                    .GroupBy(e => e.TipoDoc)
                    .Select(g => new
                    {
                        TipoDoc = g.Key,
                        Cantidad = g.Count(),
                        Total = g.Sum(x => x.Sueldo)
                    })
                    .ToList();

                dgDatos.ItemsSource = agrupado;

                var estadisticas = new
                {
                    Promedio = empleados.Average(e => e.Sueldo),
                    Maximo = empleados.Max(e => e.Sueldo),
                    Minimo = empleados.Min(e => e.Sueldo),
                    Total = empleados.Sum(e => e.Sueldo),
                    Cantidad = empleados.Count()
                };

                dgEstadisticas.ItemsSource = new List<object> { estadisticas };


                var duplicados = empleados
                    .GroupBy(e => e.NroDoc)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();


                if (duplicados.Any())
                {
                    MessageBox.Show(
                        "Documentos duplicados:\n" +
                        string.Join("\n", duplicados)
                    );
                }
                else
                {
                    MessageBox.Show("No hay duplicados");
                }

                dgTiposDocs.ItemsSource = new List<object> { estadisticas };
            }

            return empleados;
        }
        private void Filtrar_Sueldos(object sender, RoutedEventArgs e)
        {
            if (dgDatos.ItemsSource is DataView vista)
            {
                if (!filtroActivo)
                {
                    vista.RowFilter = "sueldo > 1000000";
                    btnFiltrar.Content = "Quitar Filtro";
                    filtroActivo = true;
                }
                else
                {
                    vista.RowFilter = string.Empty;
                    btnFiltrar.Content = "Filtrar por sueldo";
                    filtroActivo = false;
                }
            }
        }

        private void ExportarCSV(List<Empleado> empleados)
        {
            SaveFileDialog dlg = new SaveFileDialog();

            dlg.Filter = "CSV (*.csv)|*.csv";
            dlg.FileName = "Nomina.csv";

            if(dlg.ShowDialog() == true) 
            {
                var lineas = new List<string>();

                lineas.Add("tipo_doc, nro_doc, sueldo");

                foreach (var emp in empleados)
                {
                    lineas.Add($"{emp.TipoDoc}, {emp.NroDoc}, {emp.Sueldo}");
                }

                File.WriteAllLines(dlg.FileName, lineas);

                MessageBox.Show("Archivo exportado exitosamente.");
            }



        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExportarCSV(empleados);
        }
    }
}