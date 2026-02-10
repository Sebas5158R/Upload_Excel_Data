using System.Windows;
using Microsoft.Win32;
using System.Data;

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

        private void BtnSeleccionar_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                CargarExcel(dialog.FileName);
            }
        }

        private void CargarExcel(string ruta)
        {
            DataTable tabla = new();

            try
            {
                using (var wb = new ClosedXML.Excel.XLWorkbook(ruta))
                {
                    var hoja = wb.Worksheet(1);
                    var rango = hoja.RangeUsed();

                    if (rango == null) return; // Excel vacío

                    int columnas = rango.ColumnCount();

                    // Crear columnas genéricas si no hay encabezado
                    for (int i = 1; i <= columnas; i++)
                    {
                        tabla.Columns.Add($"Columna {i}");
                    }

                    // Leer todas las filas
                    foreach (var fila in rango.Rows())
                    {
                        DataRow dr = tabla.NewRow();
                        for (int i = 0; i < columnas; i++)
                        {
                            dr[i] = fila.Cell(i + 1).GetValue<string>();
                        }
                        tabla.Rows.Add(dr);
                    }
                }

                dgDatos.ItemsSource = tabla.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al leer Excel: " + ex.Message);
            }
        }

    }
}