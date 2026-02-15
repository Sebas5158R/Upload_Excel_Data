using System.Windows;
using Microsoft.Win32;
using System.Data;
using System.Text.RegularExpressions;

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

        private bool filtroActivo = false;
        int sueldosInvalidos = 0;

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

                    // Excel vacío
                    if (rango == null) { 
                        btnFiltrar.Visibility = Visibility.Collapsed;
                        return;
                    }

                    int columnas = rango.ColumnCount();
                    int colSueldo = -1;

                    // Cargar nombres de columnas (tipo_documento, nombres, sueldo)
                    for (int i = 1; i <= columnas; i++)
                    {
                        string nombreCol = rango.Cell(1, i).GetValue<string>();
                        tabla.Columns.Add(rango.Cell(1, i).GetValue<string>());

                        if (nombreCol.Trim().ToLower() == "sueldo")
                        {
                            colSueldo = i - 1;
                        }
                    }

                    // Expresión regular para solo dígitos
                    Regex regexNumeros = new Regex(@"^\d+$");

                    // Cargar datos
                    for (int fila = 2; fila <= rango.RowCount(); fila++)
                    {
                        string valorSueldo = rango.Cell(fila, colSueldo + 1).GetValue<string>();
                        if (!regexNumeros.IsMatch(valorSueldo))
                        {
                            sueldosInvalidos++;
                            continue; // Para no agregar fila con sueldo inválido
                        }
                        DataRow nuevaFila = tabla.NewRow();
                        for (int col = 1; col <= columnas; col++)
                        {
                            nuevaFila[col - 1] = rango.Cell(fila, col).GetValue<string>();
                        }
                        tabla.Rows.Add(nuevaFila);
                    }
                }

                dgDatos.ItemsSource = tabla.DefaultView;
                btnFiltrar.Visibility = Visibility.Visible;

                if (sueldosInvalidos > 0)
                {
                    MessageBox.Show($"Cantidad de sueldos mal digitados: {sueldosInvalidos}", "Advertencia", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al leer Excel: " + ex.Message);
            }
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
    }
}