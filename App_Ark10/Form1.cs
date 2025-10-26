using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

namespace App_Ark10
{
    public partial class Form1 : Form
    {
        private Timer timer;
        private List<string> impresorasAnteriores = new List<string>();

        private string connectionString = string.Empty;
        private DataTable productosDataTable = new DataTable();
        private DataTable carritoDataTable = new DataTable();
        private System.Windows.Forms.Timer dataRefreshTimer;

        public Form1()
        {
            InitializeComponent();

            CargarImpresoras();
            IniciarTemporizador();

            this.Width = 1234;  
            this.Height = 640;
            carritoDataTable.Columns.Add("ID", typeof(int));
            carritoDataTable.Columns.Add("Producto", typeof(string));
            carritoDataTable.Columns.Add("Color", typeof(string));
            carritoDataTable.Columns.Add("Categoría", typeof(string));
            carritoDataTable.Columns.Add("Precio", typeof(decimal));
            carritoDataTable.Columns.Add("Precio Estudiante", typeof(decimal));
            carritoDataTable.Columns.Add("Precio Proveedor", typeof(decimal));
            carritoDataTable.Columns.Add("Stock", typeof(int));
            carritoDataTable.Columns.Add("Cantidad", typeof(int));
            dataGridViewCarrito.DataSource = carritoDataTable;
            AgregarBotonesCarrito();
            dataRefreshTimer = new System.Windows.Forms.Timer();
            dataRefreshTimer.Interval = 2000; 
            dataRefreshTimer.Tick += async (sender, e) => await LoadProductosAsync();
        }

        private async void buttonconectar_Click(object sender, EventArgs e)
        {
            string host = textBoxIP.Text;
            string username = textBoxusername.Text;
            string password = textBoxpassword.Text;
            if (string.IsNullOrEmpty(host) || string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
            {
                MessageBox.Show("Por favor ingresa todas las credenciales", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            connectionString = $"Server={host};Port=3306;Database=automotriz;Uid={username};Pwd={password}";
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    MessageBox.Show("Conexión exitosa", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    panelLogin.Visible = false;
                    panelReportes.Visible = true;
                    paneldatos.Visible = true;
                    panelimpresoras.Visible = true;
                    panelcomprar.Visible =  true;
                }
                await LoadProductosAsync();

                if (!dataRefreshTimer.Enabled)
                {
                    dataRefreshTimer.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al conectar: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                buttonconectar.PerformClick();
            }
        }

        private void CargarImpresoras()
        {
            List<string> impresorasActuales = PrinterSettings.InstalledPrinters.Cast<string>().ToList();

            if (!Enumerable.SequenceEqual(impresorasAnteriores, impresorasActuales))
            {
                string impresoraSeleccionada = comboBoxImpresora.SelectedItem as string;

                comboBoxImpresora.Items.Clear();
                comboBoxImpresora.Items.AddRange(impresorasActuales.ToArray());

                if (!string.IsNullOrEmpty(impresoraSeleccionada) && impresorasActuales.Contains(impresoraSeleccionada))
                {
                    comboBoxImpresora.SelectedItem = impresoraSeleccionada;
                }
                else if (comboBoxImpresora.Items.Count > 0)
                {
                    comboBoxImpresora.SelectedIndex = 0;
                }

                impresorasAnteriores = impresorasActuales;
            }
        }

        private void IniciarTemporizador()
        {
            timer = new Timer
            {
                Interval = 3000
            };
            timer.Tick += (s, e) => CargarImpresoras();
            timer.Start();
        }

        private void ConfigurarDataGridView()
        {
            dataGridView1datos.RowHeadersVisible = false;
            dataGridView1datos.ReadOnly = true;
            dataGridView1datos.AllowUserToAddRows = false;
            dataGridView1datos.AllowUserToDeleteRows = false;
            dataGridView1datos.AllowUserToOrderColumns = false;
            dataGridViewCarrito.RowHeadersVisible = false;
            dataGridViewCarrito.ReadOnly = true;
            dataGridViewCarrito.AllowUserToAddRows = false;
            dataGridViewCarrito.AllowUserToDeleteRows = false;
            dataGridViewCarrito.AllowUserToOrderColumns = false;
        }

        private async void Form1_Load_1(object sender, EventArgs e)
        {
            ConfigurarDataGridView(); 
            textBoxIP.KeyPress += textBox_KeyPress;
            textBoxusername.KeyPress += textBox_KeyPress;
            textBoxpassword.KeyPress += textBox_KeyPress;
            textBoxBuscar.TextChanged += textBoxBuscar_TextChanged;
            AgregarBotonAgregar();
        }

        private void AgregarBotonAgregar()
        {
            if (dataGridView1datos.Columns["AgregarCarrito"] == null)
            {
                DataGridViewButtonColumn btnAgregar = new DataGridViewButtonColumn();
                btnAgregar.Name = "AgregarCarrito";
                btnAgregar.HeaderText = "Agregar";
                btnAgregar.Text = "Añadir";
                btnAgregar.UseColumnTextForButtonValue = true;
                dataGridView1datos.Columns.Add(btnAgregar);
            }
        }

        private void AgregarBotonesCarrito()
        {
            if (dataGridViewCarrito.Columns["ReducirCantidad"] == null)
            {
                DataGridViewButtonColumn btnReducir = new DataGridViewButtonColumn();
                btnReducir.Name = "ReducirCantidad";
                btnReducir.HeaderText = "Reducir";
                btnReducir.Text = "-";
                btnReducir.UseColumnTextForButtonValue = true;
                dataGridViewCarrito.Columns.Add(btnReducir);
            }
            if (dataGridViewCarrito.Columns["Eliminar"] == null)
            {
                DataGridViewButtonColumn btnEliminar = new DataGridViewButtonColumn();
                btnEliminar.Name = "Eliminar";
                btnEliminar.HeaderText = "Eliminar";
                btnEliminar.Text = "X";
                btnEliminar.UseColumnTextForButtonValue = true;
                dataGridViewCarrito.Columns.Add(btnEliminar);
            } 
        }

        private async Task LoadProductosAsync()
        {
            string query = @"
        SELECT p.id, p.nombre AS producto_nombre, c.color, cat.nombre AS categoria_nombre, 
               p.precio, p.precio_estudiante, p.precio_proveedor, p.stock, p.estado
        FROM productos p
        LEFT JOIN colors c ON p.color_id = c.id
        LEFT JOIN categorias cat ON p.categoria_id = cat.id
        WHERE p.activo = 1";
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    using (MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection))
                    {
                        DataTable updatedDataTable = new DataTable();
                        await Task.Run(() => adapter.Fill(updatedDataTable));
                        if (!DataTableEquals(productosDataTable, updatedDataTable))
                        {
                            productosDataTable = updatedDataTable;
                            dataGridView1datos.Invoke((MethodInvoker)(() => dataGridView1datos.DataSource = productosDataTable));
                            ActualizarCarrito();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lblEstadoConexion.Invoke((MethodInvoker)(() => lblEstadoConexion.Text = $"Error al cargar: {ex.Message}"));
            }
        }

        private void ActualizarCarrito()
        {
            List<DataRow> filasParaEliminar = new List<DataRow>();
            foreach (DataRow row in carritoDataTable.AsEnumerable().Reverse())
            {
                int idProducto = Convert.ToInt32(row["ID"]);
                DataRow[] productoEncontrado = productosDataTable.Select($"id = {idProducto}");

                if (productoEncontrado.Length == 0)
                {
                    filasParaEliminar.Add(row);
                    continue;
                }
                DataRow producto = productoEncontrado[0];
                int stockDisponible = producto["stock"] != DBNull.Value ? Convert.ToInt32(producto["stock"]) : 0;
                row["Stock"] = stockDisponible;
                decimal precio = producto["precio"] != DBNull.Value ? Convert.ToDecimal(producto["precio"]) : 0m;
                decimal precioEstudiante = producto["precio_estudiante"] != DBNull.Value ? Convert.ToDecimal(producto["precio_estudiante"]) : 0m;
                decimal precioProveedor = producto["precio_proveedor"] != DBNull.Value ? Convert.ToDecimal(producto["precio_proveedor"]) : 0m;
                row["Precio"] = precio;
                row["Precio Estudiante"] = precioEstudiante;
                row["Precio Proveedor"] = precioProveedor;
                int cantidadCarrito = Convert.ToInt32(row["Cantidad"]);
                if (cantidadCarrito > stockDisponible)
                {
                    row["Cantidad"] = stockDisponible;
                }
                if (stockDisponible == 0)
                {
                    filasParaEliminar.Add(row);
                }
            }
            foreach (DataRow fila in filasParaEliminar)
            {
                carritoDataTable.Rows.Remove(fila);
            }
            dataGridViewCarrito.Refresh();
            ActualizarTotales();
        }

        private decimal totalPrecio = 0m;
        private decimal totalEstudiante = 0m;
        private decimal totalProveedor = 0m;

        private void ActualizarTotales()
        {
            totalPrecio = 0m;
            totalEstudiante = 0m;
            totalProveedor = 0m;
            try
            {
                foreach (DataRow row in carritoDataTable.Rows)
                {
                    decimal precio = row["Precio"] != DBNull.Value ? Convert.ToDecimal(row["Precio"]) : 0m;
                    decimal precioEstudiante = row["Precio Estudiante"] != DBNull.Value && Convert.ToDecimal(row["Precio Estudiante"]) > 0
                        ? Convert.ToDecimal(row["Precio Estudiante"])
                        : precio;
                    decimal precioProveedor = row["Precio Proveedor"] != DBNull.Value && Convert.ToDecimal(row["Precio Proveedor"]) > 0
                        ? Convert.ToDecimal(row["Precio Proveedor"])
                        : precio;
                    int cantidad = row["Cantidad"] != DBNull.Value ? Convert.ToInt32(row["Cantidad"]) : 0;
                    totalPrecio += precio * cantidad;
                    totalEstudiante += precioEstudiante * cantidad;
                    totalProveedor += precioProveedor * cantidad;
                }
                lblTotalPrecio.Text = $"Total: {totalPrecio:N2} Bs.";
                lblTotalEstudiante.Text = $"Total Estudiante: {totalEstudiante:N2} Bs.";
                lblTotalProveedor.Text = $"Total Proveedor: {totalProveedor:N2} Bs.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al calcular totales: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool DataTableEquals(DataTable dt1, DataTable dt2)
        {
            if (dt1.Rows.Count != dt2.Rows.Count) return false;
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                for (int j = 0; j < dt1.Columns.Count; j++)
                {
                    if (!Equals(dt1.Rows[i][j], dt2.Rows[i][j]))
                        return false;
                }
            }
            return true;
        }

        private async void FiltrarProductos(string filtro)
        {
            if (productosDataTable == null || productosDataTable.Rows.Count == 0)
            {
                return;
            }
            await Task.Run(() =>
            {
                DataView vistaFiltrada = new DataView(productosDataTable);
                string filtroSeguro = filtro.Replace("'", "''");

                if (string.IsNullOrEmpty(filtroSeguro))
                {
                    vistaFiltrada.RowFilter = "";
                }
                else if (int.TryParse(filtroSeguro, out _))
                {
                    vistaFiltrada.RowFilter = $"Convert(id, 'System.String') LIKE '%{filtroSeguro}%' OR producto_nombre LIKE '%{filtroSeguro}%'";
                }
                else
                {
                    vistaFiltrada.RowFilter = $"producto_nombre LIKE '%{filtroSeguro}%'";
                }

                dataGridView1datos.Invoke((MethodInvoker)(() => dataGridView1datos.DataSource = vistaFiltrada));
            });
        }

        private void textBoxBuscar_TextChanged(object sender, EventArgs e)
        {
            FiltrarProductos(textBoxBuscar.Text.Trim());
        }

        private void dataGridView1datos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1datos.Columns["AgregarCarrito"].Index && e.RowIndex >= 0)
            {
                var dataBoundItem = dataGridView1datos.Rows[e.RowIndex].DataBoundItem;
                if (dataBoundItem is DataRowView filaSeleccionada)
                {
                    try
                    {
                        int id = filaSeleccionada["id"] != DBNull.Value ? Convert.ToInt32(filaSeleccionada["id"]) : 0;
                        string producto = filaSeleccionada["producto_nombre"]?.ToString() ?? "Desconocido";
                        string color = filaSeleccionada["color"]?.ToString() ?? "N/A";
                        string categoria = filaSeleccionada["categoria_nombre"]?.ToString() ?? "Sin categoría";
                        decimal precio = filaSeleccionada["precio"] as decimal? ?? 0m;
                        decimal? precioEstudiante = filaSeleccionada["precio_estudiante"] as decimal?;
                        decimal? precioProveedor = filaSeleccionada["precio_proveedor"] as decimal?;
                        decimal precioFinalEstudiante = (precioEstudiante.HasValue && precioEstudiante.Value > 0) ? precioEstudiante.Value : 0m;
                        decimal precioFinalProveedor = (precioProveedor.HasValue && precioProveedor.Value > 0) ? precioProveedor.Value : 0m;
                        int stock = filaSeleccionada["stock"] != DBNull.Value ? Convert.ToInt32(filaSeleccionada["stock"]) : 0;
                        if (stock == 0)
                        {
                            MessageBox.Show("Este producto está fuera de stock.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        bool encontrado = false;
                        foreach (DataRow row in carritoDataTable.Rows)
                        {
                            if ((int)row["ID"] == id)
                            {
                                int cantidadActual = (int)row["Cantidad"];
                                if (cantidadActual < stock)
                                {
                                    row["Cantidad"] = cantidadActual + 1;
                                    MessageBox.Show($"Cantidad de '{producto}' aumentada a {cantidadActual + 1}.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show("No puedes agregar más de este producto, ya que no hay más stock.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }
                                encontrado = true;
                                break;
                            }
                        }
                        if (!encontrado)
                        {
                            carritoDataTable.Rows.Add(id, producto, color, categoria, precio, precioEstudiante, precioProveedor, stock, 1);
                        }
                        dataGridViewCarrito.Refresh();
                        ActualizarTotales();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error al procesar la fila: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Error: No se pudo obtener la información del producto.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void dataGridViewCarrito_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataRow row = ((DataRowView)dataGridViewCarrito.Rows[e.RowIndex].DataBoundItem).Row;

                if (e.ColumnIndex == dataGridViewCarrito.Columns["ReducirCantidad"].Index)
                {
                    int cantidad = Convert.ToInt32(row["Cantidad"]);
                    if (cantidad > 1)
                    {
                        row["Cantidad"] = cantidad - 1;
                    }
                    else
                    {
                        MessageBox.Show("La cantidad mínima es 1. Si deseas eliminarlo, usa el botón de eliminar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else if (e.ColumnIndex == dataGridViewCarrito.Columns["Eliminar"].Index)
                {
                    carritoDataTable.Rows.Remove(row);
                }
                dataGridViewCarrito.Refresh();
                ActualizarTotales();
            }
        }

        private async void buttonvender_Click(object sender, EventArgs e)
        {
            if (carritoDataTable.Rows.Count == 0)
            {
                MessageBox.Show("El carrito está vacío.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string nombreUsuario = textBoxNombre.Text.Trim();
            string telefono = textBoxTelefono.Text.Trim();
            string ci = textBoxCI.Text.Trim();
            if (string.IsNullOrWhiteSpace(nombreUsuario))
            {
                MessageBox.Show("Por favor, ingrese un nombre.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(telefono) || !EsNumeroValido(telefono))
            {
                MessageBox.Show("Por favor, ingrese un número de teléfono válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(ci) || !EsNumeroValido(ci))
            {
                MessageBox.Show("Por favor, ingrese un número de CI válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                await connection.OpenAsync();
                bool stockActualizado = await ActualizarStock(connection);
                if (!stockActualizado)
                {
                    return;
                }
                decimal totalVenta = radioButtonPrecio.Checked ? totalPrecio :
                                     radioButtonEstudiante.Checked ? totalEstudiante :
                                     radioButtonProveedor.Checked ? totalProveedor : 0m;
                string tipoTotal = radioButtonPrecio.Checked ? "Total Normal" :
                   radioButtonEstudiante.Checked ? "Total Estudiante" :
                   radioButtonProveedor.Checked ? "Total Proveedor" : "Desconocido";
                if (totalVenta == 0)
                {
                    MessageBox.Show("Seleccione un tipo de precio (Normal, Estudiante o Proveedor).", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                int cantidadTotal = carritoDataTable.Rows.Count;
                List<Dictionary<string, object>> detalleCompra = carritoDataTable.AsEnumerable()
                    .Select(row => new Dictionary<string, object>
                    {
                { "ID", row["ID"] },
                { "Producto", row["Producto"] },
                { "Color", row["Color"] },
                { "Categoría", row["Categoría"] },
                { "Cantidad", row["Cantidad"] },
                { "Precio", row["Precio"] },
                { "Precio Estudiante", row["Precio Estudiante"] },
                { "Precio Proveedor", row["Precio Proveedor"] }
                    }).ToList();
                string detalleCompraJson = Newtonsoft.Json.JsonConvert.SerializeObject(detalleCompra);
                string query = @"
        INSERT INTO venta_locals (nombre_usuario, telefono, CI, cantidad, total, detalle_compra) 
        VALUES (@nombre, @telefono, @ci, @cantidad, @total, @detalle);
        SELECT LAST_INSERT_ID();";
                try
                {
                    using (MySqlCommand cmd = new MySqlCommand(query, connection))
                    {
                        cmd.Parameters.AddWithValue("@nombre", nombreUsuario);
                        cmd.Parameters.AddWithValue("@telefono", telefono);
                        cmd.Parameters.AddWithValue("@ci", ci);
                        cmd.Parameters.AddWithValue("@cantidad", cantidadTotal);
                        cmd.Parameters.AddWithValue("@total", totalVenta);
                        cmd.Parameters.AddWithValue("@detalle", detalleCompraJson);
                        object result = await cmd.ExecuteScalarAsync();
                        int idVenta = Convert.ToInt32(result);
                        Factura factura = new Factura
                        {
                            IDVenta = idVenta,
                            Cliente = nombreUsuario,
                            Telefono = telefono,
                            CI = ci,
                            Fecha = DateTime.Now.ToString("yyyy-MM-dd"),
                            TipoTotal = tipoTotal,
                            Total = (double)totalVenta
                        };
                        foreach (var item in detalleCompra)
                        {
                            factura.Detalles.Add(new DetalleFactura
                            {
                                ID = Convert.ToInt32(item["ID"]),
                                Producto = item["Producto"].ToString(),
                                Categoria = item["Categoría"].ToString(),
                                Color = item["Color"].ToString(),
                                Cantidad = Convert.ToInt32(item["Cantidad"]),
                                Precio = Convert.ToDouble(item["Precio"])
                            });
                        }
                        ImprimirFactura(factura);
                        MessageBox.Show($"Venta realizada con éxito. ID de venta: {idVenta}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Invoke(new Action(() =>
                        {
                            textBoxNombre.Clear();
                            textBoxTelefono.Clear();
                            textBoxCI.Clear();
                            radioButtonPrecio.Checked = true;
                            carritoDataTable.Clear();
                            dataGridViewCarrito.Refresh();
                            ActualizarTotales();
                        }));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al registrar la venta: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool EsNumeroValido(string texto)
        {
            return texto.All(char.IsDigit);
        }

        private async Task<bool> ActualizarStock(MySqlConnection connection)
        {
            using (MySqlTransaction transaction = await connection.BeginTransactionAsync())
            {
                try
                {
                    foreach (DataRow row in carritoDataTable.Rows)
                    {
                        int idProducto = Convert.ToInt32(row["ID"]);
                        int cantidadCarrito = Convert.ToInt32(row["Cantidad"]);
                        string stockQuery = "SELECT stock FROM productos WHERE id = @idProducto FOR UPDATE";
                        using (MySqlCommand cmdStock = new MySqlCommand(stockQuery, connection, transaction))
                        {
                            cmdStock.Parameters.AddWithValue("@idProducto", idProducto);
                            object result = await cmdStock.ExecuteScalarAsync();
                            int stockActual = result != DBNull.Value ? Convert.ToInt32(result) : 0;

                            if (cantidadCarrito > stockActual)
                            {
                                MessageBox.Show($"El producto '{row["Producto"]}' solo tiene {stockActual} unidades disponibles. Ajusta la cantidad antes de continuar.", "Stock insuficiente", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                await transaction.RollbackAsync();
                                return false;
                            }
                        }
                        string updateStockQuery = "UPDATE productos SET stock = stock - @cantidad WHERE id = @idProducto";
                        using (MySqlCommand cmdUpdate = new MySqlCommand(updateStockQuery, connection, transaction))
                        {
                            cmdUpdate.Parameters.AddWithValue("@cantidad", cantidadCarrito);
                            cmdUpdate.Parameters.AddWithValue("@idProducto", idProducto);
                            await cmdUpdate.ExecuteNonQueryAsync();
                        }
                    }
                    await transaction.CommitAsync();
                    return true;
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    MessageBox.Show($"Error al actualizar el stock: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
        }

        private void ImprimirFactura(Factura factura)
        {
            if (comboBoxImpresora.SelectedItem == null)
            {
                MessageBox.Show("Seleccione una impresora antes de imprimir.");
                return;
            }
            string textoFactura = factura.FormatearFactura();
            PrintDocument documento = new PrintDocument();
            documento.PrinterSettings.PrinterName = comboBoxImpresora.SelectedItem.ToString();
            documento.PrintPage += (senderPrint, ePrint) =>
            {
                ePrint.Graphics.DrawString(textoFactura,
                                            new System.Drawing.Font("Courier New", 10),
                                            Brushes.Black,
                                            50, 50);
            };
            try
            {
                documento.Print();
                MessageBox.Show("Factura impresa con éxito.", "Impresión", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                DialogResult result = MessageBox.Show(
            "Hubo un problema al intentar imprimir. ¿Desea descargar la factura como PDF?",
            "Error de impresión",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Error);

                if (result == DialogResult.Yes)
                {
                    DescargarPDFFactura(factura);
                }
                else
                {
                    MessageBox.Show("La impresión se ha cancelado.", "Impresión cancelada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void DescargarPDFFactura(Factura factura)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF Files|*.pdf",
                FileName = $"Factura_{factura.IDVenta}.pdf"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                    {
                        Document document = new Document(PageSize.A4);
                        PdfWriter writer = PdfWriter.GetInstance(document, fs);
                        document.Open();
                        iTextSharp.text.Font font = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                        string textoFactura = factura.FormatearFactura();
                        Paragraph paragraph = new Paragraph(textoFactura, font);
                        document.Add(paragraph);
                        document.Close();
                    }

                    MessageBox.Show("Factura descargada como PDF con éxito.", "PDF Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al generar el PDF: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void textBoxIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxusername_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxpassword_TextChanged(object sender, EventArgs e)
        {

        }
        //reportes-------------------------------------------------------------
        private async void buttonreporte1_Click(object sender, EventArgs e)
        {
            DataTable reporteVentas = await ObtenerReporteVentasAsync();
            if (reporteVentas.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para generar el reporte.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Reporte_Ventas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (File.Exists(saveFileDialog.FileName))
                    {
                        MessageBox.Show("El archivo ya existe. Por favor elige otro nombre o ubicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                    {
                        var ws = package.Workbook.Worksheets.Add("Reporte de Ventas");
                        ws.PrinterSettings.PaperSize = ePaperSize.Letter;
                        ws.PrinterSettings.Orientation = eOrientation.Portrait;
                        ws.PrinterSettings.FitToPage = true;
                        ws.PrinterSettings.FitToWidth = 1;
                        ws.PrinterSettings.FitToHeight = 0;
                        ws.Cells[1, 1, 1, 8].Merge = true;
                        ws.Cells[1, 1].Value = "Reporte de Ventas - POLIMÉRICOS DIAL BOLIVIA";
                        ws.Cells[1, 1].Style.Font.Size = 14;
                        ws.Cells[1, 1].Style.Font.Bold = true;
                        ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 121));
                        ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                        ws.Row(1).Height = 25;
                        ws.Cells[2, 1].Value = $"Generado el: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                        ws.Cells[2, 1].Style.Font.Italic = true;
                        ws.Cells[2, 1].Style.Font.Size = 10;
                        ws.Row(2).Height = 15;
                        string[] headers = { "Usuario", "Teléfono", "CI", "Cantidad", "Fecha Venta", "Total (Bs.)", "Estado", "Detalles" };
                        for (int col = 1; col <= headers.Length; col++)
                        {
                            ws.Cells[4, col].Value = headers[col - 1];
                            ws.Cells[4, col].Style.Font.Bold = true;
                            ws.Cells[4, col].Style.Font.Size = 10;
                            ws.Cells[4, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[4, col].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            ws.Cells[4, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[4, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[4, col].Style.WrapText = true;
                        }
                        ws.Row(4).Height = 20;
                        decimal totalVentasHoy = 0, totalVentasGlobal = 0;
                        int cantidadVentasHoy = 0, cantidadVentasGlobal = 0;
                        int row = 5;

                        foreach (DataRow dataRow in reporteVentas.Rows)
                        {
                            DateTime fechaVenta = Convert.ToDateTime(dataRow["fecha_venta"]);
                            decimal total = Convert.ToDecimal(dataRow["total"]);
                            bool esHoy = fechaVenta.Date == DateTime.Today;
                            if (esHoy)
                            {
                                totalVentasHoy += total;
                                cantidadVentasHoy++;
                            }
                            totalVentasGlobal += total;
                            cantidadVentasGlobal++;
                            string usuario = dataRow["nombre_usuario"].ToString();
                            string telefono = dataRow["telefono"].ToString();
                            string ci = dataRow["CI"].ToString();
                            string cantidad = Convert.ToInt32(dataRow["cantidad"]).ToString();
                            string fecha = fechaVenta.ToString("dd/MM/yyyy HH:mm");
                            string totalStr = total.ToString("F2");
                            string estado = dataRow["estado"].ToString();
                            string detalleFormateado = FormatearDetalleCompra(dataRow["detalle_compra"].ToString());
                            ws.Cells[row, 1].Value = usuario;
                            ws.Cells[row, 2].Value = telefono;
                            ws.Cells[row, 3].Value = ci;
                            ws.Cells[row, 4].Value = cantidad;
                            ws.Cells[row, 5].Value = fecha;
                            ws.Cells[row, 6].Value = total;
                            ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[row, 7].Value = estado;
                            ws.Cells[row, 8].Value = detalleFormateado;
                            for (int col = 1; col <= 8; col++)
                            {
                                ws.Cells[row, col].Style.WrapText = true;
                                ws.Cells[row, col].Style.Font.Size = 9;
                                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            if (esHoy)
                            {
                                ws.Cells[row, 1, row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                ws.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));
                            }
                            int maxLineas = Math.Max(
                                Math.Max(usuario.Split('\n').Length, telefono.Split('\n').Length),
                                Math.Max(ci.Split('\n').Length, detalleFormateado.Split('\n').Length)
                            );
                            ws.Row(row).Height = Math.Max(15, maxLineas * 12);

                            row++;
                        }
                        int resumenRow = row + 1;
                        ws.Cells[resumenRow, 1, resumenRow, 2].Merge = true;
                        ws.Cells[resumenRow, 1].Value = "Resumen de Ventas";
                        ws.Cells[resumenRow, 1].Style.Font.Bold = true;
                        ws.Cells[resumenRow, 1].Style.Font.Size = 12;
                        ws.Row(resumenRow).Height = 20;
                        ws.Cells[resumenRow + 1, 1].Value = "Ventas Hoy(Bs.):";
                        ws.Cells[resumenRow + 1, 2].Value = totalVentasHoy;
                        ws.Cells[resumenRow + 1, 2].Style.Numberformat.Format = "#,##0.00";
                        ws.Cells[resumenRow + 2, 1].Value = "Cant. Ventas Hoy:";
                        ws.Cells[resumenRow + 2, 2].Value = cantidadVentasHoy;
                        ws.Cells[resumenRow + 3, 1].Value = "Ventas Global(Bs.):";
                        ws.Cells[resumenRow + 3, 2].Value = totalVentasGlobal;
                        ws.Cells[resumenRow + 3, 2].Style.Numberformat.Format = "#,##0.00";
                        ws.Cells[resumenRow + 4, 1].Value = "Cant. Ventas:";
                        ws.Cells[resumenRow + 4, 2].Value = cantidadVentasGlobal;
                        for (int i = 1; i <= 4; i++)
                        {
                            ws.Cells[resumenRow + i, 1, resumenRow + i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[resumenRow + i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[resumenRow + i, 1, resumenRow + i, 2].Style.Font.Size = 10;
                            ws.Cells[resumenRow + i, 1].Style.WrapText = true;
                            ws.Row(resumenRow + i).Height = 15;
                        }
                        ws.Column(1).Width = Math.Min(20, CalcularAnchoMax(reporteVentas, "nombre_usuario", 15));
                        ws.Column(2).Width = Math.Min(15, CalcularAnchoMax(reporteVentas, "telefono", 12));  
                        ws.Column(3).Width = Math.Min(15, CalcularAnchoMax(reporteVentas, "CI", 10));   
                        ws.Column(4).Width = 8; 
                        ws.Column(5).Width = 15;
                        ws.Column(6).Width = 12;
                        ws.Column(7).Width = Math.Min(15, CalcularAnchoMax(reporteVentas, "estado", 10));   
                        ws.Column(8).Width = Math.Min(50, CalcularAnchoMaxDetalles(reporteVentas, "detalle_compra", 40));
                        var chart = ws.Drawings.AddChart("VentasChart", eChartType.ColumnClustered);
                        chart.SetPosition(resumenRow + 6, 0, 1, 0);
                        chart.SetSize(600, 200);
                        chart.Title.Text = "Resumen de Ventas";
                        chart.Legend.Position = eLegendPosition.Right;
                        var serie = chart.Series.Add(
                            ws.Cells[$"B{resumenRow + 1},B{resumenRow + 3}"],
                            ws.Cells[$"A{resumenRow + 1},A{resumenRow + 3}"]
                        );
                        serie.Header = "Ventas (Bs.)";
                        package.Save();
                    }
                    MessageBox.Show("Reporte de ventas generado con éxito.", "Excel Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al generar el reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private async Task<DataTable> ObtenerReporteVentasAsync()
        {
            DataTable dt = new DataTable();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string query = @"
                        SELECT
                            nombre_usuario,
                            telefono,
                            CI,
                            cantidad,
                            total,
                            estado,
                            detalle_compra,
                            fecha_venta
                        FROM venta_locals
                        ORDER BY fecha_venta DESC";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener los datos: {ex.Message}");
            }
            return dt;
        }
        private string FormatearDetalleCompra(string jsonDetalle)
        {
            try
            {
                var detalles = JsonConvert.DeserializeObject<dynamic[]>(jsonDetalle);
                string resultado = "";
                foreach (var item in detalles)
                {
                    resultado += $"Prod: {item.Producto}, Color: {item.Color}, Cant: {item.Cantidad}, Precio: {item.Precio} Bs.\n" +
                                 $"Cat: {item["Categoría"]}\n" +
                                 (item.PrecioProveedor != null ? $"Prov: {item.PrecioProveedor} Bs.\n" : "") +
                                 (item.PrecioEstudiante != null ? $"Est: {item.PrecioEstudiante} Bs.\n" : "") +
                                 "----------------------------------------\n";
                }
                return resultado.TrimEnd('\n');
            }
            catch
            {
                return jsonDetalle;
            }
        }
        private double CalcularAnchoMax(DataTable dt, string columna, double anchoDefault)
        {
            double maxLength = anchoDefault;
            foreach (DataRow row in dt.Rows)
            {
                string valor = row[columna].ToString();
                double length = valor.Length * 0.7;
                if (length > maxLength) maxLength = length;
            }
            return Math.Min(maxLength, 50);
        }
        private double CalcularAnchoMaxDetalles(DataTable dt, string columna, double anchoDefault)
        {
            double maxLength = anchoDefault;
            foreach (DataRow row in dt.Rows)
            {
                string detalle = FormatearDetalleCompra(row[columna].ToString());
                foreach (string linea in detalle.Split('\n'))
                {
                    double length = linea.Length * 0.7;
                    if (length > maxLength) maxLength = length;
                }
            }
            return Math.Min(maxLength, 60);
        }

        private async void buttonreporte2_Click(object sender, EventArgs e)
        {
            DataTable reporteVentasOnline = await ObtenerReporteVentasOnlineAsync();
            if (reporteVentasOnline.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para generar el reporte de ventas en línea.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Reporte_Ventas_Online_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (File.Exists(saveFileDialog.FileName))
                    {
                        MessageBox.Show("El archivo ya existe. Por favor elige otro nombre o ubicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                    {
                        var ws = package.Workbook.Worksheets.Add("Ventas en Línea");
                        ws.PrinterSettings.PaperSize = ePaperSize.Letter;
                        ws.PrinterSettings.Orientation = eOrientation.Landscape;
                        ws.PrinterSettings.FitToPage = true;
                        ws.PrinterSettings.FitToWidth = 1;
                        ws.PrinterSettings.FitToHeight = 0;
                        ws.Cells[1, 1, 1, 12].Merge = true;
                        ws.Cells[1, 1].Value = "Reporte de Ventas en Línea - POLIMÉRICOS DIAL BOLIVIA";
                        ws.Cells[1, 1].Style.Font.Size = 14;
                        ws.Cells[1, 1].Style.Font.Bold = true;
                        ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 121));
                        ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                        ws.Row(1).Height = 25;
                        ws.Cells[2, 1].Value = $"Generado el: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                        ws.Cells[2, 1].Style.Font.Italic = true;
                        ws.Cells[2, 1].Style.Font.Size = 10;
                        ws.Row(2).Height = 15;
                        string[] headers = { "ID Venta", "Nombre", "Email", "Teléfono", "Fecha Venta", "Total (Bs.)", "Método Pago", "Tipo Entrega", "Dirección", "Estado", "Comprobante Subido", "Detalles" };
                        for (int col = 1; col <= headers.Length; col++)
                        {
                            ws.Cells[4, col].Value = headers[col - 1];
                            ws.Cells[4, col].Style.Font.Bold = true;
                            ws.Cells[4, col].Style.Font.Size = 10;
                            ws.Cells[4, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[4, col].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            ws.Cells[4, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[4, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[4, col].Style.WrapText = true;
                        }
                        ws.Row(4).Height = 20;
                        decimal totalVentasOnline = 0;
                        int cantidadVentasOnline = 0;
                        int row = 5;
                        foreach (DataRow dataRow in reporteVentasOnline.Rows)
                        {
                            string idVenta = dataRow["id_venta"].ToString();
                            string nombre = dataRow["name"].ToString();
                            string email = dataRow["email"].ToString();
                            string telefono = dataRow["telefono"].ToString();
                            DateTime fechaVenta = Convert.ToDateTime(dataRow["Fecha_Venta"]);
                            decimal total = Convert.ToDecimal(dataRow["Total"]);
                            string metodoPago = dataRow["Método_Pago"].ToString();
                            string tipoEntrega = dataRow["tipo_entrega"].ToString();
                            string direccion = dataRow["Dirrecion"]?.ToString() ?? "N/A";
                            string estado = dataRow["Estado"].ToString();
                            string comprobanteSubido = dataRow["subido"] != DBNull.Value && Convert.ToBoolean(dataRow["subido"]) ? "Sí" : "No";
                            string detalleVenta = FormatearDetalleVenta(dataRow["detalle_compra"].ToString());
                            totalVentasOnline += total;
                            cantidadVentasOnline++;
                            ws.Cells[row, 1].Value = idVenta;
                            ws.Cells[row, 2].Value = nombre;
                            ws.Cells[row, 3].Value = email;
                            ws.Cells[row, 4].Value = telefono;
                            ws.Cells[row, 5].Value = fechaVenta.ToString("dd/MM/yyyy HH:mm");
                            ws.Cells[row, 6].Value = total;
                            ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[row, 7].Value = metodoPago;
                            ws.Cells[row, 8].Value = tipoEntrega;
                            ws.Cells[row, 9].Value = direccion;
                            ws.Cells[row, 10].Value = estado;
                            ws.Cells[row, 11].Value = comprobanteSubido;
                            ws.Cells[row, 12].Value = detalleVenta;
                            for (int col = 1; col <= 12; col++)
                            {
                                ws.Cells[row, col].Style.WrapText = true;
                                ws.Cells[row, col].Style.Font.Size = 9;
                                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            int maxLineas = Math.Max(
                                Math.Max(nombre.Split('\n').Length, email.Split('\n').Length),
                                Math.Max(telefono.Split('\n').Length, detalleVenta.Split('\n').Length)
                            );
                            ws.Row(row).Height = Math.Max(15, maxLineas * 12);

                            row++;
                        }
                        int resumenRow = row + 1;
                        ws.Cells[resumenRow, 1, resumenRow, 2].Merge = true;
                        ws.Cells[resumenRow, 1].Value = "Resumen de Ventas en Línea";
                        ws.Cells[resumenRow, 1].Style.Font.Bold = true;
                        ws.Cells[resumenRow, 1].Style.Font.Size = 12;
                        ws.Row(resumenRow).Height = 20;
                        ws.Cells[resumenRow + 1, 1].Value = "Total (Bs.):";
                        ws.Cells[resumenRow + 1, 2].Value = totalVentasOnline;
                        ws.Cells[resumenRow + 1, 2].Style.Numberformat.Format = "#,##0.00";
                        ws.Cells[resumenRow + 2, 1].Value = "Cant. Ventas Online:";
                        ws.Cells[resumenRow + 2, 2].Value = cantidadVentasOnline;
                        for (int i = 1; i <= 2; i++)
                        {
                            ws.Cells[resumenRow + i, 1, resumenRow + i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[resumenRow + i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[resumenRow + i, 1, resumenRow + i, 2].Style.Font.Size = 10;
                            ws.Row(resumenRow + i).Height = 15;
                        }
                        ws.Column(1).Width = 10;
                        ws.Column(2).Width = Math.Min(25, CalcularAnchoMax(reporteVentasOnline, "name", 20));
                        ws.Column(3).Width = Math.Min(25, CalcularAnchoMax(reporteVentasOnline, "email", 20));
                        ws.Column(4).Width = Math.Min(15, CalcularAnchoMax(reporteVentasOnline, "telefono", 12));
                        ws.Column(5).Width = 15;
                        ws.Column(6).Width = 12;
                        ws.Column(7).Width = 15; 
                        ws.Column(8).Width = 12;
                        ws.Column(9).Width = Math.Min(30, CalcularAnchoMax(reporteVentasOnline, "Dirrecion", 25));
                        ws.Column(10).Width = 12;
                        ws.Column(11).Width = 15; 
                        ws.Column(12).Width = Math.Min(50, CalcularAnchoMaxDetalles(reporteVentasOnline, "detalle_compra", 40));
                        var chart = ws.Drawings.AddChart("VentasOnlineChart", eChartType.ColumnClustered);
                        chart.SetPosition(resumenRow + 4, 0, 1, 0);
                        chart.SetSize(600, 200);
                        chart.Title.Text = "Resumen de Ventas en Línea";
                        chart.Legend.Position = eLegendPosition.Right;
                        var serie = chart.Series.Add(
                            ws.Cells[resumenRow + 1, 2],
                            ws.Cells[resumenRow + 1, 1]
                        );
                        serie.Header = "Ventas (Bs.)";
                        package.Save();
                    }
                    MessageBox.Show("Reporte de ventas en línea generado con éxito.", "Excel Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al generar el reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async Task<DataTable> ObtenerReporteVentasOnlineAsync()
        {
            DataTable dt = new DataTable();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string query = @"
                        SELECT 
                            v.id AS id_venta,
                            u.name,
                            u.email,
                            u.telefono,
                            v.Fecha_Venta,
                            v.Total,
                            v.Método_Pago,
                            v.tipo_entrega,
                            e.Dirrecion,
                            v.Estado,
                            cp.subido,
                            JSON_ARRAYAGG(
                                JSON_OBJECT(
                                    'Producto', p.nombre,
                                    'Color', c.color,
                                    'Cantidad', vd.cantidad,
                                    'Precio', vd.precio_unitario,
                                    'Categoría', cat.nombre,
                                    'Precio Proveedor', p.precio_proveedor,
                                    'Precio Estudiante', p.precio_estudiante
                                )
                            ) AS detalle_compra
                        FROM ventas v
                        LEFT JOIN users u ON v.ID_Usuario = u.id
                        LEFT JOIN envios e ON v.id = e.ID_Venta
                        LEFT JOIN comprobante_pagos cp ON v.id = cp.ID_Venta
                        LEFT JOIN venta_detalles vd ON v.id = vd.venta_id
                        LEFT JOIN productos p ON vd.producto_id = p.id
                        LEFT JOIN categorias cat ON p.categoria_id = cat.id
                        LEFT JOIN colors c ON p.color_id = c.id
                        GROUP BY v.id, u.name, u.email, u.telefono, v.Fecha_Venta, v.Total, v.Método_Pago, v.tipo_entrega, e.Dirrecion, v.Estado, cp.subido
                        ORDER BY v.Fecha_Venta DESC";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener los datos: {ex.Message}");
            }
            return dt;
        }
        private string FormatearDetalleVenta(string jsonDetalle)
        {
            try
            {
                var detalles = JsonConvert.DeserializeObject<dynamic[]>(jsonDetalle);
                string resultado = "";
                foreach (var item in detalles)
                {
                    resultado += $"Prod: {item.Producto}, Color: {item.Color}, Cant: {item.Cantidad}, Precio: {item.Precio} Bs.\n" +
                                 $"Cat: {item["Categoría"]}\n" +
                                 (item.PrecioProveedor != null ? $"Prov: {item.PrecioProveedor} Bs.\n" : "") +
                                 (item.PrecioEstudiante != null ? $"Est: {item.PrecioEstudiante} Bs.\n" : "") +
                                 "----------------------------------------\n";
                }
                return resultado.TrimEnd('\n');
            }
            catch
            {
                return jsonDetalle;
            }
        }

        private async void buttonreporte3_Click(object sender, EventArgs e)
        {
            DataTable reporteCarritos = await ObtenerReporteCarritosAsync();
            DataTable reporteVentasEliminadas = await ObtenerReporteVentasEliminadasAsync();
            if (reporteCarritos.Rows.Count == 0 && reporteVentasEliminadas.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para generar el reporte.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Reporte_Carritos_VentasEliminadas_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (File.Exists(saveFileDialog.FileName))
                    {
                        MessageBox.Show("El archivo ya existe. Por favor elige otro nombre o ubicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                    {
                        var ws = package.Workbook.Worksheets.Add("Reporte Combinado");
                        ws.PrinterSettings.PaperSize = ePaperSize.Letter;
                        ws.PrinterSettings.Orientation = eOrientation.Landscape;
                        ws.PrinterSettings.FitToPage = true;
                        ws.PrinterSettings.FitToWidth = 1;
                        ws.PrinterSettings.FitToHeight = 0;
                        ws.Cells[1, 1, 1, 11].Merge = true;
                        ws.Cells[1, 1].Value = "Reporte de Carritos y Ventas Eliminadas - POLIMÉRICOS DIAL BOLIVIA";
                        ws.Cells[1, 1].Style.Font.Size = 14;
                        ws.Cells[1, 1].Style.Font.Bold = true;
                        ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 121));
                        ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                        ws.Row(1).Height = 25;
                        ws.Cells[2, 1].Value = $"Generado el: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                        ws.Cells[2, 1].Style.Font.Italic = true;
                        ws.Cells[2, 1].Style.Font.Size = 10;
                        ws.Row(2).Height = 15;
                        int row = 4;
                        ws.Cells[row, 1, row, 7].Merge = true;
                        ws.Cells[row, 1].Value = "Reporte de Carritos Activos";
                        ws.Cells[row, 1].Style.Font.Bold = true;
                        ws.Cells[row, 1].Style.Font.Size = 12;
                        ws.Row(row).Height = 20;
                        row++;
                        string[] headersCarritos = { "ID Carrito", "Usuario", "Producto", "Cantidad", "Fecha Agregado", "Precio Unitario", "Detalles" };
                        for (int col = 1; col <= headersCarritos.Length; col++)
                        {
                            ws.Cells[row, col].Value = headersCarritos[col - 1];
                            ws.Cells[row, col].Style.Font.Bold = true;
                            ws.Cells[row, col].Style.Font.Size = 10;
                            ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[row, col].Style.WrapText = true;
                            ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        ws.Row(row).Height = 20;
                        row++;
                        int totalProductosCarrito = 0;
                        int cantidadCarritos = reporteCarritos.Rows.Count;
                        foreach (DataRow dataRow in reporteCarritos.Rows)
                        {
                            string idCarrito = dataRow["id_carrito"].ToString();
                            string usuario = dataRow["name"].ToString();
                            string producto = dataRow["nombre"].ToString();
                            string cantidad = dataRow["Cantidad"].ToString();
                            DateTime fechaAgregado = Convert.ToDateTime(dataRow["Fecha_Agregado"]);
                            decimal precioUnitario = Convert.ToDecimal(dataRow["precio"]);
                            string detalles = FormatearDetallesCarrito(producto, dataRow["color"]?.ToString(), precioUnitario);

                            totalProductosCarrito += Convert.ToInt32(cantidad);
                            ws.Cells[row, 1].Value = idCarrito;
                            ws.Cells[row, 2].Value = usuario;
                            ws.Cells[row, 3].Value = producto;
                            ws.Cells[row, 4].Value = cantidad;
                            ws.Cells[row, 5].Value = fechaAgregado.ToString("dd/MM/yyyy HH:mm");
                            ws.Cells[row, 6].Value = precioUnitario;
                            ws.Cells[row, 6].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[row, 7].Value = detalles;
                            for (int col = 1; col <= 7; col++)
                            {
                                ws.Cells[row, col].Style.WrapText = true;
                                ws.Cells[row, col].Style.Font.Size = 9;
                                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            int maxLineas = Math.Max(usuario.Split('\n').Length, detalles.Split('\n').Length);
                            ws.Row(row).Height = Math.Max(15, maxLineas * 12);
                            row++;
                        }
                        int resumenCarritosRow = row + 1;
                        ws.Cells[resumenCarritosRow, 1, resumenCarritosRow, 2].Merge = true;
                        ws.Cells[resumenCarritosRow, 1].Value = "Resumen de Carritos";
                        ws.Cells[resumenCarritosRow, 1].Style.Font.Bold = true;
                        ws.Cells[resumenCarritosRow, 1].Style.Font.Size = 12;
                        ws.Row(resumenCarritosRow).Height = 20;
                        ws.Cells[resumenCarritosRow + 1, 1].Value = "Total Productos en Carritos:";
                        ws.Cells[resumenCarritosRow + 1, 2].Value = totalProductosCarrito;
                        ws.Cells[resumenCarritosRow + 2, 1].Value = "Cantidad de Carritos Activos:";
                        ws.Cells[resumenCarritosRow + 2, 2].Value = cantidadCarritos;
                        for (int i = 1; i <= 2; i++)
                        {
                            ws.Cells[resumenCarritosRow + i, 1, resumenCarritosRow + i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[resumenCarritosRow + i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[resumenCarritosRow + i, 1, resumenCarritosRow + i, 2].Style.Font.Size = 10;
                            ws.Row(resumenCarritosRow + i).Height = 15;
                        }
                        row = resumenCarritosRow + 4; 
                        ws.Cells[row, 1, row, 11].Merge = true;
                        ws.Cells[row, 1].Value = "Reporte de Ventas Eliminadas";
                        ws.Cells[row, 1].Style.Font.Bold = true;
                        ws.Cells[row, 1].Style.Font.Size = 12;
                        ws.Row(row).Height = 20;
                        row++;
                        string[] headersVentasEliminadas = { "ID", "Usuario", "Cantidad", "Fecha Venta", "Total (Bs.)", "Método Pago", "Tipo Entrega", "Estado", "Detalles Venta", "Error", "Fecha Eliminación" };
                        for (int col = 1; col <= headersVentasEliminadas.Length; col++)
                        {
                            ws.Cells[row, col].Value = headersVentasEliminadas[col - 1];
                            ws.Cells[row, col].Style.Font.Bold = true;
                            ws.Cells[row, col].Style.Font.Size = 10;
                            ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[row, col].Style.WrapText = true;
                            ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        ws.Row(row).Height = 20;
                        row++;
                        decimal totalVentasEliminadas = 0;
                        int cantidadVentasEliminadas = reporteVentasEliminadas.Rows.Count;
                        foreach (DataRow dataRow in reporteVentasEliminadas.Rows)
                        {
                            string idVenta = dataRow["id"].ToString();
                            string usuario = dataRow["name"]?.ToString() ?? "N/A";
                            string cantidad = dataRow["Cantidad"]?.ToString() ?? "N/A";
                            string fechaVenta = dataRow["Fecha_Venta"] != DBNull.Value ? Convert.ToDateTime(dataRow["Fecha_Venta"]).ToString("dd/MM/yyyy HH:mm") : "N/A";
                            decimal total = dataRow["Total"] != DBNull.Value ? Convert.ToDecimal(dataRow["Total"]) : 0;
                            string metodoPago = dataRow["Método_Pago"]?.ToString() ?? "N/A";
                            string tipoEntrega = dataRow["tipo_entrega"]?.ToString() ?? "N/A";
                            string estado = dataRow["Estado"]?.ToString() ?? "N/A";
                            string detallesVenta = FormatearDetallesVenta(dataRow["detalles_venta"]?.ToString());
                            string error = dataRow["error_detalle"]?.ToString() ?? "N/A";
                            DateTime fechaEliminacion = Convert.ToDateTime(dataRow["created_at"]);
                            totalVentasEliminadas += total;
                            ws.Cells[row, 1].Value = idVenta;
                            ws.Cells[row, 2].Value = usuario;
                            ws.Cells[row, 3].Value = cantidad;
                            ws.Cells[row, 4].Value = fechaVenta;
                            ws.Cells[row, 5].Value = total;
                            ws.Cells[row, 5].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[row, 6].Value = metodoPago;
                            ws.Cells[row, 7].Value = tipoEntrega;
                            ws.Cells[row, 8].Value = estado;
                            ws.Cells[row, 9].Value = detallesVenta;
                            ws.Cells[row, 10].Value = error;
                            ws.Cells[row, 11].Value = fechaEliminacion.ToString("dd/MM/yyyy HH:mm");
                            for (int col = 1; col <= 11; col++)
                            {
                                ws.Cells[row, col].Style.WrapText = true;
                                ws.Cells[row, col].Style.Font.Size = 9;
                                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            int maxLineas = Math.Max(
                                Math.Max(usuario.Split('\n').Length, error.Split('\n').Length),
                                detallesVenta.Split('\n').Length
                            );
                            ws.Row(row).Height = Math.Max(15, maxLineas * 12);
                            row++;
                        }
                        int resumenVentasEliminadasRow = row + 1;
                        ws.Cells[resumenVentasEliminadasRow, 1, resumenVentasEliminadasRow, 2].Merge = true;
                        ws.Cells[resumenVentasEliminadasRow, 1].Value = "Resumen de Ventas Eliminadas";
                        ws.Cells[resumenVentasEliminadasRow, 1].Style.Font.Bold = true;
                        ws.Cells[resumenVentasEliminadasRow, 1].Style.Font.Size = 12;
                        ws.Row(resumenVentasEliminadasRow).Height = 20;
                        ws.Cells[resumenVentasEliminadasRow + 1, 1].Value = "Total Ventas Eliminadas (Bs.):";
                        ws.Cells[resumenVentasEliminadasRow + 1, 2].Value = totalVentasEliminadas;
                        ws.Cells[resumenVentasEliminadasRow + 1, 2].Style.Numberformat.Format = "#,##0.00";
                        ws.Cells[resumenVentasEliminadasRow + 2, 1].Value = "Cantidad Ventas Eliminadas:";
                        ws.Cells[resumenVentasEliminadasRow + 2, 2].Value = cantidadVentasEliminadas;
                        for (int i = 1; i <= 2; i++)
                        {
                            ws.Cells[resumenVentasEliminadasRow + i, 1, resumenVentasEliminadasRow + i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[resumenVentasEliminadasRow + i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[resumenVentasEliminadasRow + i, 1, resumenVentasEliminadasRow + i, 2].Style.Font.Size = 10;
                            ws.Row(resumenVentasEliminadasRow + i).Height = 15;
                        }
                        ws.Column(1).Width = 10;
                        ws.Column(2).Width = Math.Min(25, CalcularAnchoMax(reporteCarritos, "name", 20));
                        ws.Column(3).Width = Math.Min(25, CalcularAnchoMax(reporteCarritos, "nombre", 20));
                        ws.Column(4).Width = 8;
                        ws.Column(5).Width = 15; 
                        ws.Column(6).Width = 12;
                        ws.Column(7).Width = Math.Min(40, CalcularAnchoMaxDetalles(reporteCarritos, "nombre", 30));
                        ws.Column(8).Width = 15; 
                        ws.Column(9).Width = Math.Min(50, CalcularAnchoMaxDetalles(reporteVentasEliminadas, "detalles_venta", 40));
                        ws.Column(10).Width = Math.Min(30, CalcularAnchoMax(reporteVentasEliminadas, "error_detalle", 25));
                        ws.Column(11).Width = 15; 
                        package.Save();
                    }
                    MessageBox.Show("Reporte de carritos y ventas eliminadas generado con éxito.", "Excel Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al generar el reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private async Task<DataTable> ObtenerReporteCarritosAsync()
        {
            DataTable dt = new DataTable();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string query = @"
                        SELECT 
                            c.id AS id_carrito,
                            u.name,
                            p.nombre,
                            p.precio,
                            c.Cantidad,
                            c.Fecha_Agregado,
                            col.color
                        FROM carritos c
                        LEFT JOIN users u ON c.ID_Usuario = u.id
                        LEFT JOIN productos p ON c.ID_Producto = p.id
                        LEFT JOIN colors col ON p.color_id = col.id
                        ORDER BY c.Fecha_Agregado DESC";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener datos de carritos: {ex.Message}");
            }
            return dt;
        }
        private async Task<DataTable> ObtenerReporteVentasEliminadasAsync()
        {
            DataTable dt = new DataTable();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string query = @"
                        SELECT 
                            ve.id,
                            u.name,
                            ve.Cantidad,
                            ve.Fecha_Venta,
                            ve.Total,
                            ve.Método_Pago,
                            ve.tipo_entrega,
                            ve.Estado,
                            ve.detalles_venta,
                            ve.error_detalle,
                            ve.created_at
                        FROM venta_eliminadas ve
                        LEFT JOIN users u ON ve.ID_Usuario = u.id
                        ORDER BY ve.created_at DESC";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener datos de ventas eliminadas: {ex.Message}");
            }
            return dt;
        }
        private string FormatearDetallesCarrito(string producto, string color, decimal precio)
        {
            return $"Prod: {producto}\n" +
                   (string.IsNullOrEmpty(color) ? "" : $"Color: {color}\n") +
                   $"Precio: {precio:F2} Bs.";
        }
        private string FormatearDetallesVenta(string jsonDetalle)
        {
            if (string.IsNullOrEmpty(jsonDetalle)) return "N/A";
            try
            {
                var detalles = JsonConvert.DeserializeObject<dynamic[]>(jsonDetalle);
                string resultado = "";
                foreach (var item in detalles)
                {
                    resultado += $"Prod: {item.Producto}, Color: {item.Color}, Cant: {item.Cantidad}, Precio: {item.Precio} Bs.\n" +
                                 $"Cat: {item["Categoría"]}\n" +
                                 (item.PrecioProveedor != null ? $"Prov: {item.PrecioProveedor} Bs.\n" : "") +
                                 (item.PrecioEstudiante != null ? $"Est: {item.PrecioEstudiante} Bs.\n" : "") +
                                 "----------------------------------------\n";
                }
                return resultado.TrimEnd('\n');
            }
            catch
            {
                return jsonDetalle;
            }
        }
        private async void buttonreporte4_Click(object sender, EventArgs e)
        {
            DataTable reporteProductosVendidos = await ObtenerReporteProductosVendidosAsync();
            DataTable reporteStock = await ObtenerReporteStockAsync();
            if (reporteProductosVendidos.Rows.Count == 0 && reporteStock.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para generar el reporte.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = $"Reporte_ProductosVendidos_Stock_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (File.Exists(saveFileDialog.FileName))
                    {
                        MessageBox.Show("El archivo ya existe. Por favor elige otro nombre o ubicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                    {
                        var ws = package.Workbook.Worksheets.Add("Reporte Productos");
                        ws.PrinterSettings.PaperSize = ePaperSize.Letter;
                        ws.PrinterSettings.Orientation = eOrientation.Landscape;
                        ws.PrinterSettings.FitToPage = true;
                        ws.PrinterSettings.FitToWidth = 1;
                        ws.PrinterSettings.FitToHeight = 0;
                        ws.Cells[1, 1, 1, 9].Merge = true; 
                        ws.Cells[1, 1].Value = "Reporte de Productos Más Vendidos y Stock - POLIMÉRICOS DIAL BOLIVIA";
                        ws.Cells[1, 1].Style.Font.Size = 14;
                        ws.Cells[1, 1].Style.Font.Bold = true;
                        ws.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(31, 78, 121));
                        ws.Cells[1, 1].Style.Font.Color.SetColor(Color.White);
                        ws.Row(1).Height = 25;
                        ws.Cells[2, 1].Value = $"Generado el: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                        ws.Cells[2, 1].Style.Font.Italic = true;
                        ws.Cells[2, 1].Style.Font.Size = 10;
                        ws.Row(2).Height = 15;
                        int row = 4;
                        ws.Cells[row, 1, row, 5].Merge = true;
                        ws.Cells[row, 1].Value = "Reporte de Productos Más Vendidos";
                        ws.Cells[row, 1].Style.Font.Bold = true;
                        ws.Cells[row, 1].Style.Font.Size = 12;
                        ws.Row(row).Height = 20;
                        row++;
                        string[] headersProductosVendidos = { "ID Producto", "Nombre", "Cantidad Total Vendida", "Total Ingresos (Bs.)", "Número de Ventas" };
                        for (int col = 1; col <= headersProductosVendidos.Length; col++)
                        {
                            ws.Cells[row, col].Value = headersProductosVendidos[col - 1];
                            ws.Cells[row, col].Style.Font.Bold = true;
                            ws.Cells[row, col].Style.Font.Size = 10;
                            ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[row, col].Style.WrapText = true;
                            ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        ws.Row(row).Height = 20;
                        row++;
                        int totalCantidadVendida = 0;
                        decimal totalIngresos = 0;
                        int totalVentas = reporteProductosVendidos.Rows.Count;
                        foreach (DataRow dataRow in reporteProductosVendidos.Rows)
                        {
                            string idProducto = dataRow["producto_id"].ToString();
                            string nombre = dataRow["nombre"].ToString();
                            int cantidadVendida = Convert.ToInt32(dataRow["total_cantidad"]);
                            decimal ingresos = Convert.ToDecimal(dataRow["total_ingresos"]);
                            int numeroVentas = Convert.ToInt32(dataRow["numero_ventas"]);
                            totalCantidadVendida += cantidadVendida;
                            totalIngresos += ingresos;
                            ws.Cells[row, 1].Value = idProducto;
                            ws.Cells[row, 2].Value = nombre;
                            ws.Cells[row, 3].Value = cantidadVendida;
                            ws.Cells[row, 4].Value = ingresos;
                            ws.Cells[row, 4].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[row, 5].Value = numeroVentas;
                            for (int col = 1; col <= 5; col++)
                            {
                                ws.Cells[row, col].Style.WrapText = true;
                                ws.Cells[row, col].Style.Font.Size = 9;
                                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            int maxLineas = nombre.Split('\n').Length;
                            ws.Row(row).Height = Math.Max(15, maxLineas * 12);
                            row++;
                        }
                        int resumenProductosRow = row + 1;
                        ws.Cells[resumenProductosRow, 1, resumenProductosRow, 2].Merge = true;
                        ws.Cells[resumenProductosRow, 1].Value = "Resumen de Ventas de Productos";
                        ws.Cells[resumenProductosRow, 1].Style.Font.Bold = true;
                        ws.Cells[resumenProductosRow, 1].Style.Font.Size = 12;
                        ws.Row(resumenProductosRow).Height = 20;
                        ws.Cells[resumenProductosRow + 1, 1].Value = "Total Cantidad Vendida:";
                        ws.Cells[resumenProductosRow + 1, 2].Value = totalCantidadVendida;
                        ws.Cells[resumenProductosRow + 2, 1].Value = "Total Ingresos (Bs.):";
                        ws.Cells[resumenProductosRow + 2, 2].Value = totalIngresos;
                        ws.Cells[resumenProductosRow + 2, 2].Style.Numberformat.Format = "#,##0.00";
                        ws.Cells[resumenProductosRow + 3, 1].Value = "Total Ventas Realizadas:";
                        ws.Cells[resumenProductosRow + 3, 2].Value = totalVentas;
                        for (int i = 1; i <= 3; i++)
                        {
                            ws.Cells[resumenProductosRow + i, 1, resumenProductosRow + i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[resumenProductosRow + i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[resumenProductosRow + i, 1, resumenProductosRow + i, 2].Style.Font.Size = 10;
                            ws.Row(resumenProductosRow + i).Height = 15;
                        }
                        row = resumenProductosRow + 5;
                        ws.Cells[row, 1, row, 9].Merge = true;
                        ws.Cells[row, 1].Value = "Reporte de Estado del Stock";
                        ws.Cells[row, 1].Style.Font.Bold = true;
                        ws.Cells[row, 1].Style.Font.Size = 12;
                        ws.Row(row).Height = 20;
                        row++;
                        string[] headersStock = { "ID Producto", "Nombre", "Categoría", "Stock Actual", "Precio (Bs.)", "Precio Estudiante", "Precio Proveedor", "Estado", "Activo" };
                        for (int col = 1; col <= headersStock.Length; col++)
                        {
                            ws.Cells[row, col].Value = headersStock[col - 1];
                            ws.Cells[row, col].Style.Font.Bold = true;
                            ws.Cells[row, col].Style.Font.Size = 10;
                            ws.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[row, col].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[row, col].Style.WrapText = true;
                            ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        ws.Row(row).Height = 20;
                        row++;
                        int totalStock = 0;
                        int productosActivos = 0;
                        foreach (DataRow dataRow in reporteStock.Rows)
                        {
                            string idProducto = dataRow["id"].ToString();
                            string nombre = dataRow["nombre"].ToString();
                            string categoria = dataRow["categoria_nombre"].ToString();
                            int stock = Convert.ToInt32(dataRow["stock"]);
                            decimal precio = Convert.ToDecimal(dataRow["precio"]);
                            decimal? precioEstudiante = dataRow["precio_estudiante"] != DBNull.Value ? (decimal?)Convert.ToDecimal(dataRow["precio_estudiante"]) : null;
                            decimal? precioProveedor = dataRow["precio_proveedor"] != DBNull.Value ? (decimal?)Convert.ToDecimal(dataRow["precio_proveedor"]) : null;
                            string estado = dataRow["estado"].ToString();
                            bool activo = Convert.ToBoolean(dataRow["activo"]);
                            totalStock += stock;
                            if (activo) productosActivos++;
                            ws.Cells[row, 1].Value = idProducto;
                            ws.Cells[row, 2].Value = nombre;
                            ws.Cells[row, 3].Value = categoria;
                            ws.Cells[row, 4].Value = stock;
                            ws.Cells[row, 5].Value = precio;
                            ws.Cells[row, 5].Style.Numberformat.Format = "#,##0.00";
                            ws.Cells[row, 6].Value = precioEstudiante.HasValue ? precioEstudiante.Value.ToString("F2") : "N/A";
                            ws.Cells[row, 6].Style.Numberformat.Format = precioEstudiante.HasValue ? "#,##0.00" : "";
                            ws.Cells[row, 7].Value = precioProveedor.HasValue ? precioProveedor.Value.ToString("F2") : "N/A";
                            ws.Cells[row, 7].Style.Numberformat.Format = precioProveedor.HasValue ? "#,##0.00" : "";
                            ws.Cells[row, 8].Value = estado;
                            ws.Cells[row, 9].Value = activo ? "Sí" : "No";
                            for (int col = 1; col <= 9; col++)
                            {
                                ws.Cells[row, col].Style.WrapText = true;
                                ws.Cells[row, col].Style.Font.Size = 9;
                                ws.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                ws.Cells[row, col].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            }
                            int maxLineas = Math.Max(nombre.Split('\n').Length, categoria.Split('\n').Length);
                            ws.Row(row).Height = Math.Max(15, maxLineas * 12);
                            row++;
                        }
                        int resumenStockRow = row + 1;
                        ws.Cells[resumenStockRow, 1, resumenStockRow, 2].Merge = true;
                        ws.Cells[resumenStockRow, 1].Value = "Resumen de Stock";
                        ws.Cells[resumenStockRow, 1].Style.Font.Bold = true;
                        ws.Cells[resumenStockRow, 1].Style.Font.Size = 12;
                        ws.Row(resumenStockRow).Height = 20;
                        ws.Cells[resumenStockRow + 1, 1].Value = "Total Productos en Stock:";
                        ws.Cells[resumenStockRow + 1, 2].Value = totalStock;
                        ws.Cells[resumenStockRow + 2, 1].Value = "Cantidad de Productos Activos:";
                        ws.Cells[resumenStockRow + 2, 2].Value = productosActivos;
                        for (int i = 1; i <= 2; i++)
                        {
                            ws.Cells[resumenStockRow + i, 1, resumenStockRow + i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            ws.Cells[resumenStockRow + i, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            ws.Cells[resumenStockRow + i, 1, resumenStockRow + i, 2].Style.Font.Size = 10;
                            ws.Row(resumenStockRow + i).Height = 15;
                        }
                        ws.Column(1).Width = 10; 
                        ws.Column(2).Width = Math.Min(30, CalcularAnchoMax(reporteProductosVendidos, "nombre", 25));
                        ws.Column(3).Width = Math.Min(20, CalcularAnchoMax(reporteStock, "categoria_nombre", 15)); 
                        ws.Column(4).Width = 15;
                        ws.Column(5).Width = 15; 
                        ws.Column(6).Width = 15;
                        ws.Column(7).Width = 15; 
                        ws.Column(8).Width = 15;
                        ws.Column(9).Width = 10; 
                        package.Save();
                    }

                    MessageBox.Show("Reporte de productos vendidos y stock generado con éxito.", "Excel Generado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al generar el reporte: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private async Task<DataTable> ObtenerReporteProductosVendidosAsync()
        {
            DataTable dt = new DataTable();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string query = @"
                        SELECT 
                            p.id AS producto_id,
                            p.nombre,
                            SUM(COALESCE(vd.cantidad, vl.cantidad)) AS total_cantidad,
                            SUM(COALESCE(vd.subtotal, vl.total)) AS total_ingresos,
                            COUNT(DISTINCT COALESCE(vd.venta_id, vl.id)) AS numero_ventas
                        FROM productos p
                        LEFT JOIN venta_detalles vd ON p.id = vd.producto_id
                        LEFT JOIN venta_locals vl ON JSON_CONTAINS(vl.detalle_compra, JSON_OBJECT('ID', p.id))
                        GROUP BY p.id, p.nombre
                        HAVING total_cantidad > 0
                        ORDER BY total_cantidad DESC";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener datos de productos vendidos: {ex.Message}");
            }
            return dt;
        }

        private async Task<DataTable> ObtenerReporteStockAsync()
        {
            DataTable dt = new DataTable();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    await connection.OpenAsync();
                    string query = @"
                        SELECT 
                            p.id,
                            p.nombre,
                            c.nombre AS categoria_nombre,
                            p.stock,
                            p.precio,
                            p.precio_estudiante,
                            p.precio_proveedor,
                            p.estado,
                            p.activo
                        FROM productos p
                        LEFT JOIN categorias c ON p.categoria_id = c.id
                        ORDER BY p.id";
                    MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
                    adapter.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al obtener datos de stock: {ex.Message}");
            }
            return dt;
        }

        private void buttonsalir_Click(object sender, EventArgs e)
        {
            try
            {
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cerrar la aplicación: " + ex.Message);
            }
        }

        private void buttonminimizar_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panelReportes_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
