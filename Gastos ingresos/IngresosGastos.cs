using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Gastos_ingresos
{
    public partial class Ingresos_Gastos : Form
    {
        // 1. Variables a nivel de CLASE (fuera del constructor)
        private string cadenaConexion = "Data Source=.;Initial Catalog=FinanzasDB;Integrated Security=True";
        private int idSeleccionado = -1; // Para saber qué registro estamos modificando o borrando

        public Ingresos_Gastos()
        {
            InitializeComponent();
        }

        // Evento Load del formulario
        private void Ingresos_Gastos_Load(object sender, EventArgs e)
        {
            CargarGrillaYBalance();
        }

        // 2. Métodos y eventos de los botones (afuera del constructor)

        // Método para cargar todos los datos en el DataGridView (DgIngresosGastos)
        private void CargarGrilla()
        {
            try
            {
                using (SqlConnection conexion = new SqlConnection(cadenaConexion))
                {
                    string query = "SELECT Id, Tipo, Monto, Fecha, Nombre, Descripcion FROM Transacciones";
                    SqlDataAdapter adapter = new SqlDataAdapter(query, conexion);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    DgIngresosGastos.DataSource = dt;

                    // Ocultar la columna Id para que el usuario no la vea
                    if (DgIngresosGastos.Columns["Id"] != null)
                    {
                        DgIngresosGastos.Columns["Id"].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar los datos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Botón AGREGAR
        private void btnAgregar_Click(object sender, EventArgs e)
        {
            try
            {
                string tipo = "";
                decimal monto = 0;
                DateTime fecha = DateTime.Now; // <-- Inicializada aquí para evitar el error
                string nombre = "";
                string descripcion = "";

                if (!string.IsNullOrWhiteSpace(textIngreso.Text))
                {
                    tipo = "Ingreso";
                    monto = decimal.Parse(textIngreso.Text);
                    fecha = dateTimeIngreso.Value;
                    nombre = textNombreIngreso.Text;
                    descripcion = textDescripcionIngreso.Text;
                }
                else if (!string.IsNullOrWhiteSpace(textGastos.Text))
                {
                    tipo = "Gasto";
                    monto = decimal.Parse(textGastos.Text);
                    fecha = dateTimeGastos.Value;
                    nombre = textNombreGastos.Text;
                    descripcion = textDescripcionGastos.Text;
                }
                else
                {
                    MessageBox.Show("Debe ingresar un monto en Ingresos o Gastos.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using (SqlConnection conexion = new SqlConnection(cadenaConexion))
                {
                    string query = "INSERT INTO Transacciones (Tipo, Monto, Fecha, Nombre, Descripcion) VALUES (@Tipo, @Monto, @Fecha, @Nombre, @Descripcion)";
                    using (SqlCommand cmd = new SqlCommand(query, conexion))
                    {
                        cmd.Parameters.AddWithValue("@Tipo", tipo);
                        cmd.Parameters.AddWithValue("@Monto", monto);
                        cmd.Parameters.AddWithValue("@Fecha", fecha);
                        cmd.Parameters.AddWithValue("@Nombre", nombre);
                        cmd.Parameters.AddWithValue("@Descripcion", string.IsNullOrEmpty(descripcion) ? (object)DBNull.Value : descripcion);

                        conexion.Open();
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Guardado con éxito.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarGrilla();
                LimpiarCampos();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al guardar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            CargarGrillaYBalance();
        }

        // Evento al hacer clic en una celda de la grilla para rellenar los campos
        private void DgIngresosGastos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow fila = DgIngresosGastos.Rows[e.RowIndex];
                idSeleccionado = Convert.ToInt32(fila.Cells["Id"].Value);
                string tipo = fila.Cells["Tipo"].Value.ToString();

                if (tipo == "Ingreso")
                {
                    textIngreso.Text = fila.Cells["Monto"].Value.ToString();
                    dateTimeIngreso.Value = Convert.ToDateTime(fila.Cells["Fecha"].Value);
                    textNombreIngreso.Text = fila.Cells["Nombre"].Value.ToString();
                    textDescripcionIngreso.Text = fila.Cells["Descripcion"].Value.ToString();

                    LimpiarCamposGastosSolo();
                }
                else
                {
                    textGastos.Text = fila.Cells["Monto"].Value.ToString();
                    dateTimeGastos.Value = Convert.ToDateTime(fila.Cells["Fecha"].Value);
                    textNombreGastos.Text = fila.Cells["Nombre"].Value.ToString();
                    textDescripcionGastos.Text = fila.Cells["Descripcion"].Value.ToString();

                    LimpiarCamposIngresosSolo();
                }
            }
        }

        // Botón MODIFICAR
        private void btnModificar_Click(object sender, EventArgs e)
        {
            if (idSeleccionado == -1)
            {
                MessageBox.Show("Seleccione un registro de la grilla para modificar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string tipo = "";
                decimal monto = 0;
                DateTime fecha = DateTime.Now; // <-- Inicializada aquí también
                string nombre = "";
                string descripcion = "";

                if (!string.IsNullOrWhiteSpace(textIngreso.Text))
                {
                    tipo = "Ingreso";
                    monto = decimal.Parse(textIngreso.Text);
                    fecha = dateTimeIngreso.Value;
                    nombre = textNombreIngreso.Text;
                    descripcion = textDescripcionIngreso.Text;
                }
                else if (!string.IsNullOrWhiteSpace(textGastos.Text))
                {
                    tipo = "Gasto";
                    monto = decimal.Parse(textGastos.Text);
                    fecha = dateTimeGastos.Value;
                    nombre = textNombreGastos.Text;
                    descripcion = textDescripcionGastos.Text;
                }

                using (SqlConnection conexion = new SqlConnection(cadenaConexion))
                {
                    string query = "UPDATE Transacciones SET Tipo = @Tipo, Monto = @Monto, Fecha = @Fecha, Nombre = @Nombre, Descripcion = @Descripcion WHERE Id = @Id";
                    using (SqlCommand cmd = new SqlCommand(query, conexion))
                    {
                        cmd.Parameters.AddWithValue("@Tipo", tipo);
                        cmd.Parameters.AddWithValue("@Monto", monto);
                        cmd.Parameters.AddWithValue("@Fecha", fecha);
                        cmd.Parameters.AddWithValue("@Nombre", nombre);
                        cmd.Parameters.AddWithValue("@Descripcion", string.IsNullOrEmpty(descripcion) ? (object)DBNull.Value : descripcion);
                        cmd.Parameters.AddWithValue("@Id", idSeleccionado);

                        conexion.Open();
                        cmd.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("Modificado con éxito.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarGrilla();
                LimpiarCampos();
                idSeleccionado = -1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al modificar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            CargarGrillaYBalance();
        }

        // Botón BORRAR
        private void btnBorrar_Click(object sender, EventArgs e)
        {
            if (idSeleccionado == -1)
            {
                MessageBox.Show("Seleccione un registro de la grilla para borrar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("¿Está seguro de eliminar este registro?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection conexion = new SqlConnection(cadenaConexion))
                    {
                        string query = "DELETE FROM Transacciones WHERE Id = @Id";
                        using (SqlCommand cmd = new SqlCommand(query, conexion))
                        {
                            cmd.Parameters.AddWithValue("@Id", idSeleccionado);
                            conexion.Open();
                            cmd.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("Eliminado correctamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    CargarGrilla();
                    LimpiarCampos();
                    idSeleccionado = -1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al eliminar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            CargarGrillaYBalance();
        }

        // Botón ACTUALIZAR
        private void btnActualizar_Click(object sender, EventArgs e)
        {
            CargarGrilla();
            LimpiarCampos();
            idSeleccionado = -1;
        }

        // Métodos auxiliares de limpieza
        private void LimpiarCampos()
        {
            LimpiarCamposIngresosSolo();
            LimpiarCamposGastosSolo();
        }

        private void LimpiarCamposIngresosSolo()
        {
            textIngreso.Clear();
            textNombreIngreso.Clear();
            textDescripcionIngreso.Clear();
            dateTimeIngreso.Value = DateTime.Now;
        }

        private void LimpiarCamposGastosSolo()
        {
            textGastos.Clear();
            textNombreGastos.Clear();
            textDescripcionGastos.Clear();
            dateTimeGastos.Value = DateTime.Now;
        }

        // Método unificado para cargar la grilla y calcular el balance del mes actual
        private void CargarGrillaYBalance()
        {
            try
            {
                using (SqlConnection conexion = new SqlConnection(cadenaConexion))
                {
                    // 1. Cargar la grilla con todos los registros
                    string queryGrilla = "SELECT Id, Tipo, Monto, Fecha, Nombre, Descripcion FROM Transacciones";
                    SqlDataAdapter adapter = new SqlDataAdapter(queryGrilla, conexion);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    DgIngresosGastos.DataSource = dt;

                    if (DgIngresosGastos.Columns["Id"] != null)
                    {
                        DgIngresosGastos.Columns["Id"].Visible = false;
                    }

                    // 2. Calcular los totales filtrando por el mes y año actual (Octubre 2026, por ejemplo)
                    int mesActual = DateTime.Now.Month;
                    int anioActual = DateTime.Now.Year;

                    string queryTotales = @"
                SELECT 
                    ISNULL(SUM(CASE WHEN Tipo = 'Ingreso' THEN Monto ELSE 0 END), 0) AS TotalIngresos,
                    ISNULL(SUM(CASE WHEN Tipo = 'Gasto' THEN Monto ELSE 0 END), 0) AS TotalGastos
                FROM Transacciones";

                    using (SqlCommand cmd = new SqlCommand(queryTotales, conexion))
                    {
                        cmd.Parameters.AddWithValue("@Mes", mesActual);
                        cmd.Parameters.AddWithValue("@Anio", anioActual);

                        // Como el DataAdapter ya abrió y cerró la conexión, abrimos explícitamente para el DataReader
                        if (conexion.State != ConnectionState.Open)
                        {
                            conexion.Open();
                        }

                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                decimal totalIngresos = Convert.ToDecimal(reader["TotalIngresos"]);
                                decimal totalGastos = Convert.ToDecimal(reader["TotalGastos"]);
                                decimal balanceDisponible = totalIngresos - totalGastos;

                                // 3. Mostrar los resultados en las etiquetas que creaste
                                lblIngresosMes.Text = "Ingresos del Mes: $ " + totalIngresos.ToString("N2");
                                lblGastosMes.Text = "Gastos del Mes: $ " + totalGastos.ToString("N2");
                                lblBalanceMes.Text = "Disponible: $ " + balanceDisponible.ToString("N2");

                                // Opcional: Pintar el balance de verde si es positivo o cero, y rojo si da negativo
                                lblBalanceMes.ForeColor = balanceDisponible >= 0 ? Color.DarkGreen : Color.DarkRed;
                            }
                        }
                    }

                    DgIngresosGastos.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar los datos y el balance: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Ingresos_Gastos_Shown(object sender, EventArgs e)
        {
            CargarGrillaYBalance();
        }
    }
}