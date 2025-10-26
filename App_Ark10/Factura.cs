using System;
using System.Collections.Generic;
using System.Text;

public class Factura
{
    public int IDVenta { get; set; }
    public string Cliente { get; set; }
    public string Telefono { get; set; }
    public string CI { get; set; }
    public string Fecha { get; set; }
    public string TipoTotal { get; set; }
    public List<DetalleFactura> Detalles { get; set; }
    public double Total { get; set; }

    public Factura()
    {
        Detalles = new List<DetalleFactura>();
    }

    private string FormatearTexto(string texto, int ancho)
    {
        StringBuilder resultado = new StringBuilder();
        int posicion = 0;

        while (posicion < texto.Length)
        {
            int longitud = Math.Min(ancho, texto.Length - posicion);
            resultado.AppendLine("   " + texto.Substring(posicion, longitud));
            posicion += longitud;
        }

        return resultado.ToString();
    }

    public string FormatearFactura()
    {
        StringBuilder factura = new StringBuilder();
        string separador = new string('-', 32);
        string espacio = "  ";
        factura.AppendLine(espacio + "POLIMÉRICOS DIAL BOLIVIA");
        factura.AppendLine(espacio + "       NIT: 123456789");
        factura.AppendLine(separador);
        factura.AppendLine($"Venta ID: {IDVenta}");
        factura.AppendLine($"Cliente: {Cliente}");
        factura.AppendLine($"Tel: {Telefono}");
        factura.AppendLine($"CI: {CI}");
        factura.AppendLine($"Fecha: {Fecha}");
        factura.AppendLine(separador);
        factura.AppendLine("ID  Producto    | Cant.| Precio");
        factura.AppendLine(separador);
        foreach (var item in Detalles)
        {
            factura.AppendLine($"{item.ID,-3} {FormatearTexto(item.Producto, 18)}");
            factura.AppendLine($"   Categoría:");
            factura.Append(FormatearTexto(item.Categoria, 18));
            factura.AppendLine($"   Color:");
            factura.Append(FormatearTexto(item.Color, 18));
            factura.AppendLine($"                | {item.Cantidad}    | {item.Precio:F2}");
            factura.AppendLine(separador);
        }
        factura.AppendLine($"{TipoTotal}: Bs. {Total:F2}");
        factura.AppendLine(separador);
        factura.AppendLine("Términos y Condiciones:");
        factura.AppendLine("- No se aceptan devoluciones.");
        factura.AppendLine("- Gracias por su preferencia.");
        factura.AppendLine("- ¡Vuelva pronto!");
        return factura.ToString();
    }
}

public class DetalleFactura
{
    public int ID { get; set; }
    public string Producto { get; set; }
    public string Categoria { get; set; }
    public string Color { get; set; }
    public int Cantidad { get; set; }
    public double Precio { get; set; }
}
