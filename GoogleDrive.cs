using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace GoogleServices
{
    [ComVisible(true)]
    [Guid("0ddf5ab6-b797-462b-9243-8d3edec6755d")] // Un GUID único para la interfaz
    public interface IGoogleSheets
    {
        string Saludar(string name);
    }

    //// Implementar la interfaz IGoogleDrive en la clase GoogleDrive
    [ComVisible(true)]
    [Guid("4ef2c748-12c8-4101-ac68-3a1b2be23412")] // Un GUID único para la clase
    [ClassInterface(ClassInterfaceType.None)] // Evita que se genere automáticamente una interfaz COM
    public class GoogleSheets : IGoogleSheets
    {
        public string Saludar(string name)
        {
            return $"Hola {name}";
        }
    }
}
