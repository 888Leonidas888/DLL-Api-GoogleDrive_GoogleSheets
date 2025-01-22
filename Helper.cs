using System;
using System.Web;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using System.Net.Http;

namespace GoogleServices
{
    public enum ERR_UTIL
    {
        READ_FAILD = 6100,
        WRITE_FAILD = 6101,
        INTERPOLATE_FAILED = 6102
    }
    public static class Helper
    {
        /// <summary>
        /// Codifica una cadena según las reglas estándar de codificación de URL.
        /// </summary>
        /// <param name="str">La cadena que se desea codificar.</param>
        /// <returns>Cadena codificada.</returns>
        public static string URLEncode(String str)
        {
            return HttpUtility.UrlEncode(str);
        }

        /// <summary>
        /// Genera una cadena aleatoria de longitud específica utilizando solo letras mayúsculas y minúsculas.
        /// </summary>
        /// <param name="length">Longitud de la cadena generada.</param>
        /// <returns>Cadena aleatoria generada.</returns>
        public static string GenerateString(int length = 8)
        {
            var random = new Random();
            var randomString = new char[length];

            // Generar una cadena aleatoria usando códigos ASCII
            for (int i = 0; i < length; i++)
            {                
                int randomCharCode = random.Next(65, 123);       
                if (randomCharCode > 90 && randomCharCode < 97)
                {
                    randomCharCode = random.Next(97, 123);
                }
                randomString[i] = (char)randomCharCode;
            }

            return new string(randomString);
        }

        /// <summary>
        /// Escribe contenido en un archivo especificado.
        /// </summary>
        /// <param name="content">Contenido a escribir en el archivo.</param>
        /// <param name="pathTarget">Ruta donde se creará el archivo.</param>
        /// <returns>Retorna true si se escribe el archivo correctamente, false en caso de error.</returns>
        public static async Task<bool> WriteFileAsync(string content, string pathTarget)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(pathTarget, false))
                {
                    await writer.WriteAsync(content);
                }
                return true; 
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Convierte el contenido de un archivo en un arreglo de bytes.
        /// </summary>
        /// <param name="pathFile">La ruta completa del archivo.</param>
        /// <returns>Un arreglo de bytes que representa el contenido del archivo.</returns>
        public static byte[] SourceToBinary(string pathFile)
        {

            try
            {
                return System.IO.File.ReadAllBytes(pathFile);
            }
            catch
            {
                throw new Exception("No se pudo convertir a binario.");
            }
        }

        /// <summary>
        /// Lee el contenido de un archivo de texto especificado.
        /// </summary>
        /// <param name="pathResource">Ruta completa del archivo a leer.</param>
        /// <returns>Contenido del archivo como una cadena de texto.</returns>
        public static string ReadFile(string pathResource)
        {
            try
            {
                if (File.Exists(pathResource))
                {
                    return File.ReadAllText(pathResource);
                }
                else
                {
                    throw new FileNotFoundException("Recurso no encontrado.");
                }
            }
            catch
            {
                throw new Exception("Ocurrio un error al leer el recurso.");
            }
        }

        /// <summary>
        /// Convierte un arreglo de bytes a una cadena codificada en Base64.
        /// </summary>
        /// <param name="binaryData">Arreglo de bytes a convertir.</param>
        /// <returns>Cadena codificada en Base64.</returns>
        public static string BinaryTobase64(byte[] binaryData)
        {
            try
            {

                if (binaryData == null || binaryData.Length == 0)
                {
                    throw new Exception("El arreglo de bytes esta vacío o es nulo");
                }

                return Convert.ToBase64String(binaryData);
            }
            catch
            {
                throw new Exception("No se puede convertir a base64.");

            }
        }

        /// <summary>
        /// Convierte un recurso a una cadena codificada en Base64.
        /// </summary>
        /// <param name="resource">Recurso a convertir.</param>
        /// <returns>Cadena codificada en Base64.</returns>
        public static string SourceToBase64(string resource)
        {
            byte[] bytes = SourceToBinary(resource);
            string base64 = BinaryTobase64(bytes);

            return base64;
        }

        /// <summary>
        /// Crea una parte relacionada para la solicitud HTTP multipart, serializando un diccionario a formato JSON y adjuntando un archivo en base64.
        /// </summary>
        /// <param name="filePath">La ruta completa del archivo que se convertirá a base64 y se incluirá en la solicitud.</param>
        /// <param name="boundary">El delimitador utilizado para separar las partes de la solicitud multipart.</param>
        /// <param name="fileObjectJson">El objeto en formato JSON como string que contiene los datos adicionales a incluir en la solicitud, como el tipo MIME y otros atributos.</param>
        /// <returns>Una cadena que representa el cuerpo de la parte relacionada en formato multipart.</returns>
        /// <exception cref="ArgumentNullException">Lanzado cuando el parámetro <paramref name="fileObjectJson"/> es nulo o vacío.</exception>
        /// <exception cref="ArgumentException">Lanzado cuando el parámetro <paramref name="fileObjectJson"/> no es un JSON válido o no puede deserializarse correctamente en un diccionario.</exception>
        public static string CreatePartRelated(string filePath, string boundary, string fileObjectJson)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Archivo no encontrado: {filePath}");
            }

            if (string.IsNullOrEmpty(fileObjectJson))
            {
                throw new ArgumentNullException(nameof(fileObjectJson), "El parámetro fileObjectJson no puede ser nulo o vacío.");
            }

            try
            {
                var convertedDict = JsonConvert.DeserializeObject<Dictionary<string, object>>(fileObjectJson);

                string fileName = Path.GetFileName(filePath);
                convertedDict["name"] = fileName;

                string metada = JsonConvert.SerializeObject(convertedDict);
                string mimeType = convertedDict.ContainsKey("mimeType") ? convertedDict["mimeType"].ToString() : "";
                string base64 = Helper.SourceToBase64(filePath);
                string startBoundary = $"--{boundary}";
                string finishBoundary = $"{startBoundary}--";
                string strTmp = "{0}{1}Content-Type: application/json; charset=UTF-8{1}{1}{2}{1}{0}{1}";
                strTmp += "Content-Type: {3}{1}Content-Transfer-Encoding: base64{1}{1}{4}{1}{5}";

                string related = string.Format(strTmp, startBoundary, "\r\n", metada, mimeType, base64, finishBoundary);

                return related;
            }
            catch (Exception ex)
            {
                throw new ArgumentException(ex.Message);
            }
        }
        public  static HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            try
            {
                using (var client = new HttpClient())
                {                
                    HttpMethod httpMethod = new HttpMethod(method);
                    var requestMessage = new HttpRequestMessage(httpMethod, url);
                             
                    if (body != null)
                    {
                        if (body is string)
                        {
                            requestMessage.Content = new StringContent((string)body);                            
                        }
                        else if (body is byte[])
                        {
                            requestMessage.Content = new ByteArrayContent((byte[])body);
                        }
                    }
              
                    if (headers != null)
                    {
                        foreach (var header in headers)
                        {               
                            requestMessage.Headers.TryAddWithoutValidation(header.ToString(),headers.get_Item(header));
                        }
                    }               

                    return client.SendAsync(requestMessage).Result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Request: {ex.Message}");
            }
        }      
    }
}
