using System;
using System.Runtime.InteropServices;
using System.Net.Http;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;

namespace GoogleServices
{
    [ComVisible(true)]
    [Guid("39043b4c-7128-4dc8-995b-97a5e8a34a4a")] // Un GUID único para la interfaz
    public interface IGoogleDrive
    {
        /// <summary>
        /// Realiza una operación genérica.
        /// </summary>
        /// <returns>Un entero que representa el resultado de la operación.</returns>
        int Operation();

        /// <summary>
        /// Conecta el servicio utilizando el flujo OAuth.
        /// </summary>
        /// <param name="oFlowOauth">El objeto de flujo OAuth.</param>
        void ConnectionService(object oFlowOauth);

        /// <summary>
        /// Copia un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo a copiar.</param>
        /// <param name="fileObject">El objeto del archivo.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>El ID del nuevo archivo copiado.</returns>
        string Copy(string fileID, string fileObject, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Elimina un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo a eliminar.</param>
        /// <returns>Una cadena que representa el resultado de la operación.</returns>
        string Delete(string fileID);

        /// <summary>
        /// Vacía la papelera de Google Drive.
        /// </summary>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa el resultado de la operación.</returns>
        string EmptyTrash(Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Exporta un archivo de Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo a exportar.</param>
        /// <param name="queryParameters">Parámetros de consulta.</param>
        /// <param name="pathFile">La ruta del archivo de destino.</param>
        /// <returns>Un valor booleano que indica si la exportación fue exitosa.</returns>
        bool Export(string fileID, Scripting.Dictionary queryParameters, string pathFile);

        /// <summary>
        /// Genera IDs para nuevos archivos en Google Drive.
        /// </summary>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa los IDs generados.</returns>
        string GenerateIds(Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Obtiene los metadatos de un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa los metadatos del archivo.</returns>
        string GetMetadata(string fileID, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Lista los archivos en Google Drive.
        /// </summary>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa la lista de archivos.</returns>
        string List(Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Lista las etiquetas de un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa la lista de etiquetas.</returns>
        string ListLabels(string fileID, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Actualiza un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="fileObject">El objeto del archivo opcional.</param>
        /// <param name="queryParameters">Parámetros de consulta opcionales.</param>
        /// <returns>Una cadena que representa el resultado de la operación.</returns>
        string Update(string fileID, string fileObject = null, Scripting.Dictionary queryParameters = null);

        /// <summary>
        /// Descarga el contenido de un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <returns>Un valor booleano que indica si la descarga fue exitosa.</returns>
        bool DownLoadContentLink(string fileID);

        /// <summary>
        /// Descarga un archivo en Google Drive.
        /// </summary>
        /// <param name="fileID">El ID del archivo.</param>
        /// <param name="directory">El directorio de destino.</param>
        /// <returns>Un valor booleano que indica si la descarga fue exitosa.</returns>
        bool DownLoad(string fileID, string directory);

        /// <summary>
        /// Sube un archivo multimedia a Google Drive.
        /// </summary>
        /// <param name="pathFile">La ruta del archivo.</param>
        /// <param name="mimeType">El tipo MIME del archivo.</param>
        /// <returns>El ID del archivo subido.</returns>
        string UploadMedia(string pathFile = null, string mimeType = "application/octet-stream");

        /// <summary>
        /// Sube un archivo multipart a Google Drive.
        /// </summary>
        /// <param name="pathFile">La ruta del archivo.</param>
        /// <param name="fileObject">El objeto del archivo.</param>
        /// <returns>El ID del archivo subido.</returns>
        string UploadMultipart(string pathFile, string fileObject);
    }

    //// Implementar la interfaz IGoogleDrive en la clase GoogleDrive
    [ComVisible(true)]
    [Guid("7b7bfe28-2585-4411-bf47-9ad67acddc91")] // Un GUID único para la clase
    [ClassInterface(ClassInterfaceType.None)] // Evita que se genere automáticamente una interfaz COM
    public class GoogleDrive : IGoogleDrive
    {
        private const string SERVICE_END_POINT = "https://www.googleapis.com/drive/v3/files";
        private const string SERVICE_END_POINT_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files";

        private string _apiKey;
        private string _accessToken;
        private int _status;
        private string _url;

        public int Operation()
        {
            return _status;
        }
        public void ConnectionService(object oFlowOauth)
        {
            dynamic flow = oFlowOauth;

            _apiKey = flow.GetApiKey();
            _accessToken = flow.GetTokenAccess();
        }
        private string CreateQueryParameters(string pathParameters = null, Scripting.Dictionary queryParameters = null, bool endPointUpload = false)
        {
            string endPoint = endPointUpload ? SERVICE_END_POINT_UPLOAD : SERVICE_END_POINT;
            string queryString = "?";

            try
            {
                if (queryParameters != null)
                {
                    if (queryParameters is Scripting.Dictionary dict)
                    {
                        foreach (var k in dict)
                        {
                            var v = Helper.URLEncode(dict.get_Item(k));
                            queryString += $"{k}={v}&";
                        }
                    }
                    else
                    {
                        throw new Exception($"QueryParameters not is dicctionary.");
                    }
                }

                endPoint += pathParameters;
                //queryString += $"key={_apiKey} HTTP/1.1";
                queryString += $"key={_apiKey}";

                return $"{endPoint}{queryString}";
            }
            catch (Exception ex)
            {
                throw new Exception($"Error create query parameters: {ex.Message}");
            }
        }
        private HttpResponseMessage Request(string method, string url, object body = null, Scripting.Dictionary headers = null)
        {
            if (headers == null)
            {
                headers = new Scripting.Dictionary();
            }

            headers.set_Item("Authorization", $"Bearer {_accessToken}");
            headers.set_Item("Accept", "application/json");

            var response = Helper.Request(method, url, body, headers);
            _status = (int)response.StatusCode;
            return response;
        }

        private string GetNameFileForId(string fileID)
        {
            var queryParameters = new Scripting.Dictionary();
            queryParameters.Add("fields", "name");

            string json = GetMetadata(fileID, queryParameters);

            if (_status != 200)
            {
                throw new Exception("No response data");
            }

            var fileObject = JObject.Parse(json);
            string nameFile = fileObject.ContainsKey("name")
                ? fileObject["name"].ToString()
                : "unknown";

            return nameFile;
        }
        public string Copy(string fileID, string fileObject, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}/copy";
                string url = null;

                var headers = new Scripting.Dictionary();
                headers.Add("Content-Type", "application/json");

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("POST", url, fileObject, headers);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Copy {fileID} : {ex.Message}");
            }
        }
        public string Delete(string fileID)
        {
            try
            {
                string pathParameters = $"/{fileID}";
                string url = null;

                url = CreateQueryParameters(pathParameters);

                var response = Request("DELETE", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Delete {fileID} : {ex.Message}");
            }
        }
        public string EmptyTrash(Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = "/trash";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("DELETE", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error EmptyTrash : {ex.Message}");
            }
        }
        public bool Export(string fileID, Scripting.Dictionary queryParameters, string pathFile)
        {
            try
            {
                string pathParameters = $"/{fileID}/export";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var buffer = Request("GET", url).Content.ReadAsByteArrayAsync();

                buffer.Wait();

                if (buffer.Result != null)
                {
                    System.IO.File.WriteAllBytes(pathFile, buffer.Result);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Export {fileID} : {ex.Message}");
            }
        }
        public string GenerateIds(Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = "/generateIds";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Generate Ids : {ex.Message}");
            }
        }
        public string GetMetadata(string fileID, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error get metadata of {fileID} : {ex.Message}");
            }
        }
        public string List(Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string url = null;

                url = CreateQueryParameters(queryParameters: queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error list : {ex.Message}");
            }
        }
        public string ListLabels(string fileID, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}/listLabels";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("GET", url);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error list labels {fileID} : {ex.Message}");
            }
        }
        public string Update(string fileID, string fileObject = null, Scripting.Dictionary queryParameters = null)
        {
            try
            {
                string pathParameters = $"/{fileID}";
                string url = null;

                url = CreateQueryParameters(pathParameters, queryParameters);

                var response = Request("PATCH", url, fileObject);
                string data = response.Content.ReadAsStringAsync().Result;
                return data;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error update {fileID} : {ex.Message}");
            }
        }
        public bool DownLoadContentLink(string fileID)
        {
            const string WEB_CONTENT_LINK = "webContentLink";
            var fields = new Scripting.Dictionary();
            string webContentLink = null;
            string json = null;
            try
            {
                //fields["fields"] = WEB_CONTENT_LINK;
                fields.Add("fields", WEB_CONTENT_LINK);
                json = GetMetadata(fileID, fields);

                if (json != null)
                {
                    var fileObject = JObject.Parse(json);

                    if (fileObject[WEB_CONTENT_LINK] != null)
                    {
                        webContentLink = fileObject[WEB_CONTENT_LINK].ToString();
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = webContentLink,
                            UseShellExecute = true,
                        });
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error download for content link {fileID} : {ex.Message}");
            }
        }
        public bool DownLoad(string fileID, string directory)
        {
            if (!System.IO.Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException($"Not found directory {directory}");
            }

            string nameFile = GetNameFileForId(fileID);
            string fullPath = Path.Combine(directory, nameFile);
            var queryParameters = new Scripting.Dictionary();

            queryParameters.Add("alt", "media");
            string pathParameters = $"/{fileID}";

            string url = CreateQueryParameters(pathParameters, queryParameters);
            var buffer = Request("GET", url).Content.ReadAsByteArrayAsync();

            buffer.Wait();
            _url = url;
            if (buffer.Result != null)
            {
                System.IO.File.WriteAllBytes(fullPath, buffer.Result);
                return true;
            }
            return false;
        }
        public string UploadMedia(string pathFile = null, string mimeType = "application/octet-stream")
        {
            bool endPointUpload = false;
            string body = null;
            byte[] buffer = null;
            var headers = new Scripting.Dictionary();
            HttpResponseMessage response = null;

            if (pathFile != null)
            {
                if (System.IO.File.Exists(pathFile))
                {
                    endPointUpload = true;
                    buffer = Helper.SourceToBinary(pathFile);
                    var fileInfo = new FileInfo(pathFile);

                    headers.Add("Content-Type", mimeType);
                    headers.Add("Content-Length", fileInfo.Length.ToString());
                    headers.Add("Content-Transfer-Encoding", "binary");
                }
                else
                {
                    throw new FileNotFoundException($"File not exists {pathFile}");
                }
            }
            else
            {
                var fileObject = new
                {
                    mimeType = "application/vnd.google-apps.folder",
                    parents = "root"
                };

                body = JsonConvert.SerializeObject(fileObject);
            }

            var queryParameters = new Scripting.Dictionary();

            queryParameters.Add("uploadType", "media");

            string url = CreateQueryParameters(queryParameters: queryParameters,
                                                endPointUpload: endPointUpload);

            response = body != null
                ? Request("POST", url, body, headers)
                : Request("POST", url, buffer, headers);

            return response.Content.ReadAsStringAsync().Result;
        }
        public string UploadMultipart(string pathFile, string fileObject)
        {
            try
            {
                string boundary = Helper.GenerateString(15);
                var queryParameters = new Scripting.Dictionary {
                    {"uploadType", "multipart" }
                };

                string url = CreateQueryParameters(queryParameters: queryParameters, endPointUpload: true);

                using (var client = new HttpClient())
                {
                    var requestMessage = new HttpRequestMessage(HttpMethod.Post, url);
                    string body = Helper.CreatePartRelated(pathFile, boundary, fileObject);

                    requestMessage.Headers.TryAddWithoutValidation("Authorization", $"Bearer {_accessToken}");
                    requestMessage.Headers.TryAddWithoutValidation("Accept", "application/json");
                    requestMessage.Content = new StringContent(body, Encoding.UTF8);
                    requestMessage.Content.Headers.ContentType = new MediaTypeHeaderValue("multipart/related");
                    requestMessage.Content.Headers.ContentType.Parameters.Add(new NameValueHeaderValue("boundary", boundary));

                    var response = client.SendAsync(requestMessage).Result;
                    return response.Content.ReadAsStringAsync().Result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error upload {pathFile} : {ex.Message}");
            }
        }
    }
}
