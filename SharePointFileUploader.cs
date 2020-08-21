using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace gvaduha.Sharepoint
{
	/// <summary>
	/// Handle share point file operations
	/// </summary>
    public class SharePointFileUploader : IDisposable
    {
		/// <summary>
		/// This request returns authentication cookie. Ad hoc error handling.
		/// check wsdl @ severUri//_vti_bin/authentication.asmx
		/// </summary>
		/// <returns>Authentication cookie</returns>
		public static Cookie Authenticate(string serverUri, string userName, string password)
		{
			const string envelopeFormat = "<?xml version='1.0' encoding='utf-8'?>"
				+ "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
				+ "<soap:Body><Login xmlns='http://schemas.microsoft.com/sharepoint/soap/'>"
				+ "<username>{0}</username><password>{1}</password>"
				+ "</Login></soap:Body></soap:Envelope>";

			var authServiceUri = new Uri($"{serverUri}/_vti_bin/authentication.asmx");
			var req = HttpWebRequest.Create(authServiceUri) as HttpWebRequest;
			req.CookieContainer = new CookieContainer();
			req.Headers["SOAPAction"] = "http://schemas.microsoft.com/sharepoint/soap/Login";
			req.ContentType = "text/xml; charset=utf-8";
			req.Method = WebRequestMethods.Http.Post;
			using (var sw = new StreamWriter(req.GetRequestStream()))
				sw.Write(string.Format(envelopeFormat, userName, System.Security.SecurityElement.Escape(password)));
			using (var response = req.GetResponse() as HttpWebResponse)
			{
				if (response.Cookies.Count() == 0)
					throw new ApplicationException("authentication failed");

				return response.Cookies[0];
			}
		}

		/// <summary>
		/// Return form digest that is essential for share point operation
		/// https://stackoverflow.com/questions/22159609/how-to-get-request-digest-value-from-provider-hosted-app
		/// </summary>
		protected string GetFormDigest()
        {
            //_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
            var resp = _webClient.UploadData(new Uri($"{_serverRootUri}/_api/ContextInfo"), new byte[0]);
			JObject retObj = (JObject) JsonConvert.DeserializeObject(Encoding.UTF8.GetString(resp));
			return retObj.GetValue("FormDigestValue").ToString();
        }

		/// <summary>
		/// Struct for credentials
		/// </summary>
		public struct BasicCredentials
        {
			public readonly string UserName;
			public readonly string Password;
			public BasicCredentials(string userName, string password)
            {
				UserName = userName;
				Password = password;
            }
		}

		readonly string _serverRootUri;		// the "starting" part of server uri: https://busysrv/shpnt
		readonly string _serverFolderPath;  // path to actual document folder: /shared documents/myfolder
		readonly int _chunkSize;			// share point has 1M limit for file upload chunk, so files above this limit should be chunked
		readonly WebClient _webClient;		// engine for uploading files (DO NOT WORK MULTITHREADED ASYNC)
		readonly Cookie _authCookie;        // (only for forking class instances!) authorization cookie (now FedAuth)
		readonly string _formDigest;        // (only for forking class instances!) form digest for file operations

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="serverRootUri">the "starting" part of server uri: https://busysrv/shpnt</param>
		/// <param name="serverFolderPath">path to actual document folder: /shared documents/myfolder</param>
		/// <param name="credentials">Sharepoint user credentials</param>
		/// <param name="chunkSize">Sharepoint limit for upload (default 1M)</param>
		public SharePointFileUploader(string serverRootUri, string serverFolderPath, BasicCredentials credentials, int chunkSize = 1024*1024)
        {
			_serverRootUri = serverRootUri;
			_serverFolderPath = serverFolderPath;
			_chunkSize = chunkSize;

			_authCookie = Authenticate(serverRootUri, credentials.UserName, credentials.Password);
			_webClient = new WebClient();
			_webClient.Headers.Add(HttpRequestHeader.Cookie, $"{_authCookie.Name}={_authCookie.Value}");
			_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);

			_formDigest = GetFormDigest();
			_webClient.Headers.Add("x-requestdigest", _formDigest);
		}

		/// <summary>
		/// Copy constructor that creates new aggregated WebClient object.
		/// Fork call it to provide support for multithreaded operation
		/// </summary>
		/// <param name="proto">Object prototype</param>
		private SharePointFileUploader(SharePointFileUploader proto)
		{
			_serverRootUri = proto._serverRootUri;
			_serverFolderPath = proto._serverFolderPath;
			_chunkSize = proto._chunkSize;
			_authCookie = proto._authCookie;
			_formDigest = proto._formDigest;

			_webClient = new WebClient();
			_webClient.Headers.Add(HttpRequestHeader.Cookie, $"{_authCookie.Name}={_authCookie.Value}");
			_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
			_webClient.Headers.Add("x-requestdigest", _formDigest);
		}

		/// <summary>
		/// WebClient isn't support for async operations, so Fork return new instance of class with new WebClient
		/// </summary>
		/// <returns>SharePointFileUploader object with new aggregated WebClient object</returns>
		public SharePointFileUploader Fork() => new SharePointFileUploader(this);

		/// <summary>
		/// Upload files to the server
		/// </summary>
		/// <param name="filePath">local file path</param>
		/// <returns>WebClient response</returns>
		public Task<byte[]> UploadAsync(string filePath)
		{
			var fileLen = new FileInfo(filePath).Length;
			return (fileLen > _chunkSize)
				? UploadChunkedAsync(filePath, fileLen)
				: UploadMonolithAsync(filePath);
		}

		protected Task<byte[]> UploadMonolithAsync(string filePath)
        {
			return UploadImpl(GetUploadUri(filePath), File.ReadAllBytes(filePath));
		}

		/// <summary>
		/// Ad hoc regex for form digest extraction (see GetFormDigest). Otherwise we have to implement heavy soap parsing support
		/// </summary>
		static readonly Regex ServerRelativeUrlRegex = new Regex("<d:ServerRelativeUrl>([^<>\"']+)<\\/d:ServerRelativeUrl>");

		protected async Task<byte[]> UploadChunkedAsync(string filePath, long fileLen)
        {
			string serverRelativeUrl;
			{
				var resp = _webClient.UploadData(GetUploadUri(filePath), new byte[0]);
				JObject retObj = (JObject)JsonConvert.DeserializeObject(Encoding.UTF8.GetString(resp));

				serverRelativeUrl = retObj.GetValue("ServerRelativeUrl").ToString();
			}

			var uploadGuid = Guid.NewGuid();
			bool lastChunk = false;
			long currentOffset = 0L;
			var buff = new byte[_chunkSize];
			byte[] result = null;

			using (FileStream fs = File.OpenRead(filePath))
			{
				while (!lastChunk)
				{
					var readCnt = fs.Read(buff, 0, _chunkSize);
					lastChunk = readCnt < _chunkSize;

					if (lastChunk)
						Array.Resize(ref buff, readCnt);

					var chunkedUploadUri = GetChunkedUploadUri(serverRelativeUrl, uploadGuid, currentOffset, lastChunk);
					result = await UploadImpl(chunkedUploadUri, buff);
					currentOffset += readCnt;
				}
			}

			return result;
		}

		protected Task<byte[]> UploadImpl(Uri uri, byte[] data, int retries = 3)
		{
			{
				try
				{
					var resp = _webClient.UploadDataTaskAsync(uri, data);
					return resp;
				}
				catch (Exception)
				{
					if (--retries == 0) throw;
				}
			}
			while (true);
		}

		protected Uri GetUploadUri(string filePath) =>
			new Uri($"{_serverRootUri}/_api/web/getfolderbyserverrelativeurl('{_serverFolderPath}')/files/add(url='{Path.GetFileName(filePath)}',overwrite=true)");

		protected Uri GetChunkedUploadUri(string relativeUrl, Guid uploadGuid, long currentOffset, bool lastChunk)
		{
			string prefix = $"{_serverRootUri}/_api/Web/GetFileByServerRelativePath(decodedurl='{relativeUrl}')";

			if (0L == currentOffset)
				return new Uri($"{prefix}/StartUpload(uploadId=guid'{uploadGuid}')");

			if (lastChunk)
				return new Uri($"{prefix}/FinishUpload(uploadId=guid'{uploadGuid}',fileOffset={currentOffset})");

			return new Uri($"{prefix}/ContinueUpload(uploadId=guid'{uploadGuid}',fileOffset={currentOffset})");
		}

        public void Dispose()
        {
			_webClient.Dispose();
        }
    }
}
