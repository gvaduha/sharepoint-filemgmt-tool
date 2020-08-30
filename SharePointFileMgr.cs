using System;
using System.Collections.Generic;
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
    public class SharePointFileMgr : IDisposable
    {
		/// <summary>
		/// Operation to execute
		/// </summary>
		public enum Operation
        {
			Upload,
			Download,
			Remove,
			List
        }

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
			_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
            var resp = _webClient.UploadData($"{_serverRootUri}/_api/ContextInfo", new byte[0]);
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
		public SharePointFileMgr(string serverRootUri, string serverFolderPath, BasicCredentials credentials, int chunkSize = 1024*1024)
        {
			_serverRootUri = serverRootUri;
			_serverFolderPath = serverFolderPath;
			_chunkSize = chunkSize;

			_authCookie = Authenticate(serverRootUri, credentials.UserName, credentials.Password);
			_webClient = new WebClient();
			_webClient.Headers.Add(HttpRequestHeader.Cookie, $"{_authCookie.Name}={_authCookie.Value}");

			_formDigest = GetFormDigest();
			_webClient.Headers.Add("x-requestdigest", _formDigest);
		}

		/// <summary>
		/// Copy constructor that creates new aggregated WebClient object.
		/// Fork call it to provide support for multithreaded operation
		/// </summary>
		/// <param name="proto">Object prototype</param>
		private SharePointFileMgr(SharePointFileMgr proto)
		{
			_serverRootUri = proto._serverRootUri;
			_serverFolderPath = proto._serverFolderPath;
			_chunkSize = proto._chunkSize;
			_authCookie = proto._authCookie;
			_formDigest = proto._formDigest;

			_webClient = new WebClient();
			_webClient.Headers.Add(HttpRequestHeader.Cookie, $"{_authCookie.Name}={_authCookie.Value}");
			_webClient.Headers.Add("x-requestdigest", _formDigest);
		}

		/// <summary>
		/// Perform sharepoint file operation on set of files
		/// </summary>
		/// <param name="op">operation type</param>
		/// <param name="items">files for upload, directories for list, remote files mask for download and remove</param>
		/// <returns></returns>
		public async Task<string> PerformAsync(Operation op, IEnumerable<string> items)
        {
			// transform items to actual remote files for download and remove
			if (op == Operation.Download || op == Operation.Remove)
			{
				var remotefiles = await ListAsync("");
				var filter = Regex.Escape(items.First()).Replace(@"\*", ".*").Replace(@"\?", ".");
				var regex = new Regex(filter);
				items = remotefiles.ToList().Where(x => regex.IsMatch(x));
			}

			if (items.Count() == 1)
			{
				return await PerformAsync(op, items.First());
			}
			else
			{
				var schedule = items.Select(f => new {engine = Fork(), file = f});
				var uploads = schedule.Select(x => x.engine.PerformAsync(op, x.file));
				var result = await Task.WhenAll(uploads.ToArray());

				return string.Join(Environment.NewLine, result);
			}
        }

        /// <summary>
        /// Perform sharepoint file operation on file
        /// </summary>
        /// <param name="op">operation type</param>
        /// <param name="file">file name</param>
        /// <returns></returns>
        public async Task<string> PerformAsync(Operation op, string file)
        {
			Func<Func<string>, string> xxx = (fx) => fx();

			Func<string, Task<string>> fn;

			switch (op)
            {
				case Operation.Upload:
					fn = async (f) => 
					{
						var r = await UploadAsync(f);
						return Encoding.UTF8.GetString(r);
					};
					break;
				case Operation.Download:
					fn = async (f) => 
					{
						await DownloadAsync(f);
						return $"{f} - Downloaded";
					};
					break;
				case Operation.Remove:
					fn = async (f) => 
					{
						await RemoveAsync(f);
						return $"{f} - Removed";
					};
					break;
				case Operation.List:
					fn = async (f) =>
					{
						var r = await ListAsync(f, true);
						return string.Join(Environment.NewLine, r);
					};
					break;
				default:
					throw new ApplicationException("Unexpected operation");
            }

			Func<Func<string, Task<string>>, Func<string, Task<string>>> exwrap = (fn) =>
			{
				return async (arg) =>
				{
					try
					{
						return await fn(arg);
					}
					catch (Exception e)
					{
						return $"{arg.ToString()}: {e.Message}";
					}
				};
			};

			return await exwrap(fn)(file);
        }

		/// <summary>
		/// WebClient isn't support for async operations, so Fork return new instance of class with new WebClient
		/// </summary>
		/// <returns>SharePointFileMgr object with new aggregated WebClient object</returns>
		public SharePointFileMgr Fork() => new SharePointFileMgr(this);

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
				_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
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

					var uri = GetChunkedUploadUri(serverRelativeUrl, uploadGuid, currentOffset, lastChunk);
					result = await UploadImpl(uri, buff);
					currentOffset += readCnt;
				}
			}

			return result;
		}

		protected Task<byte[]> UploadImpl(string uri, byte[] data, int retries = 3)
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

		protected string GetFullFolderPath(string folder) =>
			$"{_serverRootUri}/_api/web/getfolderbyserverrelativeurl('{folder}')";

		protected string GetFullFilePath(string relativeUrl) =>
			$"{_serverRootUri}/_api/web/getfilebyserverrelativepath(decodedurl='{relativeUrl}')";

		protected string GetUploadUri(string filePath) =>
			$"{GetFullFolderPath(_serverFolderPath)}/files/add(url='{Path.GetFileName(filePath)}',overwrite=true)";

		protected string GetChunkedUploadUri(string relativeUrl, Guid uploadGuid, long currentOffset, bool lastChunk)
		{
			string prefix = GetFullFilePath(relativeUrl);

			if (0L == currentOffset)
				return $"{prefix}/startupload(uploadid=guid'{uploadGuid}')";

			if (lastChunk)
				return $"{prefix}/finishupload(uploadid=guid'{uploadGuid}',fileoffset={currentOffset})";

			return $"{prefix}/continueupload(uploadid=guid'{uploadGuid}',fileoffset={currentOffset})";
		}

		public Task DownloadAsync(string filePath)
        {
			_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
			return _webClient.DownloadFileTaskAsync($"{_serverRootUri}/{_serverFolderPath}/{filePath}", filePath);
		}

		public Task RemoveAsync(string filePath)
        {
			var req = HttpWebRequest.Create($"{_serverRootUri}/{_serverFolderPath}/{filePath}") as HttpWebRequest;
			req.Method = "DELETE";
			req.Headers.Add(HttpRequestHeader.Cookie, $"{_authCookie.Name}={_authCookie.Value}");
			req.Headers.Add("x-requestdigest", _formDigest);
			return req.GetResponseAsync();
        }

		public async Task<IEnumerable<string>> ListAsync(string filePath, bool includeFolders = false)
        {
			var uri = $"{GetFullFolderPath($"{_serverFolderPath}/{filePath}")}/files?$select=name&$orderby=name";

			_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
			var resp = await _webClient.DownloadDataTaskAsync(uri);

			JObject retObj = (JObject) JsonConvert.DeserializeObject(Encoding.UTF8.GetString(resp));

			var files =  retObj["value"].Select(x => x["Name"].ToString());

			if (includeFolders)
            {
				uri = $"{GetFullFolderPath($"{_serverFolderPath}/{filePath}")}/?$expand=folders";
				_webClient.Headers.Add(HttpRequestHeader.Accept, MediaTypeNames.Application.Json);
				resp = await _webClient.DownloadDataTaskAsync(uri);

				retObj = (JObject)JsonConvert.DeserializeObject(Encoding.UTF8.GetString(resp));

				var folders = retObj["Folders"].Select(x => $"[{x["Name"].ToString()}]");

				return folders.Concat(files);
			}

			return files;
		}

        public void Dispose()
        {
			_webClient.Dispose();
        }
    }
}
