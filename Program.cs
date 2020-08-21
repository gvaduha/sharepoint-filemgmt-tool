using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;

namespace gvaduha.Sharepoint
{
	internal class Program
	{
		private static async Task<int> Main(string[] args)
		{
			var switchMappings = new Dictionary<string, string>()
			 {
				 { "-s", "serverRootUri" },
				 { "-f", "serverFolderUri" },
				 { "-u", "userName" },
				 { "-p", "password" },
				 { "--up", "upload" },
				 { "--down", "download" },
				 { "--rm", "remove" },
				 { "--ls", "list" },
			 };

			var builder = new ConfigurationBuilder();
			builder.AddCommandLine(args, switchMappings);
			var config = builder.Build();

			var cfgProvider = config.Providers.First();

			IEnumerable<string> files = Enumerable.Empty<string>();
			var serverRootUri = "";
			var serverFolderUri = "";
			var userName = "";
			var password = "";
			var operation = SharePointFileMgr.Operation.List;

			try
			{
				if (!cfgProvider.TryGet("serverRootUri", out serverRootUri))
					throw new ApplicationException("no server root uri");
				cfgProvider.TryGet("serverFolderUri", out serverFolderUri);
				if (!cfgProvider.TryGet("userName", out userName))
					throw new ApplicationException("userName is not specified");
				if (!cfgProvider.TryGet("password", out password))
					throw new ApplicationException("password is not specified");
				var fileMask = "";
				if (cfgProvider.TryGet("upload", out fileMask))
					operation = SharePointFileMgr.Operation.Upload;
				else if(cfgProvider.TryGet("download", out fileMask))
					operation = SharePointFileMgr.Operation.Download;
				else if(cfgProvider.TryGet("remove", out fileMask))
					operation = SharePointFileMgr.Operation.Upload;
				
				cfgProvider.TryGet("list", out fileMask);

				if (operation != SharePointFileMgr.Operation.List)
                {
					files = Directory.GetDirectories(".", fileMask ?? "*");
					if (files.Count() == 0)
						files = new List<string>{""};
                }
				else
				{
					files = Directory.GetFiles(".", fileMask ?? "*");
					if (files.Count() == 0)
						throw new ApplicationException("empty file set");
				}
			}
			catch (ApplicationException e)
            {
				Console.WriteLine($"Error: {e.Message}{Environment.NewLine}");
				var module = System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName;
				Console.WriteLine($"use:{Environment.NewLine}\t{module} -s serverRootUri [-f serverFolderPath] -u userName -p password -OPERATION [fileMask]");
				Console.WriteLine($"\toperations: -U upload, -D download, -R remove, -L list (list is default if no op specified, but can take a list of directories to ls");
				Console.WriteLine($"\t\t --up upload, --down download, --rm remove, --ls list");
				Console.WriteLine($"List is default, if no op specified, but can take a list of directories to ls. If no filemask specified '*' considered");
				Console.WriteLine($"\tnote: serverFolderPath should(?) be prefixed with 'Shared Documents'");

				return 1;
            }

			try
			{
				var filemgr = new SharePointFileMgr(serverRootUri, serverFolderUri,
															new SharePointFileMgr.BasicCredentials(userName, password));

				if (files.Count() == 1)
                {
					await filemgr.UploadAsync(files.ElementAt(0));
                }
                else
                {
					var schedule = files.Select(f => new {engine = filemgr.Fork(), file = f});
					var uploads = schedule.Select(x => x.engine.UploadAsync(x.file));
					var result = await Task.WhenAll(uploads.ToArray());

					//result.ToList().ForEach(x => Console.WriteLine(Encoding.UTF8.GetString(x)));
				}
			}
			catch (Exception e)
            {
				Console.WriteLine($"Upload failed{Environment.NewLine}{e}");
				return 2;
            }

			return 0;
		}
	}
}
