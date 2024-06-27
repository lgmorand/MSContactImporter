using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace Microsoft.Internal.MSContactImporter
{
    internal class GraphUtils
    {
        private readonly string _graphAPIEndpoint = "https://graph.microsoft.com/v1.0/";
        private readonly string[] _scopes = new string[] { "user.read", "user.readbasic.all" };

        //Application name MSPhotoDownloader
        private static string ClientId = "3d53d92f-f208-4c15-8638-3d5c69661727";
        //public static PublicClientApplication PublicClientApp = new PublicClientApplication(ClientId);
        IPublicClientApplication PublicClientApp = PublicClientApplicationBuilder.Create(ClientId).Build();


        private AuthenticationResult authResult;

        public async Task SigninAsync()
        {
            authResult = null;

            try
            {
                authResult = await PublicClientApp.AcquireTokenInteractive(scopes: _scopes).ExecuteAsync();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public async Task<string> GetPhotoForUserAsync(MSFTee msftee)
        {
            Logger.LogMessageToConsole($"Getting photo for {msftee.FullName}");
            //Check authentication
            if (authResult == null)
            {   //If not signed in, try to sign a new time
                try
                {
                    await SigninAsync();
                }
                catch (Exception ex)
                {
                    throw;
                }
            }

            var httpClient = new HttpClient();
            HttpResponseMessage response;
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, _graphAPIEndpoint + $"users/{msftee.Email}/photo/$value");

                //Add the token in Authorization header
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                response = httpClient.SendAsync(request).Result; //Not using async/await pattern here because it ends up with threads deadlocks (I do not know where :o) )
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream photoStream = await response.Content.ReadAsStreamAsync();

                    if (photoStream != null)
                    {
                        //Save picture as file
                        if (photoStream.Position != 0)
                            photoStream.Position = 0;

                        //Creates Photo directory if it does not exist
                        string directoryPath = System.Windows.Forms.Application.UserAppDataPath + "\\photos";
                        if (!Directory.Exists(directoryPath))
                            Directory.CreateDirectory(directoryPath);

                        string filePath = $"{directoryPath}\\{msftee.Email}.jpeg";

                        byte[] bytesInStream = new byte[photoStream.Length];
                        await photoStream.ReadAsync(bytesInStream, 0, bytesInStream.Length);
                        using (System.IO.Stream file = System.IO.File.OpenWrite(filePath))
                        {
                            file.Write(bytesInStream, 0, bytesInStream.Length);
                            return filePath;
                        }
                    }
                    else
                        return string.Empty;
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}