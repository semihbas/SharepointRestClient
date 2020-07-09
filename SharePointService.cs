using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using RestSharp;
using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace Services
{
    public class SharePointService : ISharePointService
    {
        private readonly IOptions<AppSettings> _settings;
        private readonly IRestClient _restClient;
        private readonly ILogger<SharePointService> _logger;

        public SharePointService(IOptions<AppSettings> settings, IRestClient restClient, ILogger<SharePointService> logger)
        {
            _settings = settings;
            _restClient = restClient;
            _logger = logger;
        }

        public async Task<SharepointUploadResponse> CreateLibraryAndUploadWithMetadata(string sharePointAccessToken, string siteUrl, string folder, string fileName, string checkinComment, MemoryStream memoryStream)
        {
            var createSpLibraryRepsonse = await CreateSpLibraryAsync(sharePointAccessToken, siteUrl, folder);

            //var fieldRequest = new SharePointFieldRequest()
            //{
            //    Hidden = false,
            //    Description = "Account Code",
            //    InternalName = "AccountCode",
            //    FieldTypeKind = 2,
            //    Title = "AccountCode",
            //    __metadata = new Metadata()
            //    {
            //        type = "SP.Field"
            //    }
            //};

            //var createSpFieldRepsonse = await CreateSpFieldAsync(sharePointAccessToken, siteUrl, folder, fieldRequest);

            var uploadFileResponse = await UploadFileAsync(sharePointAccessToken, siteUrl, folder, fileName, memoryStream);

            var checkOutRepsonse = await CheckoutFileAsync(sharePointAccessToken, siteUrl, folder, fileName);

            var checkinCommentRepsonse = await CheckinCommentAsync(sharePointAccessToken, siteUrl, folder, fileName, checkinComment);

            //var fieldUpdateRequest = new
            //{
            //    AccountCode = "acc 1234",
            //    __metadata = new
            //    {
            //        type = uploadFileResponse.d.__metadata.type
            //    }
            //};

            //var updateSpFieldRepsonse = await UpdateSpFieldAsync(sharePointAccessToken, siteUrl, folder, uploadFileResponse.d.UniqueId, fieldUpdateRequest);

            return uploadFileResponse;
        }

        public async Task<SharepointUploadResponse> UploadFileAsync(string sharePointAccessToken, string siteUrl, string folder, string fileName, MemoryStream memoryStream)
        {
            var endpointUrl = $"{siteUrl}/_api/web/getfolderbyserverrelativeurl(\'{folder}\')/Files/add(url=\'{fileName}\',overwrite=true)";

            _restClient.BaseUrl = new Uri(siteUrl);

            var bytes = memoryStream.ToArray();

            var request = new RestRequest(endpointUrl, Method.POST);
            request.AddHeader("Accept", "application/json;odata=verbose");
            request.AddHeader("content-length", bytes.Length.ToString());
            request.AddHeader("Authorization", "Bearer " + sharePointAccessToken);
            request.AddFile(fileName, bytes, fileName);

            var response = await _restClient.ExecuteAsync<SharepointUploadResponse>(request);

            if (!response.IsSuccessful)
            {
                var ex = new Exception($"UploadFileAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription}");
                _logger.LogError(ex, $"UploadFileAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription} ");
                throw ex;
            }

            return response.Data;
        }

        public async Task<SharepointUploadResponse> CreateSpFieldAsync(string sharePointAccessToken, string siteUrl, string libraryTitle, SharePointFieldRequest fieldRequest)
        {
            var endpointUrl = $"{siteUrl}/_api/web/lists/getByTitle('{libraryTitle}')/fields";

            _restClient.BaseUrl = new Uri(siteUrl);

            var request = new RestRequest(endpointUrl, Method.POST);
            request.AddHeader("Accept", "application/json;odata=verbose");
            request.AddHeader("Content-Type", "application/json;odata=verbose");
            request.AddHeader("Authorization", "Bearer " + sharePointAccessToken);

            request.AddJsonBody(fieldRequest);

            var response = await _restClient.ExecuteAsync<SharepointUploadResponse>(request);

            if (!response.IsSuccessful)
            {
                var ex = new Exception($"CreateSpFieldAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription}");
                _logger.LogError(ex, $"CreateSpFieldAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription} ");
                throw ex;
            }

            return response.Data;
        }

        public async Task<SharepointUploadResponse> CreateSpLibraryAsync(string sharePointAccessToken, string siteUrl, string folder)
        {
            var endpointUrl = $"{siteUrl}/_api/web/lists/getbytitle(\'{folder}\')";

            _restClient.BaseUrl = new Uri(siteUrl);

            var request = new RestRequest(endpointUrl, Method.GET);
            request.AddHeader("Authorization", "Bearer " + sharePointAccessToken);

            var isFolderExistsResponse = await _restClient.ExecuteAsync<SharepointUploadResponse>(request);

            if (isFolderExistsResponse.IsSuccessful)
            {
                return isFolderExistsResponse.Data;
            }
            
            if (isFolderExistsResponse.StatusCode != HttpStatusCode.NotFound)
            {
                var ex = new Exception($"CreateSpLibraryAsync has error in checking that folder exists section! {isFolderExistsResponse.ErrorMessage}-{isFolderExistsResponse.StatusCode}-{isFolderExistsResponse.StatusDescription}");
                _logger.LogError(ex, $"CreateSpLibraryAsync has error in checking that folder exists section! {isFolderExistsResponse.ErrorMessage}-{isFolderExistsResponse.StatusCode}-{isFolderExistsResponse.StatusDescription} ");
                throw ex;
            }

            var postEndpointUrl = $"{siteUrl}/_api/web/lists";

            var fieldUpdateRequest = new
            {
                AllowContentTypes = "true",
                BaseTemplate = "101",
                ContentTypesEnabled = "true",
                Description = "Folder auto created by New Client Reporting",
                Title = folder,
                __metadata = new
                {
                    type = "SP.List"
                }
            };

            var postRequest = new RestRequest(postEndpointUrl, Method.POST);
            postRequest.AddHeader("Accept", "application/json;odata=verbose");
            postRequest.AddHeader("Content-Type", "application/json;odata=verbose");
            postRequest.AddHeader("Authorization", "Bearer " + sharePointAccessToken);

            postRequest.AddJsonBody(fieldUpdateRequest);

            var postResponse = await _restClient.ExecuteAsync<SharepointUploadResponse>(postRequest);

            if (!postResponse.IsSuccessful)
            {
                var ex = new Exception($"CreateSpLibraryAsync has error in folder creation section!  {postResponse.ErrorMessage}-{postResponse.StatusCode}-{postResponse.StatusDescription}");
                _logger.LogError(ex, $"CreateSpLibraryAsync has error in folder creation section!  {postResponse.ErrorMessage}-{postResponse.StatusCode}-{postResponse.StatusDescription} ");
                throw ex;
            }

            return postResponse.Data;
        }

        public async Task<SharepointUploadResponse> UpdateSpFieldAsync(string sharePointAccessToken, string siteUrl, string folder, string uniqueId, dynamic fieldUpdateRequest)
        {
            var endpointUrl = $"{siteUrl}/_api/web/Lists/GetByTitle('{folder}')/Items('{uniqueId}'))";

            _restClient.BaseUrl = new Uri(siteUrl);

            var request = new RestRequest(endpointUrl, Method.MERGE);
            request.AddHeader("Accept", "application/json;odata=verbose");
            request.AddHeader("Content-Type", "application/json;odata=verbose");
            request.AddHeader("Authorization", "Bearer " + sharePointAccessToken);
            request.AddHeader("IF-MATCH", "*");

            request.AddJsonBody(fieldUpdateRequest);

            var response = await _restClient.ExecuteAsync<SharepointUploadResponse>(request);

            if (!response.IsSuccessful)
            {
                var ex = new Exception($"UpdateSpFieldAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription}");
                _logger.LogError(ex, $"UpdateSpFieldAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription} ");
                throw ex;
            }

            return response.Data;
        }

        public async Task<SharepointUploadResponse> CheckinCommentAsync(string sharePointAccessToken, string siteUrl, string folder, string fileName, string checkinComment)
        {
            var endpointUrl = $"{siteUrl}/_api/web/getFileByServerRelativeUrl(\'/sites/Conectus_PreProd/{folder + "/" + fileName}\')/CheckIn(comment='{checkinComment}',checkintype=0)";

            _restClient.BaseUrl = new Uri(siteUrl);

            var request = new RestRequest(endpointUrl, Method.POST);
            request.AddHeader("Accept", "application/json;odata=verbose");
            request.AddHeader("Content-Type", "application/json;odata=verbose");
            request.AddHeader("IF-MATCH", "*");
            request.AddHeader("Authorization", "Bearer " + sharePointAccessToken);

            var response = await _restClient.ExecuteAsync<SharepointUploadResponse>(request);

            if (!response.IsSuccessful)
            {
                var ex = new Exception($"CheckinCommentAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription}");
                _logger.LogError(ex, $"CheckinCommentAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription} ");
                throw ex;
            }

            return response.Data;
        }

        public async Task<SharepointUploadResponse> CheckoutFileAsync(string sharePointAccessToken, string siteUrl, string folder, string fileName)
        {
            var endpointUrl = $"{siteUrl}/_api/web/getFileByServerRelativeUrl(\'/sites/Conectus_PreProd/{folder + "/" + fileName}\')/CheckOut()";

            _restClient.BaseUrl = new Uri(siteUrl);

            var request = new RestRequest(endpointUrl, Method.POST);
            request.AddHeader("Accept", "application/json;odata=verbose");
            request.AddHeader("Authorization", "Bearer " + sharePointAccessToken);

            var response = await _restClient.ExecuteAsync<SharepointUploadResponse>(request);

            if (!response.IsSuccessful)
            {
                var ex = new Exception($"CheckoutFileAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription}");
                _logger.LogError(ex, $"CheckoutFileAsync has error!  {response.ErrorMessage}-{response.StatusCode}-{response.StatusDescription} ");
                throw ex;
            }

            return response.Data;
        }
    }
}
