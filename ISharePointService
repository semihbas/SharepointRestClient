using System.IO;
using System.Threading.Tasks;

namespaceInterfaces
{
    public interface ISharePointService
    {
        Task<SharepointUploadResponse> CreateLibraryAndUploadWithMetadata(string sharePointAccessToken, string siteUrl, string folder, string fileName, string checkinComment, MemoryStream memoryStream);

        Task<SharepointUploadResponse> UploadFileAsync(string sharePointAccessToken, string siteUrl, string folder, string fileName, MemoryStream memoryStream);

        Task<SharepointUploadResponse> CreateSpFieldAsync(string sharePointAccessToken, string siteUrl, string libraryTitle, SharePointFieldRequest fieldRequest);

        Task<SharepointUploadResponse> CreateSpLibraryAsync(string sharePointAccessToken, string siteUrl, string folder);

        Task<SharepointUploadResponse> UpdateSpFieldAsync(string sharePointAccessToken, string siteUrl, string folder, string uniqueId, dynamic fieldUpdateRequest);

        Task<SharepointUploadResponse> CheckinCommentAsync(string sharePointAccessToken, string siteUrl, string folder, string fileName, string checkinComment);

        Task<SharepointUploadResponse> CheckoutFileAsync(string sharePointAccessToken, string siteUrl, string folder, string fileName);
    }
}
