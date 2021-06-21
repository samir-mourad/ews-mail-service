using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace MailService
{
    public interface IMailService
    {
        Task<string> SendMailAsync(string from, string subject, string body, string[] to);
        Task<string> SendMailAsync(string from, string subject, string body, string[] to, Dictionary<string, byte[]> attachments = null, string[] cc = null, string[] cco = null);
        Task<string> SendMailAsync(string from, string subject, string body, string[] to, string[] pathAttachments = null, string[] cc = null, string[] cco = null);
        Task<IEnumerable<MailMessage>> GetMessagesAsync(string mailbox, string[] folder, DateTime receivedDate, string[] subjectContains, string pathAttachments);
        Task<IEnumerable<MailMessage>> GetMessagesAsync(string mailbox, string from, DateTime receivedDate, string patchAttachment, bool emailUnread = false, bool ignoreFrom = false);
        Task<string> GetFolderIdAsync(string mailbox, params string[] folder);
        Task MarkMessageAsReadAsync(string mailbox, IEnumerable<string> uniqueIds);
        Task MoveMessageFolderAsync(string folderId, IEnumerable<string> uniqueId);


    }
}
