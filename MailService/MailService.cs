using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace MailService
{
    public sealed class MailService : IMailService
    {
        private const int MAX_EMAILS_RETURN = 1000;
        private readonly ILogger<MailService> _logger;
        private readonly ExchangeService _service;

        public MailService(ILogger<MailService> logger, MailCredential credentials)
        {
            _logger = logger;

            _service = new ExchangeService(ExchangeVersion.Exchange2010);
            _service.TraceEnabled = true;
            _service.Url = new Uri(credentials.Server);
            _service.WebProxy = new WebProxy();

            ServicePointManager.ServerCertificateValidationCallback =
                (sender, certificate, chain, SslPolicyErrors) => true;

            _service.Timeout = 5 * 60 * 1000;
            _service.KeepAlive = true;
        }

        public async Task<string> GetFolderIdAsync(string mailbox, params string[] folder)
        {
            var folderId = new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox);
            return await GetFolderIdAsync(folderId, folder);
        }

        public async Task<IEnumerable<MailMessage>> GetMessagesAsync(string mailbox, string[] folder, DateTime receivedDate, string[] subjectContains, string pathAttachment)
        {
            var folderId = await GetFolderIdAsync(new FolderId(WellKnownFolderName.MsgFolderRoot, mailbox), folder);
            var sharedMailbox = new FolderId(folderId);
            var itemView = new ItemView(MAX_EMAILS_RETURN);

            itemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

            var searches = new SearchFilter.SearchFilterCollection(LogicalOperator.And,
                                                                   new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, receivedDate))
                                                                   {
                                                                       new SearchFilter.IsNotEqualTo(EmailMessageSchema.From, mailbox),
                                                                       new SearchFilter.SearchFilterCollection(LogicalOperator.Or,
                                                                                                               subjectContains.Select(s => new SearchFilter.ContainsSubstring(ItemSchema.Subject,
                                                                                                                                                                              s,
                                                                                                                                                                              ContainmentMode.Substring,
                                                                                                                                                                              ComparisonMode.IgnoreCaseAndNonSpacingCharacters)))
                                                                   };
            return await FindItemsAsync(sharedMailbox, itemView, searches, pathAttachment);
        }

        public async Task<IEnumerable<MailMessage>> GetMessagesAsync(string mailbox, string from, DateTime receivedDate, string pathAttachment, bool emailUnread = false, bool ignoreFrom = false)
        {
            var sharedMailbox = new FolderId(WellKnownFolderName.Inbox, mailbox);
            var itemView = new ItemView(MAX_EMAILS_RETURN);

            itemView.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

            var searches = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, receivedDate));

            if (!ignoreFrom)
                searches.Add(new SearchFilter.SearchFilterCollection(LogicalOperator.Or,
                             new SearchFilter.IsEqualTo(EmailMessageSchema.From, from),
                             new SearchFilter.IsEqualTo(EmailMessageSchema.LastModifiedName, from)));
            else
                searches.Add(new SearchFilter.Not(
                                new SearchFilter.SearchFilterCollection(
                                    LogicalOperator.Or,
                                    new SearchFilter.IsEqualTo(EmailMessageSchema.From, from),
                                    new SearchFilter.IsEqualTo(EmailMessageSchema.LastModifiedName, from))));

            if (emailUnread)
                searches.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, !emailUnread));

            return await FindItemsAsync(sharedMailbox, itemView, searches, pathAttachment);
        }

        public async System.Threading.Tasks.Task MarkMessageAsReadAsync(string mailbox, IEnumerable<string> uniqueIds)
        {
            var folderId = new FolderId(mailbox);
            var itens = uniqueIds.Select(i => new ItemId(i));
            var messages = await _service.BindToItems(itens.ToArray(), new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.IsRead));

            foreach (var response in messages)
            {
                if (response != null && response.Item != null)
                {
                    var message = response.Item as EmailMessage;
                    message.IsRead = true;
                    await message.Update(ConflictResolutionMode.AlwaysOverwrite);
                }
            }

        }

        public async System.Threading.Tasks.Task MoveMessageFolderAsync(string folderId, IEnumerable<string> uniqueId)
            => await _service.MoveItems(uniqueId.Select(i => new ItemId(i)).ToArray(), new FolderId(folderId));

        public async Task<string> SendMailAsync(string from, string subject, string body, string[] to)
        {
            var message = CreateMessage(subject, body, to, from: from);
            return await SendMailAsync(from, message);
        }

        public async Task<string> SendMailAsync(string from, string subject, string body, string[] to, Dictionary<string, byte[]> attachments = null, string[] cc = null, string[] cco = null)
        {
            var message = CreateMessage(subject, body, to, cc, from, cco);

            if (attachments?.Any() ?? false)
                foreach (var attachment in attachments)
                    message.Attachments.AddFileAttachment(attachment.Key, attachment.Value);

            return await SendMailAsync(from, message);
        }

        public async Task<string> SendMailAsync(string from, string subject, string body, string[] to, string[] pathAttachments = null, string[] cc = null, string[] cco = null)
        {
            var message = CreateMessage(subject, body, to, cc, from, cco);

            if (pathAttachments?.Any() ?? false)
                foreach (var path in pathAttachments)
                    message.Attachments.AddFileAttachment(path);

            return await SendMailAsync(from, message);
        }


        private async Task<string> SendMailAsync(string mailbox, EmailMessage message)
        {
            try
            {
                await message.Save(new FolderId(WellKnownFolderName.Drafts, new Mailbox(mailbox)));
                await message.SendAndSaveCopy(new FolderId(WellKnownFolderName.SentItems, new Mailbox(mailbox)));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Erro ao enviar e-mail. Assunto: { message.Subject }; Destinatário - { string.Join('|', message.ToRecipients.Select(i => i.Address).ToArray()) }");
            }

            return message?.Id?.UniqueId;
        }

        private async Task<string> GetFolderIdAsync(FolderId folderId, params string[] folders)
        {
            var folderView = new FolderView(100);
            var filter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folders[0]);
            var result = await _service.FindFolders(folderId, filter, folderView);
            var folder = result.FirstOrDefault(i => i.DisplayName?.ToUpper() == folders[0]?.ToUpper());

            if (folder != null && folders.Count() > 1)
                return await GetFolderIdAsync(folder.Id, folders.Skip(1).ToArray());
            else
                return folder?.Id?.UniqueId;
        }

        private async Task<IEnumerable<MailMessage>> FindItemsAsync(FolderId folderId, ItemView itemView, SearchFilter search, string pathAttachment)
        {
            var messages = new List<MailMessage>();

            var results = await _service.FindItems(folderId, search, itemView);

            messages.AddRange(results?.Items?.Any() ?? false ? await LoadMessagesAsync(results.Items, pathAttachment) : new List<MailMessage>());

            if (results.MoreAvailable)
            {
                itemView.Offset = results.NextPageOffset.Value;
                messages.AddRange(await FindItemsAsync(folderId, itemView, search, pathAttachment));
            }

            return messages;
        }

        private async Task<IEnumerable<MailMessage>> LoadMessagesAsync(ICollection<Item> itens, string pathSaveAttachments = "")
        {
            await _service.LoadPropertiesForItems(itens, PropertySet.FirstClassProperties);

            var result = new List<MailMessage>();

            foreach (var item in itens)
            {
                var message = item as EmailMessage;

                result.Add(new MailMessage
                {
                    UniqueId = message.Id.UniqueId,
                    Body = message.Body,
                    ConversationId = message.ConversationId,
                    FromMail = message.From.Address,
                    FromName = message.From.Name,
                    Subject = message.Subject,
                    SentDate = message.DateTimeSent,
                    ReceivedDate = message.DateTimeReceived,
                    HasAttachments = message.HasAttachments,
                    Attachments = !string.IsNullOrWhiteSpace(pathSaveAttachments) && item.HasAttachments
                                ? await LoadAttachments(message, pathSaveAttachments)
                                : new List<AttachmentMessage>()
                });
            }

            return result;
        }

        private async Task<IEnumerable<AttachmentMessage>> LoadAttachments(EmailMessage message, string pathSaveAttachments)
        {
            var result = new List<AttachmentMessage>();

            try
            {
                foreach (var item in message.Attachments)
                {
                    var path = Path.Combine(pathSaveAttachments, item.Name);

                    if (!File.Exists(path))
                    {
                        var response = await item.Load();
                        var attachment = response.FirstOrDefault()?.Attachment as FileAttachment;
                        await File.WriteAllBytesAsync(path, attachment.Content);
                    }

                    result.Add(new AttachmentMessage(item));
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Erro ao carregar os anexos da mensagem - { message.Subject} ");
            }

            return result;
        }

        private EmailMessage CreateMessage(string subject, string body, string[] to, string[] cc = null, string from = null, string[] cco = null)
        {
            var message = new EmailMessage(_service)
            {
                Subject = subject,
                Body = body
            };

            message.Body.BodyType = BodyType.HTML;
            message.Importance = Importance.High;

            if (!string.IsNullOrWhiteSpace(from))
                message.From = from;

            foreach (var addressTo in to)
            {
                if (!string.IsNullOrWhiteSpace(addressTo))
                    message.ToRecipients.Add(addressTo.Trim());
            }

            if (cc?.Any() ?? false)
                foreach (var addressCc in cc)
                {
                    if (!string.IsNullOrWhiteSpace(addressCc))
                        message.ToRecipients.Add(addressCc.Trim());
                }

            if (cco?.Any() ?? false)
                foreach (var addressCco in cco)
                {
                    if (!string.IsNullOrWhiteSpace(addressCco))
                        message.ToRecipients.Add(addressCco.Trim());
                }

            return message;
        }
    }
}
