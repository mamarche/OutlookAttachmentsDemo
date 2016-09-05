using Microsoft.Exchange.WebServices.Data;
using OutlookAttachmentsDemoService.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace OutlookAttachmentsDemoService.Controllers
{
    [System.Web.Http.Cors.EnableCors(origins: "*", headers: "*", methods: "*")]
    public class AttachmentsController : ApiController
    {
        public AttachmentResponse Post([FromBody]AttachmentRequest value)
        {
            //riceve le informazioni dalla chiamata REST attraverso l'oggetto AttachmentRequest e lo passa
            //al metodo che effettua la chiamata al server Exchange
            var resp = GetAttachmentsFromExchangeServerUsingEWSManagedApi(value);
            return resp;
        }

        private AttachmentResponse GetAttachmentsFromExchangeServerUsingEWSManagedApi(AttachmentRequest request)
        {
            var attachmentsProcessedCount = 0;
            var attachmentNames = new List<string>();

            // Creo un oggetto di tipo ExchangeService
            ExchangeService service = new ExchangeService();
            //imposto il token di autenticazione ricevuto dall'add-in
            service.Credentials = new OAuthCredentials(request.attachmentToken);
            //imposto la url del server Exchange
            service.Url = new Uri(request.ewsUrl);

            // Richiede gli allegati al server
            var getAttachmentsResponse = service.GetAttachments(
                                                request.attachments.Select(a => a.id).ToArray(),
                                                null,
                                                new PropertySet(BasePropertySet.FirstClassProperties,
                                                ItemSchema.MimeContent));

            if (getAttachmentsResponse.OverallResult == ServiceResult.Success)
            {
                foreach (var attachmentResponse in getAttachmentsResponse)
                {
                    attachmentNames.Add(attachmentResponse.Attachment.Name);

                    if (attachmentResponse.Attachment is FileAttachment)
                    {
                        //mette il contenuto dell'allegato in uno stream
                        FileAttachment fileAttachment = attachmentResponse.Attachment as FileAttachment;
                        Stream s = new MemoryStream(fileAttachment.Content);
                        
                        // Qui è possibile processare il contenuto dell'allegato
                    }

                    if (attachmentResponse.Attachment is ItemAttachment)
                    {
                        //mette il contenuto dell'allegato in uno stream
                        ItemAttachment itemAttachment = attachmentResponse.Attachment as ItemAttachment;
                        Stream s = new MemoryStream(itemAttachment.Item.MimeContent.Content);

                        // Qui è possibile processare il contenuto dell'allegato
                    }

                    attachmentsProcessedCount++;
                }
            }

            // La risposta contiene il nome ed il numero degli allegati che sono stati 
            // processati dal servizio
            var response = new AttachmentResponse();
            response.attachmentNames = attachmentNames.ToArray();
            response.attachmentsProcessed = attachmentsProcessedCount;

            return response;
        }

    }
}
