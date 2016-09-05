using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OutlookAttachmentsDemoService.Models
{
    public class AttachmentDetail
    {
        public string attachmentType { get; set; }
        public string contentType { get; set; }
        public string id { get; set; }
        public bool isInline { get; set; }
        public string name { get; set; }
        public int size { get; set; }
    }
}