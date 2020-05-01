using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Docx.src.model
{
    class DocInfoOption
    {
        private string subject;
        private string title;
        private string creator;
        private string keywords;
        private string description;
        private string lastModifiedBy;
        private string revision;
        private string category;
        private string version;
        private string created;
        private string modified;

        public DocInfoOption(string subject, string title, string creator, string keywords, string description, string lastModifiedBy, string revision, string category, string version, string created, string modified)
        {
            this.subject = subject;
            this.title = title;
            this.creator = creator;
            this.keywords = keywords;
            this.description = description;
            this.lastModifiedBy = lastModifiedBy;
            this.revision = revision;
            this.category = category;
            this.version = version;
            this.created = created;
            this.modified = modified;
        }

        public string Subject { get => subject; set => subject = value; }
        public string Title { get => title; set => title = value; }
        public string Creator { get => creator; set => creator = value; }
        public string Keywords { get => keywords; set => keywords = value; }
        public string Description { get => description; set => description = value; }
        public string LastModifiedBy { get => lastModifiedBy; set => lastModifiedBy = value; }
        public string Revision { get => revision; set => revision = value; }
        public string Category { get => category; set => category = value; }
        public string Version { get => version; set => version = value; }
        public string Created { get => created; set => created = value; }
        public string Modified { get => modified; set => modified = value; }
    }
}
