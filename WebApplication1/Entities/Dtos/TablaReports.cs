using System.Collections.Generic;
using WebApplication1.Entities.DynamonDB;

namespace WebApplication1.Entities.Dtos
{
    public class TablaReports
    {
        public List<Archive> listArchive { get; set; }
        public Inspection inspection { get; set; }
        public Interviewed interviewed { get; set; }
        public Risk risk { get; set; }
        public Sinister sinister { get; set; }
    }
}
