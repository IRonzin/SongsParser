using System.Data.Entity;

namespace SongsParser
{
    class SongsContext : DbContext
    {

            public SongsContext() : base("SongsConnection")
            { }

            public DbSet<Song> Songs { get; set; }
    }
}
