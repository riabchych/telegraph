using System.Data.Entity;

namespace Telegraph
{
    public class ApplicationContext: DbContext
    {
        public ApplicationContext() : base("DefaultConnection") { }
        public DbSet<Telegram> Telegrams { get; set; }
    }
}
