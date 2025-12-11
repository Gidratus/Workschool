using System.Configuration;
using System.Data;
using System.Windows;
namespace WpfApp9
{
    public partial class App : Application
    {
        public const string ConnectionString =
            "Server=localhost\\SQLEXPRESS;Database=SchoolWork1;Trusted_Connection=True;TrustServerCertificate=True;";
    }
}
