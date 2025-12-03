using System.Configuration;
using System.Data;
using System.Windows;

namespace WpfApp9
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        //public const string ConnectionString =
        //   "Server=localhost;Database=SchoolWork1;Trusted_Connection=True;TrustServerCertificate=True;";

        public const string ConnectionString =
            "Server=localhost\\SQLEXPRESS;Database=SchoolWork1;Trusted_Connection=True;TrustServerCertificate=True;";
    }

}
