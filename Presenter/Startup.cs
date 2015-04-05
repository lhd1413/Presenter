using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Presenter.Startup))]
namespace Presenter
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
