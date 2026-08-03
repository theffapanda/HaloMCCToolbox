using System.Windows.Controls;

namespace HaloToolbox;

public partial class Mods : UserControl, IDisposable
{
    public Mods() => InitializeComponent();

    public void Dispose()
    {
        Halo3Mods.Dispose();
        HaloReachMods.Dispose();
    }
}
