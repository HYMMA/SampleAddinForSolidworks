using SolidWorks.Interop.sldworks;

/// <summary>
/// this is an example of how you would add extension methods to solidworks interfaces
/// </summary>
namespace HymmaSampleAddin
{
    /// <summary>
    /// an example of extensions 
    /// </summary>
    public static class Extensions
    {
        public static string TestMessage(this ICommandManager commandManager)
        {
            return "this is a test message from HYMMA";
        }
    }
}
