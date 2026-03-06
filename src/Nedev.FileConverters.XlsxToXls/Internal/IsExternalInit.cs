// Helper type required for record/`init` support on frameworks that don't
// include System.Runtime.CompilerServices.IsExternalInit by default
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}