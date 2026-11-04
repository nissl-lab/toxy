#if !NET5_0_OR_GREATER
namespace System.Runtime.CompilerServices
{
	/// <summary>
	/// Allows records and init-only setters on .NET Standard, the type only exists in the runtime since .NET 5.
	/// </summary>
	internal static class IsExternalInit
	{
	}
}
#endif
