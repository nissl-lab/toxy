using System;

namespace Toxy.Base
{
	// in theory we wouln't need the IEmailParser anymore if we have a Base Class but we will keep it as is for now
	/// <summary>
	/// Internal use only do not use!!!
	/// The <see cref="BaseMetaParser"/> is used to keep backward compatibility since the <see cref="Parse"/> Method could be used on the Parser directly.
	/// It should be used by every Parser, which parses the Metadata of a Stream / File.
	/// </summary>
	public abstract class BaseMetaParser : IMetadataParser
	{
		public ParserContext Context { get; set; }

		/// <summary>
		/// Initializes the <see cref="BaseMetaParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		private protected BaseMetaParser(ParserContext context)
		{
			Context = context;
		}

#nullable enable
		internal virtual void ValidateContext()
		{
			Utility.ValidateContext(Context);
		}

		// This Method is needed for backward compatibility!
		// It should prevent users from using the Parse Method directly on the Passwort to avoid code dupplication (Validating, Disposing, ...)
		/// <inheritdoc/>
		/// <remarks>
		/// Any <see cref="Exception"/> won't be catched!
		/// </remarks>
		public ToxyMetadata Parse()
		{
			ValidateContext();
			IDisposable? disposable = null;
			try
			{
				return ParseMeta(out disposable);
			}
			finally
			{
				// we don't care if it throws just dispose if needed
				disposable?.Dispose();
			}
		}

		/// <summary>
		/// Gets the <see cref="ToxyMetadata"/> of the <see cref="Context"/>
		/// </summary>
		/// <param name="disposable">An optional <see cref="IDisposable"/>, which should be disposed at the end of parsing.
		/// It will be disposed if an <see cref="Exception"/> will be thrown.</param>
		/// <returns>Returns the metadata of the <see cref="Context"/>.</returns>
		internal abstract ToxyMetadata ParseMeta(out IDisposable? disposable);
#nullable disable
	}
}
