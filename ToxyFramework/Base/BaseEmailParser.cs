using System;

namespace Toxy.Base
{
	// in theory we wouln't need the IEmailParser anymore if we have a Base Class but we will keep it as is for now
	/// <summary>
	/// Internal use only do not use!!!
	/// The <see cref="BaseEmailParser"/> is used to keep backward compatibility since the <see cref="Parse"/> Method could be used on the Parser directly.
	/// It should be used by every Parser, which parses an Email.
	/// </summary>
	public abstract class BaseEmailParser : IEmailParser
	{
		public ParserContext Context { get; set; }

		/// <summary>
		/// Initializes the <see cref="BaseEmailParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		private protected BaseEmailParser(ParserContext context)
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
		public ToxyEmail Parse()
		{
			ValidateContext();
			IDisposable? disposable = null;
			try
			{
				return ParseEmail(out disposable);
			}
			finally
			{
				// we don't care if it throws just dispose if needed
				disposable?.Dispose();
			}
		}

		/// <summary>
		/// Gets the <see cref="ToxyEmail"/> of the <see cref="Context"/>
		/// </summary>
		/// <param name="disposable">An optional <see cref="IDisposable"/>, which should be disposed at the end of parsing.
		/// It will be disposed if an <see cref="Exception"/> will be thrown.</param>
		/// <returns>Returns the <see cref="Context"/> as structured Object (<see cref="ToxyEmail"/>).</returns>
		internal abstract ToxyEmail ParseEmail(out IDisposable? disposable);
#nullable disable
	}
}
