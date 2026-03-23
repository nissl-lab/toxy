using System;

namespace Toxy.Base
{
	// in theory we wouln't need the ITextParser anymore if we have a Base Class but we will keep it as is for now
	/// <summary>
	/// Internal use only do not use!!!
	/// The <see cref="BaseTextParser"/> is used to keep backward compatibility since the <see cref="Parse"/> Method could be used on the Parser directly.
	/// </summary>
	public abstract class BaseTextParser : ITextParser
	{
		public ParserContext Context { get; set; }

		/// <summary>
		/// Initializes the <see cref="BaseTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public BaseTextParser(ParserContext context)
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
		public string Parse()
		{
		    ValidateContext();
			IDisposable? disposable = null;
			try
			{
				return ParseText(out disposable);
			}
			finally
			{
				// we don't care if it throws just dispose if needed
				disposable?.Dispose();
			}
		}

		/// <summary>
		/// Gets the Text of the <see cref="Context"/>
		/// </summary>
		/// <param name="disposable">An optional <see cref="IDisposable"/>, which should be disposed at the end of parsing.
		/// It will be disposed if an <see cref="Exception"/> will be thrown.</param>
		/// <returns>Returns the extracted Text.</returns>
		internal abstract string ParseText(out IDisposable? disposable);
#nullable disable
	}
}
