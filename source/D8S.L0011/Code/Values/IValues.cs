using System;

using R5T.T0131;


namespace D8S.L0011
{
    [ValuesMarker]
    public partial interface IValues : IValuesMarker
    {
        /// <summary>
		/// <para><value>true</value></para>
		/// </summary>
		public bool ApplicationVisibility_Default => true;
    }
}
