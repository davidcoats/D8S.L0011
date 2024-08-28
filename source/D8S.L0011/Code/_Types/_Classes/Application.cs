using System;

using R5T.T0142;

using Wd = Microsoft.Office.Interop.Word;


namespace D8S.L0011
{
    /// <summary>
    /// Encapsulates a Microsoft Word application.
    /// </summary>
    [UtilityTypeMarker]
    public class Application : IDisposable
    {
        #region IDisposable

        private bool zDisposed = false; // To detect redundant calls.


        public void Dispose()
        {
            this.Dispose(true);

            GC.SuppressFinalize(this);
        }

        // Remove the virtual call if the class is sealed (or has no plans for subclassing, in which case this should be communicated by sealing the class).
        private void Dispose(bool disposing)
        {
            if (!this.zDisposed)
            {
                if (disposing)
                {
                    // Do nothing.
                    /// The <see cref="Xl.Application"/> object itself is managed, the Excel application it is the handle to is not.
                }

                this.WdApplication.DisplayAlerts = Wd.WdAlertLevel.wdAlertsNone;
                this.WdApplication.Quit();
            }

            this.zDisposed = true;
        }

        ~Application()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            this.Dispose(false);
        }

        #endregion


        internal Wd.Application WdApplication { get; private set; }


        public Application(bool visible)
        {
            this.WdApplication = new Wd.Application()
            {
                Visible = visible,
            };
        }

        public Application()
            : this(Instances.Values.ApplicationVisibility_Default)
        {
        }

        /// <summary>
        /// Identical to <see cref="Application.Dispose()"/>, but allows for use outside of a using statment.
        /// </summary>
        public void Quit()
        {
            this.Dispose();
        }

        public Document New_Document()
        {
            var wdDocument = this.WdApplication.Documents.Add();

            var workbook = new Document(wdDocument, this);
            return workbook;
        }
    }
}
