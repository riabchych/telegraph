using System.Collections.Generic;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace Telegraph
{
    public class DocumentPaginatorWrapper : DocumentPaginator
    {
        Size m_PageSize;
        Size m_Margin;
        DocumentPaginator m_Paginator;
        Typeface m_Typeface;
        /// <summary> 
        ///     Constructor por el cual se empieza a hacer el paginador 
        ///     para el FlowDocument. 
        /// </summary> 
        /// <param name="paginator" type="System.Windows.Documents.DocumentPaginator"> 
        ///     <para> 
        ///         El paginador del documento, si se trata de un Flow document es posible hacer algo como: 
        ///         IDocumentPaginatorSource fuentePagina= miFlowDocument as IDocumentPaginatorSource; 
        ///         DocumentPaginator paginador = fuentePagina.DocumentPaginator 
        ///     </para> 
        /// </param> 
        /// <param name="pageSize" type="System.Windows.Size"> 
        ///     <para> 
        ///         El tamaño de la hoja, 8*11.5 pulgadas es carta. 
        ///     </para> 
        /// </param> 
        /// <param name="margin" type="System.Windows.Size"> 
        ///     <para> 
        ///         El Margen en las hojas 
        ///     </para> 
        /// </param> 
        public DocumentPaginatorWrapper(DocumentPaginator paginator, Size pageSize, Size margin)
        {
            m_PageSize = pageSize;
            m_Margin = margin;
            m_Paginator = paginator;
            m_Paginator.PageSize = new Size(m_PageSize.Width - margin.Width * 2,
                                            m_PageSize.Height - margin.Height * 2);
        }
        /// <summary> 
        ///     Mueve un área en específico para generar encabezados o pies de página. 
        /// </summary> 
        /// <param name="rect" type="System.Windows.Rect"> 
        ///     <para> 
        ///         Ubicación de este rectángulo. 
        ///     </para> 
        /// </param> 
        /// <returns> 
        ///     El rectangulo. 
        /// </returns> 
        Rect Move(Rect rect)
        {
            if (rect.IsEmpty)
            {
                return rect;
            }
            else
            {
                return new Rect(rect.Left + m_Margin.Width, rect.Top + m_Margin.Height,
                                rect.Width, rect.Height);
            }
        }
        /// <summary> 
        ///     Obtiene una página. 
        /// </summary> 
        /// <param name="pageNumber" type="int"> 
        ///     <para> 
        ///         El npumero de página 
        ///     </para> 
        /// </param> 
        /// <returns> 
        ///     La página solicitada del documento. 
        /// </returns> 
        public override DocumentPage GetPage(int pageNumber)
        {
            if (m_Paginator.PageCount <= pageNumber)
            {
                return null;
            }
            DocumentPage page = null;
            page = m_Paginator.GetPage(pageNumber);
            // Create a wrapper visual for transformation and add extras 
            ContainerVisual newpage = new ContainerVisual();
            DrawingVisual title = new DrawingVisual();

            using (DrawingContext ctx = title.RenderOpen())
            {
                ctx.DrawRectangle(Brushes.Transparent, null, new Rect(page.Size));
                if (m_Typeface == null)
                {
                    m_Typeface = new Typeface(new FontFamily("Arial"), FontStyles.Normal, FontWeights.Bold, FontStretches.Normal);
                }
                FormattedText text = new FormattedText("My Company Name \tPágina " + (pageNumber + 1),
                    System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    m_Typeface, 14, Brushes.Black);
                ctx.DrawText(text, new Point(48, 48)); // 1/4 inch above page content 
            }
            ContainerVisual smallerPage = new ContainerVisual();
            smallerPage.Children.Add(page.Visual);
            newpage.Children.Add(smallerPage);
            newpage.Children.Add(title);
            return new DocumentPage(newpage, m_PageSize, Move(page.BleedBox), Move(page.ContentBox));
        }
        /// <summary> 
        ///     Define si la cuenta de las páginas es o no válida 
        /// </summary> 
        /// <value> 
        ///     <para> 
        ///          
        ///     </para> 
        /// </value> 
        /// <remarks> 
        ///      
        /// </remarks> 
        public override bool IsPageCountValid
        {
            get
            {
                return m_Paginator.IsPageCountValid;
            }
        }
        /// <summary> 
        ///     Obtiene la cuenta de las páginas 
        /// </summary> 
        /// <value> 
        ///     <para> 
        ///          
        ///     </para> 
        /// </value> 
        /// <remarks> 
        ///      
        /// </remarks> 
        public override int PageCount
        {
            get
            {
                return m_Paginator.PageCount;
            }
        }
        protected Size pageSize;
        public override Size PageSize
        {
            get
            {
                return m_Paginator.PageSize;
            }
            set
            {
                if (this == m_Paginator)
                {
                    this.pageSize = value;
                }
                else
                {
                    m_Paginator.PageSize = value;
                }
            }
        }

        public override IDocumentPaginatorSource Source
        {
            get
            {
                return m_Paginator.Source;
            }
        }
    }

    public class IUSPaginatedDocument : FlowDocument, IDocumentPaginatorSource
    {
        public IUSPaginatedDocument()
            : base()
        {
        }
        #region IDocumentPaginatorSource Members 
        protected DocumentPaginator documentPaginator;
        DocumentPaginator IDocumentPaginatorSource.DocumentPaginator
        {
            get
            {
                if (documentPaginator != null)
                {
                    return documentPaginator;
                }
                else
                {
                    DocumentPaginator pgn = new IUSDocumentPaginator(this);
                    documentPaginator = new DocumentPaginatorWrapper(pgn, new Size(96 * 8.5, 96 * 11), new Size(46, 96));
                    return documentPaginator;
                }
            }
        }
        #endregion

    }
    public class IUSDocumentPaginator : DocumentPaginator
    {
        FlowDocument documento;
        FlowDocument copia;
        IDocumentPaginatorSource iDocumentPaginatorSource;
        DocumentPage[] paginas;
        public IUSDocumentPaginator(FlowDocument doc)
        {
            documento = doc;
            FlowDocument temporal = new FlowDocument();
            temporal.PageHeight = 96 * 11;
            temporal.PageWidth = 96 * 8.5;
            temporal.ColumnWidth = 96 * 6.5;
            temporal.Background = Brushes.White;
            temporal.PagePadding = new Thickness(96 * 0.5, 96 * 1, 96 * 0.5, 96 * 0.5);
            List<Block> lista = new List<Block>();
            foreach (Block item in doc.Blocks)
            {
                lista.Add(item);
            }
            foreach (Block item in lista)
            {
                temporal.Blocks.Add(item);
            }
            this.iDocumentPaginatorSource = temporal as IDocumentPaginatorSource;
            copia = temporal;
        }
        public override DocumentPage GetPage(int pageNumber)
        {
            return iDocumentPaginatorSource.DocumentPaginator.GetPage(pageNumber);
        }

        public override bool IsPageCountValid
        {
            get { return iDocumentPaginatorSource.DocumentPaginator.IsPageCountValid; }
        }

        public override int PageCount
        {
            get
            {
                return iDocumentPaginatorSource.DocumentPaginator.PageCount;
            }
        }
        protected Size pageSize;
        public override Size PageSize
        {
            get
            {
                return iDocumentPaginatorSource.DocumentPaginator.PageSize;
            }
            set
            {
                iDocumentPaginatorSource.DocumentPaginator.PageSize = value;
                pageSize = value;
            }
        }

        public override IDocumentPaginatorSource Source
        {
            get { return iDocumentPaginatorSource; }
        }

        private IDocumentPaginatorSource getSource()
        {
            return iDocumentPaginatorSource.DocumentPaginator.Source;
        }
    }
}
