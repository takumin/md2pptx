using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using MDToPPTX.PPTX.DefaultParts.SlideLayouts;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;

namespace MDToPPTX.PPTX.DefaultParts
{
	internal class DefaultPresentationDocument
	{
		public static PresentationDocument CreatePresentationDocument(
			string FilePath,
			PPTXSetting FileSettings,
			PPTXSlideLayoutGroup SlideLayouts
		)
		{
			var presentationDoc = PresentationDocument.Create(FilePath, PresentationDocumentType.Presentation);
			PresentationPart presentationPart = presentationDoc.AddPresentationPart();
			presentationPart.Presentation = new Presentation();

			DefaultPresentationPart.CreatePresentationPart(presentationPart, FileSettings, SlideLayouts);

			return presentationDoc;
		}
	}
}
