using System.Configuration;
using System.Web;
using System.Web.Mvc;
using ehmInterface;
using HeyRed.Mime;

namespace RetrieveFromHm.Controllers
{
	public class DocumentController : Controller
	{
		// GET: Document
		public ActionResult Download(int id = 0)
		{
			if (id == 0)
				throw new HttpException(400, "Bad Request");
			// clientId is set in Web.config->applicationSettings
			var clientId = RetrieveFromHm.Properties.Settings.Default.clientId;

			// Using COM object from Handyman ehmInterface/COMInterface
			var handyman = new XML();
			handyman.set_ClientId(clientId);

			// Exporting the document out of Handyman
			var filePath = handyman.GetOrderDocument(id, 1);
			if (!string.IsNullOrEmpty(filePath) && !System.IO.File.Exists(filePath))
				throw new HttpException(400, "Document error, unable to export the file");
			var mimeType = MimeTypesMap.GetMimeType(filePath);

			return new FilePathResult(filePath, mimeType);
		}
	}
}