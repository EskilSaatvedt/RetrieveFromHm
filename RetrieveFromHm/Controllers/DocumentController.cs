using System.Web;
using System.Web.Mvc;
using ehmInterface;
using HeyRed.Mime;
using RetrieveFromHm.Properties;

namespace RetrieveFromHm.Controllers
{
	public class DocumentController : Controller
	{
		/// <summary>
		///     http://ip:port/Document/Download/[HSDocumentID] will return the Document if it exisit in the Handyman
		///     The Client ID is set in the Web.config file
		/// </summary>
		/// <param name="id">HSDocumentID for give document</param>
		/// <returns>The File in question, return 404 if not found or some known error occured</returns>
		// GET: Document
		public ActionResult Download(int id = 0)
		{
			if (id == 0)
				throw new HttpException(404, "Bad Request");
			// clientId is set in Web.config->applicationSettings
			var clientId = Settings.Default.clientId;

			// Using COM object from Handyman ehmInterface/COMInterface
			var handyman = new XML();
			handyman.set_ClientId(clientId);

			var filePath = "";
			// Exporting the document out of Handyman
			try
			{
				// 0 for xml file, 1 for moving the file to the path returned as a string
				filePath = handyman.GetOrderDocument(id, 1); 
			}
			catch
			{
				throw new HttpException(404, "Unable to retrieve the file from Handyman");
			}

			if (!string.IsNullOrEmpty(filePath) && !System.IO.File.Exists(filePath))
				throw new HttpException(404, "Document error, unable to export the file");
			var mimeType = MimeTypesMap.GetMimeType(filePath);

			return new FilePathResult(filePath, mimeType);
		}
	}
}