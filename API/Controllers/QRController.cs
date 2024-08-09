using Aspose.BarCode.Generation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class QRController : ControllerBase
    {
        
        [HttpPost]
        public IActionResult Post(Dictionary<string, string>[] KeyValuePairs)
        {
            for (int i = 0; i < KeyValuePairs.Length; i++)
            {
                var generator = new BarcodeGenerator(EncodeTypes.QR);
                // Specify code text to encode
               
                    KeyValuePairs.ToList()[i].TryGetValue("URL",out string colValue);
                generator.CodeText = colValue;
                // Specify the size of the image
                generator.Parameters.Barcode.XDimension.Pixels = 8;
                generator.Parameters.Resolution = 500;
                // Save the generated QR code
                generator.Save("C:\\Ehtisham\\CCLIVE\\CCVote\\CCVote\\wwwroot\\img"+i+".jpg");
            }
            return Ok(new
            {
                message = "QR Generated Successfully!!!",
            });
        }
        private string TryGetValueFromDist(Dictionary<string, string> rowData, string key)
        {
            rowData.TryGetValue(key, out string colValue);
            return colValue;
        }
    }
    public class QRModel { 
    
    }
}
