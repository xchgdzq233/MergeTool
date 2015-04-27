using BitMiracle.LibTiff.Classic;

namespace MergeTool
{
    class MyTiffErrorHandler : TiffErrorHandler
    {
        public override void WarningHandler(Tiff tif, string method, string format, params object[] args)
        {
        }

        public override void WarningHandlerExt(Tiff tif, object clientData, string method, string format, params object[] args)
        {
        }

        public override void ErrorHandler(Tiff tif, string method, string format, params object[] args)
        {
        }
    }
}
