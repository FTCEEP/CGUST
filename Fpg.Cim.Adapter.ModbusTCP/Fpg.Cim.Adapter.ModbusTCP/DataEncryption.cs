using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.IO.Compression;
using NLog;

namespace Fpg.Cim.Adapter.ModbusTCP.DataEncryption
{
    class DataEncryption
    {
        private static Logger gLogger = NLog.LogManager.GetCurrentClassLogger();
        public byte[] DataSetToByte(DataSet ds)
        {
            try
            {
                ds.RemotingFormat = SerializationFormat.Binary;
                BinaryFormatter ser = new BinaryFormatter();
                MemoryStream unMS = new MemoryStream();
                ser.Serialize(unMS, ds);

                byte[] bytes = unMS.ToArray();
                int lenbyte = bytes.Length;

                MemoryStream compMs = new MemoryStream();
                GZipStream compStream = new GZipStream(compMs, CompressionMode.Compress, true);
                compStream.Write(bytes, 0, lenbyte);

                compStream.Close();
                unMS.Close();
                compMs.Close();
                byte[] zipData = compMs.ToArray();

                return zipData;
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("DataSetToByte--", ex);
                throw ex;
            }
        }

        public DataSet ByteToDataset(byte[] da)
        {
            try
            {
                MemoryStream input = new MemoryStream();
                input.Write(da, 0, da.Length);
                input.Position = 0;
                GZipStream gzip = new GZipStream(input, CompressionMode.Decompress, true);

                MemoryStream output = new MemoryStream();
                byte[] buff = new byte[4096];
                int read = -1;
                read = gzip.Read(buff, 0, buff.Length);

                while (read > 0)
                {
                    output.Write(buff, 0, read);
                    read = gzip.Read(buff, 0, buff.Length);
                }
                gzip.Close();
                byte[] rebytes = output.ToArray();
                output.Close();
                input.Close();

                MemoryStream ms = new MemoryStream(rebytes);
                BinaryFormatter bf = new BinaryFormatter();
                Object obj = bf.Deserialize(ms);

                return (DataSet)obj;
            }
            catch (Exception ex)
            {
                gLogger.ErrorException("ByteToDataset--", ex);
                throw ex;
            }
        }
    }
}
