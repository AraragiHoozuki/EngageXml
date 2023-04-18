using AssetsTools.NET;
using AssetsTools.NET.Extra;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EngageXml
{
    internal class BundleIO
    {
        static AssetsManager AM = new AssetsManager();

        public static void UnloadAll()
        {
            AM.UnloadAll();
        }
        public static void UnoadBundleFile(string path)
        {
            AM.UnloadBundleFile(path);
        }
        /// <summary>
        /// Read an Fe engage data bundle's first asset as byte[]
        /// </summary>
        /// <param name="bundlePath"></param>
        /// <returns></returns>
        public static byte[] ReadBundleAssetData(string bundlePath)
        {
            var bun = AM.LoadBundleFile(bundlePath);

            //load first asset from bundle
            var inst = AM.LoadAssetsFileFromBundle(bun, 0);
            if (!inst.file.typeTree.hasTypeTree)
                AM.LoadClassDatabaseFromPackage(inst.file.typeTree.unityVersion);
            var inf = inst.table.assetFileInfo[0].index == 1 ? inst.table.assetFileInfo[1] : inst.table.assetFileInfo[0];
            var baseField = AM.GetTypeInstance(inst, inf).GetBaseField();
            byte[] data = baseField.Get("m_Script").GetValue().AsStringBytes();

            AM.UnloadAll();
            return data;
        }

        /// <summary>
        /// Insert data to bundle, replacing it's first asset
        /// </summary>
        /// <param name="data"></param>
        /// <param name="bundlePath"></param>
        public static void InsertAsset(byte[] data, string bundlePath)
        {
            var bun = AM.LoadBundleFile(bundlePath);

            //load first asset from bundle
            var inst = AM.LoadAssetsFileFromBundle(bun, 0);
            if (!inst.file.typeTree.hasTypeTree)
                AM.LoadClassDatabaseFromPackage(inst.file.typeTree.unityVersion);
            var inf = inst.table.assetFileInfo[0].index == 1 ? inst.table.assetFileInfo[1] : inst.table.assetFileInfo[0];
            var baseField = AM.GetTypeInstance(inst, inf).GetBaseField();
            baseField.Get("m_Script").GetValue().Set(data);

            var newGoBytes = baseField.WriteToByteArray();
            var repl = new AssetsReplacerFromMemory(0, inf.index, (int)inf.curFileType, 0xffff, newGoBytes);

            //write changes to memory
            byte[] newAssetData;
            using (var stream = new MemoryStream())
            using (var writer = new AssetsFileWriter(stream))
            {
                inst.file.Write(writer, 0, new List<AssetsReplacer>() { repl }, 0);
                newAssetData = stream.ToArray();
            }

            //get new bundle binary data
            var bunRepl = new BundleReplacerFromMemory(inst.name, null, true, newAssetData, -1);
            byte[] newBundleData;
            using (var stream = new MemoryStream())
            using (var bunWriter = new AssetsFileWriter(stream))
            {
                bun.file.Write(bunWriter, new List<BundleReplacer>() { bunRepl });
                newBundleData = stream.ToArray();
            }

            //write new bundle binary data
            MemoryStream newBundleStream = new MemoryStream(newBundleData);
            bun = AM.LoadBundleFile(newBundleStream, $"{bundlePath}.mod");
            AM.UnloadBundleFile(bundlePath);

            using (var stream = File.OpenWrite(bundlePath))
            using (var writer = new AssetsFileWriter(stream))
            {
                bun.file.Pack(bun.file.reader, writer, AssetBundleCompressionType.LZ4);
            }

            AM.UnloadAll();
        }
    }
}
