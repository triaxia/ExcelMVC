using Function.Definitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelMvc.Rtd
{
    /// <summary>
    /// A naive way of creating RTD servers... until I can do COM server factory...
    /// </summary>
    public static class RtdServers
    {
        private static readonly string MutexName = $"Global\\{nameof(RtdServers)}";
        private static Mutex SystemMutex;
        private static readonly TimeSpan Timeout = TimeSpan.FromSeconds(5);
        public static (Type type, string progId) Acquire(IRtdServerImpl impl)
        {
            lock (MutexName)
            {
                var pair = Impls.FirstOrDefault(x => x.Value == impl);
                if (pair.Key != null)
                    return (pair.Key, RtdRegistration.GetProgId(pair.Key));

                try
                {
                    SystemMutex = new Mutex(false, MutexName);
                    if (!SystemMutex.WaitOne(Timeout))
                        throw new TimeoutException($"Wait on {MutexName} timed out ({Timeout})");
                }
                catch (AbandonedMutexException)
                {
                    // the owning process abandoned, e.g. terminated with ReleaseMutex
                }
                catch
                {
                    SystemMutex?.Dispose();
                    throw;
                }
                var type = Impls.FirstOrDefault(x => x.Value == null).Key;
                if (type == null)
                    // not really going to occur...
                    throw new ArgumentOutOfRangeException($"Rtd server limit {Impls.Count} exceeded ");

                Impls[type] = impl;
                RtdRegistration.RegisterType(type);
                return (type, RtdRegistration.GetProgId(type));
            }
        }

        public static IRtdServerImpl GetImpl(Type rtdType)
        {
            lock (MutexName)
            {
                var impl = Impls[rtdType];
                RtdRegistration.DeleteProgId(RtdRegistration.GetProgId(rtdType));
                SystemMutex?.ReleaseMutex();
                SystemMutex?.Dispose();
                return impl;
            };
        }

        public static void Release(Type rtdType)
        {
            lock (MutexName)
            {
                Impls[rtdType] = null;
            }
        }

        private static readonly Dictionary<Type, IRtdServerImpl> Impls
            = new Dictionary<Type, IRtdServerImpl>()
            {
                {typeof(Rtd101), null},
                {typeof(Rtd102), null},
                {typeof(Rtd103), null},
                {typeof(Rtd104), null},
                {typeof(Rtd105), null},
                {typeof(Rtd106), null},
                {typeof(Rtd107), null},
                {typeof(Rtd108), null},
                {typeof(Rtd109), null},
                {typeof(Rtd110), null},
                {typeof(Rtd111), null},
                {typeof(Rtd112), null},
                {typeof(Rtd113), null},
                {typeof(Rtd114), null},
                {typeof(Rtd115), null},
                {typeof(Rtd116), null},
                {typeof(Rtd117), null},
                {typeof(Rtd118), null},
                {typeof(Rtd119), null},
                {typeof(Rtd120), null},
                {typeof(Rtd121), null},
                {typeof(Rtd122), null},
                {typeof(Rtd123), null},
                {typeof(Rtd124), null},
                {typeof(Rtd125), null},
                {typeof(Rtd126), null},
                {typeof(Rtd127), null},
                {typeof(Rtd128), null},
                {typeof(Rtd129), null},
                {typeof(Rtd130), null},
                {typeof(Rtd131), null},
                {typeof(Rtd132), null},
                {typeof(Rtd133), null},
                {typeof(Rtd134), null},
                {typeof(Rtd135), null},
                {typeof(Rtd136), null},
                {typeof(Rtd137), null},
                {typeof(Rtd138), null},
                {typeof(Rtd139), null},
                {typeof(Rtd140), null},
                {typeof(Rtd141), null},
                {typeof(Rtd142), null},
                {typeof(Rtd143), null},
                {typeof(Rtd144), null},
                {typeof(Rtd145), null},
                {typeof(Rtd146), null},
                {typeof(Rtd147), null},
                {typeof(Rtd148), null},
                {typeof(Rtd149), null},
                {typeof(Rtd150), null},
                {typeof(Rtd151), null},
                {typeof(Rtd152), null},
                {typeof(Rtd153), null},
                {typeof(Rtd154), null},
                {typeof(Rtd155), null},
                {typeof(Rtd156), null},
                {typeof(Rtd157), null},
                {typeof(Rtd158), null},
                {typeof(Rtd159), null},
                {typeof(Rtd160), null},
                {typeof(Rtd161), null},
                {typeof(Rtd162), null},
                {typeof(Rtd163), null},
                {typeof(Rtd164), null},
                {typeof(Rtd165), null},
                {typeof(Rtd166), null},
                {typeof(Rtd167), null},
                {typeof(Rtd168), null},
                {typeof(Rtd169), null},
                {typeof(Rtd170), null},
                {typeof(Rtd171), null},
                {typeof(Rtd172), null},
                {typeof(Rtd173), null},
                {typeof(Rtd174), null},
                {typeof(Rtd175), null},
                {typeof(Rtd176), null},
                {typeof(Rtd177), null},
                {typeof(Rtd178), null},
                {typeof(Rtd179), null},
                {typeof(Rtd180), null},
                {typeof(Rtd181), null},
                {typeof(Rtd182), null},
                {typeof(Rtd183), null},
                {typeof(Rtd184), null},
                {typeof(Rtd185), null},
                {typeof(Rtd186), null},
                {typeof(Rtd187), null},
                {typeof(Rtd188), null},
                {typeof(Rtd189), null},
                {typeof(Rtd190), null},
                {typeof(Rtd191), null},
                {typeof(Rtd192), null},
                {typeof(Rtd193), null},
                {typeof(Rtd194), null},
                {typeof(Rtd195), null},
                {typeof(Rtd196), null},
                {typeof(Rtd197), null},
                {typeof(Rtd198), null},
                {typeof(Rtd199), null},
                {typeof(Rtd200), null}
            };
    }

    [Guid("b92c4d6a-0586-435c-a6a6-053bbd1ae1b7")][ComVisible(true)][ProgId("ExcelMvc.Rtd101")] public class Rtd101 : RtdServer { };
    [Guid("d473afa0-94b1-44a0-9915-ea31618ed346")][ComVisible(true)][ProgId("ExcelMvc.Rtd102")] public class Rtd102 : RtdServer { };
    [Guid("1b01b885-fd41-4d4a-8858-6097c6312961")][ComVisible(true)][ProgId("ExcelMvc.Rtd103")] public class Rtd103 : RtdServer { };
    [Guid("b77a40d4-6684-4ede-9386-474b48d6566e")][ComVisible(true)][ProgId("ExcelMvc.Rtd104")] public class Rtd104 : RtdServer { };
    [Guid("14231272-47bf-45c5-8d14-0a9735ea5fca")][ComVisible(true)][ProgId("ExcelMvc.Rtd105")] public class Rtd105 : RtdServer { };
    [Guid("f473f0cd-ed88-42de-b402-7b3ba423c70d")][ComVisible(true)][ProgId("ExcelMvc.Rtd106")] public class Rtd106 : RtdServer { };
    [Guid("16a1ca58-3348-462d-9080-994be6c6a511")][ComVisible(true)][ProgId("ExcelMvc.Rtd107")] public class Rtd107 : RtdServer { };
    [Guid("1cc71efe-e5e6-49cd-9487-c96b391951a6")][ComVisible(true)][ProgId("ExcelMvc.Rtd108")] public class Rtd108 : RtdServer { };
    [Guid("4b1dc311-95c7-46e6-a3dd-de08674f9726")][ComVisible(true)][ProgId("ExcelMvc.Rtd109")] public class Rtd109 : RtdServer { };
    [Guid("cc8e967d-cd42-4272-956a-2bf5568ec18d")][ComVisible(true)][ProgId("ExcelMvc.Rtd110")] public class Rtd110 : RtdServer { };
    [Guid("415d429c-375a-40ec-8e89-a9e7fe3bc25a")][ComVisible(true)][ProgId("ExcelMvc.Rtd111")] public class Rtd111 : RtdServer { };
    [Guid("a6239d8b-4cd9-4f22-b9ae-441a66ba4c89")][ComVisible(true)][ProgId("ExcelMvc.Rtd112")] public class Rtd112 : RtdServer { };
    [Guid("e323c9a3-dd8d-49fd-9b20-49824c750440")][ComVisible(true)][ProgId("ExcelMvc.Rtd113")] public class Rtd113 : RtdServer { };
    [Guid("51f0e1d8-39f6-438c-b824-65bf92423292")][ComVisible(true)][ProgId("ExcelMvc.Rtd114")] public class Rtd114 : RtdServer { };
    [Guid("c9aea142-869b-477e-93a5-3b007bfffb9c")][ComVisible(true)][ProgId("ExcelMvc.Rtd115")] public class Rtd115 : RtdServer { };
    [Guid("d85100a9-8a1a-4203-9efc-25bd92423925")][ComVisible(true)][ProgId("ExcelMvc.Rtd116")] public class Rtd116 : RtdServer { };
    [Guid("0b449eec-7e7e-4d52-9ff4-9c4a63dcc198")][ComVisible(true)][ProgId("ExcelMvc.Rtd117")] public class Rtd117 : RtdServer { };
    [Guid("3901d993-bc5d-4461-9bf3-874f94fda277")][ComVisible(true)][ProgId("ExcelMvc.Rtd118")] public class Rtd118 : RtdServer { };
    [Guid("a13c9958-109e-400a-a796-0fe4a7db06b8")][ComVisible(true)][ProgId("ExcelMvc.Rtd119")] public class Rtd119 : RtdServer { };
    [Guid("b878efaa-d8f9-4aa0-9fc0-5f8a65ce2887")][ComVisible(true)][ProgId("ExcelMvc.Rtd120")] public class Rtd120 : RtdServer { };
    [Guid("e88a34c1-9a05-4c6d-9ee1-bba181830f14")][ComVisible(true)][ProgId("ExcelMvc.Rtd121")] public class Rtd121 : RtdServer { };
    [Guid("657e5092-1fca-445f-9d86-8ef14c7d363e")][ComVisible(true)][ProgId("ExcelMvc.Rtd122")] public class Rtd122 : RtdServer { };
    [Guid("e8847a56-2d9c-41aa-9893-e20a0617ce36")][ComVisible(true)][ProgId("ExcelMvc.Rtd123")] public class Rtd123 : RtdServer { };
    [Guid("69d52fd3-1caf-4c2f-af0c-151251d05ee7")][ComVisible(true)][ProgId("ExcelMvc.Rtd124")] public class Rtd124 : RtdServer { };
    [Guid("552141f1-23c5-4b25-bb57-765694b913ec")][ComVisible(true)][ProgId("ExcelMvc.Rtd125")] public class Rtd125 : RtdServer { };
    [Guid("9df89eba-88ae-4305-b29f-ace5ecb19462")][ComVisible(true)][ProgId("ExcelMvc.Rtd126")] public class Rtd126 : RtdServer { };
    [Guid("1d899dd9-960e-436d-ab4a-22b7d61184e8")][ComVisible(true)][ProgId("ExcelMvc.Rtd127")] public class Rtd127 : RtdServer { };
    [Guid("483ead92-4faf-4662-8e83-fbc4a30ab797")][ComVisible(true)][ProgId("ExcelMvc.Rtd128")] public class Rtd128 : RtdServer { };
    [Guid("bc914a7e-0620-4761-9ee3-ff18dcfeb12f")][ComVisible(true)][ProgId("ExcelMvc.Rtd129")] public class Rtd129 : RtdServer { };
    [Guid("38734fc1-60bf-4c52-bb29-66110bbc02c4")][ComVisible(true)][ProgId("ExcelMvc.Rtd130")] public class Rtd130 : RtdServer { };
    [Guid("2b42732f-dd4f-4201-a1bb-8040a9293980")][ComVisible(true)][ProgId("ExcelMvc.Rtd131")] public class Rtd131 : RtdServer { };
    [Guid("f0d53a5a-b4fb-448d-93ab-64acaed144a8")][ComVisible(true)][ProgId("ExcelMvc.Rtd132")] public class Rtd132 : RtdServer { };
    [Guid("02b1b15a-963f-4c51-95a8-758e10698bd8")][ComVisible(true)][ProgId("ExcelMvc.Rtd133")] public class Rtd133 : RtdServer { };
    [Guid("c6ff2d97-8830-4824-b4c1-5204cabf96b8")][ComVisible(true)][ProgId("ExcelMvc.Rtd134")] public class Rtd134 : RtdServer { };
    [Guid("2d24c46f-0d2f-461e-81d9-d87f3d89fb19")][ComVisible(true)][ProgId("ExcelMvc.Rtd135")] public class Rtd135 : RtdServer { };
    [Guid("34390d7f-d34a-4667-9d3e-2467c30d52d4")][ComVisible(true)][ProgId("ExcelMvc.Rtd136")] public class Rtd136 : RtdServer { };
    [Guid("52db2c0b-a9b9-4410-a1e8-85ab541e1d7b")][ComVisible(true)][ProgId("ExcelMvc.Rtd137")] public class Rtd137 : RtdServer { };
    [Guid("6e6d501b-fcc7-4afc-8910-bdd0e98e0ffa")][ComVisible(true)][ProgId("ExcelMvc.Rtd138")] public class Rtd138 : RtdServer { };
    [Guid("67425f1b-0354-41ea-98f4-690da471edde")][ComVisible(true)][ProgId("ExcelMvc.Rtd139")] public class Rtd139 : RtdServer { };
    [Guid("65c7c3d1-6479-4dc9-a30e-e13c6eeabe92")][ComVisible(true)][ProgId("ExcelMvc.Rtd140")] public class Rtd140 : RtdServer { };
    [Guid("bb2d391b-bcd0-46d8-ba88-77a11009d055")][ComVisible(true)][ProgId("ExcelMvc.Rtd141")] public class Rtd141 : RtdServer { };
    [Guid("90f7e95d-6d13-4fae-bf9f-db5f5a28ca4f")][ComVisible(true)][ProgId("ExcelMvc.Rtd142")] public class Rtd142 : RtdServer { };
    [Guid("9b4a12ea-4a1e-4fd1-b46b-8af0eeecd49f")][ComVisible(true)][ProgId("ExcelMvc.Rtd143")] public class Rtd143 : RtdServer { };
    [Guid("042f35d2-b7dc-4125-9ad7-04d4d11ccc91")][ComVisible(true)][ProgId("ExcelMvc.Rtd144")] public class Rtd144 : RtdServer { };
    [Guid("5eff7e25-c3ad-4a20-b755-0106a628cd54")][ComVisible(true)][ProgId("ExcelMvc.Rtd145")] public class Rtd145 : RtdServer { };
    [Guid("da09c246-f702-481e-b51c-99d3340852b6")][ComVisible(true)][ProgId("ExcelMvc.Rtd146")] public class Rtd146 : RtdServer { };
    [Guid("21033feb-26ef-41f6-b9ae-5360bbe5d712")][ComVisible(true)][ProgId("ExcelMvc.Rtd147")] public class Rtd147 : RtdServer { };
    [Guid("fe57af7e-bae4-47b3-8769-5a6c3c018ec9")][ComVisible(true)][ProgId("ExcelMvc.Rtd148")] public class Rtd148 : RtdServer { };
    [Guid("caec3f36-442f-4bda-9058-0f78e770fe13")][ComVisible(true)][ProgId("ExcelMvc.Rtd149")] public class Rtd149 : RtdServer { };
    [Guid("15f42dae-c798-4fc3-adc0-e26b867107a1")][ComVisible(true)][ProgId("ExcelMvc.Rtd150")] public class Rtd150 : RtdServer { };
    [Guid("df824be3-d0c8-412e-b72b-34255b716ef6")][ComVisible(true)][ProgId("ExcelMvc.Rtd151")] public class Rtd151 : RtdServer { };
    [Guid("9873bf26-aeab-458d-b26f-4e258134062d")][ComVisible(true)][ProgId("ExcelMvc.Rtd152")] public class Rtd152 : RtdServer { };
    [Guid("4b1fa237-1727-4cff-af50-101779921433")][ComVisible(true)][ProgId("ExcelMvc.Rtd153")] public class Rtd153 : RtdServer { };
    [Guid("69a3caef-aadc-49c0-9dd7-7ba2c35196c6")][ComVisible(true)][ProgId("ExcelMvc.Rtd154")] public class Rtd154 : RtdServer { };
    [Guid("f23fbfeb-2847-4e27-a3f9-cf5668e58435")][ComVisible(true)][ProgId("ExcelMvc.Rtd155")] public class Rtd155 : RtdServer { };
    [Guid("739ce140-009c-4afa-bae7-322c7e5a090f")][ComVisible(true)][ProgId("ExcelMvc.Rtd156")] public class Rtd156 : RtdServer { };
    [Guid("e55d273d-bc37-4bd9-89df-e48a809e91fe")][ComVisible(true)][ProgId("ExcelMvc.Rtd157")] public class Rtd157 : RtdServer { };
    [Guid("9bf195e0-2fc0-4ea4-b1c8-eef73aa715cc")][ComVisible(true)][ProgId("ExcelMvc.Rtd158")] public class Rtd158 : RtdServer { };
    [Guid("2c47c541-5311-46ea-b5a6-083bda87635f")][ComVisible(true)][ProgId("ExcelMvc.Rtd159")] public class Rtd159 : RtdServer { };
    [Guid("5d83e72f-ef9d-4d89-8644-b44a0ecd15f2")][ComVisible(true)][ProgId("ExcelMvc.Rtd160")] public class Rtd160 : RtdServer { };
    [Guid("6d174a02-40bd-4cc6-918f-a5787ba35d9b")][ComVisible(true)][ProgId("ExcelMvc.Rtd161")] public class Rtd161 : RtdServer { };
    [Guid("58d5bf52-7dce-441b-92ae-1cd7681c148a")][ComVisible(true)][ProgId("ExcelMvc.Rtd162")] public class Rtd162 : RtdServer { };
    [Guid("3dd71ea3-5d1e-4cbc-bb60-62e2a7197d1d")][ComVisible(true)][ProgId("ExcelMvc.Rtd163")] public class Rtd163 : RtdServer { };
    [Guid("6a09cb82-8995-4ed9-9d0e-c3c1df2ecff1")][ComVisible(true)][ProgId("ExcelMvc.Rtd164")] public class Rtd164 : RtdServer { };
    [Guid("74bfcad8-3ac3-4676-a592-408252ed7fcf")][ComVisible(true)][ProgId("ExcelMvc.Rtd165")] public class Rtd165 : RtdServer { };
    [Guid("a5ab5b07-2eb9-4e5c-b638-5a03ec0b08a1")][ComVisible(true)][ProgId("ExcelMvc.Rtd166")] public class Rtd166 : RtdServer { };
    [Guid("d7ed99fc-8ebd-43c2-a25f-f3a162e95e95")][ComVisible(true)][ProgId("ExcelMvc.Rtd167")] public class Rtd167 : RtdServer { };
    [Guid("bd5bb457-7ba1-4fea-94a6-a6181cfde4e7")][ComVisible(true)][ProgId("ExcelMvc.Rtd168")] public class Rtd168 : RtdServer { };
    [Guid("2b3331d7-e3cb-400c-94e2-ee9d049a9b10")][ComVisible(true)][ProgId("ExcelMvc.Rtd169")] public class Rtd169 : RtdServer { };
    [Guid("2d454736-cf9b-4df0-a802-5997c1e49bb2")][ComVisible(true)][ProgId("ExcelMvc.Rtd170")] public class Rtd170 : RtdServer { };
    [Guid("aead3e12-2f6a-40ff-9650-4a8fa751f0fd")][ComVisible(true)][ProgId("ExcelMvc.Rtd171")] public class Rtd171 : RtdServer { };
    [Guid("fa1c7c5d-6be3-4c38-8269-17ea93cd533a")][ComVisible(true)][ProgId("ExcelMvc.Rtd172")] public class Rtd172 : RtdServer { };
    [Guid("32e907ff-65d6-47f6-8bc3-55eb966457a7")][ComVisible(true)][ProgId("ExcelMvc.Rtd173")] public class Rtd173 : RtdServer { };
    [Guid("f3647579-33c8-48c5-a6bc-d692c7781522")][ComVisible(true)][ProgId("ExcelMvc.Rtd174")] public class Rtd174 : RtdServer { };
    [Guid("b6d7ed0a-aa12-473f-bcc1-2beb537b11f5")][ComVisible(true)][ProgId("ExcelMvc.Rtd175")] public class Rtd175 : RtdServer { };
    [Guid("5a3c7472-4505-4f1f-8978-40d4583ad1d5")][ComVisible(true)][ProgId("ExcelMvc.Rtd176")] public class Rtd176 : RtdServer { };
    [Guid("10895e59-43df-40a2-973d-9fc8a94fee14")][ComVisible(true)][ProgId("ExcelMvc.Rtd177")] public class Rtd177 : RtdServer { };
    [Guid("b13608c9-52d1-4c80-a37a-b0443fff514d")][ComVisible(true)][ProgId("ExcelMvc.Rtd178")] public class Rtd178 : RtdServer { };
    [Guid("b189c9b3-036c-46a5-aeeb-2ab169ad45f0")][ComVisible(true)][ProgId("ExcelMvc.Rtd179")] public class Rtd179 : RtdServer { };
    [Guid("2763ce64-e5e3-439e-9af3-c5381bef43bc")][ComVisible(true)][ProgId("ExcelMvc.Rtd180")] public class Rtd180 : RtdServer { };
    [Guid("476b4924-4182-4be4-a8f1-28736f486e28")][ComVisible(true)][ProgId("ExcelMvc.Rtd181")] public class Rtd181 : RtdServer { };
    [Guid("1a281c34-632d-4e41-9e40-3b78380c56e0")][ComVisible(true)][ProgId("ExcelMvc.Rtd182")] public class Rtd182 : RtdServer { };
    [Guid("6db5cdac-7988-4ae1-a45b-b11a282871e0")][ComVisible(true)][ProgId("ExcelMvc.Rtd183")] public class Rtd183 : RtdServer { };
    [Guid("b4aefdc3-e756-4f64-9558-9613bcebf8bb")][ComVisible(true)][ProgId("ExcelMvc.Rtd184")] public class Rtd184 : RtdServer { };
    [Guid("ec9d31cb-0858-40f1-8d4e-602d3a073d44")][ComVisible(true)][ProgId("ExcelMvc.Rtd185")] public class Rtd185 : RtdServer { };
    [Guid("6cc6f540-d6af-47c3-9b7e-8e5839c88f64")][ComVisible(true)][ProgId("ExcelMvc.Rtd186")] public class Rtd186 : RtdServer { };
    [Guid("a7134135-4cf5-458a-8687-ec65941f4e54")][ComVisible(true)][ProgId("ExcelMvc.Rtd187")] public class Rtd187 : RtdServer { };
    [Guid("fe1b8429-64c8-4a2f-bcec-65ec7588f061")][ComVisible(true)][ProgId("ExcelMvc.Rtd188")] public class Rtd188 : RtdServer { };
    [Guid("fea90c44-8245-4b1c-a513-9368c21e59cc")][ComVisible(true)][ProgId("ExcelMvc.Rtd189")] public class Rtd189 : RtdServer { };
    [Guid("711cb1c4-9d3c-4a8b-bf39-384f32cfe5d8")][ComVisible(true)][ProgId("ExcelMvc.Rtd190")] public class Rtd190 : RtdServer { };
    [Guid("18a06297-4fdc-49a0-abd3-67498f1c7d20")][ComVisible(true)][ProgId("ExcelMvc.Rtd191")] public class Rtd191 : RtdServer { };
    [Guid("7b509e27-b458-4ac0-8198-235424fec579")][ComVisible(true)][ProgId("ExcelMvc.Rtd192")] public class Rtd192 : RtdServer { };
    [Guid("8729cb94-2b2b-4081-ab36-f56f03073957")][ComVisible(true)][ProgId("ExcelMvc.Rtd193")] public class Rtd193 : RtdServer { };
    [Guid("165c563e-91fb-45da-ab9d-0a773befd81e")][ComVisible(true)][ProgId("ExcelMvc.Rtd194")] public class Rtd194 : RtdServer { };
    [Guid("7d43afec-b3ae-4572-b683-a25a5a351fad")][ComVisible(true)][ProgId("ExcelMvc.Rtd195")] public class Rtd195 : RtdServer { };
    [Guid("3258efe7-2f9b-489c-bef7-5e17599ddae9")][ComVisible(true)][ProgId("ExcelMvc.Rtd196")] public class Rtd196 : RtdServer { };
    [Guid("cdf88bcf-7498-420d-b7e2-04b6e6d27480")][ComVisible(true)][ProgId("ExcelMvc.Rtd197")] public class Rtd197 : RtdServer { };
    [Guid("65684c2a-98dc-4b19-bb02-26254d067586")][ComVisible(true)][ProgId("ExcelMvc.Rtd198")] public class Rtd198 : RtdServer { };
    [Guid("d439a944-8ade-4af7-bb36-7158aa18762f")][ComVisible(true)][ProgId("ExcelMvc.Rtd199")] public class Rtd199 : RtdServer { };
    [Guid("89b8ef4f-b57a-43e2-9212-bfe1f6f33db1")][ComVisible(true)][ProgId("ExcelMvc.Rtd200")] public class Rtd200 : RtdServer { };
}
