using System;

namespace OneNoteCodeHelper.Services
{
    internal static class PageEditCoordinator
    {
        // 有界锁池，页面之间偶尔串行可以接受，不永久持有页面 ID。
        private static readonly object[] Gates = CreateGates();
        private static object[] CreateGates()
        {
            var gates = new object[64];
            for (var i = 0; i < gates.Length; i++) gates[i] = new object();
            return gates;
        }
        internal static object ForPage(string pageId) => Gates[(uint)(pageId ?? "").GetHashCode() % Gates.Length];
    }
}
