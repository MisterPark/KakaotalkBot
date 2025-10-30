using System.Diagnostics;

namespace KakaotalkBot
{
    public static class Time
    {
        private static Stopwatch stopwatch = new Stopwatch();
        private static long lastTick = 0;
        private static long nowTick = 0;
        private static long deltaTime = 0;
        public static long DeltaTime
        {
            get; set;
        }

        public static void Initialize()
        {
            stopwatch.Start();
        }


        public static void Update()
        {
            nowTick = stopwatch.ElapsedMilliseconds;
            deltaTime = nowTick - lastTick;
            lastTick = nowTick;
            DeltaTime = deltaTime;
        }
    }
}
