namespace KakaotalkBot
{
    public class CustomTimer
    {
        private long tick = 0;
        public long Delay { get; set; }
        public CustomTimer(long delay)
        {
            Delay = delay;
        }

        public bool Check(long deltaTime)
        {
            tick += deltaTime;
            if (tick >= Delay)
            {
                tick = 0;
                return true;
            }

            return false;
        }

        public void Reset()
        {
            tick = 0;
        }

    }
}
