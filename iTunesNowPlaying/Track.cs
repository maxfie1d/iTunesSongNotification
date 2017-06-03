namespace iTunesNowPlaying
{
    public struct Track
    {
        public string Title { get; }
        public int PlayedCount { get; }

        public Track(string title, int playedCount)
        {
            this.Title = title;
            this.PlayedCount = playedCount;
        }
    }
}
