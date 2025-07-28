namespace EyesOfficeTesterCli
{
    internal class CommandLineParser
    {
        private readonly List<string> _args;

        public CommandLineParser(string[] args)
        {
            _args = args.ToList();
        }

        public string? GetStringArgument(string key, char shortKey)
        {
            var index = _args.IndexOf("--" + key);

            if (index >= 0 && _args.Count > index)
            {
                return _args[index + 1];
            }

            index = _args.IndexOf("-" + shortKey);

            if (index >= 0 && _args.Count > index)
            {
                return _args[index + 1];
            }

            return null;
        }

        public bool GetSwitchArgument(string value, char shortKey)
        {
            return _args.Contains("--" + value) || _args.Contains("-" + shortKey);
        }
    }
}
