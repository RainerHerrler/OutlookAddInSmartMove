using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SmartMove.OutlookAddIn
{
    internal static class SubjectWordExtractor
    {
        private static readonly Regex WordPattern = new Regex(
            @"[\p{L}\p{M}]+",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // Based on the German Snowball stop-word list (BSD license), extended
        // with common low-information e-mail greetings and subject prefixes.
        private static readonly HashSet<string> StopWords = new HashSet<string>(
            new[]
            {
                "aber", "alle", "allem", "allen", "aller", "alles", "also", "ander", "andere",
                "anderem", "anderen", "anderer", "anderes", "anderm", "andern", "anderr", "anders",
                "auch", "dabei", "damit", "dann", "dass", "daß", "dazu", "dein", "deine", "deinem",
                "deinen", "deiner", "deines", "denn", "derer", "derselbe", "derselben", "denselben",
                "desselben", "demselben", "dessen", "dasselbe", "dich", "dies", "diese", "dieselbe",
                "dieselben", "diesem", "diesen", "dieser",
                "dieses", "doch", "dort", "durch", "eine", "einem", "einen", "einer", "eines",
                "einig", "einige", "einigem", "einigen", "einiger", "einiges", "einmal", "etwas",
                "euch", "euer", "eure", "eurem", "euren", "eurer", "eures", "gegen", "gewesen", "habe",
                "haben", "hatte", "hatten", "hier", "hinter", "ihnen", "indem", "jede", "jedem",
                "ihre", "ihrem", "ihren", "ihrer", "ihres", "jeden", "jeder", "jedes", "jene", "jenem",
                "jenen", "jener", "jenes", "jetzt", "kann", "kein", "keine", "keinem", "keinen",
                "keiner", "keines", "können", "könnte",
                "machen", "manche", "manchem", "manchen", "mancher", "manches", "mein", "meine",
                "meinem", "meinen", "meiner", "meines", "mich", "muss", "musste", "nach", "nicht", "nichts",
                "noch", "oder", "ohne", "schon", "sehr", "sein", "seine", "seinem", "seinen", "seiner",
                "seines", "selbst", "sich", "sind", "solche", "solchem", "solchen", "solcher",
                "solches", "soll", "sollte", "sondern", "sonst", "über", "unter", "unse", "unsem", "unsen",
                "unser", "unses", "unsere", "unserem", "unseren", "unserer", "unseres", "viel", "viele",
                "während", "waren", "warst", "weil", "weiter", "welche", "welchem", "welchen", "welcher",
                "welches", "wenn", "werde", "werden", "wieder", "wird", "wirst", "will", "willst", "wollen", "wollte",
                "würde", "würden", "zwar", "zwischen",

                // SmartMove e-mail additions.
                "antwort", "bitte", "danke", "email", "freundliche", "freundlichen", "grüsse", "grüße",
                "guten", "hallo", "liebe", "lieben", "lieber", "nachricht", "vielen", "weitergeleitet",
                "weitergeleitete"
            },
            StringComparer.OrdinalIgnoreCase);

        public static IReadOnlyCollection<string> Extract(string subject)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(subject))
            {
                return result;
            }

            foreach (Match match in WordPattern.Matches(subject))
            {
                string word = match.Value.ToLowerInvariant();
                if (word.Length > 3 && !StopWords.Contains(word))
                {
                    result.Add(word);
                }
            }

            return result;
        }
    }
}
