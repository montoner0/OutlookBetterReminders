using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

// Copyright (c) 2016 Ben Spiller.

namespace BetterReminders
{
    public struct SnoozeTime : IComparable<SnoozeTime>
    {
        /// <summary>
        /// Number of seconds after now or after start time to wakeup
        /// </summary>
        public int Secs;
        /// <summary>
        /// if true, secs is measured from now, if false from the meeting start time
        /// </summary>
        public bool FromNow;

        /// <summary>
        ///
        /// </summary>
        /// <param name="secs">Seconds after now or after start time (negative indicates seconds before start time)</param>
        /// <param name="fromNow">true if relative to now, false if relative to meeting start time</param>
        public SnoozeTime(int secs, bool fromNow)
        {
            Secs = secs;
            FromNow = fromNow;
        }

        public DateTime GetNextReminderTime(DateTime startTime)
        {
            return (FromNow ? DateTime.Now : startTime).AddSeconds(Secs);
        }

        public static SnoozeTime Parse(string snoozeTime)
        {
            Match m = Regex.Match(snoozeTime, @"([\d.]+) *(s|h|m)", RegexOptions.IgnoreCase);
            if (!m.Success)
                throw new ArgumentException($"Invalid snooze time '{snoozeTime}': must contain <number> s|m|h");

            bool startTimeRelative = new[] { "start", "after", "before" }.Any(s => snoozeTime.IndexOf(s, StringComparison.InvariantCultureIgnoreCase) >= 0);
            bool fromNow = !startTimeRelative;

            float secs = float.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            switch (m.Groups[2].Value) {
                case "m": secs *= 60; break;
                case "h": secs *= 60 * 60; break;
                case "s": break;
            }
            if (secs < 1)
                throw new ArgumentException("Invalid snooze time, must be positive");
            if (snoozeTime.IndexOf("after", StringComparison.InvariantCultureIgnoreCase) < 0 && startTimeRelative)
                secs = -secs;

            return new SnoozeTime(Convert.ToInt32(secs), fromNow);
        }

        public override string ToString()
        {
            string t;
            int absSecs = (Secs > 0) ? Secs : -Secs;

            t = absSecs >= 60 && absSecs % 60 == 0
                ? $"{absSecs / 60} minute{(absSecs == 60 ? "" : "s")}"
                : $"{absSecs} second{(absSecs == 1 ? "" : "s")}";
            t = FromNow
                ? $"Remind in {t}"
                : $"Remind {t} {(Secs < 0 ? "before start time" : "after start time")}";
            // sanity check assertion
            return Parse(t).Equals(this) ? t : throw new Exception($"Error in snooze time ToString/Parse for: {t}");
        }

        public static List<SnoozeTime> ParseList(string list)
        {
            return new List<SnoozeTime>(list.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(t => Parse(t)));
        }

        public static string ListToString(List<SnoozeTime> list)
        {
            return string.Join(",", list);
        }

        #region IComparable<SnoozeTime> Members

        public int CompareTo(SnoozeTime other)
        {
            return FromNow == other.FromNow
                // small/earlier last, far times and start times at top of list
                ? other.Secs - Secs
                // fromnow items first
                : FromNow ? +1 : -1;
        }

        #endregion
    }
}
