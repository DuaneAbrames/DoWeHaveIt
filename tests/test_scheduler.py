import datetime as dt
import os
from pathlib import Path
import sys
import unittest
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'container'))
# Docker build installs the same files beside tests, rather than under container/.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from runtime import MAX_MINUTES, boundary, configuration, integer
from scheduler import Schedule


def ns(value):
    return int(dt.datetime.fromisoformat(value).timestamp()) * 1_000_000_000


class SchedulerTests(unittest.TestCase):
    def test_standard_20_minutes(self):
        start = ns('2026-01-01T00:00:00+00:00')
        self.assertEqual(boundary(start, 20, 0, True), start)
        self.assertEqual(boundary(start + 1, 20, 0), start + 1200 * 10**9)

    def test_offsets(self):
        start = ns('2026-01-01T00:00:00+00:00')
        for frequency, offset in ((20, 5), (15, 7), (60, 10)):
            self.assertEqual(boundary(start, frequency, offset), start + offset * 60 * 10**9)
        self.assertEqual(boundary(start, 20, 45), boundary(start, 20, 5))

    def test_arbitrary_long_intervals(self):
        for frequency in (137, 1501, 2880, MAX_MINUTES):
            current = ns('2026-12-31T23:59:59+00:00')
            first = boundary(current, frequency, 12)
            second = boundary(first, frequency, 12)
            self.assertGreater(first, current)
            self.assertEqual(second - first, frequency * 60 * 10**9)
            self.assertEqual((first // 10**9 - 12 * 60) % (frequency * 60), 0)

    def test_midnight_month_year(self):
        for date in ('2026-01-31', '2026-12-31', '2028-02-29'):
            current = ns(date + 'T23:59:59+00:00')
            self.assertEqual(boundary(current, 20, 0), current + 10**9)

    def test_daylight_saving(self):
        zone = ZoneInfo('America/New_York')
        for date in ('2026-03-08T06:40:00+00:00', '2026-11-01T05:40:00+00:00'):
            first = ns(date)
            values = [boundary(first + n * 1200 * 10**9, 20, 0, True) for n in range(4)]
            self.assertTrue(all(b-a == 1200 * 10**9 for a, b in zip(values, values[1:])))
            local = [dt.datetime.fromtimestamp(v // 10**9, zone) for v in values]
            self.assertNotEqual(local[0].utcoffset(), local[-1].utcoffset())

    def test_restart_and_no_replay(self):
        first = ns('2026-01-01T00:03:00+00:00')
        schedule = Schedule(first, 20, 0)
        self.assertIsNone(schedule.tick(first))
        expected = ns('2026-01-01T00:20:00+00:00')
        self.assertEqual(schedule.next, expected)
        restart = Schedule(ns('2026-01-02T12:03:00+00:00'), 20, 0)
        self.assertEqual(restart.next, ns('2026-01-02T12:20:00+00:00'))

    def test_long_scan_does_not_shift_schedule(self):
        first = ns('2026-01-01T00:00:00+00:00')
        schedule = Schedule(first, 20, 0)
        for index in range(4):
            self.assertEqual(schedule.tick(first + index * 1200 * 10**9), ('scheduled', first + index * 1200 * 10**9))

    def test_clock_adjustments(self):
        first = ns('2026-01-01T00:00:00+00:00')
        schedule = Schedule(first, 20, 0)
        schedule.tick(first)
        self.assertIsNone(schedule.tick(first - 10**9))
        self.assertIsNone(schedule.tick(first))
        self.assertEqual(schedule.tick(first + 5 * 3600 * 10**9 + 3 * 10**9)[0], 'missed')
        self.assertEqual(schedule.next, first + (5 * 3600 + 1200) * 10**9)
        self.assertIsNone(schedule.tick(first + 1200 * 10**9))

    def test_disabled(self):
        self.assertIsNone(boundary(0, 0, 123))
        self.assertIsNone(Schedule(0, 0, 123).tick(10**30))
        self.assertEqual(configuration({'SCHEDULE_FREQUENCY': '0', 'SCHEDULE_OFFSET': '999'})[:3], (0, 999, 999))

    def test_invalid_config(self):
        for name in ('SCHEDULE_FREQUENCY', 'SCHEDULE_OFFSET'):
            for value in ('-1', '+1', '', ' 20', '20 ', '1.2', 'no', '2e3', '9' * 10000, str(MAX_MINUTES + 1)):
                with self.subTest(name=name, value=value[:20]), self.assertRaises(ValueError):
                    configuration({name: value})
        for value in ('0', '-1', '36501'):
            with self.assertRaises(ValueError):
                configuration({'LOG_RETENTION_DAYS': value})


if __name__ == '__main__':
    unittest.main()
