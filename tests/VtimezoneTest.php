<?php

/**
 * libcalendaring_vcalendar::get_vtimezone() tests
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Affero General Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with this program. If not, see <http://www.gnu.org/licenses/>.
 */

class VtimezoneTest extends PHPUnit\Framework\TestCase
{
    public function setUp(): void
    {
        require_once __DIR__ . '/../lib/libcalendaring_vcalendar.php';
    }

    private function transitions($vt)
    {
        $out = [];

        foreach ($vt->children() as $child) {
            if (in_array($child->name, ['STANDARD', 'DAYLIGHT'])) {
                $out[] = [
                    'name'    => $child->name,
                    'dtstart' => (string) $child->DTSTART,
                    'from'    => (string) $child->TZOFFSETFROM,
                    'to'      => (string) $child->TZOFFSETTO,
                ];
            }
        }

        return $out;
    }

    /**
     * A zone whose last transition predates the exported period must still be described by
     * that transition, not by the start of the period.
     *
     * Europe/Moscow has had no transition since 2014-10-26. getTransitions() over a window
     * around a 2026 event therefore returns a single element, which PHP uses to describe the
     * state at the window's start: its 'time' is that boundary, not a change. Writing it as
     * STANDARD claims a change that never happened, and every datetime before it resolves
     * with TZOFFSETFROM (+0400 here) instead of the zone's actual offset.
     */
    public function test_vtimezone_without_transitions_in_period()
    {
        $from = strtotime('2026-10-26 10:30:00 Europe/Moscow');
        $to   = strtotime('2026-11-09 16:30:00 Europe/Moscow');

        $vt = libcalendaring_vcalendar::get_vtimezone('Europe/Moscow', $from, $to);

        $this->assertNotFalse($vt);

        $transitions = $this->transitions($vt);

        $this->assertCount(1, $transitions, 'one component for a zone with no DST');
        $this->assertSame('STANDARD', $transitions[0]['name']);
        $this->assertSame('+0400', $transitions[0]['from']);
        $this->assertSame('+0300', $transitions[0]['to']);

        // the 2014-10-26 change, which is the only one this zone has had
        $this->assertSame('20141025T220000', $transitions[0]['dtstart']);
    }

    /**
     * A zone that does change clocks inside the period keeps describing both directions.
     */
    public function test_vtimezone_with_transitions_in_period()
    {
        $from = strtotime('2026-06-01 12:00:00 UTC');
        $to   = strtotime('2026-12-01 12:00:00 UTC');

        $vt = libcalendaring_vcalendar::get_vtimezone('Europe/Berlin', $from, $to);

        $this->assertNotFalse($vt);

        $names = array_column($this->transitions($vt), 'name');

        $this->assertContains('STANDARD', $names);
        $this->assertContains('DAYLIGHT', $names);
    }
}
