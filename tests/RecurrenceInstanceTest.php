<?php

/**
 * libcalendaring plugin's recurrence instance identifier tests
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

class RecurrenceInstanceTest extends PHPUnit\Framework\TestCase
{
    public function setUp(): void
    {
        require_once __DIR__ . '/../libcalendaring.php';
    }

    /**
     * Read an identifier back the way identify_recurrence_instance() does, against an
     * event whose own start is stated in $tz.
     */
    private function read_back($instance, $start)
    {
        $object = [
            'uid'           => 'test',
            'recurrence_id' => 'test',
            'start'         => $start,
            '_instance'     => $instance,
        ];

        libcalendaring::identify_recurrence_instance($object);

        return $object['recurrence_date']
            ->setTimezone(new DateTimeZone('UTC'))
            ->format('Y-m-d\TH:i:s\Z');
    }

    /**
     * The identifier is persisted and read back in a context where the only zone available
     * is the event's own start. That is not necessarily the zone it was formatted in: an
     * organiser states RECURRENCE-ID with a TZID while the master's DTSTART may be in UTC
     * form. Both describe one instant, so the round trip has to agree in both cases.
     *
     * libcalendaring::recurrence_instance_identifier()
     * libcalendaring::identify_recurrence_instance()
     */
    public function test_recurrence_instance_identifier_round_trip()
    {
        $moscow = new DateTimeZone('Europe/Moscow');
        $utc    = new DateTimeZone('UTC');

        // occurrence 10:00 Europe/Moscow == 07:00Z, moved to 10:30 Moscow == 07:30Z
        $exception = [
            'uid'             => 'test',
            'start'           => new DateTime('2026-10-19 10:30:00', $moscow),
            'recurrence_date' => new DateTime('2026-10-19 10:00:00', $moscow),
        ];

        $instance = libcalendaring::recurrence_instance_identifier($exception);

        $this->assertSame('20261019T070000Z', $instance, 'identifier states the instant, in UTC');

        $this->assertSame(
            '2026-10-19T07:00:00Z',
            $this->read_back($instance, new DateTime('2026-10-19 10:30:00', $moscow)),
            'master stated with a TZID'
        );

        $this->assertSame(
            '2026-10-19T07:00:00Z',
            $this->read_back($instance, new DateTime('2026-10-19 07:30:00', $utc)),
            'master stated in UTC form, same instant'
        );
    }

    /**
     * An identifier already in a store carries no marker, and reinterpreting it would
     * change what existing data means. It keeps being read against the event's start.
     */
    public function test_identifier_without_a_marker_is_read_as_before()
    {
        $moscow = new DateTimeZone('Europe/Moscow');
        $utc    = new DateTimeZone('UTC');

        $this->assertSame(
            '2026-10-19T07:00:00Z',
            $this->read_back('20261019T100000', new DateTime('2026-10-19 10:30:00', $moscow))
        );

        $this->assertSame(
            '2026-10-19T10:00:00Z',
            $this->read_back('20261019T100000', new DateTime('2026-10-19 07:30:00', $utc))
        );
    }

    /**
     * All-day instances are dates, not instants (RFC 5545 3.8.4.4), so no zone applies.
     */
    public function test_allday_identifier_stays_a_date()
    {
        $event = [
            'uid'             => 'test',
            'allday'          => true,
            'start'           => new DateTime('2026-10-19 00:00:00', new DateTimeZone('Europe/Moscow')),
            'recurrence_date' => new DateTime('2026-10-19 00:00:00', new DateTimeZone('Europe/Moscow')),
        ];

        $this->assertSame('20261019', libcalendaring::recurrence_instance_identifier($event));
    }
}
